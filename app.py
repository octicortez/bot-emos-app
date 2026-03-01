import streamlit as st
import pandas as pd
import time
import os
import urllib.request
import ssl
import shutil
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# --- 1.  IDs DE EMOS 
ID_CASILLERO_1 = "vCIRNUME"   
ID_CASILLERO_2 = "vSCCNUME"   
ID_CASILLERO_3 = "vMZANUME"   
ID_CASILLERO_4 = "vPARNUME"   
ID_CASILLERO_5 = "vPHONUME"   
ID_BOTON_BUSCAR_ABAJO = "BUTTON1" 
ID_BOTON_IMPRIMIR_BOLETA = "BUTTON1"

# --- FUNCIÓN DEL BOT (El Cerebro) ---
def consultar_propiedad(driver, wait, nomenclatura, periodo_buscado, carpeta_destino):
    partes = str(nomenclatura).split("-")
    if len(partes) != 5:
        return {"Nomenclatura": nomenclatura, "Periodo": periodo_buscado, "Importe Total": "-", "Vencimiento": "-", "Estado": "Formato Incorrecto"}

    datos_extraidos = {
        "Nomenclatura": nomenclatura,
        "Periodo": periodo_buscado,
        "Importe Total": "No encontrado",
        "Vencimiento": "-",
        "Estado": "Sin Deuda"
    }

    try:
        driver.delete_all_cookies() 
        driver.get("https://emosvirtual.riocuarto.gov.ar:9090/emosweb/servlet/com.emosweb.login")
        time.sleep(3) 
        
        # Escritura
        c1 = wait.until(EC.presence_of_element_located((By.ID, ID_CASILLERO_1)))
        c1.clear(); c1.send_keys(partes[0]); time.sleep(0.5)
        c2 = driver.find_element(By.ID, ID_CASILLERO_2)
        c2.clear(); c2.send_keys(partes[1]); time.sleep(0.5)
        c3 = driver.find_element(By.ID, ID_CASILLERO_3)
        c3.clear(); c3.send_keys(partes[2]); time.sleep(0.5)
        c4 = driver.find_element(By.ID, ID_CASILLERO_4)
        c4.clear(); c4.send_keys(partes[3]); time.sleep(0.5)
        c5 = driver.find_element(By.ID, ID_CASILLERO_5)
        c5.clear(); c5.send_keys(partes[4]); time.sleep(1)
        
        boton_buscar = driver.find_element(By.ID, ID_BOTON_BUSCAR_ABAJO)
        driver.execute_script("arguments[0].click();", boton_buscar)
        time.sleep(5) 
        
        filas = driver.find_elements(By.TAG_NAME, "tr")
        
        for fila in filas:
            texto_fila = fila.text.strip()
            if texto_fila.startswith(str(periodo_buscado)):
                datos = texto_fila.split()
                if len(datos) >= 6:
                    datos_extraidos["Importe Total"] = datos[-1]
                    datos_extraidos["Vencimiento"] = datos[1]
                    datos_extraidos["Estado"] = "IMPAGO, PDF Descargado"
                    
                    # Tildar casilla y apretar imprimir
                    casilla = fila.find_element(By.TAG_NAME, "input")
                    driver.execute_script("arguments[0].click();", casilla)
                    time.sleep(1)
                    
                    boton_imprimir = driver.find_element(By.ID, ID_BOTON_IMPRIMIR_BOLETA)
                    driver.execute_script("arguments[0].click();", boton_imprimir)
                    time.sleep(5) 
                    
                    # Cazando la ventanita flotante
                    try:
                        cuadritos = driver.find_elements(By.CSS_SELECTOR, "iframe, embed, object")
                        pdf_url = None
                        for cuadrito in cuadritos:
                            link = cuadrito.get_attribute("src") or cuadrito.get_attribute("data")
                            if link:
                                pdf_url = link
                                break
                        
                        if pdf_url:
                            periodo_limpio = str(periodo_buscado).replace('/', '-')
                            nombre_pdf = f"Boleta_{nomenclatura}_{periodo_limpio}.pdf"
                            ruta_final = os.path.join(carpeta_destino, nombre_pdf)
                            
                            cookies = driver.get_cookies()
                            texto_cookies = "; ".join([f"{c['name']}={c['value']}" for c in cookies])
                            
                            ctx = ssl.create_default_context()
                            ctx.check_hostname = False
                            ctx.verify_mode = ssl.CERT_NONE
                            
                            req = urllib.request.Request(pdf_url)
                            req.add_header("Cookie", texto_cookies)
                            req.add_header("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)")
                            
                            with urllib.request.urlopen(req, context=ctx) as response:
                                with open(ruta_final, "wb") as f:
                                    f.write(response.read())
                        else:
                            datos_extraidos["Estado"] = "Error: PDF no encontrado en la ventana"
                    except Exception as e:
                        datos_extraidos["Estado"] = f"Error de descarga: {e}"
                    break 
        return datos_extraidos

    except Exception as e:
        datos_extraidos["Estado"] = f"ERROR TÉCNICO"
        return datos_extraidos

# --- INTERFAZ WEB (La Cara) ---
st.set_page_config(page_title="Asistente EMOS", page_icon="🤖", layout="centered")

st.title("🤖 Asistente de Boletas EMOS")
st.markdown("Sube tu archivo Excel con las nomenclaturas y los periodos. El bot buscará las deudas automáticamente y descargará los PDFs.")
st.divider()

archivo_subido = st.file_uploader("Sube tu archivo Excel (ej: lista_propiedades.xlsx)", type=["xlsx"])

if archivo_subido is not None:
    df_propiedades = pd.read_excel(archivo_subido)
    df_propiedades.columns = df_propiedades.columns.str.strip() # Limpiamos espacios
    
    st.write(f"📄 Se detectaron **{len(df_propiedades)}** propiedades en tu lista:")
    st.dataframe(df_propiedades)
    st.divider()
    
    if st.button("🚀 Buscar Boletas y Descargar PDFs", use_container_width=True):
        
        carpeta_temp = "Boletas_Temporales"
        os.makedirs(carpeta_temp, exist_ok=True)
        resultados = []
        
        # Barra de progreso visual
        barra_progreso = st.progress(0)
        texto_estado = st.empty()
        
        try:
            texto_estado.text("Iniciando navegador en el servidor...")
            chrome_options = Options()
            
            # 1. Le decimos exactamente dónde está instalado Chrome en el servidor Linux
            chrome_options.binary_location = "/usr/bin/chromium"
            
            chrome_options.add_argument("--headless=new") 
            chrome_options.add_argument("--no-sandbox") 
            chrome_options.add_argument("--disable-dev-shm-usage") 
            chrome_options.add_argument("--window-size=1920,1080")
            chrome_options.add_argument("--disable-gpu")
            
            # 2. Le decimos que use el conductor local en vez de descargarlo
            servicio = Service("/usr/bin/chromedriver")
            driver = webdriver.Chrome(service=servicio, options=chrome_options)
            
            wait = WebDriverWait(driver, 10)
            
            total_filas = len(df_propiedades)
            
            for index, row in df_propiedades.iterrows():
                nomenclatura = row.iloc[0]
                periodo = row.iloc[1]
                
                if pd.isna(nomenclatura) or pd.isna(periodo):
                    continue
                
                texto_estado.text(f"Consultando: {nomenclatura} ({index + 1}/{total_filas})...")
                resultado = consultar_propiedad(driver, wait, nomenclatura, periodo, carpeta_temp)
                resultados.append(resultado)
                
                # Actualizar barra
                progreso_actual = int(((index + 1) / total_filas) * 100)
                barra_progreso.progress(progreso_actual)

            driver.quit()
            
            # Guardar el Excel de resultados en la misma carpeta temporal
            df_resultados = pd.DataFrame(resultados)
            ruta_excel = os.path.join(carpeta_temp, "Reporte_Resultados.xlsx")
            df_resultados.to_excel(ruta_excel, index=False)
            
            # Comprimir todo en un archivo ZIP
            texto_estado.text("Empaquetando archivos...")
            shutil.make_archive("Boletas_EMOS", 'zip', carpeta_temp)
            
            st.success("✅ ¡Proceso terminado con éxito!")
            texto_estado.empty() # Borramos el texto de carga
            
            # Mostrar botón de descarga del ZIP
            with open("Boletas_EMOS.zip", "rb") as f:
                st.download_button(
                    label="📥 Descargar Reporte y PDFs (.ZIP)",
                    data=f,
                    file_name="Boletas_EMOS_Terminadas.zip",
                    mime="application/zip",
                    use_container_width=True
                )
                
        except Exception as e:
            st.error(f"Ocurrió un error inesperado: {e}")
            
        finally:
            # Limpieza: borramos la carpeta temporal para no llenar el disco
            if os.path.exists(carpeta_temp):
                shutil.rmtree(carpeta_temp)
