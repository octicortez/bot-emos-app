import streamlit as st
import pandas as pd
import time
import os
import urllib.request
import ssl
import shutil
import glob
from pypdf import PdfWriter
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
    
    # --- LA MEMORIA DEL BOT ---
    # Le enseñamos a recordar si ya terminó el trabajo para no borrar los botones
    if "proceso_terminado" not in st.session_state:
        st.session_state.proceso_terminado = False

    # --- EL BOTÓN MÁGICO ---
    if st.button("🚀 Buscar Boletas y Descargar PDFs", use_container_width=True):
        
        carpeta_temp = "Boletas_Temporales"
        os.makedirs(carpeta_temp, exist_ok=True)
        resultados = []
        barra_progreso = st.progress(0)
        texto_estado = st.empty()
        
        try:
            texto_estado.text("Iniciando navegador...")
            chrome_options = Options()
            chrome_options.add_argument("--window-size=1920,1080")
            
            # Detector Inteligente (Nube vs Mac)
            if os.path.exists("/usr/bin/chromium"):
                chrome_options.binary_location = "/usr/bin/chromium"
                chrome_options.add_argument("--headless=new") 
                chrome_options.add_argument("--no-sandbox") 
                chrome_options.add_argument("--disable-dev-shm-usage") 
                chrome_options.add_argument("--disable-gpu")
                servicio = Service("/usr/bin/chromedriver")
            else:
                chrome_options.add_argument("--headless=new")
                servicio = Service(ChromeDriverManager().install())
                
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
                barra_progreso.progress(int(((index + 1) / total_filas) * 100))

            driver.quit()
            
            # 1. Guardar el Excel de resultados (Adentro y afuera de la carpeta temporal)
            df_resultados = pd.DataFrame(resultados)
            df_resultados.to_excel(os.path.join(carpeta_temp, "Reporte_Resultados.xlsx"), index=False)
            df_resultados.to_excel("Reporte_Resultados_Final.xlsx", index=False) # Copia segura afuera
            
            # 2. Crear el PDF Maestro
            texto_estado.text("Uniendo todas las boletas para imprimir...")
            archivos_pdf = glob.glob(os.path.join(carpeta_temp, "*.pdf"))
            if archivos_pdf:
                fusionador = PdfWriter()
                for pdf_path in archivos_pdf:
                    fusionador.append(pdf_path)
                fusionador.write("Boletas_Unidas_Para_Imprimir.pdf")
                fusionador.close()
            
            # 3. Comprimir todo en un archivo ZIP
            texto_estado.text("Empaquetando archivos finales...")
            shutil.make_archive("Boletas_EMOS", 'zip', carpeta_temp)
            
            # Limpieza de la carpeta temporal
            if os.path.exists(carpeta_temp):
                shutil.rmtree(carpeta_temp)
                
            texto_estado.empty()
            
            # ¡Activamos la memoria! El bot recordará que ya terminó.
            st.session_state.proceso_terminado = True
            
        except Exception as e:
            st.error(f"Ocurrió un error inesperado: {e}")

    # --- LOS BOTONES DE DESCARGA PERMANENTES ---
    # Como están afuera del botón "Buscar", nunca van a desaparecer
    if st.session_state.proceso_terminado:
        st.success("✅ ¡Proceso terminado con éxito!")
        st.info("Tus archivos están listos. Haz clic en cualquiera de los botones, no desaparecerán:")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            if os.path.exists("Boletas_EMOS.zip"):
                with open("Boletas_EMOS.zip", "rb") as f_zip:
                    st.download_button("📦 Bajar Todo (.ZIP)", data=f_zip, file_name="Boletas_EMOS_Terminadas.zip", mime="application/zip", use_container_width=True)
                    
        with col2:
            if os.path.exists("Boletas_Unidas_Para_Imprimir.pdf"):
                with open("Boletas_Unidas_Para_Imprimir.pdf", "rb") as f_pdf:
                    st.download_button("🖨️ Bajar PDF Unido", data=f_pdf, file_name="Boletas_Para_Imprimir.pdf", mime="application/pdf", use_container_width=True)
                    
        with col3:
            if os.path.exists("Reporte_Resultados_Final.xlsx"):
                with open("Reporte_Resultados_Final.xlsx", "rb") as f_xls:
                    st.download_button("📊 Bajar Solo Excel", data=f_xls, file_name="Reporte_Resultados.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
