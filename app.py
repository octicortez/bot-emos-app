import streamlit as st
import pandas as pd
import time
import os
import urllib.request
import ssl
import shutil
import glob
import base64
import datetime
from pypdf import PdfWriter
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager

# ==========================================
# --- 1. CONFIGURACIÓN DE IDs ---
# ==========================================

# --- TUS IDs DE EMOS (¡REEMPLAZA ESTO CON TUS DATOS!) ---
ID_CASILLERO_1_EMOS = "vCIRNUME"   
ID_CASILLERO_2_EMOS = "vSCCNUME"   
ID_CASILLERO_3_EMOS = "vMZANUME"   
ID_CASILLERO_4_EMOS = "vPARNUME"   
ID_CASILLERO_5_EMOS = "vPHONUME"   
ID_BOTON_BUSCAR_EMOS = "BUTTON1"  
ID_BOTON_IMPRIMIR_EMOS = "BUTTON1"
ID_FECHA_ACTUALIZACION_EMOS = "vFECHAACTUALIZACION"

# --- IDs DE LA MUNICIPALIDAD (Ya configurados) ---
ID_C1_MUNI = "vEIWCIR"
ID_C2_MUNI = "vEIWSEC"
ID_C3_MUNI = "vEIWMZA"
ID_C4_MUNI = "vEIWPA"
ID_C5_MUNI = "vEIWPH"
ID_FECHA_MUNI = "vEIWVTO"
ID_BUSCAR_MUNI = "BUSCAR"
ID_CHECKALL_MUNI = "vCHECKALL_GRIDSDTBDEUDAONLINE"
ID_TOTAL_MUNI = "span_vTOTDEU"
ID_BOLETA_MUNI = "BOLETADEPAGO1"
ID_IMPRIMIR_MUNI = "IMPRIMIR"

# ==========================================
# --- 2. LOS MOTORES DEL BOT ---
# ==========================================

def consultar_emos(driver, wait, nomenclatura, periodo_buscado, carpeta_destino, fecha_pago_obj):
    # EMOS necesita el año corto (ej: 20/03/26)
    fecha_pago_str = fecha_pago_obj.strftime("%d/%m/%y")
    partes = str(nomenclatura).split("-")
    
    datos_extraidos = {
        "Nomenclatura": nomenclatura,
        "Periodo": periodo_buscado,
        "Importe Total": "No encontrado",
        "Vencimiento": "-",
        "Estado": "Sin Deuda"
    }

    if len(partes) != 5:
        datos_extraidos["Estado"] = "Formato Incorrecto"
        return datos_extraidos

    try:
        driver.delete_all_cookies() 
        driver.get("https://emosvirtual.riocuarto.gov.ar:9090/emosweb/servlet/com.emosweb.login")
        time.sleep(3) 
        
        # Ingresar Nomenclatura
        c1 = wait.until(EC.presence_of_element_located((By.ID, ID_CASILLERO_1_EMOS)))
        c1.clear(); c1.send_keys(partes[0]); time.sleep(0.5)
        c2 = driver.find_element(By.ID, ID_CASILLERO_2_EMOS)
        c2.clear(); c2.send_keys(partes[1]); time.sleep(0.5)
        c3 = driver.find_element(By.ID, ID_CASILLERO_3_EMOS)
        c3.clear(); c3.send_keys(partes[2]); time.sleep(0.5)
        c4 = driver.find_element(By.ID, ID_CASILLERO_4_EMOS)
        c4.clear(); c4.send_keys(partes[3]); time.sleep(0.5)
        c5 = driver.find_element(By.ID, ID_CASILLERO_5_EMOS)
        c5.clear(); c5.send_keys(partes[4]); time.sleep(1)
        
        boton_buscar = driver.find_element(By.ID, ID_BOTON_BUSCAR_EMOS)
        driver.execute_script("arguments[0].click();", boton_buscar)
        time.sleep(5) 
        
        # Actualizar Fecha EMOS
        try:
            casillero_fecha = wait.until(EC.element_to_be_clickable((By.ID, ID_FECHA_ACTUALIZACION_EMOS)))
            casillero_fecha.click()
            time.sleep(0.5)
            casillero_fecha.send_keys(Keys.END)
            for _ in range(10): casillero_fecha.send_keys(Keys.BACKSPACE)
            time.sleep(0.5)
            casillero_fecha.send_keys(fecha_pago_str)
            time.sleep(1)
            boton_confirmar_fecha = driver.find_element(By.ID, "BUTTON5")
            driver.execute_script("arguments[0].click();", boton_confirmar_fecha)
            time.sleep(6) 
        except Exception as e:
            print(f"Advertencia: No se pudo cambiar la fecha en EMOS. {e}")
        
        # Buscar la deuda y descargar
        filas = driver.find_elements(By.TAG_NAME, "tr")
        for fila in filas:
            texto_fila = fila.text.strip()
            if texto_fila.startswith(str(periodo_buscado)):
                datos = texto_fila.split()
                if len(datos) >= 6:
                    datos_extraidos["Importe Total"] = datos[-1]
                    datos_extraidos["Vencimiento"] = datos[1]
                    datos_extraidos["Estado"] = "IMPAGO, PDF Descargado"
                    
                    casilla = fila.find_element(By.TAG_NAME, "input")
                    driver.execute_script("arguments[0].click();", casilla)
                    time.sleep(1)
                    
                    boton_imprimir = driver.find_element(By.ID, ID_BOTON_IMPRIMIR_EMOS)
                    driver.execute_script("arguments[0].click();", boton_imprimir)
                    time.sleep(5) 
                    
                    try:
                        cuadritos = driver.find_elements(By.CSS_SELECTOR, "iframe, embed, object")
                        pdf_url = None
                        for cuadrito in cuadritos:
                            link = cuadrito.get_attribute("src") or cuadrito.get_attribute("data")
                            if link: pdf_url = link; break
                        
                        if pdf_url:
                            periodo_limpio = str(periodo_buscado).replace('/', '-')
                            nombre_pdf = f"Boleta_EMOS_{nomenclatura}_{periodo_limpio}.pdf"
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
                            datos_extraidos["Estado"] = "Error: PDF no encontrado en EMOS"
                    except Exception as e:
                        datos_extraidos["Estado"] = f"Error de descarga: {e}"
                    break 
        return datos_extraidos

    except Exception as e:
        datos_extraidos["Estado"] = f"ERROR TÉCNICO"
        return datos_extraidos


def consultar_muni(driver, wait, nomenclatura, carpeta_destino, fecha_pago_obj):
    # MUNI necesita el año largo (ej: 20/03/2026)
    fecha_pago_str = fecha_pago_obj.strftime("%d/%m/%Y")
    partes = str(nomenclatura).split("-")
    
    datos_extraidos = {
        "Nomenclatura": nomenclatura,
        "Periodo": "Toda la deuda",
        "Importe Total": "No encontrado",
        "Vencimiento": fecha_pago_str,
        "Estado": "Sin Deuda"
    }

    if len(partes) != 5:
        datos_extraidos["Estado"] = "Formato Incorrecto"
        return datos_extraidos

    try:
        driver.delete_all_cookies()
        driver.get("https://app.riocuarto.gov.ar:8443/gestiontributaria/servlet/com.recursos.hceduimpmul?Inmo")
        time.sleep(2) 

        # Ingresar Nomenclatura
        c1 = wait.until(EC.presence_of_element_located((By.ID, ID_C1_MUNI)))
        c1.clear(); c1.send_keys(partes[0]); time.sleep(0.5)
        c2 = driver.find_element(By.ID, ID_C2_MUNI)
        c2.clear(); c2.send_keys(partes[1]); time.sleep(0.5)
        c3 = driver.find_element(By.ID, ID_C3_MUNI)
        c3.clear(); c3.send_keys(partes[2]); time.sleep(0.5)
        c4 = driver.find_element(By.ID, ID_C4_MUNI)
        c4.clear(); c4.send_keys(partes[3]); time.sleep(0.5)
        c5 = driver.find_element(By.ID, ID_C5_MUNI)
        c5.clear(); c5.send_keys(partes[4]); time.sleep(1)

        # Configurar Fecha
        casillero_fecha = driver.find_element(By.ID, ID_FECHA_MUNI)
        casillero_fecha.click()
        time.sleep(0.5)
        casillero_fecha.send_keys(Keys.END)
        for _ in range(12): casillero_fecha.send_keys(Keys.BACKSPACE)
        time.sleep(0.5)
        casillero_fecha.send_keys(fecha_pago_str)
        time.sleep(1)

        # Buscar
        boton_buscar = driver.find_element(By.ID, ID_BUSCAR_MUNI)
        driver.execute_script("arguments[0].click();", boton_buscar)

        # Seleccionar todo y leer total
        try:
            checkbox_todas = wait.until(EC.element_to_be_clickable((By.ID, ID_CHECKALL_MUNI)))
            driver.execute_script("arguments[0].click();", checkbox_todas)
            time.sleep(1)
            
            elemento_total = driver.find_element(By.ID, ID_TOTAL_MUNI)
            datos_extraidos["Importe Total"] = elemento_total.text.strip()
        except:
            datos_extraidos["Estado"] = "No se encontraron deudas para seleccionar"
            return datos_extraidos

        # Generar e imprimir
        boton_boleta = driver.find_element(By.ID, ID_BOLETA_MUNI)
        driver.execute_script("arguments[0].click();", boton_boleta)
        
        boton_imprimir = wait.until(EC.element_to_be_clickable((By.ID, ID_IMPRIMIR_MUNI)))
        driver.execute_script("arguments[0].click();", boton_imprimir)
        time.sleep(5) 

        # Descarga Ninja JS
        cuadritos = driver.find_elements(By.CSS_SELECTOR, "iframe, embed, object")
        pdf_url = None
        for cuadrito in cuadritos:
            link = cuadrito.get_attribute("src") or cuadrito.get_attribute("data")
            if link: pdf_url = link; break

        if pdf_url:
            base64_pdf = driver.execute_async_script("""
                var uri = arguments[0];
                var callback = arguments[1];
                var xhr = new XMLHttpRequest();
                xhr.open('GET', uri, true);
                xhr.responseType = 'arraybuffer';
                xhr.onload = function() {
                    var arrayBuffer = xhr.response;
                    var byteArray = new Uint8Array(arrayBuffer);
                    var byteString = '';
                    for (var i = 0; i < byteArray.byteLength; i++) {
                        byteString += String.fromCharCode(byteArray[i]);
                    }
                    var base64 = btoa(byteString);
                    callback(base64);
                };
                xhr.onerror = function() { callback(null); };
                xhr.send();
            """, pdf_url)
            
            if base64_pdf:
                nombre_pdf = f"Boleta_MUNI_{nomenclatura}.pdf"
                ruta_final = os.path.join(carpeta_destino, nombre_pdf)
                with open(ruta_final, "wb") as f:
                    f.write(base64.b64decode(base64_pdf))
                datos_extraidos["Estado"] = "PDF Descargado Exitosamente"
            else:
                datos_extraidos["Estado"] = "Error al descargar PDF (JS Falló)"
        else:
            datos_extraidos["Estado"] = "PDF no encontrado en pantalla"
            
        return datos_extraidos

    except Exception as e:
        datos_extraidos["Estado"] = f"ERROR TÉCNICO"
        return datos_extraidos


# ==========================================
# --- 3. LA INTERFAZ WEB (STREAMLIT) ---
# ==========================================

st.set_page_config(page_title="Gestor de Impuestos", page_icon="🧾", layout="centered")

st.title("🧾 Gestor Automático de Impuestos")
st.markdown("Central de descargas para boletas de propiedades.")
st.divider()

# --- SELECTOR DE OPCIÓN B ---
tipo_impuesto = st.selectbox(
    "¿Qué impuesto deseas consultar hoy?", 
    ["💧 EMOS (Agua y Cloacas)", "🏛️ Municipalidad (Contribución Inmobiliaria)"]
)
es_emos = "EMOS" in tipo_impuesto

st.divider()
st.markdown("### 📅 Configuración de Pago")
fecha_seleccionada = st.date_input("Selecciona la fecha en la que se realizará el pago:", datetime.date.today())
if es_emos:
    st.write(f"👉 *Formato interno EMOS:* **{fecha_seleccionada.strftime('%d/%m/%y')}**")
else:
    st.write(f"👉 *Formato interno MUNI:* **{fecha_seleccionada.strftime('%d/%m/%Y')}**")

st.divider()
st.markdown("### 📄 Carga de Datos")
if es_emos:
    st.info("Para EMOS, el Excel debe tener la columna Nomenclatura y la columna Periodo.")
else:
    st.info("Para la Municipalidad, el bot seleccionará y sumará toda la deuda pendiente de la Nomenclatura.")

archivo_subido = st.file_uploader(f"Sube tu archivo Excel para {tipo_impuesto}", type=["xlsx"])

# --- LA MEMORIA DEL BOT ---
if "proceso_terminado" not in st.session_state:
    st.session_state.proceso_terminado = False

if archivo_subido is not None:
    df_propiedades = pd.read_excel(archivo_subido)
    df_propiedades.columns = df_propiedades.columns.str.strip() 
    
    st.write(f"Se detectaron **{len(df_propiedades)}** filas en tu lista:")
    st.dataframe(df_propiedades)
    st.divider()
    
    if st.button(f"🚀 Iniciar Búsqueda en {tipo_impuesto}", use_container_width=True):
        
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
                chrome_options.page_load_strategy = 'eager' 
                servicio = Service(ChromeDriverManager().install())
                
            driver = webdriver.Chrome(service=servicio, options=chrome_options)
            driver.set_page_load_timeout(180)
            wait = WebDriverWait(driver, 15)
            total_filas = len(df_propiedades)
            
            for index, row in df_propiedades.iterrows():
                nomenclatura = row.iloc[0]
                periodo = row.iloc[1] if es_emos and len(row) > 1 else "-"
                
                if pd.isna(nomenclatura):
                    continue
                    
                texto_estado.text(f"Consultando: {nomenclatura} ({index + 1}/{total_filas})...")
                
                # Desvío de tráfico según la opción elegida
                if es_emos:
                    resultado = consultar_emos(driver, wait, nomenclatura, periodo, carpeta_temp, fecha_seleccionada)
                else:
                    resultado = consultar_muni(driver, wait, nomenclatura, carpeta_temp, fecha_seleccionada)
                
                resultados.append(resultado)
                barra_progreso.progress(int(((index + 1) / total_filas) * 100))

            driver.quit()
            
            # 1. Guardar el Excel
            df_resultados = pd.DataFrame(resultados)
            nombre_excel = f"Reporte_{'EMOS' if es_emos else 'MUNI'}.xlsx"
            df_resultados.to_excel(os.path.join(carpeta_temp, nombre_excel), index=False)
            df_resultados.to_excel(nombre_excel, index=False) 
            
            # 2. Crear el PDF Maestro
            texto_estado.text("Uniendo todas las boletas para imprimir...")
            archivos_pdf = glob.glob(os.path.join(carpeta_temp, "*.pdf"))
            if archivos_pdf:
                fusionador = PdfWriter()
                for pdf_path in archivos_pdf:
                    fusionador.append(pdf_path)
                fusionador.write("Boletas_Unidas.pdf")
                fusionador.close()
            
            # 3. Comprimir todo
            texto_estado.text("Empaquetando archivos finales...")
            shutil.make_archive("Boletas_Finales", 'zip', carpeta_temp)
            
            # Limpieza
            if os.path.exists(carpeta_temp):
                shutil.rmtree(carpeta_temp)
                
            texto_estado.empty()
            st.session_state.proceso_terminado = True
            
        except Exception as e:
            st.error(f"Ocurrió un error inesperado: {e}")

# --- LOS BOTONES DE DESCARGA PERMANENTES ---
if st.session_state.proceso_terminado:
    st.success("✅ ¡Proceso terminado con éxito!")
    st.info("Tus archivos están listos. Haz clic en cualquiera de los botones:")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        if os.path.exists("Boletas_Finales.zip"):
            with open("Boletas_Finales.zip", "rb") as f_zip:
                st.download_button("📦 Bajar Todo (.ZIP)", data=f_zip, file_name="Boletas_Completas.zip", mime="application/zip", use_container_width=True)
                
    with col2:
        if os.path.exists("Boletas_Unidas.pdf"):
            with open("Boletas_Unidas.pdf", "rb") as f_pdf:
                st.download_button("🖨️ Bajar PDF Unido", data=f_pdf, file_name="Boletas_Para_Imprimir.pdf", mime="application/pdf", use_container_width=True)
                
    with col3:
        # Busca el excel que se haya generado (EMOS o MUNI)
        archivos_excel = glob.glob("Reporte_*.xlsx")
        if archivos_excel:
            archivo_reciente = max(archivos_excel, key=os.path.getmtime)
            with open(archivo_reciente, "rb") as f_xls:
                st.download_button("📊 Bajar Reporte Excel", data=f_xls, file_name=archivo_reciente, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
