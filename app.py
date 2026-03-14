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
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
import datetime

# --- TUS IDs DE EMOS ---
ID_CASILLERO_1 = "vCIRNUME"   
ID_CASILLERO_2 = "vSCCNUME"   
ID_CASILLERO_3 = "vMZANUME"   
ID_CASILLERO_4 = "vPARNUME"   
ID_CASILLERO_5 = "vPHONUME"   
ID_BOTON_BUSCAR_ABAJO = "BUTTON1"  
ID_BOTON_IMPRIMIR_BOLETA = "BUTTON1"
ID_FECHA_ACTUALIZACION = "vFECHAACTUALIZACION"

def consultar_emos(driver, wait, nomenclatura, periodo_buscado, carpeta_destino, fecha_pago_obj):
    fecha_pago_str = fecha_pago_obj.strftime("%d/%m/%y")
    partes = str(nomenclatura).split("-")
    
    datos_extraidos = {"Nomenclatura": nomenclatura, "Periodo": periodo_buscado, "Importe Total": "No encontrado", "Vencimiento": "-", "Estado": "Sin Deuda"}
    if len(partes) != 5:
        datos_extraidos["Estado"] = "Formato Incorrecto"
        return datos_extraidos

    try:
        driver.delete_all_cookies() 
        driver.get("https://emosvirtual.riocuarto.gov.ar:9090/emosweb/servlet/com.emosweb.login")
        time.sleep(4) 
        
        c1 = wait.until(EC.presence_of_element_located((By.ID, ID_CASILLERO_1)))
        c1.clear(); c1.send_keys(partes[0]); time.sleep(0.5)
        c2 = driver.find_element(By.ID, ID_CASILLERO_2); c2.clear(); c2.send_keys(partes[1]); time.sleep(0.5)
        c3 = driver.find_element(By.ID, ID_CASILLERO_3); c3.clear(); c3.send_keys(partes[2]); time.sleep(0.5)
        c4 = driver.find_element(By.ID, ID_CASILLERO_4); c4.clear(); c4.send_keys(partes[3]); time.sleep(0.5)
        c5 = driver.find_element(By.ID, ID_CASILLERO_5); c5.clear(); c5.send_keys(partes[4]); time.sleep(1)
        
        boton_buscar = driver.find_element(By.ID, ID_BOTON_BUSCAR_ABAJO)
        driver.execute_script("arguments[0].click();", boton_buscar)
        time.sleep(5) 
        
        try:
            casillero_fecha = wait.until(EC.element_to_be_clickable((By.ID, ID_FECHA_ACTUALIZACION)))
            casillero_fecha.click(); time.sleep(0.5)
            casillero_fecha.send_keys(Keys.END)
            for _ in range(10): casillero_fecha.send_keys(Keys.BACKSPACE)
            time.sleep(0.5); casillero_fecha.send_keys(fecha_pago_str); time.sleep(1)
            boton_confirmar_fecha = driver.find_element(By.ID, "BUTTON5")
            driver.execute_script("arguments[0].click();", boton_confirmar_fecha)
            time.sleep(6) 
        except Exception: pass
        
        filas = driver.find_elements(By.TAG_NAME, "tr")
        for fila in filas:
            texto_fila = fila.text.strip()
            if texto_fila.startswith(str(periodo_buscado)):
                datos = texto_fila.split()
                if len(datos) >= 6:
                    datos_extraidos["Importe Total"] = datos[-1]; datos_extraidos["Vencimiento"] = datos[1]
                    casilla = fila.find_element(By.TAG_NAME, "input")
                    driver.execute_script("arguments[0].click();", casilla); time.sleep(1)
                    boton_imprimir = driver.find_element(By.ID, ID_BOTON_IMPRIMIR_BOLETA)
                    driver.execute_script("arguments[0].click();", boton_imprimir); time.sleep(5) 
                    
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
                            ctx = ssl.create_default_context(); ctx.check_hostname = False; ctx.verify_mode = ssl.CERT_NONE
                            req = urllib.request.Request(pdf_url)
                            req.add_header("Cookie", texto_cookies)
                            req.add_header("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)")
                            with urllib.request.urlopen(req, context=ctx) as response:
                                with open(ruta_final, "wb") as f: f.write(response.read())
                            datos_extraidos["Estado"] = "PDF Descargado"
                        else: datos_extraidos["Estado"] = "Error: PDF no encontrado"
                    except Exception: datos_extraidos["Estado"] = "Error de descarga"
                    break 
        return datos_extraidos
    except Exception as e:
        try:
            driver.save_screenshot(os.path.join(carpeta_destino, f"ERROR_{nomenclatura}.png"))
            datos_extraidos["Estado"] = "Error (Ver foto en ZIP)"
        except: datos_extraidos["Estado"] = "Error Crítico"
        return datos_extraidos

st.set_page_config(page_title="Gestor EMOS", page_icon="💧", layout="centered")
st.title("💧 Gestor Automático - EMOS")
fecha_seleccionada = st.date_input("Fecha de pago:", datetime.date.today())
st.info("El Excel debe tener: Columna 1 (Nomenclatura) y Columna 2 (Periodo ej: 03/2026).")
archivo_subido = st.file_uploader("Sube tu archivo Excel para EMOS", type=["xlsx"])

if "proceso_terminado" not in st.session_state: st.session_state.proceso_terminado = False

if archivo_subido is not None:
    df = pd.read_excel(archivo_subido)
    st.write(f"Filas detectadas: {len(df)}")
    
    if st.button("🚀 Iniciar Búsqueda EMOS", use_container_width=True):
        carpeta_temp = "Boletas_EMOS_Temp"
        if os.path.exists(carpeta_temp): shutil.rmtree(carpeta_temp, ignore_errors=True)
        os.makedirs(carpeta_temp, exist_ok=True)
        for f_viejo in ["Boletas_EMOS.zip", "EMOS_Unidas.pdf"] + glob.glob("Reporte_EMOS*.xlsx"):
            if os.path.exists(f_viejo): os.remove(f_viejo)
                    
        resultados = []; barra = st.progress(0); estado = st.empty()
        
        try:
            estado.text("Iniciando motor EMOS...")
            chrome_options = Options(); chrome_options.add_argument("--window-size=1920,1080")
            if os.path.exists("/usr/bin/chromium"):
                chrome_options.binary_location = "/usr/bin/chromium"; chrome_options.add_argument("--headless=new") 
                chrome_options.add_argument("--no-sandbox"); chrome_options.add_argument("--disable-dev-shm-usage") 
                chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
                servicio = Service("/usr/bin/chromedriver")
            else:
                chrome_options.add_argument("--headless=new"); chrome_options.page_load_strategy = 'eager' 
                servicio = Service(ChromeDriverManager().install())
                
            driver = webdriver.Chrome(service=servicio, options=chrome_options)
            driver.set_page_load_timeout(180); wait = WebDriverWait(driver, 15)
            
            for index, row in df.iterrows():
                if pd.isna(row.iloc[0]): continue
                estado.text(f"Consultando: {row.iloc[0]}...")
                resultados.append(consultar_emos(driver, wait, row.iloc[0], row.iloc[1], carpeta_temp, fecha_seleccionada))
                barra.progress(int(((index + 1) / len(df)) * 100))

            driver.quit()
            
            df_res = pd.DataFrame(resultados); df_res.to_excel("Reporte_EMOS.xlsx", index=False)
            estado.text("Uniendo PDFs...")
            pdfs = glob.glob(os.path.join(carpeta_temp, "*.pdf"))
            if pdfs:
                fusionador = PdfWriter()
                for p in pdfs: fusionador.append(p)
                fusionador.write("EMOS_Unidas.pdf"); fusionador.close()
            
            shutil.make_archive("Boletas_EMOS", 'zip', carpeta_temp)
            st.session_state.proceso_terminado = True
            shutil.rmtree(carpeta_temp, ignore_errors=True); estado.empty()
            
        except Exception as e: st.error(f"Error: {e}")

if st.session_state.proceso_terminado:
    st.success("✅ ¡Terminado!")
    col1, col2, col3 = st.columns(3)
    with col1:
        if os.path.exists("Boletas_EMOS.zip"): st.download_button("📦 Bajar .ZIP", data=open("Boletas_EMOS.zip", "rb"), file_name="Boletas_EMOS.zip", mime="application/zip")
    with col2:
        if os.path.exists("EMOS_Unidas.pdf"): st.download_button("🖨️ Bajar PDF", data=open("EMOS_Unidas.pdf", "rb"), file_name="EMOS_Unidas.pdf", mime="application/pdf")
    with col3:
        if os.path.exists("Reporte_EMOS.xlsx"): st.download_button("📊 Bajar Excel", data=open("Reporte_EMOS.xlsx", "rb"), file_name="Reporte_EMOS.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
