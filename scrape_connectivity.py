from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import time, requests, csv, os
from datetime import datetime

# Ruta a tu ChromeDriver
CHROMEDRIVER_PATH = "C:/Users/Alvaro/Desktop/chromedriver-win64/chromedriver.exe"

# URL base
BASE_URL = "https://connectivity.office.com/"

# Configurar opciones de Chrome (modo incógnito y ventana completa)
options = Options()
options.add_argument("--window-size=1920,1080")
options.add_argument("--user-data-dir=C:/ChromeSeleniumProfile")

#options.add_argument("--incognito")  # Navegación privada
options.add_experimental_option("prefs", {
    "profile.default_content_setting_values.geolocation": 1  # permitir ubicación
})
# Servicio de ChromeDriver
service = Service(CHROMEDRIVER_PATH)

# Número máximo de reintentos para obtener el reporte
MAX_RETRIES = 3
final_url = None

for attempt in range(MAX_RETRIES):
    driver = webdriver.Chrome(service=service, options=options)
    try:
        print(f"➡️ [Intento {attempt+1}] Abriendo la página de Microsoft Connectivity Test...")
        driver.get(BASE_URL)

        try:
        # Esperar a que aparezca el popup de permisos de ubicación
            allow_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//button[.='Allow while visiting the site']"))
            )
            allow_button.click()
            print("✅ Popup de ubicación aceptado.")
        except:
            print("ℹ️ No apareció popup de ubicación o ya fue gestionado.")

        # Esperar a que el botón esté presente
        run_button = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//button[.//span[text()='Run test']]"))
        )

        # Scroll y clic via JavaScript
        driver.execute_script("arguments[0].scrollIntoView(true);", run_button)
        time.sleep(1)
        driver.execute_script("arguments[0].click();", run_button)

        print("✅ Botón detectado y clickeado correctamente...")

        # Esperar hasta que la URL contenga un ID real
        WebDriverWait(driver, 90).until(
            lambda d: "/report/" in d.current_url and not d.current_url.endswith("/temp")
        )

                # Esperar hasta que el test termine (título 'Network Test Report' aparece)
        try:
            WebDriverWait(driver, 120).until(
                EC.presence_of_element_located((By.XPATH, "//div[contains(text(), 'Network Test Report')]"))
            )
            print("✅ Test completado. Página de reporte cargada.")
        except:
            print("❌ Timeout esperando el resultado del test.")
            driver.quit()
            exit(1)

        final_url = driver.current_url
        print("📎 URL del nuevo reporte:", final_url)

        # Guardar el HTML completo (opcional)
        with open("reporte_connectivity_office.html", "w", encoding="utf-8") as f:  
            f.write(driver.page_source)
            print("💾 HTML del reporte guardado como 'reporte_connectivity_office.html'")

    except Exception as e:
        print(f"❌ Intento {attempt+1} – No se pudo obtener el reporte: {e}")
        driver.quit()
        if attempt < MAX_RETRIES - 1:
            print("🔄 Reintentando en unos segundos...")
            time.sleep(5)  # Espera breve antes de reintentar
            continue
        else:
            print("❌ No se pudo obtener el reporte tras varios intentos. ¿Falló la prueba?")
            exit(1)
    else:
        # Si llegamos aquí, el intento fue exitoso; salir del bucle
        
        break

# Verificar que obtuvimos una URL de reporte
if not final_url:
    # Si no hay URL final, significa que todos los intentos fallaron
    # (Ya habríamos hecho exit(1) en el else del último intento, por lo que este bloque es precautorio)
    raise SystemExit("Error: No se obtuvo ningún reporte de conectividad.")

# Extraer token desde el navegador
access_token = driver.execute_script("return window.localStorage.getItem('access_token');")
if not access_token:
    print("❌ No se pudo obtener el access_token del localStorage.")
    driver.quit()
    exit(1)

driver.quit()


# Extraer el ID del reporte desde la URL obtenida
report_id = final_url.split("/")[-1]
print("🆔 Report ID generado:", report_id)

# Consultar la API para obtener los datos del reporte
api_url = f"https://connectivity.office.com/api/NetworkTestReport/ReadReport/{report_id}"
headers = {
    "Authorization": f"Bearer {access_token}"
}
try:
    response = requests.get(api_url, headers=headers, timeout=10)

except requests.exceptions.RequestException as e:
    print("❌ Error de conexión al consultar la API:", e)
    exit(1)

print(f"🌐 Código de respuesta API: {response.status_code}")
if response.status_code != 200:
    print(f"❌ Error al consultar la API con el ID proporcionado (HTTP {response.status_code})")
    exit(1)

# Parsear respuesta JSON
data = {}
try:
    data = response.json()
except ValueError as e:
    print("❌ Error al parsear la respuesta de la API como JSON:", e)
    exit(1)

# Extraer métricas de interés del JSON
date_time = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC")
latency_rtt = data.get("skype", {}).get("rttLatency")
jitter = data.get("skype", {}).get("averageJitter")
packet_loss = data.get("skype", {}).get("packetLossRate")
packet_loss_pct = packet_loss * 100 if packet_loss is not None else None
teams_score = data.get("scoreInfo", {}).get("teamsScore")
media_success = data.get("skype", {}).get("mediaConnectivitySuccessful")
call_success = data.get("skype", {}).get("callSuccessful")

# Contar hops hasta Teams en el traceroute (si existe)
teams_trace = next((t for t in data.get("traceroutes", {}).get("tracerouteResults", []) 
                    if t.get("hostname") and "teams" in t.get("hostname")), None)
hops = len(teams_trace["path"]) if teams_trace and teams_trace.get("path") else None

# Guardar los resultados en CSV
csv_file = "teams_connectivity_metrics.csv"
headers = ["FechaHora", "RTT_Latency", "Average_Jitter", "Packet_Loss_Rate", "Teams_Score",
           "Media_Connectivity_Successful", "Call_Successful", "Traceroute_Hops"]

write_header = not os.path.exists(csv_file)
with open(csv_file, mode="a", newline='', encoding="utf-8") as f:
    writer = csv.writer(f)
    if write_header:
        writer.writerow(headers)
    writer.writerow([
        date_time, latency_rtt, jitter, packet_loss_pct, teams_score,
        media_success, call_success, hops
    ])

print(f"✅ Prueba finalizada. Resultados guardados en: {csv_file}")
