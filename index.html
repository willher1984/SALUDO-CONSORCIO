import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# 📄 Ruta del archivo Excel
excel_path = r"C:\Users\LENOVO\Desktop\William Herrrera\SEPTIEMBRE\CARNETS\CARNETIZACION.xlsx"

# 📄 Ruta del archivo HTML
html_path = r"C:\Users\LENOVO\Desktop\HTML navidad consorcio.html"

# 📘 Cargar archivo
df = pd.read_excel(excel_path)

# ⚙️ Configuración de Chrome
options = Options()
options.add_experimental_option("detach", True)
options.add_argument("--start-maximized")

# 🚀 Iniciar navegador
driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=options
)

# 🌐 Abrir WhatsApp Web
driver.get("https://web.whatsapp.com/")
print("📱 Escanea el código QR con tu celular para iniciar sesión...")
time.sleep(25)

# 🔁 Enviar mensajes
for i, row in df.iterrows():
    numero = str(row["NUMERO DE CEL"]).strip()
    nombre = str(row.iloc[2]).strip()  # Columna C

    if not numero.isdigit():
        print(f"⚠️ Número inválido en fila {i+2}: {numero}")
        df.at[i, "ENVIADO"] = "❌ Número inválido"
        continue

    mensaje = (
        f"Hola {nombre} 👋\n\n"
        f"Te saludamos desde *Consorcio BFTEL* 🎄✨\n\n"
        f"Queremos compartir contigo un mensaje especial preparado "
        f"para esta temporada.\n\n"
        f"📎 *Abre el archivo adjunto en tu navegador (Chrome o Edge)* "
        f"para verlo correctamente.\n\n"
        f"*Feliz Navidad y próspero Año Nuevo* 🎁🎆\n"
        f"*Consorcio BFTEL*"
    )

    link = (
        "https://web.whatsapp.com/send?"
        f"phone=57{numero}&"
        f"text={mensaje.replace(' ', '%20').replace(chr(10), '%0A')}"
    )

    driver.get(link)

    try:
        # ⌨️ Esperar caja de texto y enviar mensaje
        textbox = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//div[@role="textbox"]'))
        )
        time.sleep(2)
        textbox.send_keys("\n")

        # 📎 CLIP
        clip = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, "//span[@data-icon='plus-rounded']"))
        )
        clip.click()
        time.sleep(1)

        # 📄 DOCUMENTO
        documento_btn = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, "//span[text()='Documento']"))
        )
        documento_btn.click()
        time.sleep(1)

        # 📥 INPUT DOCUMENTO
        file_input = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, "//input[@type='file']"))
        )
        file_input.send_keys(html_path)

        time.sleep(3)

        # 📤 ENVIAR DOCUMENTO
        send_button_doc = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable(
                (By.XPATH, "//div[@aria-label='Enviar' or @aria-label='Send']")
            )
        )
        send_button_doc.click()

        print(f"✅ Mensaje y HTML enviados a {nombre} ({numero})")
        df.at[i, "ENVIADO"] = "✅ Enviado"

    except Exception as e:
        print(f"❌ Error enviando a {numero}: {e}")
        df.at[i, "ENVIADO"] = "❌ Error"

    time.sleep(5)

# 💾 Guardar resultados
df.to_excel(excel_path, index=False)
print("📄 Envíos finalizados. Revisa la columna 'ENVIADO' en tu Excel.")

