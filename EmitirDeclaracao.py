from selenium import webdriver  # pip install selenium
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import csv
from PyPDF2 import PdfReader
from fpdf import FPDF  # pip install fpdf2
from googletrans import Translator  # pip install googletrans==4.0.0-rc1

# =========================
# FUNÇÕES AUXILIARES
# =========================
def esperar_loading_sumir(navegador, timeout=40):
    """Espera o overlay de loading sumir"""
    WebDriverWait(navegador, timeout).until(
        EC.invisibility_of_element_located((By.ID, "loading-screen"))
    )

# =========================
# CONFIGURAÇÕES DO CHROME
# =========================
download_dir = r"C:\Users\Bruno Daniel\Desktop\Relatorios"

chrome_options = Options()
chrome_options.add_experimental_option(
    "prefs",
    {
        "download.default_directory": download_dir,
        "plugins.always_open_pdf_externally": True
    }
)
# chrome_options.add_argument("--headless=new")  # opcional

navegador = webdriver.Chrome(options=chrome_options)
navegador.maximize_window()
wait = WebDriverWait(navegador, 60)  # timeout maior

# =========================
# ACESSO AO SISTEMA
# =========================
navegador.get(
    "https://sistemas.portaledu.com.br/FrameHTML/web/app/edu/PortalEducacional/login/"
)

# =========================
# LOGIN
# =========================
campo_usuario = wait.until(EC.element_to_be_clickable((By.ID, "User")))
campo_usuario.clear()
campo_usuario.send_keys("2251346")

campo_senha = wait.until(EC.element_to_be_clickable((By.ID, "Pass")))
campo_senha.clear()
campo_senha.send_keys("110598")
campo_senha.send_keys(Keys.ENTER)

# =========================
# CONFIRMAÇÃO
# =========================
esperar_loading_sumir(navegador)
btn_confirmar = wait.until(EC.presence_of_element_located((By.ID, "btnConfirmar")))
navegador.execute_script("arguments[0].click();", btn_confirmar)

# =========================
# MENU → RELATÓRIOS
# =========================
esperar_loading_sumir(navegador)
menu = wait.until(EC.presence_of_element_located((By.ID, "show-menu")))
navegador.execute_script("arguments[0].click();", menu)

esperar_loading_sumir(navegador)
relatorios = wait.until(EC.presence_of_element_located((By.ID, "EDU_PORTAL_ACADEMICO_RELATORIOS")))
navegador.execute_script("arguments[0].click();", relatorios)

# =========================
# BOTÃO → EMITIR RELATÓRIO
# =========================
esperar_loading_sumir(navegador)
xpath_emitir = "//a[contains(normalize-space(.), 'Emitir relatório')]"
btn_emitir = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_emitir)))
navegador.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_emitir)
navegador.execute_script("arguments[0].click();", btn_emitir)

# =========================
# AGUARDAR DOWNLOAD DO PDF
# =========================
esperar_loading_sumir(navegador)
time.sleep(5)  # ajuste se PDFs forem grandes

# Busca o PDF mais recente
pdf_files = [os.path.join(download_dir, f) for f in os.listdir(download_dir) if f.lower().endswith(".pdf")]
pdf_path = max(pdf_files, key=os.path.getctime)

# =========================
# EXTRAI TEXTO DO PDF
# =========================
reader = PdfReader(pdf_path)
texto = ""
for page in reader.pages:
    page_text = page.extract_text()
    if page_text:
        texto += page_text + "\n"

# =========================
# TRADUZ TEXTO PARA INGLÊS
# =========================
translator = Translator()
texto_traduzido = translator.translate(texto, src='pt', dest='en').text

# =========================
# CRIA NOVO PDF TRADUZIDO
# =========================
pdf_traduzido = FPDF()
pdf_traduzido.add_page()
pdf_traduzido.set_auto_page_break(auto=True, margin=15)
pdf_traduzido.set_font("Arial", size=12)

for linha in texto_traduzido.split("\n"):
    pdf_traduzido.multi_cell(0, 10, linha)

novo_pdf_path = os.path.join(download_dir, "Relatorio_traduzido.pdf")
pdf_traduzido.output(novo_pdf_path)
print("PDF traduzido criado em:", novo_pdf_path)

# =========================
# SALVAR TEXTO TRADUZIDO EM CSV
# =========================
csv_path = os.path.join(download_dir, "Relatorio_traduzido.csv")
linhas = texto_traduzido.split("\n")

with open(csv_path, mode='w', newline='', encoding='utf-8') as csvfile:
    writer = csv.writer(csvfile)
    for linha in linhas:
        writer.writerow([linha])

print("Texto traduzido salvo em CSV:", csv_path)

input("Pressione ENTER para encerrar...")
navegador.quit()
