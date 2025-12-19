import os
import time
import re
import subprocess

import pyautogui
import pyperclip
import pdfplumber


# ================== CONFIG GERAL ==================

# ✅ desacelera todas as ações (ajuda MUITO em PC mais lento)
pyautogui.PAUSE = 0.25

# ⚠️ você desativou o failsafe (ok, mas cuidado)
pyautogui.FAILSAFE = False

# Caminho do Chrome
CHROME_PATH = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

# Imagens de referência
PROFILE_IMAGE = os.path.join("assets", "profile_card.png")
DOWNLOAD_BUTTON = os.path.join("assets", "download_button_full.png")

# URL da planilha
SHEET_URL = "https://docs.google.com/spreadsheets/d/1EXUku-9f4YYVLluBsxdOwRLtOqUCsemhsgtGodsGcP8/edit?gid=0#gid=0"

# Pasta padrão de downloads
DOWNLOAD_FOLDER = os.path.join(os.path.expanduser("~"), "Downloads")

# Tempo necessário para download
DOWNLOAD_WAIT_TIME = 7

# Tempo de espera para carregar páginas
META_PAGE_WAIT = 12
SHEET_WAIT = 12

# Contas Meta Ads da Syngenta
ACCOUNTS = [
    {"id": "1650759615721192", "nome": "Fungicidas"},
    {"id": "385162694606413", "nome": "FeirasMarketing"},
    {"id": "1158167895477076", "nome": "MaisAgro"},
    {"id": "1019413338731101", "nome": "Inseticidas"},

    {"id": "1099735580201582", "nome": "Macfor | Syngenta Feiras/Marketing (BR92AGM204)"},
    {"id": "462092151343521", "nome": "Macfor | Syngenta Atua Agro/Sustentação (BR92AGM210)"},
    {"id": "852980785172515", "nome": "Macfor | Syngenta Herbicidas (BR92AGM2AN)"},
    {"id": "172360363998930", "nome": "Macfor | Syngenta Fungicidas (BR92AGM20O)"},
    {"id": "1771904982934146", "nome": "Macfor | Syngenta Inseticidas (BR92AGM20Q)"},
    {"id": "950703088657879", "nome": "Macfor | Syngenta Seedcare (BR92AGM2BA)"},
    {"id": "1020067378368871", "nome": "Macfor | Syngenta Acessa Agro (BR92AGM10C)"},
    {"id": "765684883958323", "nome": "Macfor | Syngenta Estação Conhecimento (BR92AGM20N)"},
    {"id": "571942740337298", "nome": "Macfor | Syngenta Digital Agriculture (BR92AGM211)"},
    {"id": "133228928129814", "nome": "Macfor | Syngenta Evento - RAD"},
    {"id": "542477196378439", "nome": "Macfor | Syngenta Ornamentais"},
    {"id": "842449349581233", "nome": "Macfor | Syngenta PPM (BR92PPM100)"},
    {"id": "258653771927366", "nome": "Macfor | Syngenta Stewardship (BR92AGR600)"},
    {"id": "699596767278840", "nome": "Macfor | Syngenta Debate Algodão (BR92AGM106)"},
    {"id": "553400571903200", "nome": "Syngenta | Cultura Cana"},
    {"id": "3405215566172828", "nome": "Macfor | Syngenta Cultura Cana (BR92AGM10F)"},
    {"id": "590995028486503", "nome": "Macfor | Syngenta Asset Trigo (BR92AGM10E)"},
    {"id": "668193860452874", "nome": "Macfor | Syngenta Certificação Seedcare (BR92AGM20H)"},
    {"id": "636805440581086", "nome": "Macfor | Syngenta Webinar Soja (BR92AGM10G)"},
    {"id": "1235187726838699", "nome": "Macfor | Syngenta Seedcare Institute (BR92AGM2BB)"},
    {"id": "1028559497664734", "nome": "20anosSyngenta"},
    {"id": "396654441512296", "nome": "Macfor | Syngenta Mitos e Verdades (BR92AGM20H)"},
    {"id": "710129026283488", "nome": "Macfor | Syngenta Institucional (BR92AGM10G)"},
    {"id": "701118754125401", "nome": "Macfor | Syngenta Acessa Agro Revus (BR92AGM209)"},
    {"id": "320602035639647", "nome": "Macfor | Syngenta Nucoffee (BR92AGM200)"},
    {"id": "1062535854251486", "nome": "Macfor | Syngenta Bela Safra (BR92AGM204)"},
    {"id": "488321232300961", "nome": "Macfor | Syngenta Cultura Café (BR92AGM10G)"},
    {"id": "329192268651342", "nome": "Macfor | Syngenta Cultura Soja (BR92AGM10G)"},
    {"id": "1750540025140665", "nome": "Macfor | Syngenta Dipagro"},
    {"id": "563177111583972", "nome": "Macfor | Syngenta Mais Agro (BR92XXE81A)"},
    {"id": "1224315364723889", "nome": "Macfor | Syngenta Minor Crops"},
    {"id": "238882792157244", "nome": "Macfor | Syngenta Dia das Mães (BR92AGM203)"},
    {"id": "127479630380237", "nome": "Macfor | Syngenta Synap (BR92AGS9BZ)"},
    {"id": "220997594237645", "nome": "Macfor | Syngenta Verdavis (BR92AGM2AI)"},
    {"id": "951052662908228", "nome": "Syngenta Biológicos (BR92AGM2BJ)"},
]


# ================== HELPERS ==================

def build_url(account_id: str) -> str:
    return (
        "https://business.facebook.com/billing_hub/payment_activity"
        f"?asset_id={account_id}"
        "&business_id=150272408869988"
        "&placement=ads_manager"
        f"&payment_account_id={account_id}"
    )


def safe_locate(image_path: str, confidence: float = 0.85):
    """Evita crash do ImageNotFoundException."""
    try:
        return pyautogui.locateOnScreen(image_path, confidence=confidence)
    except pyautogui.ImageNotFoundException:
        return None


def focus_address_bar_and_open(url: str, wait_after: float):
    """Abre URL na aba atual com Ctrl+L + paste."""
    pyperclip.copy(url)
    pyautogui.hotkey("ctrl", "l")
    time.sleep(0.5)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.5)
    pyautogui.press("enter")
    time.sleep(wait_after)


# ================== CHROME / META ==================

def open_chrome():
    print("🟦 Abrindo o Chrome...")
    subprocess.Popen([CHROME_PATH])
    time.sleep(5)


def select_profile():
    print("🟨 Procurando o card do perfil...")
    for _ in range(40):  # mais tempo
        location = safe_locate(PROFILE_IMAGE, confidence=0.80)
        if location:
            print("✅ Perfil encontrado!")
            pyautogui.click(pyautogui.center(location))
            time.sleep(6)
            return True
        time.sleep(1.5)

    print("❌ Não foi possível encontrar o card do perfil.")
    return False


def navigate_to_meta_account(url: str):
    print(f"➡️ Indo para a conta (Meta): {url}")
    focus_address_bar_and_open(url, wait_after=META_PAGE_WAIT)

    # buffer extra para garantir que o DOM terminou
    time.sleep(5)


def click_download_button():
    print("🟩 Procurando botão de download da última transação...")

    for _ in range(35):  # tempo maior
        location = safe_locate(DOWNLOAD_BUTTON, confidence=0.85)
        if location:
            print("✅ Botão de download encontrado!")
            pyautogui.click(pyautogui.center(location))
            print(f"⏳ Aguardando {DOWNLOAD_WAIT_TIME}s para download...")
            time.sleep(DOWNLOAD_WAIT_TIME)
            return True
        time.sleep(1)

    print("⚠️ Nenhum botão encontrado — conta pode não ter transações no período.")
    return False


# ================== PDF ==================

def get_latest_pdf():
    pdfs = [f for f in os.listdir(DOWNLOAD_FOLDER) if f.lower().endswith(".pdf")]
    if not pdfs:
        raise FileNotFoundError("Nenhum PDF encontrado na pasta Downloads.")

    pdfs.sort(key=lambda x: os.path.getmtime(os.path.join(DOWNLOAD_FOLDER, x)), reverse=True)
    latest = os.path.join(DOWNLOAD_FOLDER, pdfs[0])
    print("📄 PDF mais recente:", latest)
    return latest


def extract_metodo(text: str) -> str:
    # pega a linha que contém a bandeira (evita o bug de retornar "Pago")
    m = re.search(r"(MasterCard|Visa|American Express)\s+.*", text, re.IGNORECASE)
    return m.group(0).strip() if m else ""


def extract_from_pdf(pdf_path):
    print("📤 Extraindo dados do PDF...")

    with pdfplumber.open(pdf_path) as pdf:
        text = (pdf.pages[0].extract_text() or "")

    def rx(pattern, default=""):
        m = re.search(pattern, text)
        return m.group(1).strip() if m else default

    data = rx(r"Data de pagamento\s+(.+)")
    referencia = rx(r"Número de referência:\s*([A-Z0-9]+)")
    id_transacao = rx(r"Identificação da transação\s+(\d+-\d+)")
    valor = rx(r"R\$[\s]?([\d\.\,]+)")
    conta = rx(r"Número de identificação da conta:\s*(\d+)")

    metodo = extract_metodo(text)
    status = "Pago" if re.search(r"\bPago\b", text) else ("Falha" if re.search(r"\bFalha\b", text) else "")

    dados = {
        "data": data,
        "valor": valor,
        "metodo": metodo,
        "status": status,
        "id": id_transacao,
        "conta": conta,
        "referencia": referencia,
        "nome_pdf": os.path.basename(pdf_path),
    }

    print("✅ Dados extraídos:", dados)
    return dados


# ================== SHEETS ==================

def open_sheet(url: str):
    print("📊 Abrindo Google Sheets...")
    focus_address_bar_and_open(url, wait_after=SHEET_WAIT)

    # Força foco na grade (isso evita colar na barra do navegador)
    pyautogui.press("esc")
    time.sleep(0.5)
    pyautogui.hotkey("ctrl", "home")
    time.sleep(0.6)

    print("📌 Planilha carregada e grade focada.")


def go_to_first_empty_row_column_a():
    """
    Encontra a primeira linha vazia na coluna A.
    FUNCIONA inclusive quando a planilha só tem cabeçalho.
    Estratégia:
      - vai para A1
      - desce para A2
      - seleciona a célula (F2) e copia (Ctrl+C)
    """
    print("📄 Procurando primeira linha vazia (coluna A)...")

    # topo
    pyautogui.hotkey("ctrl", "home")
    time.sleep(0.5)

    # A2
    pyautogui.press("down")
    time.sleep(0.4)

    # varre pra baixo até achar vazia
    for _ in range(2000):  # limite de segurança
        pyautogui.press("f2")          # garante modo edição da célula (clipboard pega o conteúdo)
        time.sleep(0.2)
        pyautogui.hotkey("ctrl", "a")  # seleciona conteúdo dentro da célula
        time.sleep(0.15)
        pyautogui.hotkey("ctrl", "c")
        time.sleep(0.25)

        cell_value = pyperclip.paste().strip()

        # sai do modo edição
        pyautogui.press("esc")
        time.sleep(0.15)

        if cell_value == "":
            print("✅ Linha vazia encontrada!")
            return

        pyautogui.press("down")
        time.sleep(0.12)

    raise RuntimeError("Não encontrei linha vazia (limite atingido).")


def write_to_sheet(values):
    # garante que não estamos na barra do navegador / modo edição estranho
    pyautogui.press("esc")
    time.sleep(0.3)

    campos = [
        values["data"],
        values["valor"],
        values["metodo"],
        values["status"],
        values["id"],
        values["conta"],
        values["referencia"],
        values["nome_pdf"],
    ]

    print("📝 Preenchendo linha na planilha...")

    for i, campo in enumerate(campos):
        pyperclip.copy(campo)
        time.sleep(0.2)
        pyautogui.hotkey("ctrl", "v")
        time.sleep(0.35)
        pyautogui.press("tab")
        time.sleep(0.35)

    # finaliza linha
    pyautogui.press("enter")
    time.sleep(1.5)

    # sai do modo edição e dá tempo pro autosave
    pyautogui.press("esc")
    time.sleep(1.2)

    # clique neutro (fora da grade) para parar cursor piscando
    w, _ = pyautogui.size()
    pyautogui.click(w - 60, 120)
    time.sleep(2.0)

    print("✅ Linha adicionada (aguardei autosave).")


# ================== MAIN ==================

def main():
    open_chrome()

    if not select_profile():
        return

    for conta in ACCOUNTS:
        print(f"\n🔵 Processando conta: {conta['nome']} (ID {conta['id']})")

        # 1) abrir conta no Meta
        url = build_url(conta["id"])
        navigate_to_meta_account(url)

        # 2) tentar baixar PDF
        ok = click_download_button()
        if not ok:
            print(f"➡️ Sem transações / sem botão — pulando {conta['nome']}.")
            continue

        # 3) pega PDF e extrai
        pdf_path = get_latest_pdf()
        dados = extract_from_pdf(pdf_path)

        # 4) abre planilha, acha primeira linha vazia e escreve
        open_sheet(SHEET_URL)
        go_to_first_empty_row_column_a()
        write_to_sheet(dados)

        # 5) buffer extra antes da próxima conta (evita corrida de aba/URL)
        time.sleep(2.5)

    print("\n🎉 Todas as contas foram processadas!\n")


if __name__ == "__main__":
    main()
