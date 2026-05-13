# -*- coding: utf-8 -*-
"""
Orquestrador: DEJT TST + JT Juris (headless) + envio WhatsApp
==============================================================
Executa dejt_tst_daily_headless.py -> jt_juris_teste_headless.py
e envia o DOCX resultante via WhatsApp Web (Selenium).

Variáveis de ambiente (ou .env na mesma pasta):
    WHATSAPP_PHONE       - número com DDI (ex: 5531996414518)
    JT_TRT_SELECTION     - "1"=TRT3, "2"=TRT24, "3"=Ambos (padrão: "3")
    JT_DOCX_PATH         - caminho do DOCX (padrão: Diario_J_TST_com_variaveis.docx)
    DEJT_BASE_DIR        - diretório base (padrão: pasta deste script)
    SKIP_WHATSAPP        - "1" para pular envio WhatsApp
"""
from __future__ import annotations

import os
import sys
import time
import subprocess
import logging
from pathlib import Path
from datetime import datetime

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
SCRIPT_DIR = Path(__file__).resolve().parent
BASE_DIR = Path(os.getenv("DEJT_BASE_DIR", str(SCRIPT_DIR)))

try:
    from dotenv import load_dotenv
    load_dotenv(SCRIPT_DIR / ".env")
except ImportError:
    pass

PHONE = os.getenv("WHATSAPP_PHONE", "5531996414518")
SKIP_WHATSAPP = os.getenv("SKIP_WHATSAPP", "0") == "1"
PYTHON = sys.executable

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
logger = logging.getLogger("SEXTA")

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def run_script(script_name: str, extra_env: dict | None = None) -> int:
    script = SCRIPT_DIR / script_name
    if not script.exists():
        logger.error(f"Script nao encontrado: {script}")
        return 1
    env = os.environ.copy()
    env["DEJT_BASE_DIR"] = str(BASE_DIR)
    if extra_env:
        env.update(extra_env)
    logger.info(f"Executando {script_name}...")
    result = subprocess.run(
        [PYTHON, str(script)],
        cwd=str(SCRIPT_DIR),
        env=env,
    )
    logger.info(f"{script_name} finalizado (exit={result.returncode})")
    return result.returncode


def find_docx() -> Path:
    """Encontra o DOCX gerado mais recente."""
    # Prioridade: variável de ambiente
    env_path = os.getenv("JT_DOCX_PATH")
    if env_path and Path(env_path).exists():
        return Path(env_path)
    # Busca padrão
    default = BASE_DIR / "Diario_J_TST_com_variaveis.docx"
    if default.exists():
        return default
    # Busca por data
    today = datetime.now().strftime("%Y-%m-%d")
    dated = BASE_DIR / f"Diario_J_TST_com_variaveis_{today}.docx"
    if dated.exists():
        return dated
    # Glob
    candidates = sorted(BASE_DIR.glob("Diario_J_TST_com_variaveis*.docx"),
                        key=lambda p: p.stat().st_mtime, reverse=True)
    if candidates:
        return candidates[0]
    raise FileNotFoundError("Nenhum DOCX gerado encontrado.")


def send_whatsapp(phone: str, docx_path: Path) -> bool:
    """Envia arquivo via WhatsApp Web usando Selenium."""
    try:
        from selenium import webdriver
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
    except ImportError:
        logger.error("selenium nao instalado. pip install selenium")
        return False

    logger.info(f"Enviando {docx_path.name} para {phone} via WhatsApp Web...")

    options = webdriver.ChromeOptions()
    options.add_argument(r"--user-data-dir=C:\Users\paollo\AppData\Local\Google\Chrome\User Data")
    options.add_argument("--profile-directory=Default")
    options.add_argument("--start-maximized")

    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 60)
    try:
        driver.get(f"https://web.whatsapp.com/send?phone={phone}")
        time.sleep(8)

        # Aguarda sessao autenticada
        try:
            wait.until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, "div[contenteditable='true']")))
        except Exception:
            logger.error("WhatsApp Web nao autenticado. Faca login manualmente primeiro.")
            return False

        # Mensagem
        msg_box = driver.find_element(By.CSS_SELECTOR, "div[contenteditable='true']")
        msg_box.click()
        msg_box.send_keys("Documento gerado automaticamente.")

        # Anexar arquivo
        attach_selectors = [
            "span[data-icon='plus']",
            "span[data-icon='clip']",
            "div[aria-label*='Anexar']",
            "div[title*='Anexar']",
        ]
        clicked = False
        for sel in attach_selectors:
            elems = driver.find_elements(By.CSS_SELECTOR, sel)
            if elems:
                elems[0].click()
                clicked = True
                break
        if not clicked:
            logger.error("Botao de anexar nao encontrado.")
            return False

        file_input = wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, "input[type='file']")))
        file_input.send_keys(str(docx_path.resolve()))
        time.sleep(5)

        # Enviar
        send_selectors = [
            "span[data-icon='send']",
            "button[aria-label='Send']",
            "div[aria-label='Send']",
        ]
        sent = False
        for sel in send_selectors:
            elems = driver.find_elements(By.CSS_SELECTOR, sel)
            if elems:
                elems[0].click()
                sent = True
                break
        if not sent:
            logger.error("Botao de enviar nao encontrado.")
            return False

        time.sleep(3)
        logger.info("WhatsApp: arquivo enviado com sucesso!")
        return True
    except Exception as e:
        logger.error(f"Erro no WhatsApp: {e}")
        return False
    finally:
        driver.quit()


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> int:
    logger.info("=" * 60)
    logger.info("DEJT + JT JURIS + WHATSAPP — Sexta-feira")
    logger.info(f"Base: {BASE_DIR}")
    logger.info("=" * 60)

    # 1) DEJT TST Daily
    rc1 = run_script("dejt_tst_daily_headless.py")

    # 2) JT Juris (appenda no mesmo DOCX)
    rc2 = run_script("jt_juris_teste_headless.py")

    # 3) Localizar DOCX gerado
    try:
        docx = find_docx()
        logger.info(f"DOCX encontrado: {docx}")
    except FileNotFoundError as e:
        logger.error(str(e))
        return 1

    # 4) WhatsApp
    if not SKIP_WHATSAPP:
        send_whatsapp(PHONE, docx)
    else:
        logger.info("Envio WhatsApp pulado (SKIP_WHATSAPP=1)")

    logger.info("Concluido!")
    return 0


if __name__ == "__main__":
    sys.exit(main())
