"""Ajustes do Chrome/Selenium por ambiente (Windows local x Linux/Docker).

No Windows nada muda. No Linux (servidor, sem monitor) só o modo headless
funciona, o Chromium vem do pacote do sistema e o chromedriver precisa ser
apontado por variável de ambiente (CHROME_BINARY_PATH / CHROMEDRIVER_PATH).
"""
import os
import sys

IS_LINUX = sys.platform.startswith("linux")


def sem_tela():
    """True num servidor Linux sem display: só headless consegue abrir o Chrome."""
    return IS_LINUX and not os.environ.get("DISPLAY")


def aplicar_opcoes(options):
    """Aplica binário do Chromium e flags exigidas em container Linux."""
    binary = os.environ.get("CHROME_BINARY_PATH")
    if binary:
        options.binary_location = binary
    if IS_LINUX:
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        if sem_tela():
            options.add_argument("--headless=new")
    return options


def servico():
    """Service do chromedriver: usa CHROMEDRIVER_PATH quando definido."""
    from selenium.webdriver.chrome.service import Service

    driver_path = os.environ.get("CHROMEDRIVER_PATH")
    return Service(executable_path=driver_path) if driver_path else Service()
