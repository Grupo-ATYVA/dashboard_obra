"""Coleta assistida dos resultados de rompimento de concreto no Melius Qualidade.

A API pública da Melius (usada em app.py / sync_melius_data) só expõe dados de
FVS (Ficha de Verificação de Serviço) e não cobre a tela "Rastreabilidade >
Concretagem", onde ficam os resultados reais de rompimento dos corpos de
prova. Esse módulo abre essa tela dentro de uma sessão logada de verdade
(Selenium) e extrai os mesmos dados que um usuário veria manualmente.

A senha existe somente em memória durante a execução e não é gravada em
disco, seguindo o mesmo modelo de segurança do sienge_sync.py.
"""

from __future__ import annotations

import os
import re
import time
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import Callable

MELIUS_LOGIN_URL = "https://app.meliusqualidade.com.br/apex/r/mq/sgq/login"


def _with_stale_retry(func, attempts: int = 4, delay: float = 0.6):
    """Reexecuta func() se a APEX re-renderizar a região via AJAX no meio da
    leitura (StaleElementReferenceException) — comum ao trocar de aba ou
    paginar um relatório interativo, quando o DOM antigo é substituído
    enquanto ainda estamos lendo os elementos dele."""
    from selenium.common.exceptions import StaleElementReferenceException

    last_exc = None
    for attempt in range(attempts):
        try:
            return func()
        except StaleElementReferenceException as exc:
            last_exc = exc
            time.sleep(delay)
    raise last_exc


def _norm(value: object) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    return "".join(ch for ch in text if not unicodedata.combining(ch)).strip().lower()


def _number(value):
    if value in (None, ""):
        return None
    match = re.search(r"-?\d+(?:[.,]\d+)?", str(value))
    if not match:
        return None
    try:
        return float(match.group(0).replace(",", "."))
    except ValueError:
        return None


def _clear_chrome_profile_locks(profile_dir: Path) -> None:
    for name in ("SingletonCookie", "SingletonLock", "SingletonSocket", "DevToolsActivePort"):
        try:
            path = profile_dir / name
            if path.exists():
                path.unlink()
        except Exception:
            pass


def _chrome_options(profile_dir: Path, *, headless: bool = True):
    from selenium.webdriver.chrome.options import Options

    import chrome_env

    options = Options()
    chrome_binary = os.environ.get("CHROME_BINARY_PATH")
    if chrome_binary:
        options.binary_location = chrome_binary
    if headless or chrome_env.sem_tela():
        options.add_argument("--headless=new")
    options.add_argument("--window-size=1440,1000")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--no-sandbox")
    options.add_argument("--no-first-run")
    options.add_argument("--no-default-browser-check")
    options.add_argument(f"--user-data-dir={profile_dir}")
    options.page_load_strategy = "eager"
    return options


def _chrome_service():
    """Usa o chromedriver do CHROMEDRIVER_PATH quando definido (deploy em
    Linux/ARM com Chromium); sem essa variável, comportamento inalterado —
    o Selenium Manager resolve o driver sozinho, como já fazia no Windows."""
    from selenium.webdriver.chrome.service import Service

    driver_path = os.environ.get("CHROMEDRIVER_PATH")
    return Service(executable_path=driver_path) if driver_path else Service()


def _start_chrome_driver(webdriver, data_path: Path, status: Callable):
    """Retorna (driver, perfil_temporario_pra_apagar_no_fim | None). O perfil
    temporário (um por execução, com timestamp) nunca era limpo depois do
    driver.quit() — foram se acumulando dezenas de pastas em dados/ ao longo
    das sincronizações, chegando a travar operações no Windows Explorer que
    tentassem mexer na pasta inteira. O "perfil salvo" (.melius_profile) é
    intencionalmente persistente entre execuções e nunca é apagado aqui."""
    from selenium.common.exceptions import WebDriverException

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    attempts = [
        ("perfil temporario", data_path / f".melius_profile_runtime_{stamp}", True),
        ("perfil salvo", data_path / ".melius_profile", False),
    ]
    errors = []
    for label, profile, headless in attempts:
        try:
            profile.mkdir(parents=True, exist_ok=True)
            _clear_chrome_profile_locks(profile)
            status("opening", f"Abrindo Chrome automatico ({label})...")
            driver = webdriver.Chrome(service=_chrome_service(), options=_chrome_options(profile, headless=headless))
            status("opening", f"Chrome aberto ({label}). Acessando a Melius...")
            return driver, (profile if label == "perfil temporario" else None)
        except WebDriverException as exc:
            errors.append(f"{label}: {exc.msg or exc}")
    raise RuntimeError("Não foi possível abrir o Chrome automático. " + " | ".join(errors))


def _cleanup_profile_dir(profile) -> None:
    if not profile:
        return
    import shutil
    try:
        shutil.rmtree(profile, ignore_errors=True)
    except Exception:
        pass


def _login(driver, email: str, password: str, status: Callable) -> None:
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    status("login", "Abrindo o login da Melius...")
    driver.get(MELIUS_LOGIN_URL)
    wait = WebDriverWait(driver, 25)
    user_field = wait.until(EC.presence_of_element_located((By.ID, "P101_USERNAME")))
    pass_field = driver.find_element(By.ID, "P101_PASSWORD")
    user_field.clear()
    user_field.send_keys(email)
    pass_field.clear()
    pass_field.send_keys(password)
    status("login", "Enviando credenciais...")
    button = next(
        b for b in driver.find_elements(By.TAG_NAME, "button")
        if b.text.strip() == "Acessar"
    )
    button.click()
    time.sleep(1.5)
    if "senha incorreta" in driver.page_source.lower() or "usu" in driver.page_source.lower() and "incorreto" in driver.page_source.lower():
        raise RuntimeError("A Melius recusou o usuário/senha informados para o login web.")
    try:
        wait.until(lambda d: "P101_USERNAME" not in d.page_source)
    except Exception:
        raise RuntimeError("Não foi possível confirmar o login na Melius (tela não mudou).")


def _open_concretagem_list(driver, status: Callable) -> None:
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    status("navigating", "Abrindo Rastreabilidade > Concretagem...")
    wait = WebDriverWait(driver, 20)
    rastreabilidade = wait.until(
        EC.presence_of_element_located((By.XPATH, "//span[@role='treeitem' and normalize-space()='Rastreabilidade']"))
    )
    li = rastreabilidade.find_element(By.XPATH, "./ancestor::li[1]")
    toggle = li.find_element(By.CSS_SELECTOR, ".a-TreeView-toggle")
    driver.execute_script("arguments[0].click();", toggle)
    time.sleep(0.6)
    link = wait.until(
        EC.presence_of_element_located((By.XPATH, "//a[normalize-space()='Concretagem']"))
    )
    driver.execute_script("arguments[0].click();", link)
    wait.until(EC.presence_of_element_located((By.XPATH, "//th//*[contains(text(),'Data Aplica')]")))
    time.sleep(0.5)


def _find_body_table(driver):
    from selenium.webdriver.common.by import By

    tables = driver.find_elements(By.TAG_NAME, "table")
    for t in tables:
        tid = t.get_attribute("id") or ""
        if tid.endswith("_orig"):
            return t
    return None


def _cell_text(cells, idx):
    try:
        return cells[idx].text.strip()
    except IndexError:
        return ""


def _scrape_concretagem_page(driver) -> list:
    from selenium.webdriver.common.by import By

    table = _find_body_table(driver)
    if table is None:
        return []
    rows = table.find_elements(By.TAG_NAME, "tr")
    items = []
    for row in rows:
        cells = row.find_elements(By.TAG_NAME, "td")
        if len(cells) < 10:
            continue
        det_link = None
        for a in row.find_elements(By.TAG_NAME, "a"):
            href = a.get_attribute("href") or ""
            if "rastreabilidade-itens" in href:
                det_link = href
                break
        if not det_link:
            continue
        items.append({
            "data_aplicacao": _cell_text(cells, 1),
            "fornecedor": _cell_text(cells, 2),
            "fck_projeto": _number(_cell_text(cells, 3)),
            "slump_previsto": _cell_text(cells, 4),
            "modulo_elasticidade": _cell_text(cells, 5),
            "status": _cell_text(cells, 6),
            "local": _cell_text(cells, 7),
            "qtde_concretagem": _number(_cell_text(cells, 8)),
            "volume_projetado": _number(_cell_text(cells, 9)),
            "volume_realizado": _number(_cell_text(cells, 10)) if len(cells) > 10 else None,
            "detalhe_url": det_link,
        })
    return items


def _scrape_concretagem_summaries(driver, status: Callable) -> list:
    """Percorre TODAS as páginas do relatório interativo de Concretagem —
    a versão anterior só lia a página 1 (a mais recente, por ordenação
    padrão da APEX), fazendo concretagens mais antigas desaparecerem
    silenciosamente do dashboard. Mesma paginação usada em
    _scrape_diario_summaries, adaptada pra essa tabela."""
    from selenium.webdriver.common.by import By

    items = []
    seen_links = set()
    page = 1
    while True:
        page_items = _with_stale_retry(lambda: _scrape_concretagem_page(driver))
        novos = 0
        for it in page_items:
            if it["detalhe_url"] in seen_links:
                continue
            seen_links.add(it["detalhe_url"])
            items.append(it)
            novos += 1
        status("navigating", f"{len(items)} concretagens encontradas (página {page})...")
        next_btn = None
        for b in driver.find_elements(By.CSS_SELECTOR, "button.a-IRR-button--pagination"):
            title = _norm(b.get_attribute("title"))
            if "prox" in title:
                next_btn = b
                break
        if next_btn is None or not next_btn.is_enabled():
            break
        if novos == 0 and page > 1:
            # segurança: se uma página não trouxe nenhum item novo, para —
            # evita loop infinito caso o botão "próxima" fique habilitado
            # sem realmente avançar (fim de lista mal sinalizado pela APEX).
            break
        driver.execute_script("arguments[0].click();", next_btn)
        time.sleep(1.0)
        page += 1
        if page > 40:
            break
    return items


def _scrape_corpo_prova_modal(driver, status: Callable) -> list:
    """Clica no link de Corpos de Prova (já localizado) e le o iframe da caixa de dialogo."""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    wait = WebDriverWait(driver, 12)
    try:
        iframe = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "iframe[src*='rastreabilidade-ensaio']"))
        )
    except Exception:
        return []
    resultados = []
    try:
        driver.switch_to.frame(iframe)
        body_table = None
        for t in driver.find_elements(By.TAG_NAME, "table"):
            if "Dias Resultado" in t.text or "FCK Resultado" in t.text:
                body_table = t
                break
        if body_table is not None:
            # As colunas variam de posição conforme a tabela esteja em modo edição
            # (fichas "Aguardando Resultados", com uma coluna extra de seletor) ou
            # somente leitura (fichas já aprovadas). Por isso mapeamos os índices
            # pelo texto do cabeçalho em vez de usar posições fixas.
            header_cells = body_table.find_elements(By.TAG_NAME, "th")
            col_idx = {}
            for i, th in enumerate(header_cells):
                label = _norm(th.text)
                if "identifica" in label:
                    col_idx["id"] = i
                elif "dias" in label:
                    col_idx["dias"] = i
                elif "fck" in label:
                    col_idx["fck"] = i
                elif "observa" in label:
                    col_idx["obs"] = i
            if "fck" not in col_idx or "dias" not in col_idx:
                return []
            for row in body_table.find_elements(By.TAG_NAME, "tr"):
                cells = row.find_elements(By.TAG_NAME, "td")
                if len(cells) <= col_idx["fck"]:
                    continue
                resultado = _number(_cell_text(cells, col_idx["fck"]))
                if resultado is None:
                    continue
                resultados.append({
                    "corpo_prova": _cell_text(cells, col_idx.get("id", -1)) if "id" in col_idx else "",
                    "idade_dias": _number(_cell_text(cells, col_idx["dias"])),
                    "resultado_mpa": resultado,
                    "observacao": _cell_text(cells, col_idx["obs"]) if "obs" in col_idx and len(cells) > col_idx["obs"] else "",
                })
    finally:
        driver.switch_to.default_content()
    return resultados


def _close_dialog(driver) -> None:
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys

    try:
        driver.switch_to.default_content()
        closers = driver.find_elements(By.CSS_SELECTOR, ".ui-dialog:not([style*='display: none']) .ui-dialog-titlebar-close")
        if closers:
            driver.execute_script("arguments[0].click();", closers[-1])
            time.sleep(0.3)
            return
    except Exception:
        pass
    try:
        driver.switch_to.default_content()
        driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
        time.sleep(0.3)
    except Exception:
        pass


def _scrape_caminhoes(driver, status: Callable) -> list:
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    wait = WebDriverWait(driver, 20)
    wait.until(EC.presence_of_element_located((By.XPATH, "//th//*[contains(text(),'Caminh')]")))
    table = _find_body_table(driver)
    if table is None:
        return []
    caminhoes = []
    rows = table.find_elements(By.TAG_NAME, "tr")
    row_count = len(rows)
    for idx in range(row_count):
        table = _find_body_table(driver)
        if table is None:
            break
        rows = table.find_elements(By.TAG_NAME, "tr")
        if idx >= len(rows):
            break
        row = rows[idx]
        cells = row.find_elements(By.TAG_NAME, "td")
        if len(cells) < 7:
            continue
        caminhao = {
            "referencia": _cell_text(cells, 0),
            "identificacao": _cell_text(cells, 1),
            "nota_fiscal": _cell_text(cells, 2),
            "lacre": _cell_text(cells, 3),
            "volume_m3": _number(_cell_text(cells, 4)),
            "slump_medido": _cell_text(cells, 5),
            "qualidade": _cell_text(cells, 6),
            "local_detalhado": _cell_text(cells, 7) if len(cells) > 7 else "",
            "resultados_idade": [],
        }
        cp_link = None
        for a in row.find_elements(By.TAG_NAME, "a"):
            href = a.get_attribute("href") or ""
            if "a-dialog-open" in href and "rastreabilidade-ensaio" in href:
                cp_link = a
                break
        if cp_link is not None:
            try:
                driver.execute_script("arguments[0].click();", cp_link)
                caminhao["resultados_idade"] = _scrape_corpo_prova_modal(driver, status)
            except Exception as exc:
                status("navigating", f"Falha ao ler corpos de prova do caminhão {caminhao['identificacao']}: {exc}")
            finally:
                _close_dialog(driver)
        caminhoes.append(caminhao)
    return caminhoes


def _open_diario_list(driver, status: Callable) -> None:
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    status("navigating", "Abrindo Gestão da Obra > Diário de Obra...")
    wait = WebDriverWait(driver, 20)
    gestao = wait.until(
        EC.presence_of_element_located((By.XPATH, "//span[@role='treeitem' and normalize-space()='Gestão da Obra']"))
    )
    li = gestao.find_element(By.XPATH, "./ancestor::li[1]")
    toggle = li.find_element(By.CSS_SELECTOR, ".a-TreeView-toggle")
    driver.execute_script("arguments[0].click();", toggle)
    time.sleep(0.6)
    link = wait.until(EC.presence_of_element_located((By.XPATH, "//a[normalize-space()='Diário de Obra']")))
    driver.execute_script("arguments[0].click();", link)
    wait.until(EC.presence_of_element_located((By.XPATH, "//th//*[contains(text(),'Data')]")))
    time.sleep(0.5)


def _scrape_diario_page(driver) -> list:
    from selenium.webdriver.common.by import By

    table = _find_body_table(driver)
    if table is None:
        return []
    page_items = []
    for row in table.find_elements(By.TAG_NAME, "tr"):
        cells = row.find_elements(By.TAG_NAME, "td")
        if len(cells) < 8:
            continue
        det_link = None
        for a in row.find_elements(By.TAG_NAME, "a"):
            href = a.get_attribute("href") or ""
            if "rel-sgq-obra-diario-v2-detalhe" in href:
                det_link = href
                break
        if not det_link:
            continue
        id_match = re.search(r"p262_id_sgq_obra_diario=(\d+)", det_link)
        diario_id = id_match.group(1) if id_match else ""
        if not diario_id:
            continue
        page_items.append({
            "id": diario_id,
            "data": _cell_text(cells, 1),
            "numero": _cell_text(cells, 2),
            "status": _cell_text(cells, 3),
            "dias_trabalhados": _number(_cell_text(cells, 4)),
            "qtde_colaboradores": _number(_cell_text(cells, 5)),
            "comentarios_gerais": _cell_text(cells, 6),
            "aprovadores": _cell_text(cells, 7),
            "detalhe_url": det_link,
        })
    return page_items


def _scrape_diario_summaries(driver, status: Callable) -> list:
    from selenium.webdriver.common.by import By

    items = []
    seen_ids = set()
    page = 1
    while True:
        page_items = _with_stale_retry(lambda: _scrape_diario_page(driver))
        for it in page_items:
            if it["id"] in seen_ids:
                continue
            seen_ids.add(it["id"])
            items.append(it)
        status("navigating", f"{len(items)} registros de diário listados...")
        next_btn = None
        for b in driver.find_elements(By.CSS_SELECTOR, "button.a-IRR-button--pagination"):
            title = _norm(b.get_attribute("title"))
            if "prox" in title:
                next_btn = b
                break
        if next_btn is None or not next_btn.is_enabled():
            break
        driver.execute_script("arguments[0].click();", next_btn)
        time.sleep(1.0)
        page += 1
        if page > 20:
            break
    return items


def _field_text(driver, field_id: str) -> str:
    try:
        return driver.execute_script(
            """
            const c = document.getElementById(arguments[0] + '_CONTAINER');
            if (!c) return '';
            const clone = c.cloneNode(true);
            const lbl = clone.querySelector('.t-Form-labelContainer');
            if (lbl) lbl.remove();
            return clone.textContent.trim();
            """,
            field_id,
        ) or ""
    except Exception:
        return ""


def _read_diario_atividades_table(driver) -> list:
    from selenium.webdriver.common.by import By

    atividades = []
    for t in driver.find_elements(By.TAG_NAME, "table"):
        tid = t.get_attribute("id") or ""
        if not tid.endswith("_orig"):
            continue
        try:
            header_table = driver.find_element(By.ID, tid[:-len("_orig")])
            headers = [_norm(th.text) for th in header_table.find_elements(By.TAG_NAME, "th")]
        except Exception:
            headers = []
        if not any("etapa" in h for h in headers):
            continue
        for row in t.find_elements(By.TAG_NAME, "tr"):
            cells = row.find_elements(By.TAG_NAME, "td")
            if len(cells) < 4:
                continue
            etapa, servicos, local, comentario = (_cell_text(cells, i) for i in range(4))
            if not (etapa or servicos or local or comentario):
                continue
            atividades.append({"etapa": etapa, "servicos": servicos, "local": local, "comentario": comentario})
        break
    return atividades


def _scrape_diario_atividades(driver, status: Callable) -> tuple:
    from selenium.webdriver.common.by import By

    try:
        tab = driver.find_element(
            By.XPATH, "//a[@role='tab' and .//span[normalize-space()='Atividades Realizadas']]"
        )
        driver.execute_script("arguments[0].click();", tab)
        time.sleep(0.8)
    except Exception:
        return [], _field_text(driver, "P262_REGISTRO_ATIVIDADES")
    atividades = _with_stale_retry(lambda: _read_diario_atividades_table(driver))
    return atividades, _field_text(driver, "P262_REGISTRO_ATIVIDADES")


def _scrape_diario_detail(driver, status: Callable) -> dict:
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID, "P262_CLIMA_MANHA_CONTAINER")))
    detail = {
        "clima_manha": _field_text(driver, "P262_CLIMA_MANHA"),
        "clima_tarde": _field_text(driver, "P262_CLIMA_TARDE"),
        "clima_noite": _field_text(driver, "P262_CLIMA_NOITE"),
        "pluviometria": _field_text(driver, "P262_PLUVIOMETRIA"),
        "engenheiro_responsavel": _field_text(driver, "P262_USU_IN_ENGENHEIRO"),
        "ocorrencias": _field_text(driver, "P262_REGISTRO_SINISTRO"),
    }
    atividades, atividades_adicionais = _scrape_diario_atividades(driver, status)
    detail["atividades"] = atividades
    detail["atividades_adicionais"] = atividades_adicionais
    return detail


def run_diario_sync(email: str, password: str, data_path: Path, status: Callable, cache_anterior: dict = None) -> dict:
    """Loga na Melius, percorre Gestão da Obra > Diário de Obra e devolve os registros.

    Mesma abordagem do run_sync (rompimentos): a API pública não expõe o
    Diário de Obra, então a coleta é feita navegando a tela real com Selenium.

    `cache_anterior` (opcional): resultado de uma sincronização anterior. Um
    diário de um dia já passado não muda mais depois de lançado (é um
    registro histórico), então diários já lidos antes são reaproveitados do
    cache em vez de reabrir a página de detalhe de novo — só diários novos
    (id ainda não visto) são realmente buscados no navegador.
    """
    from selenium import webdriver

    registros_anteriores = {
        r.get("id"): r for r in (cache_anterior or {}).get("registros", []) or [] if r.get("id")
    }

    driver = None
    profile_to_clean = None
    try:
        driver, profile_to_clean = _start_chrome_driver(webdriver, data_path, status)
        _login(driver, email, password, status)
        _open_diario_list(driver, status)
        summaries = _scrape_diario_summaries(driver, status)
        total = len(summaries)
        registros = []
        reaproveitados = 0
        for i, item in enumerate(summaries, start=1):
            anterior = registros_anteriores.get(item.get("id"))
            if anterior:
                reaproveitados += 1
                status("scraping", f"({i}/{total}) diário já lido antes — reaproveitando do cache...")
                registros.append(anterior)
                continue
            status("scraping", f"Lendo diário {i}/{total} ({item['data']})...")
            driver.get(item["detalhe_url"])
            time.sleep(0.5)
            try:
                detail = _scrape_diario_detail(driver, status)
            except Exception as exc:
                status("scraping", f"Falha ao ler detalhe do diário {i}/{total}: {exc}")
                detail = {}
            registros.append({**{k: v for k, v in item.items() if k != "detalhe_url"}, **detail})
        if reaproveitados:
            status("scraping", f"{reaproveitados} de {total} diários reaproveitados do cache (registro histórico não muda).")
        status("done", f"{len(registros)} registros de diário sincronizados com sucesso.")
        return {
            "ok": True,
            "updated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
            "registros": registros,
        }
    finally:
        if driver is not None:
            try:
                driver.quit()
            except Exception:
                pass
        _cleanup_profile_dir(profile_to_clean)


MELIUS_IDADE_FINAL = 28  # idade final do ensaio — acima disso o resultado nunca muda mais


def _concretagem_completa(c: dict) -> bool:
    """Uma concretagem cujo resultado de 28 dias já chegou nunca mais muda no
    Melius — reaproveitar o que já foi lido antes evita reabrir a página de
    detalhe (navegação real no navegador) pra cada uma das dezenas de
    concretagens a cada sincronização, que é o que deixava o "Atualizar
    Melius completo" lento."""
    for r in c.get("resultados_idade", []) or []:
        try:
            idade = float(r.get("idade_dias") or 0)
            resultado = float(r.get("resultado_mpa"))
        except (TypeError, ValueError):
            continue
        if idade >= MELIUS_IDADE_FINAL and resultado == resultado:  # resultado==resultado descarta NaN
            return True
    return False


def run_sync(email: str, password: str, data_path: Path, status: Callable, cache_anterior: dict = None) -> dict:
    """Loga na Melius, percorre Rastreabilidade > Concretagem e devolve os dados coletados.

    `status(state, message)` é chamado a cada etapa para permitir acompanhamento
    (ex.: gravar em um arquivo de status lido pelo dashboard).

    `cache_anterior` (opcional): resultado da sincronização anterior (mesmo
    formato do retorno desta função). Concretagens que já têm resultado de
    28 dias são reaproveitadas do cache em vez de reabrir a página de
    detalhe de novo — só concretagens novas ou ainda aguardando resultado
    final são realmente buscadas no navegador, tornando reexecuções bem
    mais rápidas que a primeira sincronização completa.
    """
    from selenium import webdriver

    concretagens_anteriores = {
        c.get("id"): c for c in (cache_anterior or {}).get("concretagens", []) or [] if c.get("id")
    }

    driver = None
    profile_to_clean = None
    try:
        driver, profile_to_clean = _start_chrome_driver(webdriver, data_path, status)
        _login(driver, email, password, status)
        _open_concretagem_list(driver, status)
        summaries = _scrape_concretagem_summaries(driver, status)
        total = len(summaries)
        concretagens = []
        reaproveitadas = 0
        for i, item in enumerate(summaries, start=1):
            id_match = re.search(r"p277_id_sgq_rastreabilidade=(\d+)", item.get("detalhe_url", ""))
            item_id = id_match.group(1) if id_match else ""
            anterior = concretagens_anteriores.get(item_id) if item_id else None
            if anterior and _concretagem_completa(anterior):
                reaproveitadas += 1
                status("scraping", f"({i}/{total}) já com resultado final — reaproveitando do cache...")
                concretagens.append(anterior)
                continue
            status("scraping", f"Lendo concretagem {i}/{total} ({item['data_aplicacao']})...")
            driver.get(item["detalhe_url"])
            time.sleep(0.4)
            try:
                caminhoes = _scrape_caminhoes(driver, status)
            except Exception as exc:
                status("scraping", f"Falha ao ler caminhões da concretagem {i}/{total}: {exc}")
                caminhoes = []
            resultados_idade = []
            for c in caminhoes:
                for r in c.get("resultados_idade", []):
                    resultados_idade.append({**r, "caminhao": c.get("identificacao", "")})
            concretagens.append({
                "id": item_id,
                **{k: v for k, v in item.items() if k != "detalhe_url"},
                "caminhoes": caminhoes,
                "resultados_idade": resultados_idade,
            })
        if reaproveitadas:
            status("scraping", f"{reaproveitadas} de {total} concretagens reaproveitadas do cache (resultado final já conhecido).")
        status("done", f"{len(concretagens)} concretagens sincronizadas com sucesso.")
        return {
            "ok": True,
            "updated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
            "concretagens": concretagens,
        }
    finally:
        if driver is not None:
            try:
                driver.quit()
            except Exception:
                pass
        _cleanup_profile_dir(profile_to_clean)
