"""Atualizacao assistida dos relatorios Sienge usados pelo dashboard.

O login e o eventual codigo de verificacao existem somente em memoria durante
a execucao. Os arquivos atuais so sao substituidos depois que todos os novos
relatorios passam pela validacao XLSX.
"""

from __future__ import annotations

import os
import json
import shutil
import tempfile
import time
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import Callable

from openpyxl import load_workbook


SIENGE_HOME = "https://gamadias.sienge.com.br/sienge/8/index.html"
REPORTS = [
    {
        "id": "apropriacoes",
        "name": "Relatório Analítico de Apropriações por Obra.xlsx",
        "url": SIENGE_HOME + "#/common/page/588",
        "title": "RELATÓRIO ANALÍTICO DE APROPRIAÇÕES POR OBRA",
        "date_label": "Período",
        "min_rows": 10,
        "min_cols": 10,
        "checkboxes": [
            ("Considerar os valores de descontos e acréscimos", True),
            ("Descontar os valores de impostos retidos", True),
            ("Apresentar títulos com valores de apropriações zerados", True),
            ("Apresentar dados em colunas não mescladas", True),
        ],
    },
    {
        "id": "a_pagar",
        "name": "Relatório de Contas a Pagar.xlsx",
        "url": SIENGE_HOME + "#/common/page/1674",
        "title": "RELATÓRIO DE CONTAS A PAGAR",
        "date_label": "Vencimento",
        "advanced": True,
        "direct_csv": True,
        "min_rows": 5,
        "min_cols": 8,
        "checkboxes": [
            ("Considerar impostos retidos", True),
        ],
    },
    {
        "id": "pagas",
        "name": "Relatório de Contas Pagas.xlsx",
        "url": SIENGE_HOME + "#/common/page/1843",
        "title": "RELATÓRIO DE CONTAS PAGAS",
        "date_label": "Período",
        "advanced": True,
        "advanced_required": False,
        "min_rows": 5,
        "min_cols": 10,
        "checkboxes": [],
    },
    {
        "id": "distorcoes",
        "name": "Relatório de Distorções das Apropriações por Obra.xlsx",
        "url": SIENGE_HOME + "#/common/page/609",
        "title": "RELATÓRIO DE DISTORÇÕES DAS APROPRIAÇÕES POR OBRA",
        "date_label": "Período",
        "min_rows": 2,
        "min_cols": 8,
        "checkboxes": [
            ("Considerar os valores de descontos e acréscimos", True),
            ("Descontar os valores de impostos retidos", True),
        ],
    },
]

_PERSISTENT_DRIVER = None
_PERSISTENT_DRIVER_PROFILE = ""


def _norm(value: object) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    return "".join(ch for ch in text if not unicodedata.combining(ch)).strip().lower()


def _validate_xlsx(path: Path, report: dict) -> dict:
    if not path.exists() or path.stat().st_size < 5_000:
        raise ValueError(f"Arquivo ausente ou pequeno demais: {path.name}")
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        if not wb.sheetnames:
            raise ValueError(f"{path.name} não contém planilhas.")
        ws = wb[wb.sheetnames[0]]
        if ws.max_row < report["min_rows"] or ws.max_column < report["min_cols"]:
            raise ValueError(
                f"{path.name} incompleto: {ws.max_row} linhas e {ws.max_column} colunas."
            )
        return {
            "file": report["name"],
            "rows": ws.max_row,
            "columns": ws.max_column,
            "size": path.stat().st_size,
        }
    finally:
        wb.close()


def _clear_chrome_profile_locks(profile_dir: Path) -> None:
    """Remove travas deixadas por uma execucao anterior do Chrome."""
    for name in ("SingletonCookie", "SingletonLock", "SingletonSocket", "DevToolsActivePort"):
        try:
            path = profile_dir / name
            if path.exists():
                path.unlink()
        except Exception:
            pass


def _chrome_options(profile_dir: Path, download_dir: Path, *, headless: bool = True):
    from selenium.webdriver.chrome.options import Options

    import chrome_env

    options = Options()
    chrome_binary = os.environ.get("CHROME_BINARY_PATH")
    if chrome_binary:
        options.binary_location = chrome_binary
    if headless or chrome_env.sem_tela():
        options.add_argument("--headless=new")
    options.add_argument("--window-size=1600,1000")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--no-sandbox")
    options.add_argument("--no-first-run")
    options.add_argument("--no-default-browser-check")
    options.add_argument(f"--user-data-dir={profile_dir}")
    options.page_load_strategy = "eager"
    options.add_experimental_option(
        "prefs",
        {
            "download.default_directory": str(download_dir),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
        },
    )
    return options


def _chrome_service():
    """Usa o chromedriver do CHROMEDRIVER_PATH quando definido (deploy em
    Linux/ARM com Chromium); sem essa variável, comportamento inalterado —
    o Selenium Manager resolve o driver sozinho, como já fazia no Windows."""
    from selenium.webdriver.chrome.service import Service

    driver_path = os.environ.get("CHROMEDRIVER_PATH")
    return Service(executable_path=driver_path) if driver_path else Service()


def _profile_slug(email: str) -> str:
    slug = "".join(ch if ch.isalnum() else "_" for ch in _norm(email))
    return slug.strip("_")[:80] or "usuario"


def _start_chrome_driver(webdriver, data_path: Path, stage_dir: Path, status: Callable, email: str):
    """Abre o Chrome com retries curtos e perfil alternativo se o perfil salvo travar."""
    from selenium.common.exceptions import WebDriverException

    profile_dir = data_path / f".sienge_profile_{_profile_slug(email)}"
    profile_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    attempts = [
        ("perfil temporario visivel", data_path / f".sienge_profile_visible_{stamp}", False, True),
        ("perfil temporario limpo", data_path / f".sienge_profile_runtime_{stamp}", True, True),
        ("perfil salvo limpo", profile_dir, False, True),
    ]
    errors = []
    for label, profile, headless, clear_locks in attempts:
        try:
            profile.mkdir(parents=True, exist_ok=True)
            if clear_locks:
                _clear_chrome_profile_locks(profile)
            status("opening", f"Abrindo Chrome automatico ({label})...", None)
            driver = webdriver.Chrome(service=_chrome_service(), options=_chrome_options(profile, stage_dir, headless=headless))
            status("opening", f"Chrome aberto ({label}). Acessando o Sienge...", None)
            return driver, str(profile)
        except WebDriverException as exc:
            errors.append(f"{label}: {exc.msg or exc}")
            _clear_chrome_profile_locks(profile)
            time.sleep(1)
        except Exception as exc:
            errors.append(f"{label}: {exc}")
            _clear_chrome_profile_locks(profile)
            time.sleep(1)
    raise RuntimeError(
        "CHROME_START_FAILED: Nao consegui abrir o Chrome automatico para acessar o Sienge. "
        "Feche as janelas do Chrome abertas e tente novamente. Detalhe: "
        + " | ".join(errors[-2:])
    )


def _new_download(download_dir: Path, before: set[Path], timeout: int = 150) -> Path:
    deadline = time.time() + timeout
    while time.time() < deadline:
        partial = list(download_dir.glob("*.crdownload")) + list(download_dir.glob("*.tmp"))
        candidates = [
            p for p in download_dir.iterdir()
            if p.is_file() and p not in before and p.suffix.lower() in {".xlsx", ".xls", ".csv"}
        ]
        if candidates and not partial:
            newest = max(candidates, key=lambda p: p.stat().st_mtime)
            if newest.stat().st_size > 5_000:
                return newest
        time.sleep(1)
    raise TimeoutError("O Sienge não concluiu o download do Excel dentro do tempo esperado.")


def _switch_to_page(driver, text: str, depth: int = 0) -> bool:
    from selenium.common.exceptions import WebDriverException

    wanted = _norm(text)
    try:
        if wanted in _norm(driver.find_element("tag name", "body").text):
            return True
    except WebDriverException:
        pass
    if depth >= 4:
        return False
    frames = driver.find_elements("css selector", "iframe,frame")
    for frame in frames:
        try:
            driver.switch_to.frame(frame)
            if _switch_to_page(driver, text, depth + 1):
                return True
            driver.switch_to.parent_frame()
        except WebDriverException:
            try:
                driver.switch_to.default_content()
            except WebDriverException:
                pass
    return False


def _wait_page(driver, title: str, timeout: int = 45) -> None:
    deadline = time.time() + timeout
    while time.time() < deadline:
        driver.switch_to.default_content()
        if _switch_to_page(driver, title):
            return
        time.sleep(1)
    raise TimeoutError(f"A página '{title}' não carregou.")


def _visible(driver, selector: str):
    for element in driver.find_elements("css selector", selector):
        try:
            if element.is_displayed() and element.is_enabled():
                return element
        except Exception:
            continue
    return None


def _wait_visible(driver, selector: str, timeout: int = 30):
    deadline = time.time() + timeout
    while time.time() < deadline:
        element = _visible(driver, selector)
        if element:
            return element
        time.sleep(0.5)
    return None


def _wait_not_loading(driver, timeout: int = 60) -> None:
    deadline = time.time() + timeout
    last_body = ""
    while time.time() < deadline:
        try:
            driver.switch_to.default_content()
            body = _norm(driver.find_element("tag name", "body").text)
            last_body = body
            if body and body.strip() not in {"carregando", "loading"} and "carregando..." not in body:
                return
        except Exception:
            pass
        time.sleep(0.7)


def _click_text(driver, texts: list[str], timeout: int = 15) -> bool:
    from selenium.webdriver.common.by import By

    wanted = [_norm(text) for text in texts]
    deadline = time.time() + timeout
    while time.time() < deadline:
        for element in driver.find_elements(
            By.CSS_SELECTOR,
            "button,a,input[type=button],input[type=submit],[role=button],div[tabindex]",
        ):
            try:
                value = element.text or element.get_attribute("value") or element.get_attribute("title") or ""
                normalized = _norm(value)
                if element.is_displayed() and element.is_enabled() and any(item in normalized for item in wanted):
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", element)
                    element.click()
                    return True
            except Exception:
                continue
        try:
            clicked = driver.execute_script(
                """
                const norm=s=>(s||'').normalize('NFD').replace(/[\\u0300-\\u036f]/g,'')
                  .replace(/\\s+/g,' ').trim().toLowerCase();
                const wanted=arguments[0];
                for(const node of document.querySelectorAll('button,a,[role=button],div,span')){
                  const text=norm(node.textContent);
                  if(!text || !wanted.some(item=>text===item || text.includes(item))) continue;
                  const target=node.closest('button,a,[role=button],[tabindex]')||node;
                  const rect=target.getBoundingClientRect();
                  if(rect.width>0 && rect.height>0){target.click();return true;}
                }
                return false;
                """,
                wanted,
            )
            if clicked:
                return True
        except Exception:
            pass
        time.sleep(0.5)
    return False


def _click_text_any_context(driver, texts: list[str], timeout: int = 15) -> bool:
    deadline = time.time() + timeout
    def visit(depth=0):
        if _click_text(driver, texts, timeout=1):
            return True
        if depth >= 4:
            return False
        for frame in driver.find_elements("css selector", "iframe,frame"):
            try:
                driver.switch_to.frame(frame)
                if visit(depth + 1):
                    return True
                driver.switch_to.parent_frame()
            except Exception:
                driver.switch_to.default_content()
        return False

    while time.time() < deadline:
        for handle in driver.window_handles:
            try:
                driver.switch_to.window(handle)
                driver.switch_to.default_content()
                if visit():
                    return True
            except Exception:
                continue
        time.sleep(0.5)
    return False


def _click_sienge_id_login(driver, timeout: int = 25) -> bool:
    """Clica especificamente em ENTRAR COM SIENGE ID, sem confundir com ENTRAR."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        for handle in driver.window_handles:
            try:
                driver.switch_to.window(handle)
                driver.switch_to.default_content()
                clicked = driver.execute_script(
                    """
                    const norm=s=>(s||'').normalize('NFD').replace(/[\\u0300-\\u036f]/g,'')
                      .replace(/\\s+/g,' ').trim().toLowerCase();
                    const nodes=[...document.querySelectorAll('button,a,input,[role=button],div[tabindex]')];
                    for(const node of nodes){
                      const text=norm([node.textContent,node.value,node.title,node.getAttribute('aria-label')]
                        .filter(Boolean).join(' '));
                      if(text !== 'entrar com sienge id') continue;
                      const target=node.closest('button,a,[role=button],[tabindex]') || node;
                      const rect=target.getBoundingClientRect();
                      if(rect.width>0 && rect.height>0){
                        target.scrollIntoView({block:'center'});
                        target.click();
                        return true;
                      }
                    }
                    return false;
                    """
                )
                if clicked:
                    return True
            except Exception:
                pass
        time.sleep(0.5)
    return False


def _click_connected_account_card(driver, email: str = "") -> bool:
    """Seleciona o cartão de conta do Sienge ID sem depender do nome do usuário."""
    try:
        return bool(driver.execute_script(
            """
            const norm=s=>(s||'').normalize('NFD').replace(/[\\u0300-\\u036f]/g,'')
              .replace(/\\s+/g,' ').trim().toLowerCase();
            const email=norm(arguments[0]);
            const nodes=[...document.querySelectorAll(
              'button,a,[role=button],[tabindex],article,section,li,div'
            )].filter(node=>{
              const rect=node.getBoundingClientRect();
              return rect.width>0 && rect.height>0;
            });
            const candidates=nodes.filter(node=>{
              const text=norm(node.textContent);
              return text && ((email && text.includes(email)) || text.includes('conectado'));
            }).sort((a,b)=>{
              const textA=norm(a.textContent);
              const textB=norm(b.textContent);
              const emailA=email && textA.includes(email) ? 0 : 1;
              const emailB=email && textB.includes(email) ? 0 : 1;
              if(emailA!==emailB) return emailA-emailB;
              return a.getBoundingClientRect().width*a.getBoundingClientRect().height
                - b.getBoundingClientRect().width*b.getBoundingClientRect().height;
            });
            for(const node of candidates){
              const target=node.closest('button,a,[role=button],[tabindex]') || node;
              const text=norm(target.textContent);
              if(text.includes('desconectar') || text.includes('cancelar')) continue;
              target.scrollIntoView({block:'center'});
              target.click();
              return true;
            }
            return false;
            """,
            email,
        ))
    except Exception:
        return False


def _expand_section_for_field(driver, section_text: str, field_id: str) -> bool:
    return bool(
        driver.execute_script(
            """
            const norm=s=>(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'')
              .replace(/\\s+/g,' ').trim().toLowerCase();
            const wanted=norm(arguments[0]);
            const field=document.getElementById(arguments[1]);
            const visible=el=>!!el && el.offsetParent!==null && el.getClientRects().length>0;
            if(visible(field)) return true;
            const headings=[...document.querySelectorAll('div,span,td,th,label')]
              .filter(el=>norm(el.textContent)===wanted);
            for(const heading of headings){
              const containers=[heading,heading.parentElement,heading.parentElement?.parentElement].filter(Boolean);
              for(const container of containers){
                const controls=[...container.querySelectorAll('button,a,[role=button],[onclick],img,input[type=button]')];
                for(const control of controls){
                  if(!visible(control)) continue;
                  control.click();
                  if(visible(field)) return true;
                }
                if(visible(container) && typeof container.click==='function'){
                  container.click();
                  if(visible(field)) return true;
                }
              }
            }
            return visible(field);
            """,
            section_text,
            field_id,
        )
    )


def _expand_section_for_field_any_context(driver, section_text: str, field_id: str) -> bool:
    def visit(depth=0):
        if _expand_section_for_field(driver, section_text, field_id):
            return True
        if depth >= 5:
            return False
        for frame in driver.find_elements("css selector", "iframe,frame"):
            try:
                driver.switch_to.frame(frame)
                if visit(depth + 1):
                    return True
                driver.switch_to.parent_frame()
            except Exception:
                driver.switch_to.default_content()
        return False

    driver.switch_to.default_content()
    return visit()


def _inputs_by_label(driver, label: str):
    return driver.execute_script(
        """
        const norm=s=>(s||'').normalize('NFD').replace(/[\\u0300-\\u036f]/g,'')
          .replace(/\\s+/g,' ').trim().toLowerCase();
        const wanted=norm(arguments[0]);
        const nodes=[...document.querySelectorAll('label,td,th,span,div')];
        for(const node of nodes){
          const txt=norm(node.textContent);
          if(!txt || !txt.startsWith(wanted) || txt.length>wanted.length+35) continue;
          const scopes=[node.closest('tr'),node.parentElement,node.closest('fieldset'),node.closest('section')].filter(Boolean);
          for(const scope of scopes){
            const inputs=[...scope.querySelectorAll('input:not([type=hidden]):not([type=checkbox]):not([type=radio])')]
              .filter(e=>!e.disabled && !e.readOnly && e.offsetParent!==null && e.getClientRects().length>0);
            if(inputs.length) return inputs;
          }
        }
        return [];
        """,
        label,
    )


def _set_input(element, value: str, tab: bool = False) -> None:
    from selenium.webdriver.common.keys import Keys

    element.click()
    element.send_keys(Keys.CONTROL, "a")
    element.send_keys(value)
    if tab:
        element.send_keys(Keys.TAB)


def _set_labeled_values(driver, label: str, values: list[str]) -> bool:
    elements = _inputs_by_label(driver, label)
    if not elements:
        return False
    for element, value in zip(elements, values):
        _set_input(element, value, tab=True)
        time.sleep(0.7)
    return True


def _set_labeled_values_any_context(driver, label: str, values: list[str]) -> bool:
    def visit(depth=0):
        if _set_labeled_values(driver, label, values):
            return True
        if depth >= 5:
            return False
        for frame in driver.find_elements("css selector", "iframe,frame"):
            try:
                driver.switch_to.frame(frame)
                if visit(depth + 1):
                    return True
                driver.switch_to.parent_frame()
            except Exception:
                driver.switch_to.default_content()
        return False

    driver.switch_to.default_content()
    return visit()


def _set_input_by_id_any_context(driver, field_id: str, value: str) -> bool:
    def visit(depth=0):
        try:
            element = driver.find_element("id", field_id)
            if element.is_displayed() and element.is_enabled():
                _set_input(element, value, tab=True)
                time.sleep(1)
                return True
        except Exception:
            pass
        if depth >= 5:
            return False
        for frame in driver.find_elements("css selector", "iframe,frame"):
            try:
                driver.switch_to.frame(frame)
                if visit(depth + 1):
                    return True
                driver.switch_to.parent_frame()
            except Exception:
                driver.switch_to.default_content()
        return False

    driver.switch_to.default_content()
    return visit()


def _set_checkbox(driver, label: str, checked: bool) -> bool:
    element = driver.execute_script(
        """
        const norm=s=>(s||'').normalize('NFD').replace(/[\\u0300-\\u036f]/g,'')
          .replace(/\\s+/g,' ').trim().toLowerCase();
        const wanted=norm(arguments[0]);
        const labels=[...document.querySelectorAll('label')];
        for(const label of labels){
          if(!norm(label.textContent).includes(wanted)) continue;
          if(label.htmlFor){
            const e=document.getElementById(label.htmlFor); if(e) return e;
          }
          const inside=label.querySelector('input[type=checkbox]'); if(inside) return inside;
          const near=label.parentElement && label.parentElement.querySelector('input[type=checkbox]'); if(near) return near;
        }
        const nodes=[...document.querySelectorAll('td,span,div')];
        for(const node of nodes){
          if(!norm(node.textContent).includes(wanted)) continue;
          const scope=node.closest('tr')||node.parentElement;
          const e=scope && scope.querySelector('input[type=checkbox]'); if(e) return e;
        }
        return null;
        """,
        label,
    )
    if not element:
        return False
    if bool(element.is_selected()) != checked:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", element)
        element.click()
    return True


def _set_checkbox_any_context(driver, label: str, checked: bool) -> bool:
    def visit(depth=0):
        if _set_checkbox(driver, label, checked):
            return True
        if depth >= 5:
            return False
        for frame in driver.find_elements("css selector", "iframe,frame"):
            try:
                driver.switch_to.frame(frame)
                if visit(depth + 1):
                    return True
                driver.switch_to.parent_frame()
            except Exception:
                driver.switch_to.default_content()
        return False

    driver.switch_to.default_content()
    return visit()


def _capture_debug(driver, output_dir: Path, report_id: str) -> None:
    contexts = []

    def visit(path="top", depth=0):
        try:
            body = driver.find_element("tag name", "body").text[:8_000]
            inputs = driver.execute_script(
                """
                return [...document.querySelectorAll('input,button,a,[role=button]')].map(e=>({
                  tag:e.tagName,type:e.type||'',id:e.id||'',name:e.name||'',
                  text:(e.innerText||e.value||e.title||'').slice(0,300),
                  visible:!!(e.offsetParent!==null && e.getClientRects().length),
                  disabled:!!e.disabled,readOnly:!!e.readOnly
                }));
                """
            )
            contexts.append({"path": path, "body": body, "inputs": inputs})
        except Exception as exc:
            contexts.append({"path": path, "error": str(exc)})
        if depth >= 5:
            return
        for index, frame in enumerate(driver.find_elements("css selector", "iframe,frame")):
            try:
                driver.switch_to.frame(frame)
                visit(f"{path}/frame[{index}]", depth + 1)
                driver.switch_to.parent_frame()
            except Exception:
                driver.switch_to.default_content()

    try:
        driver.switch_to.default_content()
        visit()
        output_dir.mkdir(parents=True, exist_ok=True)
        with open(output_dir / f"sienge_debug_{report_id}.json", "w", encoding="utf-8") as fh:
            json.dump(contexts, fh, ensure_ascii=False, indent=2)
        driver.switch_to.default_content()
        driver.save_screenshot(str(output_dir / f"sienge_debug_{report_id}.png"))
    except Exception:
        pass


def _login(driver, email: str, password: str, status: Callable, code_provider: Callable[[], str]) -> None:
    from selenium.webdriver.common.keys import Keys

    def logged_in():
        try:
            body = _norm(driver.find_element("tag name", "body").text)
            return (
                "gamadias.sienge.com.br" in driver.current_url
                and "bem-vindo!" not in body
                and (
                    "pesquise uma funcionalidade" in body
                    or "base: rgd incorporacao" in body
                    or "ola, elvis henrique" in body
                    or "proximos eventos e webinars" in body
                    or "#/common/" in driver.current_url
                    or "#/home" in driver.current_url
                )
            )
        except Exception:
            return False

    def invalid_credentials(body: str) -> bool:
        markers = (
            "senha incorreta",
            "senha invalida",
            "senha inválida",
            "email ou senha incorretos",
            "e-mail ou senha incorretos",
            "email ou senha invalidos",
            "e-mail ou senha invalidos",
            "usuario/e-mail ou senha invalidos",
            "usuário/e-mail ou senha inválidos",
            "usuario ou senha invalidos",
            "usuário ou senha inválidos",
            "credenciais invalidas",
            "credenciais inválidas",
            "nao foi possivel entrar",
            "não foi possível entrar",
        )
        return any(word in body for word in markers) or ("senha" in body and "inval" in body)

    _accept_pending_alerts(driver)
    try:
        driver.get(REPORTS[0]["url"])
        _accept_pending_alerts(driver)
        _wait_page(driver, REPORTS[0]["title"], timeout=8)
        return
    except Exception:
        _accept_pending_alerts(driver)

    try:
        driver.get(SIENGE_HOME)
    except Exception:
        _accept_pending_alerts(driver)
        driver.get(SIENGE_HOME)
    _accept_pending_alerts(driver)
    _wait_not_loading(driver, timeout=12)
    if logged_in():
        return
    try:
        driver.get(REPORTS[0]["url"])
        _accept_pending_alerts(driver)
        _wait_page(driver, REPORTS[0]["title"], timeout=8)
        return
    except Exception:
        _accept_pending_alerts(driver)
        try:
            driver.get(SIENGE_HOME)
            _wait_not_loading(driver, timeout=12)
        except Exception:
            pass

    if "gamadias.sienge.com.br" in driver.current_url:
        status("login", "Clicando em Entrar com Sienge ID...", None)
        if not _click_sienge_id_login(driver, timeout=30):
            if logged_in():
                return
            _capture_debug(driver, Path(tempfile.gettempdir()), "login_home")
            try:
                screen = _norm(driver.find_element("tag name", "body").text)[:500]
            except Exception:
                screen = ""
            raise RuntimeError(
                f"O botão Entrar com Sienge ID não foi localizado. URL: {driver.current_url}. Tela: {screen}"
            )
        deadline = time.time() + 30
        while time.time() < deadline and "login.sienge.com.br" not in driver.current_url:
            time.sleep(0.5)

    email_input = _wait_visible(driver, "input[type=email],input[name*=email i],input[id*=email i]", timeout=30)
    if email_input:
        if not password:
            raise RuntimeError(
                "AUTH_PASSWORD_REQUIRED: A sessão salva do Sienge expirou. Informe a senha para renovar o acesso."
            )
        status("login", "Entrando pelo Sienge ID...")
        _set_input(email_input, email)
        email_input.send_keys(Keys.ENTER)
        time.sleep(1)

    password_input = _wait_visible(driver, "input[type=password]", timeout=30)
    if password_input:
        _set_input(password_input, password)
        password_input.send_keys(Keys.ENTER)
        time.sleep(2)

    deadline = time.time() + 90
    while time.time() < deadline:
        if logged_in():
            return
        body = _norm(driver.find_element("tag name", "body").text)
        invalid_password_words = (
            "senha incorreta", "senha invalida", "senha inválida",
            "email ou senha incorretos", "e-mail ou senha incorretos",
            "credenciais invalidas", "credenciais inválidas",
            "nao foi possivel entrar", "não foi possível entrar",
            "usuario ou senha invalidos", "usuário ou senha inválidos",
        )
        if invalid_credentials(body) or any(word in body for word in invalid_password_words):
            raise RuntimeError(
                "AUTH_PASSWORD_INVALID: A senha informada está incorreta. Digite a senha correta e tente novamente."
            )
        if "ja esta conectado ao sistema" in body:
            status("login", "Sessão ativa encontrada. Clicando em Prosseguir...")
            if not _click_text(driver, ["prosseguir"], timeout=10):
                raise RuntimeError("O botão Prosseguir da sessão ativa não foi localizado.")
            time.sleep(3)
            continue
        if "login.sienge.com.br" in driver.current_url and (
            "conectado" in body or _norm(email) in body
        ) and not any(word in body for word in ("verifique o codigo", "codigo de 6 digitos")):
            status("login", "Selecionando a conta conectada no Sienge ID...")
            if not _click_connected_account_card(driver, email):
                raise RuntimeError("O cartão da conta conectada no Sienge ID não foi localizado.")
            time.sleep(3)
            continue
        code_input = _visible(
            driver,
            "input[autocomplete=one-time-code],input[name*=code i],input[id*=code i],input[name*=token i],input[id*=token i]",
        )
        if code_input or any(word in body for word in ("codigo de verificacao", "verifique o codigo", "codigo de 6 digitos", "verification code")):
            status("awaiting_code", "Digite o código de verificação enviado pelo Sienge.")
            code = code_provider()
            if not code:
                raise TimeoutError("Código de verificação não informado.")
            code_inputs = [
                item for item in driver.find_elements("css selector", "input[type=text],input[type=tel],input[type=number]")
                if item.is_displayed() and item.is_enabled()
            ]
            if code_input:
                code_inputs = [code_input]
            if not code_inputs:
                raise RuntimeError("O campo do código de verificação não foi localizado.")
            if len(code_inputs) >= len(code) > 1:
                for item, digit in zip(code_inputs, code):
                    _set_input(item, digit)
            else:
                _set_input(code_inputs[0], code)
            time.sleep(1)
            try:
                if not logged_in():
                    _click_text(driver, ["verificar", "continuar"], timeout=3)
            except Exception:
                pass
            status("login", "Validando o código de verificação...")
            verify_deadline = time.time() + 60
            while time.time() < verify_deadline:
                if logged_in():
                    return
                verify_body = _norm(driver.find_element("tag name", "body").text)
                if any(word in verify_body for word in ("codigo invalido", "codigo incorreto", "codigo expirado")):
                    break
                if "ja esta conectado ao sistema" in verify_body:
                    status("login", "Sessão ativa encontrada. Clicando em Prosseguir...")
                    if not _click_text(driver, ["prosseguir"], timeout=10):
                        raise RuntimeError("O botão Prosseguir da sessão ativa não foi localizado.")
                    status("login", "Prosseguindo com a nova sessão do Sienge...")
                elif "login.sienge.com.br" in driver.current_url and (
                    "conectado" in verify_body or _norm(email) in verify_body
                ) and not any(word in verify_body for word in ("verifique o codigo", "codigo de 6 digitos")):
                    status("login", "Selecionando a conta conectada no Sienge ID...")
                    if not _click_connected_account_card(driver, email):
                        raise RuntimeError("O cartão da conta conectada no Sienge ID não foi localizado.")
                if "login.sienge.com.br" in driver.current_url and not any(
                    word in verify_body for word in ("verifique o codigo", "codigo de 6 digitos")
                ):
                    if not _click_connected_account_card(driver, email):
                        _click_text(
                            driver,
                            ["rgd incorporacao", "gamadias", "continuar", "acessar", "entrar"],
                            timeout=1,
                        )
                time.sleep(1)
        time.sleep(1)
    try:
        screen = _norm(driver.find_element("tag name", "body").text)[:350]
        url = driver.current_url
    except Exception:
        screen, url = "", ""
    if invalid_credentials(screen):
        raise RuntimeError(
            "AUTH_PASSWORD_INVALID: A senha informada está incorreta. Digite a senha correta e tente novamente."
        )
    if "auth/login" in url and "senha" in screen and ("usuario" in screen or "entrar" in screen):
        raise RuntimeError(
            "AUTH_PASSWORD_INVALID: Usuário/e-mail ou senha inválidos. Confira os dados e tente novamente."
        )
    raise TimeoutError(f"O login do Sienge não foi concluído. URL: {url}. Tela: {screen}")


def _click_excel_export_any_context(driver) -> bool:
    def visit(depth=0):
        try:
            clicked = driver.execute_script(
                """
                const norm=s=>(s||'').normalize('NFD').replace(/[\\u0300-\\u036f]/g,'')
                  .replace(/\\s+/g,' ').trim().toLowerCase();
                for(const node of document.querySelectorAll('button,a,[role=button],input,img,span')){
                  const signal=norm([node.textContent,node.title,node.alt,node.value,
                    node.getAttribute('aria-label'),node.getAttribute('href'),
                    node.getAttribute('onclick')].filter(Boolean).join(' '));
                  if(!/(xlsx|excel|exportar)/.test(signal)) continue;
                  const target=node.closest('button,a,[role=button],[onclick]')||node;
                  const rect=target.getBoundingClientRect();
                  if(rect.width>0 && rect.height>0){target.click();return true;}
                }
                return false;
                """
            )
            if clicked:
                return True
        except Exception:
            pass
        if depth >= 5:
            return False
        for frame in driver.find_elements("css selector", "iframe,frame"):
            try:
                driver.switch_to.frame(frame)
                if visit(depth + 1):
                    return True
                driver.switch_to.parent_frame()
            except Exception:
                driver.switch_to.default_content()
        return False

    for handle in driver.window_handles:
        try:
            driver.switch_to.window(handle)
            driver.switch_to.default_content()
            if visit():
                return True
        except Exception:
            continue
    return False


def _restore_report_window(driver, origin_handle: str, start_handles: set[str]) -> None:
    for handle in set(driver.window_handles) - start_handles:
        try:
            driver.switch_to.window(handle)
            driver.close()
        except Exception:
            pass
    handles = driver.window_handles
    if origin_handle in handles:
        driver.switch_to.window(origin_handle)
    elif handles:
        driver.switch_to.window(handles[0])
    driver.switch_to.default_content()


def _accept_pending_alerts(driver) -> None:
    for _ in range(3):
        try:
            driver.switch_to.alert.accept()
            time.sleep(0.3)
        except Exception:
            break


def _try_excel_export(driver, download_dir: Path, before: set[Path], direct_csv: bool = False) -> Path:
    origin_handle = driver.current_window_handle
    start_handles = set(driver.window_handles)
    if direct_csv:
        if not _click_text_any_context(driver, ["gerar relatório csv", "gerar relatorio csv"]):
            raise RuntimeError("O botão Gerar Relatório CSV não foi encontrado.")
    elif not _click_text_any_context(driver, ["visualizar"]):
        raise RuntimeError("O botão Visualizar não foi encontrado.")

    deadline = time.time() + 45
    while time.time() < deadline:
        try:
            downloaded = _new_download(download_dir, before, timeout=1)
            _restore_report_window(driver, origin_handle, start_handles)
            return downloaded
        except TimeoutError:
            pass
        if not direct_csv and _click_excel_export_any_context(driver):
            try:
                downloaded = _new_download(download_dir, before, timeout=8)
                _restore_report_window(driver, origin_handle, start_handles)
                return downloaded
            except TimeoutError:
                pass
        time.sleep(0.4)

    _restore_report_window(driver, origin_handle, start_handles)
    raise TimeoutError("Não foi possível acionar a exportação em Excel no visualizador do Sienge.")


def _download_report(driver, download_dir: Path, report: dict, status: Callable) -> Path:
    status("downloading", f"Preparando {report['name']}...")
    _accept_pending_alerts(driver)
    try:
        driver.get(report["url"])
    except Exception:
        _accept_pending_alerts(driver)
        driver.get(report["url"])
    _accept_pending_alerts(driver)
    if report["url"].split("#", 1)[-1] not in driver.current_url:
        driver.get(report["url"])
        _accept_pending_alerts(driver)
    _wait_page(driver, report["title"])
    # Alguns relatórios exibem validações atrasadas depois da navegação.
    for _ in range(10):
        _accept_pending_alerts(driver)
        time.sleep(0.25)

    if report.get("advanced"):
        try:
            expanded = _expand_section_for_field_any_context(
                driver, "Parâmetros avançados", "obra.empreend.cdEmpreendView"
            )
        except Exception:
            _accept_pending_alerts(driver)
            expanded = _expand_section_for_field_any_context(
                driver, "Parâmetros avançados", "obra.empreend.cdEmpreendView"
            )
        if not expanded and report.get("advanced_required", True):
            raise RuntimeError(f"Se????o Par??metros avan??ados n??o encontrada em {report['name']}.")
        if not expanded:
            status("downloading", f"Par??metros avan??ados n??o apareceram em {report['name']}; tentando preencher a Obra mesmo assim.")
        time.sleep(2)
    obra_set = _set_input_by_id_any_context(driver, "obra.empreend.cdEmpreendView", "1016")
    if not obra_set:
        obra_set = _set_labeled_values_any_context(driver, "Obra", ["1016"])
    if not obra_set:
        _capture_debug(driver, download_dir.parent, report["id"])
        raise RuntimeError(f"Campo Obra não encontrado em {report['name']}.")
    end_date = datetime.now().strftime("%d/%m/%Y")
    _set_labeled_values_any_context(driver, report["date_label"], ["01/01/2000", end_date])
    _set_labeled_values_any_context(driver, "Data do indicador para correção", [end_date])
    _set_labeled_values_any_context(driver, "Correção até", [end_date])
    _set_labeled_values_any_context(driver, "Indexador", ["0"])
    for label, checked in report["checkboxes"]:
        _set_checkbox_any_context(driver, label, checked)

    before = set(download_dir.iterdir())
    status("downloading", f"Baixando {report['name']}...")
    try:
        downloaded = _try_excel_export(driver, download_dir, before, report.get("direct_csv", False))
    except Exception:
        _capture_debug(driver, download_dir.parent, report["id"] + "_export")
        raise
    final_stage = download_dir / report["name"]
    if downloaded.suffix.lower() == ".csv":
        import csv
        from openpyxl import Workbook

        converted = download_dir / (downloaded.stem + ".xlsx")
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Relatório"
        with open(downloaded, "r", encoding="cp1252", newline="") as source:
            for row in csv.reader(source, delimiter=";", quotechar='"'):
                sheet.append(row)
        workbook.save(converted)
        downloaded.unlink(missing_ok=True)
        downloaded = converted
    if downloaded != final_stage:
        if final_stage.exists():
            final_stage.unlink()
        downloaded.replace(final_stage)
    _accept_pending_alerts(driver)
    return final_stage


def run_sync(
    data_dir: str,
    email: str,
    password: str,
    status: Callable[[str, str, dict | None], None],
    code_provider: Callable[[], str],
) -> dict:
    """Baixa, valida e publica os quatro relatorios."""
    from selenium import webdriver

    data_path = Path(data_dir)
    data_path.mkdir(parents=True, exist_ok=True)
    stage_dir = Path(tempfile.mkdtemp(prefix="sienge_reports_", dir=str(data_path)))
    backup_dir = data_path / "backup" / "relatorios_sienge" / datetime.now().strftime("%Y%m%d_%H%M%S")
    # Reusa a sessão autenticada para que o MFA seja solicitado apenas quando
    # o próprio Sienge expirar ou revogar o acesso.
    global _PERSISTENT_DRIVER, _PERSISTENT_DRIVER_PROFILE
    expected_profile = str(data_path / f".sienge_profile_{_profile_slug(email)}")
    driver = _PERSISTENT_DRIVER
    try:
        status("opening", "Abrindo o acesso seguro ao Sienge...", None)
        if driver and _PERSISTENT_DRIVER_PROFILE and _PERSISTENT_DRIVER_PROFILE != expected_profile:
            try:
                driver.quit()
            except Exception:
                pass
            driver = None
            _PERSISTENT_DRIVER = None
            _PERSISTENT_DRIVER_PROFILE = ""
        if driver:
            try:
                _ = driver.current_url
            except Exception:
                driver = None
                _PERSISTENT_DRIVER = None
                _PERSISTENT_DRIVER_PROFILE = ""
        if not driver:
            driver, profile_used = _start_chrome_driver(webdriver, data_path, stage_dir, status, email)
            _PERSISTENT_DRIVER = driver
            _PERSISTENT_DRIVER_PROFILE = profile_used
        driver.execute_cdp_cmd(
            "Page.setDownloadBehavior",
            {"behavior": "allow", "downloadPath": str(stage_dir)},
        )
        _login(driver, email, password, status, code_provider)
        try:
            driver.minimize_window()
        except Exception:
            pass
        staged = []
        validation = []
        for index, report in enumerate(REPORTS, start=1):
            status(
                "downloading",
                f"Baixando relatório {index} de {len(REPORTS)}: {report['name']}",
                {"current": index, "total": len(REPORTS)},
            )
            path = _download_report(driver, stage_dir, report, status)
            validation_result = _validate_xlsx(path, report)
            validation.append(validation_result)
            staged.append((report, path))
            backup_dir.mkdir(parents=True, exist_ok=True)
            target = data_path / report["name"]
            if target.exists() and not (backup_dir / target.name).exists():
                shutil.copy2(target, backup_dir / target.name)
            shutil.copy2(path, target)
            status(
                "downloading",
                f"{report['name']} salvo e validado.",
                {"current": index, "total": len(REPORTS), "files": validation},
            )

        backup_dir.mkdir(parents=True, exist_ok=True)
        for report, staged_path in staged:
            target = data_path / report["name"]
            os.replace(staged_path, target)
        return {
            "ok": True,
            "files": validation,
            "backup": str(backup_dir),
            "updated_at": datetime.now().isoformat(timespec="seconds"),
        }
    finally:
        try:
            if driver:
                driver.quit()
        except Exception:
            pass
        _PERSISTENT_DRIVER = None
        _PERSISTENT_DRIVER_PROFILE = ""
        shutil.rmtree(stage_dir, ignore_errors=True)
