"""Integração com o AutoDoc (plataforma.autodoc.com.br) — sincroniza a
análise de projetos (revisões, status, cronograma de projetos) direto da
pasta de gestão da obra, via automação de navegador (Selenium).

Não existe API pública documentada do AutoDoc, e o site exige login
completo (e-mail + senha) — sem 2FA visível no fluxo testado. As
credenciais (AUTODOC_EMAIL / AUTODOC_SENHA) ficam só no iniciar_dashboard.bat
do usuário e são lidas por variável de ambiente, nunca digitadas por mim
nem expostas em log.

Sem credencial configurada, as funções de sincronização retornam um erro
claro, sem travar o dashboard.
"""

from __future__ import annotations

import json
import os
import time

BASE_URL = "https://plataforma.autodoc.com.br"
CRONOGRAMA_FOLDER_ID = "692bbb70a6a1b6254b068576"  # "7. CRONOGRAMA DE PROJETOS"
PROJETOS_FOLDER_ID = "692bbaf4a6a1b6254b0680dc"  # raiz "PROJETOS" (disciplinas)
IFC_MODELOS_FOLDER_ID = "69c6847090779ef45a8e4762"  # pasta "MODELO" com os .ifc por disciplina

# ids das 30 pastas de disciplina dentro de PROJETOS — extraídos uma vez dos
# hrefs da listagem (ver get_folder_ids), não mudam a não ser que a
# estrutura de pastas no AutoDoc seja reorganizada.
DISCIPLINA_FOLDER_IDS = {
    "00_GESTÃO": "692bbb130cbceb10980e186f",
    "ACÚSTICO": "692bbb050cbceb10980e1795",
    "AR CONDICIONADO - EXAUSTÃO": "692bbb040cbceb10980e178b",
    "ARQUITETURA": "692bbb04cf0d6b15d005a314",
    "BLOCOS DE FUNDAÇÃO": "69ea550d447531f13f246b2f",
    "COMPATIBILIZAÇÃO": "692bbb13cf0d6b15d005a40f",
    "CONTENÇÃO": "692bbb05cf0d6b15d005a32d",
    "ELÉTRICO": "692bbb04b5a2b031a409c58a",
    "ELEVADORES": "692bbb060cbceb10980e179d",
    "ESQUADRIAS": "692bbb06b5a2b031a409c5aa",
    "ESTRUTURA CONCRETO": "692bbb040cbceb10980e177f",
    "ESTRUTURA METÁLICA": "692bbb06cf0d6b15d005a337",
    "EXTENSÃO DE REDE": "692bbb10b5a2b031a409c65b",
    "FUNDAÇÃO": "692bbb04cf0d6b15d005a319",
    "GÁS": "692bbb05b5a2b031a409c594",
    "HIDROSSANITÁRIO": "692bbb04b5a2b031a409c58f",
    "IMPERMEABILIZAÇÃO": "692bbb05cf0d6b15d005a328",
    "INCÊNDIO": "692bbb04b5a2b031a409c587",
    "INTERIORES": "692bbb05b5a2b031a409c5a5",
    "IRRIGAÇÃO": "692bbb05b5a2b031a409c599",
    "LUMÍNICO - TÉRMICO": "692bbb05cf0d6b15d005a332",
    "PAISAGISMO": "692bbb050cbceb10980e1790",
    "PERSONALIZAÇÕES": "692bbb06b5a2b031a409c5ad",
    "PRESSURIZAÇÃO": "692bbb05cf0d6b15d005a323",
    "SEGURANÇA": "692bbb12cf0d6b15d005a3fe",
    "SONDAGEM": "692bbb0fb5a2b031a409c647",
    "SPDA": "692bbb04cf0d6b15d005a31e",
    "TELEFÔNICO - SISTEMAS": "692bbb040cbceb10980e1786",
    "TOPOGRAFIA": "692bbb050cbceb10980e1798",
    "UNIDADE DECORADA": "692bbb110cbceb10980e1855",
}

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dados", "synergia")
STATE_PATH = os.path.join(DATA_DIR, "autodoc_state.json")
IFC_DIR = os.path.join(DATA_DIR, "ifc_models")
IFC_INDEX_PATH = os.path.join(DATA_DIR, "ifc_models_index.json")
IFC_TMP_DIR = os.path.join(DATA_DIR, "ifc_downloads_tmp")


def credenciais_status():
    email = os.environ.get("AUTODOC_EMAIL", "").strip()
    senha = os.environ.get("AUTODOC_SENHA", "")
    if not email or not senha:
        return {"ok": False, "error": "AUTODOC_EMAIL/AUTODOC_SENHA não configurados no iniciar_dashboard.bat."}
    return {"ok": True, "email": email}


def _load_state():
    if not os.path.exists(STATE_PATH):
        return {}
    try:
        with open(STATE_PATH, "r", encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return {}


def _save_state(state):
    os.makedirs(DATA_DIR, exist_ok=True)
    tmp = STATE_PATH + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(state, fh, ensure_ascii=False, indent=2)
    os.replace(tmp, STATE_PATH)


_STATUS_EM_DIA = {"APROVADO"}


def calcular_atrasados(state=None):
    """Varre os documentos ATUAIS (não o histórico — revisão obsoleta não
    conta como pendência) de todas as disciplinas já sincronizadas e separa
    o que está com status diferente de "Aprovado" (ex.: Bloqueado, Pendente,
    Em análise, Reprovado) — é o sinal mais direto que temos vindo do
    próprio AutoDoc de que algo está parado."""
    state = state if state is not None else _load_state()
    det = state.get("disciplinas_detalhe") or {}
    atrasados, em_dia = [], []

    def _avaliar(nome, docs):
        for doc in docs or []:
            status = str(doc.get("status") or "").strip()
            if not status or status == "--":
                continue  # pasta, não documento — sem status pra avaliar
            item = {"disciplina": nome, **doc}
            if status.upper() in _STATUS_EM_DIA:
                em_dia.append(item)
            else:
                atrasados.append(item)

    for nome, d in det.items():
        _avaliar(nome, d.get("atuais"))
    _avaliar("00_GESTÃO (cronograma)", state.get("cronograma_atual"))

    return {"atrasados": atrasados, "em_dia_qtd": len(em_dia), "total_avaliado": len(atrasados) + len(em_dia)}


# ── ANÁLISE DE PROJETOS — registro de documentos, retrabalho e curva de
# entregas. Modelo inspirado em como o mercado de AEC acompanha entrega de
# projeto: um registro único de documentos (equivalente ao "document
# register" de uma CDE ISO 19650), índice de retrabalho (nº de revisões até
# aprovação — quanto maior, mais idas e vindas de coordenação/comentários) e
# curva de entregas acumuladas (S-curve de projeto, mesma lógica da curva
# físico-financeira de obra, aplicada aos documentos de projeto).

def _base_doc_id(cod):
    """Remove o sufixo de revisão do nome do arquivo (ex.: -R00, _R01) pra
    agrupar todas as revisões do MESMO documento ao longo do tempo — sem
    isso não dá pra saber quantas vezes um documento foi revisado."""
    import re
    return re.sub(r'[-_]R\d+', '', cod, flags=re.IGNORECASE)


def _parse_data_br(s):
    try:
        partes = str(s).strip().split('/')
        if len(partes) != 3:
            return None
        d, m, y = partes
        return f"{int(y):04d}-{int(m):02d}-{int(d):02d}"
    except Exception:
        return None


def analisar_documentos(state=None):
    state = state if state is not None else _load_state()
    det = state.get("disciplinas_detalhe") or {}

    registro = []                # revisão atual de cada documento, achatado, pra tabela única
    por_doc_base = {}            # base_id -> {disciplina, revisoes:[...]}
    status_por_disciplina = {}   # disciplina -> {status: contagem}

    for nome, d in det.items():
        contagem_status = {}
        for doc in d.get("atuais") or []:
            cod = doc.get("cod") or ""
            status = (doc.get("status") or "").strip()
            if not cod or not status or status == "--":
                continue
            registro.append({"disciplina": nome, **doc})
            contagem_status[status] = contagem_status.get(status, 0) + 1
            base = _base_doc_id(cod)
            por_doc_base.setdefault(base, {"disciplina": nome, "revisoes": []})
            por_doc_base[base]["revisoes"].append(doc)
        for doc in d.get("historico") or []:
            cod = doc.get("cod") or ""
            if not cod:
                continue
            base = _base_doc_id(cod)
            por_doc_base.setdefault(base, {"disciplina": nome, "revisoes": []})
            por_doc_base[base]["revisoes"].append(doc)
        if contagem_status:
            status_por_disciplina[nome] = contagem_status

    # ── índice de retrabalho: nº de revisões por documento ───────────────
    retrabalho = []
    for base, info in por_doc_base.items():
        n = len(info["revisoes"])
        datas = sorted(d2 for d2 in (_parse_data_br(r.get("data")) for r in info["revisoes"]) if d2)
        dias_total = None
        if len(datas) >= 2:
            from datetime import date
            dias_total = (date.fromisoformat(datas[-1]) - date.fromisoformat(datas[0])).days
        retrabalho.append({
            "disciplina": info["disciplina"], "documento": base,
            "revisoes": n, "dias_entre_revisoes": dias_total,
        })
    retrabalho.sort(key=lambda x: (-x["revisoes"], x["documento"]))

    ranking_map = {}
    for r in retrabalho:
        agg = ranking_map.setdefault(r["disciplina"], {"disciplina": r["disciplina"], "documentos": 0, "revisoes_total": 0})
        agg["documentos"] += 1
        agg["revisoes_total"] += r["revisoes"]
    ranking = list(ranking_map.values())
    for r in ranking:
        r["media_revisoes"] = round(r["revisoes_total"] / r["documentos"], 2) if r["documentos"] else 0
    ranking.sort(key=lambda x: -x["media_revisoes"])

    # ── curva de entregas: acumulado de documentos aprovados por data ────
    # (S-curve de projeto — mesma ideia da curva físico-financeira de obra,
    # aplicada à 1ª aprovação registrada de cada documento.)
    primeiras_aprovacoes = []
    for base, info in por_doc_base.items():
        aprovados = [r for r in info["revisoes"] if (r.get("status") or "").strip().upper() == "APROVADO"]
        datas_aprov = sorted(d2 for d2 in (_parse_data_br(r.get("data")) for r in aprovados) if d2)
        if datas_aprov:
            primeiras_aprovacoes.append(datas_aprov[0])
    primeiras_aprovacoes.sort()
    curva_map = {}
    acumulado = 0
    for dt in primeiras_aprovacoes:
        acumulado += 1
        curva_map[dt] = acumulado
    curva_entregas = [{"data": k, "acumulado": v} for k, v in sorted(curva_map.items())]

    total_docs = len(registro)
    total_aprovados = sum(1 for r in registro if (r.get("status") or "").upper() == "APROVADO")

    return {
        "registro": registro,
        "status_por_disciplina": status_por_disciplina,
        "retrabalho_documentos": retrabalho[:80],
        "retrabalho_ranking": ranking,
        "curva_entregas": curva_entregas,
        "total_documentos": total_docs,
        "pct_aprovados": round(100 * total_aprovados / total_docs, 1) if total_docs else 0,
    }


def _novo_driver():
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options

    import chrome_env

    opts = Options()
    opts.add_argument("--headless=new")
    opts.add_argument("--window-size=1440,1000")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--log-level=3")
    chrome_env.aplicar_opcoes(opts)
    return webdriver.Chrome(service=chrome_env.servico(), options=opts)


def _login(driver, email, senha):
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    driver.get(BASE_URL + "/")
    wait = WebDriverWait(driver, 20)
    try:
        WebDriverWait(driver, 4).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Eu aceito')]"))
        ).click()
    except Exception:
        pass
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='email']"))).send_keys(email)
    wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Continuar')]"))).click()
    senha_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='password']")))
    senha_input.send_keys(senha)
    wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Entrar')]"))).click()
    WebDriverWait(driver, 20).until(lambda d: "/suite" in d.current_url or "/projetos" in d.current_url)
    # a conta/empresa ativa é resolvida de forma assíncrona depois do login —
    # navegar pra uma pasta antes disso terminar cai numa tela de "sem
    # permissão" mesmo com o login certo. Espera aparecer um nome de conta
    # de verdade (não "Selecione") antes de seguir.
    def _conta_pronta(d):
        try:
            texto = d.find_element(By.TAG_NAME, "body").text
            return "Conta:" in texto and "Selecione" not in texto.split("Conta:")[1][:40]
        except Exception:
            return False
    try:
        WebDriverWait(driver, 15).until(_conta_pronta)
    except Exception:
        pass


def _esperar_linhas(driver, timeout=15):
    """Espera a tabela (Ant Design, carregada via AJAX) realmente ter
    linhas COM CONTEÚDO — as linhas aparecem no DOM (tr.ant-table-row) antes
    do texto ser preenchido (esqueleto de carregamento), então só checar
    presença da linha lê célula vazia. Também espera a QUANTIDADE de linhas
    parar de crescer (a lista carrega aos poucos), não só a primeira."""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait

    def _tem_conteudo(d):
        rows = d.find_elements(By.CSS_SELECTOR, "tr.ant-table-row")
        if not rows:
            return False
        primeira_com_texto = any(c.text.strip() for c in rows[0].find_elements(By.CSS_SELECTOR, "td.ant-table-cell"))
        return primeira_com_texto

    try:
        WebDriverWait(driver, timeout).until(_tem_conteudo)
    except Exception:
        return  # pasta pode genuinamente estar vazia

    # espera a contagem de linhas estabilizar (carregamento incremental)
    anterior = -1
    for _ in range(10):
        atual = len(driver.find_elements(By.CSS_SELECTOR, "tr.ant-table-row"))
        if atual == anterior:
            break
        anterior = atual
        time.sleep(0.4)


def _parse_tabela_documentos(driver):
    """Lê a tabela de documentos (cód/rev/título/tam/data/status) da pasta
    atualmente aberta — é uma tabela Ant Design (tr.ant-table-row >
    td.ant-table-cell), célula 1 = nome, 2 = rev, 3 = título, 4 = tamanho,
    5 = data, 6 = status. O AutoDoc não tem API pública — sensível a
    mudança de layout do site."""
    from selenium.webdriver.common.by import By

    linhas = []
    try:
        rows = driver.find_elements(By.CSS_SELECTOR, "tr.ant-table-row")
    except Exception:
        rows = []
    for row in rows:
        try:
            cells = row.find_elements(By.CSS_SELECTOR, "td.ant-table-cell")
        except Exception:
            continue
        textos = [c.text.strip() for c in cells]
        if not textos or not any(textos):
            continue
        linhas.append(textos)
    return linhas


def _parse_documento_linha(cel):
    """cel = lista de textos das células da linha. Linhas de ARQUIVO têm
    uma célula de checkbox vazia na frente (offset 1: nome, rev, título,
    tamanho, data, status). Linhas de PASTA não têm essa célula de
    checkbox — o nome já vem em [0] (offset 0). Detecta pelo primeiro
    campo estar vazio ou não."""
    if not cel:
        return {"cod": "", "rev": "", "titulo": "", "tamanho": "", "data": "", "status": ""}
    offset = 1 if not cel[0].strip() else 0

    def _get(i):
        idx = i + offset
        return cel[idx].strip() if idx < len(cel) else ""

    return {
        "cod": _get(0),
        "rev": _get(1),
        "titulo": _get(2),
        "tamanho": _get(3),
        "data": _get(4),
        "status": _get(5),
    }


def sincronizar_projetos():
    """Loga no AutoDoc, lê a pasta de Cronograma de Projetos (revisão atual
    + histórico completo de obsoletos) e a listagem de disciplinas dentro
    de PROJETOS. Demora ~30-60s (login + navegação real de página)."""
    cred = credenciais_status()
    if not cred.get("ok"):
        return {"ok": False, "error": cred.get("error")}

    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    driver = _novo_driver()
    try:
        _login(driver, cred["email"], os.environ.get("AUTODOC_SENHA", ""))
        wait = WebDriverWait(driver, 20)

        # ── pasta do cronograma: revisão atual ──────────────────────────
        driver.get(f"{BASE_URL}/projetos/directories/folder/{CRONOGRAMA_FOLDER_ID}")
        _esperar_linhas(driver)
        linhas_atuais = _parse_tabela_documentos(driver)
        if not linhas_atuais:
            time.sleep(2.5)
            linhas_atuais = _parse_tabela_documentos(driver)
        atuais = [_parse_documento_linha(p) for p in linhas_atuais]
        atuais = [d for d in atuais if d["cod"]]

        # ── histórico (obsoletos) ────────────────────────────────────────
        historico = []
        try:
            obs_btns = driver.find_elements(By.XPATH, "//*[contains(text(),'Obsoletos')]")
            if obs_btns:
                driver.execute_script("arguments[0].click();", obs_btns[0])
                time.sleep(1)
                _esperar_linhas(driver)
                historico = [_parse_documento_linha(p) for p in _parse_tabela_documentos(driver)]
                historico = [d for d in historico if d["cod"]]
        except Exception:
            pass

        # ── disciplinas (pasta PROJETOS raiz) ────────────────────────────
        driver.get(f"{BASE_URL}/projetos/directories/folder/{PROJETOS_FOLDER_ID}")
        _esperar_linhas(driver)
        disciplinas_raw = _parse_tabela_documentos(driver)
        if not disciplinas_raw:
            # a app as vezes some com as linhas por um instante na troca de
            # pasta (renderização assíncrona) — uma segunda tentativa depois
            # de uma pausa curta resolve.
            time.sleep(2.5)
            disciplinas_raw = _parse_tabela_documentos(driver)
        disciplinas = []
        for partes in disciplinas_raw:
            d = _parse_documento_linha(partes)
            if not d["cod"]:
                continue
            disciplinas.append({"nome": d["cod"], "ultima_atualizacao": d["data"]})

        # carrega o estado existente e só atualiza os campos desta sincronização
        # — sobrescrever tudo apagaria a checklist de compatibilização (dado
        # digitado manualmente, não vem do AutoDoc) numa próxima sincronização.
        state = _load_state()
        state["cronograma_atual"] = atuais
        state["cronograma_historico"] = historico
        state["disciplinas"] = disciplinas
        state["sync_at"] = time.strftime("%Y-%m-%d %H:%M:%S")
        _save_state(state)
        return {"ok": True, "revisoes_atuais": len(atuais), "revisoes_historicas": len(historico), "disciplinas": len(disciplinas)}
    except Exception as exc:
        return {"ok": False, "error": f"Falha ao sincronizar com o AutoDoc: {exc}"}
    finally:
        try:
            driver.quit()
        except Exception:
            pass


def _ler_pasta_completa(driver, folder_id):
    """Navega pra uma pasta e lê os documentos atuais + histórico
    (obsoletos). Reaproveitado tanto pelo cronograma quanto pelo crawl de
    disciplinas — mesma lógica de espera/retry pros dois casos."""
    from selenium.webdriver.common.by import By

    driver.get(f"{BASE_URL}/projetos/directories/folder/{folder_id}")
    _esperar_linhas(driver)
    linhas = _parse_tabela_documentos(driver)
    if not linhas:
        time.sleep(2.5)
        linhas = _parse_tabela_documentos(driver)
    atuais = [_parse_documento_linha(p) for p in linhas]
    atuais = [d for d in atuais if d["cod"]]

    historico = []
    try:
        obs_btns = driver.find_elements(By.XPATH, "//*[contains(text(),'Obsoletos')]")
        if obs_btns:
            driver.execute_script("arguments[0].click();", obs_btns[0])
            time.sleep(1)
            _esperar_linhas(driver)
            historico = [_parse_documento_linha(p) for p in _parse_tabela_documentos(driver)]
            historico = [d for d in historico if d["cod"]]
    except Exception:
        pass

    return atuais, historico


def sincronizar_projetos_completo(disciplinas=None):
    """Igual sincronizar_projetos(), mas entra em CADA pasta de disciplina
    (não só lista os nomes) e lê os documentos atuais + histórico de cada
    uma. Demora bem mais (~5-10 min pras 30 disciplinas) — pensado pra rodar
    sob demanda (script/relatório), não como botão de UI comum.

    disciplinas: lista opcional de nomes (chaves de DISCIPLINA_FOLDER_IDS)
    pra restringir o crawl só a algumas — por padrão lê todas."""
    cred = credenciais_status()
    if not cred.get("ok"):
        return {"ok": False, "error": cred.get("error")}

    alvo = disciplinas or list(DISCIPLINA_FOLDER_IDS.keys())
    driver = _novo_driver()
    try:
        _login(driver, cred["email"], os.environ.get("AUTODOC_SENHA", ""))

        atuais_crono, historico_crono = _ler_pasta_completa(driver, CRONOGRAMA_FOLDER_ID)

        disciplinas_detalhe = {}
        for nome in alvo:
            fid = DISCIPLINA_FOLDER_IDS.get(nome)
            if not fid:
                continue
            try:
                atuais_d, historico_d = _ler_pasta_completa(driver, fid)
            except Exception as exc:
                atuais_d, historico_d = [], []
                print(f"  [AUTODOC] falha ao ler disciplina {nome}: {exc}")
            disciplinas_detalhe[nome] = {"atuais": atuais_d, "historico": historico_d}

        state = _load_state()
        state["cronograma_atual"] = atuais_crono
        state["cronograma_historico"] = historico_crono
        state["disciplinas_detalhe"] = disciplinas_detalhe
        state["disciplinas"] = [
            {"nome": nome, "ultima_atualizacao": (d["atuais"][0]["data"] if d["atuais"] else "")}
            for nome, d in disciplinas_detalhe.items()
        ]
        state["sync_completo_at"] = time.strftime("%Y-%m-%d %H:%M:%S")
        _save_state(state)
        return {"ok": True, "disciplinas_lidas": len(disciplinas_detalhe)}
    except Exception as exc:
        return {"ok": False, "error": f"Falha ao sincronizar com o AutoDoc: {exc}"}
    finally:
        try:
            driver.quit()
        except Exception:
            pass


# ── MODELOS IFC (para compatibilização e levantamento de materiais) ────────
# A pasta IFC_MODELOS_FOLDER_ID guarda um .ifc por disciplina (ex.: estrutural,
# elétrico, hidráulico...) — diferente da pasta COMPATIBILIZAÇÃO, que guarda o
# modelo federado em .nwd (Navisworks, formato proprietário/binário que o
# visualizador 3D desta dashboard não consegue abrir). Só os .ifc daqui
# entram no visualizador.

def _load_ifc_index():
    if not os.path.exists(IFC_INDEX_PATH):
        return []
    try:
        with open(IFC_INDEX_PATH, "r", encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return []


def _save_ifc_index(idx):
    os.makedirs(DATA_DIR, exist_ok=True)
    tmp = IFC_INDEX_PATH + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(idx, fh, ensure_ascii=False, indent=2)
    os.replace(tmp, IFC_INDEX_PATH)


def _slug(nome):
    import re
    base = re.sub(r"[^A-Za-z0-9_.-]+", "_", nome).strip("_")
    return base or "modelo"


def _upsert_ifc_index(nome, **campos):
    idx = _load_ifc_index()
    entry = next((e for e in idx if e["nome"] == nome), None)
    if entry is None:
        entry = {"id": _slug(nome), "nome": nome}
        idx.append(entry)
    entry.update(campos)
    _save_ifc_index(idx)
    return entry


def listar_modelos_ifc_autodoc():
    """Lista (sem baixar) os arquivos .ifc disponíveis na pasta dedicada do
    AutoDoc — rápido (uma navegação), pra UI mostrar o que dá pra buscar."""
    cred = credenciais_status()
    if not cred.get("ok"):
        return {"ok": False, "error": cred.get("error")}

    driver = _novo_driver()
    try:
        _login(driver, cred["email"], os.environ.get("AUTODOC_SENHA", ""))
        driver.get(f"{BASE_URL}/projetos/directories/folder/{IFC_MODELOS_FOLDER_ID}")
        _esperar_linhas(driver)
        linhas = _parse_tabela_documentos(driver)
        if not linhas:
            time.sleep(2.5)
            linhas = _parse_tabela_documentos(driver)
        docs = [_parse_documento_linha(p) for p in linhas]
        docs = [d for d in docs if d["cod"] and d["cod"].lower().endswith(".ifc")]

        baixados = {e["nome"]: e for e in _load_ifc_index()}
        for d in docs:
            info = baixados.get(d["cod"])
            d["baixado"] = bool(info and info.get("status") == "pronto")
            d["model_id"] = info["id"] if info else None
        return {"ok": True, "modelos": docs}
    except Exception as exc:
        return {"ok": False, "error": f"Falha ao listar modelos IFC no AutoDoc: {exc}"}
    finally:
        try:
            driver.quit()
        except Exception:
            pass


def _marcar_baixando(nome):
    return _upsert_ifc_index(nome, status="baixando", erro=None)


def _marcar_erro(nome, mensagem):
    return _upsert_ifc_index(nome, status="erro", erro=mensagem)


def baixar_modelo_ifc(nome_arquivo):
    """Baixa um .ifc específico da pasta dedicada do AutoDoc (login real,
    clique no menu "..." da linha > "Download (com título)"), descompacta o
    .zip que o AutoDoc entrega e guarda o .ifc em IFC_DIR. Demora ~15-30s —
    pensado pra rodar em thread de fundo (ver rota /api/ifc/autodoc/baixar)."""
    import shutil
    import zipfile

    cred = credenciais_status()
    if not cred.get("ok"):
        _marcar_erro(nome_arquivo, cred.get("error"))
        return {"ok": False, "error": cred.get("error")}

    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.action_chains import ActionChains

    os.makedirs(IFC_TMP_DIR, exist_ok=True)
    os.makedirs(IFC_DIR, exist_ok=True)
    pasta_dl = os.path.join(IFC_TMP_DIR, _slug(nome_arquivo) + "_" + str(int(time.time())))
    os.makedirs(pasta_dl, exist_ok=True)

    opts = Options()
    opts.add_argument("--headless=new")
    opts.add_argument("--window-size=1440,1000")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--log-level=3")
    opts.add_experimental_option("prefs", {
        "download.default_directory": pasta_dl,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    })
    import chrome_env
    chrome_env.aplicar_opcoes(opts)
    driver = webdriver.Chrome(service=chrome_env.servico(), options=opts)
    try:
        driver.execute_cdp_cmd("Page.setDownloadBehavior", {"behavior": "allow", "downloadPath": pasta_dl})
        _login(driver, cred["email"], os.environ.get("AUTODOC_SENHA", ""))
        driver.get(f"{BASE_URL}/projetos/directories/folder/{IFC_MODELOS_FOLDER_ID}")
        _esperar_linhas(driver)

        linha_alvo = None
        for row in driver.find_elements(By.CSS_SELECTOR, "tr.ant-table-row"):
            cells = row.find_elements(By.CSS_SELECTOR, "td.ant-table-cell")
            textos = [c.text.strip() for c in cells]
            if any(t == nome_arquivo for t in textos):
                linha_alvo = row
                break
        if linha_alvo is None:
            raise RuntimeError(f'Arquivo "{nome_arquivo}" não encontrado na pasta do AutoDoc.')

        cells = linha_alvo.find_elements(By.CSS_SELECTOR, "td.ant-table-cell")
        ellipsis = cells[-1].find_element(By.CSS_SELECTOR, ".table-action-ellipsis-v")
        ActionChains(driver).move_to_element(ellipsis).click().perform()
        time.sleep(1)
        dl_item = driver.find_element(By.XPATH, "//*[contains(text(),'Download (com')]")
        dl_item.click()

        zip_path = None
        for _ in range(40):
            time.sleep(1)
            arquivos = os.listdir(pasta_dl)
            prontos = [f for f in arquivos if not f.endswith(".crdownload")]
            if prontos:
                zip_path = os.path.join(pasta_dl, prontos[0])
                break
        if not zip_path:
            raise RuntimeError("Download não concluiu a tempo (timeout de 40s).")

        with zipfile.ZipFile(zip_path) as zf:
            membros = [n for n in zf.namelist() if n.lower().endswith(".ifc")]
            if not membros:
                raise RuntimeError("O .zip baixado não contém nenhum arquivo .ifc.")
            with zf.open(membros[0]) as src:
                destino = os.path.join(IFC_DIR, _slug(nome_arquivo))
                with open(destino, "wb") as dst:
                    shutil.copyfileobj(src, dst)

        tamanho = os.path.getsize(destino)
        entry = _upsert_ifc_index(
            nome_arquivo,
            status="pronto",
            erro=None,
            arquivo=os.path.basename(destino),
            tamanho_bytes=tamanho,
            baixado_em=time.strftime("%Y-%m-%d %H:%M:%S"),
        )
        return {"ok": True, "item": entry}
    except Exception as exc:
        _marcar_erro(nome_arquivo, str(exc))
        return {"ok": False, "error": f"Falha ao baixar modelo IFC do AutoDoc: {exc}"}
    finally:
        try:
            driver.quit()
        except Exception:
            pass
        try:
            shutil.rmtree(pasta_dl, ignore_errors=True)
        except Exception:
            pass
