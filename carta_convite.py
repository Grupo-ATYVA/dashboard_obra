# -*- coding: utf-8 -*-
"""Robô da Carta Convite — gera a carta convite (.docx/.pdf) sozinho, sem IA.

Fluxo:
  1. Loga no AutoDoc (mesmo login do iniciar_dashboard.bat, via
     autodoc_sync) e varre as pastas configuradas em
     dados/synergia/templates/carta_convite_<pacote>.json, pegando a ÚLTIMA
     REVISÃO APROVADA de cada prancha (revisão mais nova ainda "Bloqueado"
     fica sinalizada, não entra).
  2. Lê os prazos por pavimento do Cronograma Base do PCP (xlsx).
  3. Monta o .docx completo (textos/cláusulas vêm do JSON de configuração —
     editar lá, não aqui) e, se o Word estiver instalado, exporta o PDF.

Uso (linha de comando / agendador do Windows):
  python carta_convite.py                    # sincroniza AutoDoc + gera
  python carta_convite.py --sem-autodoc      # usa as pranchas da última leitura
  python carta_convite.py --saida "C:\\pasta" --sem-pdf

O dashboard usa as mesmas funções (aba "Cartas Convite").
"""

from __future__ import annotations

import argparse
import copy
import datetime as dt
import glob
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
from io import BytesIO

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "dados", "synergia")
TEMPLATES_DIR = os.path.join(DATA_DIR, "templates")
SAIDA_DIR = os.path.join(DATA_DIR, "cartas_convite")
STATE_PATH = os.path.join(DATA_DIR, "carta_convite_state.json")
LOG_PATH = os.path.join(BASE_DIR, "logs", "robo_carta_convite.log")

MESES = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho",
         "agosto", "setembro", "outubro", "novembro", "dezembro"]

# Sufixo de local no código da prancha (SYB-ALV-PE-10100-SS2) -> descrição.
LOCAIS = {
    "SS2": "Subsolo 02", "SS1": "Subsolo 01", "TER": "Térreo", "PG1": "Garagem 01",
    "PG2": "Garagem 02", "RES": "Reservatório inferior", "TIP": "Pavimento tipo",
    "OPC": "Pavimento tipo – opção de planta", "LAZ": "Lazer", "COB": "Cobertura",
    "CTA": "Cortes A", "CTB": "Cortes B", "CTC": "Cortes de detalhe",
    "NOR": "Geral – quadros e tabelas", "GER": "Geral", "RSP": "Reservatório superior",
    "BAR": "Barrilete", "CMQ": "Casa de máquinas", "FCH": "Fachadas",
}


def _local_desc(local):
    m = re.fullmatch(r"(\d{2})P", local or "")
    if m:
        return f"{int(m.group(1))}º pavimento"
    return LOCAIS.get(local, local)


def _log(msg):
    linha = f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {msg}"
    try:
        print(linha, flush=True)
    except UnicodeEncodeError:
        print(linha.encode("ascii", "replace").decode(), flush=True)
    try:
        os.makedirs(os.path.dirname(LOG_PATH), exist_ok=True)
        with open(LOG_PATH, "a", encoding="utf-8") as fh:
            fh.write(linha + "\n")
    except Exception:
        pass


# ─── configuração / estado ────────────────────────────────────────────────
def config_path(pacote):
    return os.path.join(TEMPLATES_DIR, f"carta_convite_{pacote}.json")


def listar_pacotes():
    out = []
    for p in sorted(glob.glob(os.path.join(TEMPLATES_DIR, "carta_convite_*.json"))):
        pacote = os.path.basename(p)[len("carta_convite_"):-len(".json")]
        try:
            with open(p, "r", encoding="utf-8") as fh:
                nome = json.load(fh).get("nome", pacote)
        except Exception:
            nome = pacote
        out.append({"pacote": pacote, "nome": nome})
    return out


def carregar_config(pacote):
    with open(config_path(pacote), "r", encoding="utf-8") as fh:
        return json.load(fh)


def _load_state():
    try:
        with open(STATE_PATH, "r", encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return {}


def _save_state(state):
    os.makedirs(DATA_DIR, exist_ok=True)
    tmp = STATE_PATH + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(state, fh, ensure_ascii=False, indent=1)
    os.replace(tmp, STATE_PATH)


def estado_pacote(pacote):
    return _load_state().get(pacote) or {}


def _atualizar_estado(pacote, **campos):
    state = _load_state()
    st = state.setdefault(pacote, {})
    st.update(campos)
    _save_state(state)
    return st


# ─── AutoDoc: disciplinas, pranchas atuais e download ─────────────────────
# Cada linha da tabela do AutoDoc: o nome é o texto do link (a raiz PROJETOS
# não tem a coluna de checkbox, as subpastas têm — pelo link não importa).
_JS_LINHAS = """return Array.from(document.querySelectorAll('tr.ant-table-row')).map(r => {
  const a = r.querySelector('a');
  return {
    nome: a ? (a.innerText || '').trim() : '',
    href: a ? (a.getAttribute('href') || '') : '',
    cells: Array.from(r.querySelectorAll('td.ant-table-cell')).map(c => (c.innerText || '').trim())
  };
});"""
_RE_PASTA = re.compile(r"/projetos/directories/folder/([0-9a-f]{24})$")
_RE_ARQUIVO = re.compile(r"/projetos/directories/folder/([0-9a-f]{24})/file/([0-9a-f]{24})")
PRANCHAS_DIR = os.path.join(SAIDA_DIR, "pranchas")


def _carregar_credenciais_do_bat():
    """O robô roda fora do dashboard (agendador do Windows), então não herda
    as variáveis do iniciar_dashboard.bat — lê as linhas AUTODOC_* de lá pra
    não precisar guardar a senha em mais um lugar."""
    if os.environ.get("AUTODOC_EMAIL") and os.environ.get("AUTODOC_SENHA"):
        return
    bat = os.path.join(BASE_DIR, "iniciar_dashboard.bat")
    try:
        with open(bat, "r", encoding="utf-8", errors="ignore") as fh:
            for line in fh:
                m = re.match(r'\s*set\s+"(AUTODOC_\w+)=(.*)"\s*$', line, re.I)
                if m:
                    os.environ.setdefault(m.group(1), m.group(2))
    except Exception:
        pass


def _autodoc():
    _carregar_credenciais_do_bat()
    if BASE_DIR not in sys.path:
        sys.path.insert(0, BASE_DIR)
    import autodoc_sync as ad
    return ad


def _listar_pasta(driver, ad, folder_id):
    driver.get(f"{ad.BASE_URL}/projetos/directories/folder/{folder_id}")
    ad._esperar_linhas(driver)
    time.sleep(0.8)
    linhas = driver.execute_script(_JS_LINHAS) or []
    if not linhas:
        time.sleep(2.5)
        linhas = driver.execute_script(_JS_LINHAS) or []
    return linhas


def _ler_disciplinas(driver, ad):
    out = []
    for ln in _listar_pasta(driver, ad, ad.PROJETOS_FOLDER_ID):
        m = _RE_PASTA.search(ln.get("href") or "")
        if m and ln.get("nome"):
            out.append({"nome": ln["nome"], "id": m.group(1)})
    return out


def _varrer(driver, ad, folder_id, caminho, extensoes, profundidade=0, max_prof=4, progresso=None):
    """Lê a pasta e desce nas subpastas. caminho = lista de nomes de pasta
    abaixo da disciplina. Retorna só arquivos com as extensões pedidas."""
    if progresso:
        progresso(" / ".join(caminho))
    arquivos, subpastas = [], []
    for ln in _listar_pasta(driver, ad, folder_id):
        nome = ln.get("nome") or ""
        if not nome:
            continue
        href = ln.get("href") or ""
        m = _RE_PASTA.search(href)
        if m:
            subpastas.append((m.group(1), nome))
            continue
        if extensoes and not any(nome.lower().endswith(e.lower()) for e in extensoes):
            continue
        c = ln.get("cells") or []
        i = c.index(nome) if nome in c else 1
        c = c + [""] * 8
        mf = _RE_ARQUIVO.search(href)
        arquivos.append({"arquivo": nome, "rev": c[i + 1], "titulo": c[i + 2], "tamanho": c[i + 3],
                         "data": c[i + 4], "status": c[i + 5], "caminho": list(caminho),
                         "folder_id": mf.group(1) if mf else folder_id, "file_id": mf.group(2) if mf else ""})
    if profundidade < max_prof:
        for sub_id, sub_nome in subpastas:
            arquivos += _varrer(driver, ad, sub_id, caminho + [sub_nome], extensoes, profundidade + 1, max_prof, progresso)
    return arquivos


# ── leitura rápida: API interna que a própria tela do AutoDoc usa ──
# Depois do login no navegador, o cookie idToken autentica as chamadas a
# projetos-ng-bff (JSON, ~1 s por pasta, várias pastas em paralelo) — bem
# mais rápido que abrir cada pasta como página. Se a API mudar/falhar, o
# robô volta sozinho para a leitura pelo navegador (_varrer).
API_BFF = "https://projetos-ng-bff.autodoc.com.br/v1"
AUTODOC_CUSTOMER_ID = "6604808c5b2241654b0f1a5f"  # conta Temazec no AutoDoc
_STATUS_PADRAO = {"11600": "Bloqueado", "11601": "Aprovado", "11602": "Liberado para Obra", "11603": "Bloqueado Parcialmente"}


class _ApiAutodoc:
    def __init__(self, driver):
        import requests
        ck = {c["name"]: c["value"] for c in driver.get_cookies()}
        if not ck.get("idToken"):
            raise RuntimeError("sessão do AutoDoc sem idToken")
        self.s = requests.Session()
        self.s.headers.update({"Authorization": "Bearer " + ck["idToken"], "X-Customer-Id": AUTODOC_CUSTOMER_ID,
                               "X-Product-value": "projetos", "Accept": "application/json, text/plain, */*",
                               "Referer": "https://plataforma.autodoc.com.br/"})
        self.status = dict(_STATUS_PADRAO)
        try:
            for s in self.s.get(f"{API_BFF}/status", timeout=30).json().get("data") or []:
                self.status[str(s["id"])] = s["name"]
        except Exception:
            pass

    def listar(self, folder_id):
        r = self.s.get(f"{API_BFF}/folders/{folder_id}/subitems/light", params={"filter[obsolete]": "false"}, timeout=60)
        r.raise_for_status()
        j = r.json()
        return ([(f["id"], f["name"]) for f in j.get("subFolders") or []], j.get("documents") or [])


def _data_api(s):
    try:
        return dt.datetime.strptime(str(s)[:10], "%Y-%m-%d").strftime("%d/%m/%Y")
    except ValueError:
        return ""


def _varrer_api(api, folder_id, extensoes, max_prof=4, progresso=None, trabalhadores=8):
    """Mesmo resultado de _varrer, lendo as pastas em paralelo pela API."""
    from concurrent.futures import ThreadPoolExecutor
    arquivos, nivel = [], [(folder_id, [])]
    with ThreadPoolExecutor(max_workers=trabalhadores) as ex:
        for prof in range(max_prof + 1):
            if not nivel:
                break
            if progresso:
                progresso(f"{len(nivel)} pasta(s) no nível {prof + 1}")
            resultados = list(ex.map(lambda item: (item, api.listar(item[0])), nivel))
            proximo = []
            for (fid, caminho), (subs, docs) in resultados:
                for d in docs:
                    nome = d.get("name") or ""
                    if extensoes and not any(nome.lower().endswith(e.lower()) for e in extensoes):
                        continue
                    arquivos.append({"arquivo": nome, "rev": d.get("revision") or "",
                                     "titulo": (d.get("description") or "").strip(), "tamanho": d.get("size") or "",
                                     "data": _data_api(d.get("status_updated_at") or d.get("updated_at")),
                                     "status": api.status.get(str(d.get("status")), str(d.get("status") or "")),
                                     "caminho": list(caminho), "folder_id": d.get("folder_id") or fid,
                                     "file_id": d.get("id") or "", "codigo": (d.get("code") or "").strip()})
                proximo += [(sid, caminho + [snome]) for sid, snome in subs]
            nivel = proximo
    return arquivos


def _chave_prancha(arquivo):
    """SYB-ALV-PE-10100-SS2-R02.pdf -> (base 'SYB-ALV-PE-10100-SS2', rev 2)."""
    stem = os.path.splitext(arquivo)[0].strip()
    rev = None
    # alguns arquivos vêm com a revisão repetida: ...-R00_R00
    while True:
        m = re.search(r"[-_ ]R(\d+)$", stem, re.I)
        if not m:
            break
        rev = int(m.group(1)) if rev is None else rev
        stem = stem[:m.start()]
    return stem, rev


def consolidar_pranchas(arquivos, prefixo=""):
    """Agrupa as revisões do mesmo documento e escolhe a última APROVADA.
    Revisão mais nova com outro status (ex.: Bloqueado) vai em 'pendente'."""
    grupos = {}
    for a in arquivos:
        base, rev = _chave_prancha(a["arquivo"])
        if a.get("codigo"):  # código oficial do AutoDoc, quando a API informa
            base = _chave_prancha(a["codigo"])[0]
        if rev is None:
            try:
                rev = int(a.get("rev") or 0)
            except ValueError:
                rev = 0
        grupos.setdefault(base, []).append({**a, "rev_n": rev})
    out = []
    for base, revs in grupos.items():
        revs.sort(key=lambda r: r["rev_n"])
        aprovadas = [r for r in revs if (r.get("status") or "").strip().upper() in ("APROVADO", "LIBERADO PARA OBRA")]
        atual = aprovadas[-1] if aprovadas else revs[-1]
        mais_nova = revs[-1]
        partes = base.split("-")
        local = partes[-1] if len(partes) >= 5 else ""
        titulo = (atual.get("titulo") or "").strip()
        if titulo in ("", "--") or (len(titulo) > 60 and local in LOCAIS):
            descricao = prefixo + (_local_desc(local) if local else base)
        else:
            titulo = re.sub(r"\s+", " ", titulo)
            descricao = prefixo + (titulo if len(titulo) <= 90 else titulo[:87].rstrip() + "…")
        out.append({
            "documento": base, "local": local, "numero": partes[3] if len(partes) >= 5 else base,
            "descricao": descricao, "rev": f"R{atual['rev_n']:02d}", "data": atual.get("data", ""),
            "status": atual.get("status", ""), "arquivo": atual["arquivo"],
            "folder_id": atual.get("folder_id", ""), "file_id": atual.get("file_id", ""),
            "tamanho": atual.get("tamanho", ""),
            "aprovada": bool(aprovadas),
            "pendente": (f"R{mais_nova['rev_n']:02d}" if mais_nova is not atual else ""),
            "pendente_status": (mais_nova.get("status", "") if mais_nova is not atual else ""),
            "pendente_data": (mais_nova.get("data", "") if mais_nova is not atual else ""),
        })
    out.sort(key=lambda d: (d["numero"], d["documento"]))
    return out


def _rotulo_grupo(disciplina, subpasta):
    sub = re.sub(r"^\s*\d+\s*[.\-_)]\s*", "", subpasta or "").strip()
    return f"{disciplina} – {sub}" if sub else disciplina


def disciplinas_selecionadas(pacote, cfg=None, st=None):
    st = st if st is not None else estado_pacote(pacote)
    if st.get("disciplinas_selecionadas") is not None:
        return st["disciplinas_selecionadas"]
    cfg = cfg or carregar_config(pacote)
    return list(cfg["documentos"].get("disciplinas_padrao") or [])


def chaves_selecionadas(pacote, cfg=None, st=None):
    """Chaves (grupo|documento) das pranchas marcadas. Sem escolha salva,
    vale o padrão do JSON (subpastas_padrao)."""
    st = st if st is not None else estado_pacote(pacote)
    if st.get("pranchas_selecionadas") is not None:
        return set(st["pranchas_selecionadas"])
    cfg = cfg or carregar_config(pacote)
    padrao = set(cfg["documentos"].get("grupos_padrao") or [])
    return {f"{g['grupo']}|{d['documento']}" for g in (st.get("grupos") or []) if g["grupo"] in padrao
            for d in g["pranchas"]}


def grupos_selecionados(pacote, cfg=None, st=None):
    """Só os grupos/pranchas marcados, na ordem do AutoDoc."""
    st = st if st is not None else estado_pacote(pacote)
    sel = chaves_selecionadas(pacote, cfg, st)
    disc = set(disciplinas_selecionadas(pacote, cfg, st))
    out = []
    for g in st.get("grupos") or []:
        if g.get("disciplina") and g["disciplina"] not in disc:
            continue
        pr = [d for d in g["pranchas"] if f"{g['grupo']}|{d['documento']}" in sel]
        if pr:
            out.append({**g, "pranchas": pr})
    # grupos do escopo principal (ordem do JSON) primeiro, o resto depois
    ordem = list(reversed((cfg or carregar_config(pacote))["documentos"].get("grupos_padrao") or []))
    out.sort(key=lambda g: -ordem.index(g["grupo"]) if g["grupo"] in ordem else 1)
    return out


def salvar_selecao(pacote, disciplinas=None, pranchas=None):
    campos = {}
    if disciplinas is not None:
        cfg = carregar_config(pacote)
        st = estado_pacote(pacote)
        novas = [str(d) for d in disciplinas]
        frentes = _sincronizar_frentes_disciplinas(pacote, disciplinas_selecionadas(pacote, cfg, st), novas, cfg, st)
        if frentes is not None:
            campos["frentes"] = frentes
        campos["disciplinas_selecionadas"] = novas
    if pranchas is not None:
        campos["pranchas_selecionadas"] = [str(p) for p in pranchas]
    return _atualizar_estado(pacote, **campos)


def atualizar_disciplinas(pacote):
    """Só lista os tipos de projeto (pastas da raiz PROJETOS) — rápido."""
    ad = _autodoc()
    cred = ad.credenciais_status()
    if not cred.get("ok"):
        return {"ok": False, "error": cred.get("error")}
    driver = ad._novo_driver()
    try:
        ad._login(driver, cred["email"], os.environ.get("AUTODOC_SENHA", ""))
        try:
            disc = [{"nome": n, "id": i} for i, n in _ApiAutodoc(driver).listar(ad.PROJETOS_FOLDER_ID)[0]]
        except Exception:
            disc = _ler_disciplinas(driver, ad)
        _atualizar_estado(pacote, disciplinas_autodoc=disc, disciplinas_at=time.strftime("%Y-%m-%d %H:%M:%S"))
        return {"ok": True, "disciplinas": disc}
    except Exception as exc:
        return {"ok": False, "error": f"Falha ao ler o AutoDoc: {exc}"}
    finally:
        try:
            driver.quit()
        except Exception:
            pass


def sincronizar_pranchas(pacote):
    """Loga no AutoDoc, atualiza a lista de disciplinas e varre as
    disciplinas marcadas (todas as subpastas). ~1-2 min por disciplina."""
    ad = _autodoc()
    cred = ad.credenciais_status()
    if not cred.get("ok"):
        return {"ok": False, "error": cred.get("error")}
    cfg = carregar_config(pacote)
    dcfg = cfg["documentos"]
    extensoes = dcfg.get("extensoes") or [".pdf"]
    prefixos = dcfg.get("prefixos") or {}
    st0 = estado_pacote(pacote)
    alvo = disciplinas_selecionadas(pacote, cfg, st0)
    anterior = {f"{g['grupo']}|{d['documento']}": d["rev"] for g in (st0.get("grupos") or []) for d in g["pranchas"]}

    _atualizar_estado(pacote, sync_status="sincronizando", sync_erro=None, sync_progresso="Entrando no AutoDoc...")
    driver = ad._novo_driver()
    try:
        ad._login(driver, cred["email"], os.environ.get("AUTODOC_SENHA", ""))
        api = None
        try:
            api = _ApiAutodoc(driver)
            subs, _ = api.listar(ad.PROJETOS_FOLDER_ID)
            disciplinas = [{"nome": n, "id": i} for i, n in subs]
            _log("AutoDoc: leitura rápida (API) ativa.")
        except Exception as exc:
            api = None
            _log(f"AutoDoc: API indisponível ({exc}); lendo pelo navegador (mais lento).")
            disciplinas = _ler_disciplinas(driver, ad)
        ids = {d["nome"]: d["id"] for d in disciplinas}
        grupos = []
        for nome in alvo:
            fid = ids.get(nome)
            if not fid:
                _log(f"AutoDoc: disciplina '{nome}' não existe mais na raiz PROJETOS — ignorada.")
                continue
            _log(f"AutoDoc: lendo {nome} ...")

            def _prog(txt, nome=nome):
                _atualizar_estado(pacote, sync_progresso=f"Lendo {nome}" + (f" / {txt}" if txt else ""))
            arqs = None
            if api:
                try:
                    arqs = _varrer_api(api, fid, extensoes, progresso=_prog)
                except Exception as exc:
                    _log(f"  API falhou em {nome} ({exc}); lendo pelo navegador.")
            if arqs is None:
                arqs = _varrer(driver, ad, fid, [], extensoes, progresso=_prog)
            por_sub = {}
            for a in arqs:
                por_sub.setdefault(a["caminho"][0] if a["caminho"] else "", []).append(a)
            for sub in sorted(por_sub, key=lambda s: [int(x) if x.isdigit() else x for x in re.split(r"(\d+)", s)]):
                rot = _rotulo_grupo(nome, sub)
                prefixo = prefixos.get(rot, nome.capitalize() + " – ")
                pr = consolidar_pranchas(por_sub[sub], prefixo)
                grupos.append({"grupo": rot, "disciplina": nome, "subpasta": sub,
                               "caminho": f"PROJETOS / {nome}" + (f" / {sub}" if sub else ""), "pranchas": pr})
            _log(f"  {len(arqs)} arquivos em {len(por_sub)} subpasta(s)")
        mudancas = []
        for g in grupos:
            for d in g["pranchas"]:
                old = anterior.get(f"{g['grupo']}|{d['documento']}")
                if anterior and old and old != d["rev"]:
                    mudancas.append(f"{d['documento']}: {old} -> {d['rev']}")
        agora = time.strftime("%Y-%m-%d %H:%M:%S")
        hist = (estado_pacote(pacote).get("historico_mudancas") or [])
        if mudancas:
            hist = ([{"data": agora, "mudancas": mudancas}] + hist)[:50]
        _atualizar_estado(pacote, grupos=grupos, disciplinas_autodoc=disciplinas, disciplinas_at=agora,
                          sync_at=agora, sync_status="ok", sync_erro=None, sync_progresso="",
                          ultimas_mudancas=mudancas, historico_mudancas=hist)
        return {"ok": True, "grupos": grupos, "mudancas": mudancas, "sync_at": agora,
                "primeira_leitura": not anterior}
    except Exception as exc:
        _atualizar_estado(pacote, sync_status="erro", sync_erro=str(exc), sync_progresso="")
        return {"ok": False, "error": f"Falha ao ler o AutoDoc: {exc}"}
    finally:
        try:
            driver.quit()
        except Exception:
            pass


def _clicar_download(driver, arquivo):
    """Menu "..." da linha > Download > Download (cópia controlada, com QR)."""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.action_chains import ActionChains

    linha = None
    for row in driver.find_elements(By.CSS_SELECTOR, "tr.ant-table-row"):
        links = row.find_elements(By.CSS_SELECTOR, "a")
        if links and links[0].text.strip() == arquivo:
            linha = row
            break
    if linha is None:
        raise RuntimeError(f"{arquivo} não encontrado na pasta do AutoDoc")
    reticencias = linha.find_elements(By.CSS_SELECTOR, "td.ant-table-cell")[-1].find_element(
        By.CSS_SELECTOR, ".table-action-ellipsis-v")
    ActionChains(driver).move_to_element(reticencias).click().perform()
    time.sleep(1.0)
    menu = [e for e in driver.find_elements(By.XPATH, "//div[normalize-space(text())='Download']") if e.is_displayed()]
    if not menu:
        raise RuntimeError("menu de download do AutoDoc não apareceu")
    ActionChains(driver).move_to_element(menu[0]).perform()
    time.sleep(1.2)
    sub = [e for e in driver.find_elements(By.XPATH, "//div[normalize-space(text())='Download']") if e.is_displayed()]
    sub[-1].click()


def baixar_pranchas(pranchas, progresso=None):
    """Baixa (com cache local — o nome do arquivo já traz a revisão) os PDFs
    das pranchas. Retorna {arquivo: caminho_local} e lista de erros."""
    import zipfile
    os.makedirs(PRANCHAS_DIR, exist_ok=True)
    locais, faltando, erros = {}, [], []
    for d in pranchas:
        destino = os.path.join(PRANCHAS_DIR, d["arquivo"])
        if os.path.exists(destino) and os.path.getsize(destino) > 0:
            locais[d["arquivo"]] = destino
        else:
            faltando.append(d)
    if not faltando:
        return locais, erros

    ad = _autodoc()
    cred = ad.credenciais_status()
    if not cred.get("ok"):
        return locais, [cred.get("error")]
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    tmp = tempfile.mkdtemp(prefix="carta_dl_")
    opts = Options()
    for a in ("--headless=new", "--window-size=1440,1000", "--disable-gpu", "--no-sandbox", "--log-level=3"):
        opts.add_argument(a)
    opts.add_experimental_option("prefs", {"download.default_directory": tmp, "download.prompt_for_download": False,
                                           "plugins.always_open_pdf_externally": True})
    import chrome_env
    chrome_env.aplicar_opcoes(opts)
    driver = webdriver.Chrome(service=chrome_env.servico(), options=opts)
    try:
        driver.execute_cdp_cmd("Page.setDownloadBehavior", {"behavior": "allow", "downloadPath": tmp})
        ad._login(driver, cred["email"], os.environ.get("AUTODOC_SENHA", ""))
        for n, d in enumerate(faltando, 1):
            if progresso:
                progresso(f"Baixando prancha {n}/{len(faltando)}: {d['arquivo']}")
            try:
                for f in os.listdir(tmp):
                    os.remove(os.path.join(tmp, f))
                driver.get(f"{ad.BASE_URL}/projetos/directories/folder/{d['folder_id']}")
                ad._esperar_linhas(driver)
                time.sleep(0.8)
                _clicar_download(driver, d["arquivo"])
                baixado = None
                for _ in range(60):
                    time.sleep(1)
                    prontos = [f for f in os.listdir(tmp) if not f.endswith((".crdownload", ".tmp"))]
                    if prontos:
                        baixado = os.path.join(tmp, prontos[0])
                        break
                if not baixado:
                    raise RuntimeError("download não terminou em 60 s")
                destino = os.path.join(PRANCHAS_DIR, d["arquivo"])
                if zipfile.is_zipfile(baixado):
                    with zipfile.ZipFile(baixado) as zf:
                        membro = next((m for m in zf.namelist() if m.lower().endswith(os.path.splitext(d["arquivo"])[1].lower())), None)
                        if not membro:
                            raise RuntimeError("zip do AutoDoc veio sem o arquivo")
                        with zf.open(membro) as src, open(destino, "wb") as dst:
                            shutil.copyfileobj(src, dst)
                else:
                    shutil.move(baixado, destino)
                locais[d["arquivo"]] = destino
            except Exception as exc:
                erros.append(f"{d['arquivo']}: {exc}")
                _log(f"Falha ao baixar {d['arquivo']}: {exc}")
    finally:
        try:
            driver.quit()
        except Exception:
            pass
        shutil.rmtree(tmp, ignore_errors=True)
    return locais, erros


# ─── escopo por frente de serviço (editável na aba) ───────────────────────
# O escopo é preenchido na aba (começa em branco); o JSON traz só um MODELO
# ("frentes") que pode ser carregado com um clique. No estado:
#   frentes [{id,nome,nota,incluida,servicos:[{servico,un,qtd}]}] ·
#   matriz_resp {"g.i": "A"|"B"} · matriz_frente {"g.i": id|"geral"} ·
#   anexo_valores {"g.i": valor}
def frentes_modelo(pacote, cfg=None):
    cfg = cfg or carregar_config(pacote)
    return [{"id": f["id"], "nome": f["nome"], "nota": f.get("nota", ""), "incluida": True,
             "servicos": [{"servico": n, "un": u, "qtd": q} for n, u, q in f["servicos"]]}
            for f in cfg.get("frentes") or []]


def frentes_da_disciplina(cfg, disciplina):
    """Frentes sugeridas para um tipo de projeto (escopo_por_disciplina no
    JSON). 'origem' marca de onde vieram, pra sair junto se desmarcar."""
    mapa = cfg.get("escopo_por_disciplina") or {}
    item = mapa.get(disciplina)
    if item == "frentes":
        item = cfg.get("frentes") or []
    out = []
    for f in item or []:
        out.append({"id": f["id"], "nome": f["nome"], "nota": f.get("nota", ""), "incluida": True, "origem": disciplina,
                    "servicos": [{"servico": n, "un": u, "qtd": q} for n, u, q in f["servicos"]]})
    return out


def frentes_brutas(pacote, cfg=None, st=None):
    """Frentes como estão no editor. Sem nada salvo, vêm das disciplinas
    marcadas no card 1."""
    cfg = cfg or carregar_config(pacote)
    st = st if st is not None else estado_pacote(pacote)
    if st.get("frentes") is not None:
        return st["frentes"]
    out, ids = [], set()
    for d in disciplinas_selecionadas(pacote, cfg, st):
        for f in frentes_da_disciplina(cfg, d):
            if f["id"] not in ids:
                ids.add(f["id"])
                out.append(f)
    return out


def _sincronizar_frentes_disciplinas(pacote, antes, depois, cfg, st):
    """Marcou disciplina: entram as frentes dela que ainda não estão lá.
    Desmarcou: saem as que vieram dela. O resto (editado à mão) fica."""
    if st.get("frentes") is None:
        return None  # ainda automático — frentes_brutas já resolve
    frentes = [dict(f) for f in st["frentes"]]
    # frente salva sem origem (versão anterior da aba): descobre pelo id
    id_origem = {fr["id"]: d for d in set(antes) | set(depois) for fr in frentes_da_disciplina(cfg, d)}
    for f in frentes:
        if not f.get("origem") and f.get("id") in id_origem:
            f["origem"] = id_origem[f["id"]]
    removidas = set(antes) - set(depois)
    frentes = [f for f in frentes if f.get("origem") not in removidas]
    ids = {f["id"] for f in frentes}
    for d in depois:
        if d in antes:
            continue
        for f in frentes_da_disciplina(cfg, d):
            if f["id"] not in ids:
                ids.add(f["id"])
                frentes.append(f)
    return frentes


def frentes_efetivas(pacote, cfg=None, st=None):
    cfg = cfg or carregar_config(pacote)
    st = st if st is not None else estado_pacote(pacote)
    out = []
    for f in frentes_brutas(pacote, cfg, st):
        servs = [{"chave": f"{f['id']}|{i}", "servico": s.get("servico", ""), "un": s.get("un", ""), "qtd": s.get("qtd")}
                 for i, s in enumerate(f.get("servicos") or []) if (s.get("servico") or "").strip()]
        out.append({"id": f["id"], "nome": f.get("nome", ""), "nota": f.get("nota", ""),
                    "incluida": f.get("incluida", True) is not False and bool((f.get("nome") or "").strip()),
                    "servicos": servs, "origem": f.get("origem") or _origem_por_id(cfg).get(f["id"], "")})
    return out


def _origem_por_id(cfg):
    return {fr["id"]: d for d in (cfg.get("escopo_por_disciplina") or {}) if not d.startswith("_")
            for fr in frentes_da_disciplina(cfg, d)}


def contexto_escopo(pacote, cfg=None, st=None):
    """Tudo que muda conforme o tipo de projeto / frentes marcadas e é usado
    nos textos e quadros: título, lista de frentes, equipamentos e as
    disciplinas que de fato têm frente no escopo."""
    cfg = cfg or carregar_config(pacote)
    st = st if st is not None else estado_pacote(pacote)
    pd = cfg.get("por_disciplina") or {}
    frentes = [f for f in frentes_efetivas(pacote, cfg, st) if f["incluida"]]
    disciplinas, manuais = [], []
    for f in frentes:
        o = f.get("origem")
        if o and o in pd:
            if o not in disciplinas:
                disciplinas.append(o)
        else:
            manuais.append(f["nome"])
    titulos = [pd[d].get("titulo") or d.title() for d in disciplinas] + manuais
    equip = [pd[d]["equipamentos"] for d in disciplinas if pd[d].get("equipamentos")]
    titulo = " / ".join(titulos) or "Carta Convite"
    return {
        "titulo": titulo, "titulo_maiusculo": titulo.upper(),
        "frentes_texto": _lista_pt([f["nome"].lower() for f in frentes]),
        "equipamentos": "; ".join(equip) or "as ferramentas e equipamentos necessários à execução dos serviços",
        "disciplinas_ativas": disciplinas,
    }


def _slug_frente(nome, usados):
    import unicodedata
    base = unicodedata.normalize("NFKD", nome or "").encode("ascii", "ignore").decode().lower()
    base = re.sub(r"[^a-z0-9]+", "_", base).strip("_")[:30] or "frente"
    s, n = base, 2
    while s in usados or s == "geral":
        s, n = f"{base}_{n}", n + 1
    return s


def _num(v):
    if v in (None, ""):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).strip().replace(".", "").replace(",", "."))
    except ValueError:
        return None


def matriz_efetiva(pacote, cfg=None, st=None, so_escopo=True):
    """Itens da matriz com a frente de cada um; 'geral' vale sempre, os de
    frente só entram se a frente estiver no escopo."""
    cfg = cfg or carregar_config(pacote)
    st = st if st is not None else estado_pacote(pacote)
    frentes = frentes_efetivas(pacote, cfg, st)
    incl = {f["id"] for f in frentes if f["incluida"]}
    presentes = {f["id"] for f in frentes}  # no escopo, marcadas ou não
    resp = st.get("matriz_resp") or {}
    frente_ov = st.get("matriz_frente") or {}
    visivel = (lambda fr: fr == "geral" or fr in incl) if so_escopo else (lambda fr: fr == "geral" or fr in presentes)
    # itens próprios de cada frente sugerida pelo tipo de projeto (matriz_por_frente)
    por_frente = cfg.get("matriz_por_frente") or {}
    extras = []
    for f in frentes:
        for i, (txt, r) in enumerate(por_frente.get(f["id"]) or [], 1):
            k = f"f.{f['id']}.{i}"
            extras.append((k, txt, r, frente_ov.get(k, f["id"])))
    out = []
    for gi, (grupo, linhas) in enumerate(cfg["matriz"]["grupos"], 1):
        itens = []
        base = [(f"{gi}.{li}", txt, r, frente_ov.get(f"{gi}.{li}", frente)) for li, (txt, r, frente) in enumerate(linhas, 1)]
        for k, txt, r, frente in base + (extras if gi == 1 else []):
            if not visivel(frente):
                continue
            itens.append({"chave": k, "texto": txt, "resp": resp.get(k, r), "frente": frente})
        if itens:
            out.append({"grupo": grupo, "itens": itens})
    return out


def anexo_efetivo(pacote, cfg=None, st=None):
    cfg = cfg or carregar_config(pacote)
    st = st if st is not None else estado_pacote(pacote)
    """Quadro resumo: padrão do JSON → valores sugeridos pelas disciplinas
    ativas (por_disciplina.valores) → o que foi marcado na aba. Itens
    extras de cada disciplina entram no último grupo."""
    vals = st.get("anexo_valores") or {}
    ctx = contexto_escopo(pacote, cfg, st)
    pd = cfg.get("por_disciplina") or {}
    sug = {}
    for d in ctx["disciplinas_ativas"]:
        for rot, v in ((pd[d].get("valores") or {}).items()):
            if isinstance(v, list):  # múltipla escolha: une o que cada disciplina pede
                sug[rot] = list(dict.fromkeys((sug.get(rot) or []) + v))
            else:
                sug.setdefault(rot, v)
    grupos = cfg["anexo"]["grupos"]
    out = []
    for gi, g in enumerate(grupos, 1):
        brutos = [(f"{gi}.{li}", it) for li, it in enumerate(g["itens"], 1)]
        if gi == len(grupos):
            vistos = {it["rotulo"] for _, it in brutos}
            for d in ctx["disciplinas_ativas"]:
                for i, it in enumerate(pd[d].get("itens_extras") or [], 1):
                    if it["rotulo"] not in vistos:
                        vistos.add(it["rotulo"])
                        brutos.append((f"d.{_slug_frente(d, set())}.{i}", it))
        itens = []
        for k, it in brutos:
            v = vals[k] if k in vals else sug.get(it["rotulo"], it.get("valor"))
            if isinstance(v, str):
                v = _fmt(v, {"equipamentos": ctx["equipamentos"]})
            itens.append({**it, "chave": k, "valor": v})
        out.append({"grupo": g["grupo"], "itens": itens})
    return out


def salvar_escopo(pacote, frentes=None, matriz=None, matriz_frente=None, anexo=None):
    campos = {}
    if frentes is not None:
        limpas, usados = [], set()
        for f in frentes:
            nome = str(f.get("nome") or "").strip()
            fid = str(f.get("id") or "").strip()
            if not re.fullmatch(r"[a-z0-9_]+", fid) or fid in usados or fid == "geral":
                fid = _slug_frente(nome, usados)
            usados.add(fid)
            servs = [{"servico": str(s.get("servico") or "").strip(), "un": str(s.get("un") or "").strip(),
                      "qtd": _num(s.get("qtd"))} for s in (f.get("servicos") or [])]
            limpas.append({"id": fid, "nome": nome, "nota": str(f.get("nota") or "").strip(),
                           "incluida": f.get("incluida", True) is not False, "servicos": servs,
                           "origem": str(f.get("origem") or "")})
        campos["frentes"] = limpas
    if matriz is not None:
        campos["matriz_resp"] = {str(k): v for k, v in matriz.items() if v in ("A", "B")}
    if matriz_frente is not None:
        campos["matriz_frente"] = {str(k): str(v) for k, v in matriz_frente.items() if v}
    if anexo is not None:
        campos["anexo_valores"] = anexo
    return _atualizar_estado(pacote, **campos)


def _lista_pt(nomes):
    nomes = [n for n in nomes if n]
    return nomes[0] if len(nomes) == 1 else (", ".join(nomes[:-1]) + " e " + nomes[-1] if nomes else "")


def gerar_planilha_proposta(pacote, opcoes=None):
    """Planilha Excel que vai anexa à carta: a proponente preenche os
    preços unitários (células amarelas) e devolve junto com a proposta."""
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.worksheet.datavalidation import DataValidation  # noqa: F401 (mantém openpyxl completo)

    cfg = carregar_config(pacote)
    st = estado_pacote(pacote)
    op = opcoes or {}
    frentes = [f for f in frentes_efetivas(pacote, cfg, st) if f["incluida"]]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Proposta"
    fino = Side(style="thin", color="808080")
    borda = Border(left=fino, right=fino, top=fino, bottom=fino)
    cinza = PatternFill("solid", fgColor="D9D9D9")
    claro = PatternFill("solid", fgColor="F2F2F2")
    amarelo = PatternFill("solid", fgColor="FFF2CC")
    for col, w in zip("ABCDEFG", (7, 26, 52, 7, 13, 16, 18)):
        ws.column_dimensions[col].width = w

    # logo da Temazec (A1:B2) e título ao lado
    logo = os.path.join(BASE_DIR, cfg.get("logo", "static/temazec_logo.png"))
    if os.path.exists(logo):
        try:
            from openpyxl.drawing.image import Image as XLImage
            img = XLImage(logo)
            escala = 52 / img.height
            img.width, img.height = int(img.width * escala), 52
            ws.add_image(img, "A1")
        except Exception as exc:  # sem Pillow a planilha sai sem logo, mas sai
            _log(f"Logo não inserida na planilha: {exc}")
    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 18
    ws["C1"] = "PLANILHA DE PROPOSTA – " + contexto_escopo(pacote, cfg, st)["titulo_maiusculo"]
    ws["C1"].font = Font(bold=True, size=13)
    ws["C1"].alignment = Alignment(vertical="bottom")
    ws["C2"] = f"Obra {cfg.get('obra', '')} – " + cfg["rodape"]
    ws["C2"].font = Font(size=9, color="595959")
    ws["C2"].alignment = Alignment(vertical="top")
    linha = 4
    for rot, val in (("Proponente (razão social)", op.get("empresa", "")), ("CNPJ", ""), ("Responsável / contato", op.get("ac", "")),
                     ("Telefone / e-mail", ""), ("Data da proposta", "")):
        ws.cell(linha, 1, rot).font = Font(bold=True)
        ws.merge_cells(start_row=linha, start_column=1, end_row=linha, end_column=2)
        c = ws.cell(linha, 3, val)
        ws.merge_cells(start_row=linha, start_column=3, end_row=linha, end_column=7)
        for col in range(1, 8):
            ws.cell(linha, col).border = borda
        c.fill = amarelo
        linha += 1

    linha += 1
    cab = ["Item", "Frente de serviço", "Serviço", "Un.", "Quantidade", "Preço unit. (R$)", "Total (R$)"]
    for i, t in enumerate(cab, 1):
        c = ws.cell(linha, i, t)
        c.font = Font(bold=True)
        c.fill = cinza
        c.border = borda
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = ws.cell(linha + 1, 1)
    linha += 1
    primeira = linha
    n = 0
    for f in frentes:
        for s in f["servicos"]:
            n += 1
            vals = [n, f["nome"], s["servico"], s["un"], s["qtd"], None, f"=IF(OR(E{linha}=\"\",F{linha}=\"\"),\"\",E{linha}*F{linha})"]
            for i, v in enumerate(vals, 1):
                c = ws.cell(linha, i, v)
                c.border = borda
                c.alignment = Alignment(vertical="center", wrap_text=i in (2, 3), horizontal="center" if i in (1, 4) else None)
            ws.cell(linha, 5).number_format = '#,##0.00'
            if s["qtd"] is None:
                ws.cell(linha, 5).fill = amarelo
            ws.cell(linha, 6).number_format = '"R$" #,##0.00'
            ws.cell(linha, 6).fill = amarelo
            ws.cell(linha, 7).number_format = '"R$" #,##0.00'
            linha += 1
    ultima = linha - 1
    ws.cell(linha, 1, "VALOR TOTAL DA PROPOSTA").font = Font(bold=True)
    ws.merge_cells(start_row=linha, start_column=1, end_row=linha, end_column=6)
    ws.cell(linha, 1).alignment = Alignment(horizontal="right")
    c = ws.cell(linha, 7, f"=SUM(G{primeira}:G{ultima})")
    c.number_format = '"R$" #,##0.00'
    c.font = Font(bold=True)
    for col in range(1, 8):
        ws.cell(linha, col).border = borda
        ws.cell(linha, col).fill = claro
    linha += 2
    ws.cell(linha, 1, "Células em amarelo: preencher. Quantidades em branco: informar a quantidade considerada na proposta.").font = Font(italic=True, size=9, color="595959")
    linha += 2
    for rot in (cfg.get("proposta") or {}).get("condicoes") or []:
        ws.cell(linha, 1, rot).font = Font(bold=True)
        ws.merge_cells(start_row=linha, start_column=1, end_row=linha, end_column=3)
        ws.merge_cells(start_row=linha, start_column=4, end_row=linha, end_column=7)
        for col in range(1, 8):
            ws.cell(linha, col).border = borda
        ws.cell(linha, 4).fill = amarelo
        linha += 1
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def _vars_nome(pacote, cfg):
    ctx = contexto_escopo(pacote, cfg)
    # título vira nome de arquivo: sem caracteres proibidos no Windows
    return {"obra": cfg.get("obra", ""), "titulo": re.sub(r'[\\/:*?"<>|]+', "-", ctx["titulo"]).strip()}


def nome_planilha(pacote):
    cfg = carregar_config(pacote)
    modelo = (cfg.get("proposta") or {}).get("arquivo") or "Planilha de proposta - {titulo}.xlsx"
    return _fmt(modelo, _vars_nome(pacote, cfg))


# ─── prazos (Cronograma Base do PCP) ──────────────────────────────────────
def _achar_planilha(padroes):
    """Padrões em ordem de prioridade (a pasta dados vem primeiro); dentro de
    um mesmo padrão, o arquivo mais recente."""
    for p in padroes or []:
        p = p if os.path.isabs(p) else os.path.join(BASE_DIR, p)
        achados = [f for f in glob.glob(p) if not os.path.basename(f).startswith("~$")]
        if achados:
            return max(achados, key=os.path.getmtime)
    return None


def ler_prazos(pacote, cfg=None):
    """Lê do PCP (aba Cronograma Base) as datas de início/fim por pavimento
    para cada atividade configurada. Sem a planilha, usa a última leitura."""
    cfg = cfg or carregar_config(pacote)
    pc = cfg["prazos"]
    arq = _achar_planilha(pc.get("planilha"))
    if not arq:
        st = estado_pacote(pacote)
        if st.get("prazos"):
            return {"ok": True, "cache": True, **st["prazos"]}
        return {"ok": False, "error": "Planilha do PCP (Cronograma Base) não encontrada."}
    import openpyxl
    import warnings
    tarefas = [t for t, _ in pc["colunas"]]
    por_pav = {}
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb = openpyxl.load_workbook(arq, read_only=True, data_only=True)
    try:
        ws = wb[pc.get("aba", "Cronograma Base")]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if len(row) < 7 or row[1] != pc.get("linha_producao") or row[2] not in tarefas:
                continue
            ini, fim = row[5], row[6]
            if not isinstance(ini, dt.datetime) or not isinstance(fim, dt.datetime):
                continue
            por_pav.setdefault(str(row[3]).strip(), {})[row[2]] = [ini.strftime("%d/%m/%Y"), fim.strftime("%d/%m/%Y")]
    finally:
        wb.close()

    def _ordem(p):
        m = re.match(r"(\d+)", p)
        return int(m.group(1)) if m else 999
    linhas = [{"pavimento": p, "datas": por_pav[p]} for p in sorted(por_pav, key=_ordem)]
    nome = os.path.basename(arq)
    m = re.search(r"(\d{2})\.(\d{4})", nome)
    ref = f"{m.group(1)}/{m.group(2)}" if m else time.strftime("%m/%Y", time.localtime(os.path.getmtime(arq)))
    prazos = {"linhas": linhas, "arquivo": arq, "fonte": f"PCP {cfg.get('obra', '')} {ref}".strip(),
              "lido_em": time.strftime("%Y-%m-%d %H:%M:%S")}
    _atualizar_estado(pacote, prazos=prazos)
    return {"ok": True, "cache": False, **prazos}


# ─── geração do .docx ─────────────────────────────────────────────────────
def _data_extenso(d):
    return f"{d.day:02d} de {MESES[d.month - 1]} de {d.year}"


def _br(v):
    return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _fmt(texto, vars_):
    return re.sub(r"\{(\w+)\}", lambda m: str(vars_.get(m.group(1), m.group(0))), texto or "")


def gerar_docx(pacote, opcoes=None, grupos=None, prazos=None):
    """Monta a carta. opcoes (tudo opcional): data (date), empresa, ac,
    prazo_proposta, eng_responsavel, telefone, revisao. Retorna
    (BytesIO, avisos:list[str])."""
    from docx import Document
    from docx.shared import Pt, Cm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
    from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    cfg = carregar_config(pacote)
    op = {k: v for k, v in (opcoes or {}).items() if v not in (None, "")}
    st = estado_pacote(pacote)
    grupos = grupos if grupos is not None else grupos_selecionados(pacote, cfg, st)
    if prazos is None:
        prazos = ler_prazos(pacote, cfg)
    avisos = []
    hoje = op.get("data_documento") or dt.date.today()

    pad = cfg.get("padroes", {})
    V = {"obra": cfg.get("obra", ""), "ano": str(hoje.year)}
    for k in ("eng_responsavel", "prazo_proposta", "telefone", "revisao"):
        V[k] = op.get(k) or _fmt(pad.get(k, ""), V)
    V["data_pranchas"] = (st.get("sync_at") or "")[:10]
    if V["data_pranchas"]:
        y, m_, d_ = V["data_pranchas"].split("-")
        V["data_pranchas"] = f"{d_}/{m_}/{y}"
    else:
        V["data_pranchas"] = "[[—]]"
    frentes = [f for f in frentes_efetivas(pacote, cfg, st) if f["incluida"]]
    ctx = contexto_escopo(pacote, cfg, st)
    V.update({k: v for k, v in ctx.items() if isinstance(v, str)})
    V["frentes_texto"] = ctx["frentes_texto"] or "[[escopo a definir]]"
    if not frentes:
        avisos.append("Nenhuma frente de serviço marcada no escopo.")

    # prazos -> variáveis de mobilização
    linhas_prazo = prazos.get("linhas") if prazos.get("ok") else []
    col_tarefas = cfg["prazos"]["colunas"]
    if linhas_prazo:
        primeira = linhas_prazo[0]
        inicios = [ab[0] for ln in linhas_prazo for ab in ln["datas"].values() if ab]
        ini = min(inicios, key=lambda s: dt.datetime.strptime(s, "%d/%m/%Y")) if inicios else ""
        V["mobilizacao"] = ini
        V["primeiro_pav"] = primeira["pavimento"].replace("Pav", "pavimento")
        try:
            dd = dt.datetime.strptime(ini, "%d/%m/%Y")
            V["mobilizacao_mes"] = f"{MESES[dd.month - 1].capitalize()}/{dd.year}"
        except ValueError:
            V["mobilizacao_mes"] = ini
        V["fonte_cronograma"] = prazos.get("fonte", "")
    else:
        V.update(mobilizacao="[[__/__/____]]", primeiro_pav="[[__]]", mobilizacao_mes="[[a definir]]",
                 fonte_cronograma="[[cronograma não encontrado]]")
        avisos.append("Prazos por pavimento não encontrados (planilha do PCP).")

    # ── documento base ──
    BLACK = RGBColor(0, 0, 0)
    GRAY = RGBColor(0x59, 0x59, 0x59)
    HEAD_HEX, LIGHT_HEX = "D9D9D9", "F2F2F2"
    FONT = "Arial"
    doc = Document()
    stl = doc.styles["Normal"]
    stl.font.name = FONT
    stl.font.size = Pt(9.5)
    stl.font.color.rgb = BLACK
    stl.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    stl.paragraph_format.space_after = Pt(4)
    stl.paragraph_format.line_spacing = 1.15
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(2)
    sec.top_margin, sec.bottom_margin = Cm(3.3), Cm(2)
    sec.header_distance, sec.footer_distance = Cm(0.9), Cm(0.9)
    CONTENT_W = 17.0

    def shade(cell, hex_):
        tcPr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), hex_)
        tcPr.append(shd)

    def cell_margins(table, top=50, bottom=50, left=80, right=80):
        m = OxmlElement("w:tblCellMar")
        for k, v in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
            e = OxmlElement(f"w:{k}"); e.set(qn("w:w"), str(v)); e.set(qn("w:type"), "dxa"); m.append(e)
        table._tbl.tblPr.append(m)

    def borders(table, color="000000", size=2, inside=True):
        b = OxmlElement("w:tblBorders")
        edges = ["top", "left", "bottom", "right"] + (["insideH", "insideV"] if inside else [])
        for e in ["top", "left", "bottom", "right", "insideH", "insideV"]:
            el = OxmlElement(f"w:{e}")
            if e in edges:
                el.set(qn("w:val"), "single"); el.set(qn("w:sz"), str(size)); el.set(qn("w:color"), color)
            else:
                el.set(qn("w:val"), "nil")
            b.append(el)
        table._tbl.tblPr.append(b)

    def no_borders(table):
        borders(table, inside=False, size=0, color="FFFFFF")
        for el in table._tbl.tblPr.find(qn("w:tblBorders")):
            el.set(qn("w:val"), "nil")

    def widths(table, cms):
        table.autofit = False
        tbl = table._tbl
        for gc, w in zip(tbl.tblGrid.findall(qn("w:gridCol")), cms):
            gc.set(qn("w:w"), str(int(w * 567)))
        for row in table.rows:
            for c, w in zip(row.cells, cms):
                c.width = Cm(w)
        tblW = tbl.tblPr.find(qn("w:tblW"))
        if tblW is None:
            tblW = OxmlElement("w:tblW"); tbl.tblPr.append(tblW)
        tblW.set(qn("w:w"), str(int(sum(cms) * 567))); tblW.set(qn("w:type"), "dxa")
        lay = OxmlElement("w:tblLayout"); lay.set(qn("w:type"), "fixed"); tbl.tblPr.append(lay)

    def repeat_header(row):
        e = OxmlElement("w:tblHeader"); e.set(qn("w:val"), "true"); row._tr.get_or_add_trPr().append(e)

    def cant_split(row):
        e = OxmlElement("w:cantSplit"); e.set(qn("w:val"), "true"); row._tr.get_or_add_trPr().append(e)

    def para_border(p, side, color, size, space):
        pb = OxmlElement("w:pBdr")
        e = OxmlElement(f"w:{side}")
        e.set(qn("w:val"), "single"); e.set(qn("w:sz"), str(size)); e.set(qn("w:space"), str(space)); e.set(qn("w:color"), color)
        pb.append(e); p._p.get_or_add_pPr().append(pb)

    def add_runs(p, text, size=None, color=None, bold_all=False, italic=False):
        """Texto com **negrito** e [[pendência destacada em amarelo]]."""
        for tok in re.split(r"(\*\*.+?\*\*|\[\[.+?\]\])", _fmt(text, V)):
            if not tok:
                continue
            b, hl = bold_all, False
            if tok.startswith("**"):
                tok, b = tok[2:-2], True
            elif tok.startswith("[["):
                tok, hl, b = tok[2:-2], True, True
            r = p.add_run(tok)
            r.bold = b; r.italic = italic
            if size: r.font.size = Pt(size)
            if color is not None: r.font.color.rgb = color
            if hl:
                h = OxmlElement("w:highlight"); h.set(qn("w:val"), "yellow"); r._r.get_or_add_rPr().append(h)
        return p

    def P(text="", align=WD_ALIGN_PARAGRAPH.JUSTIFY, size=None, color=None, bold=False, italic=False,
          before=0, after=4):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(before); p.paragraph_format.space_after = Pt(after)
        add_runs(p, text, size, color, bold, italic)
        return p

    def H1(num, text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(12); p.paragraph_format.space_after = Pt(5)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(f"{num}.  {text.upper()}"); r.bold = True; r.font.size = Pt(10.5)
        return p

    def H2(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8); p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text); r.bold = True; r.font.size = Pt(9.5)
        return p

    def item(num, text, indent=1.1):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pf = p.paragraph_format
        pf.left_indent = Cm(indent); pf.first_line_indent = Cm(-indent); pf.space_after = Pt(4)
        pf.tab_stops.add_tab_stop(Cm(indent))
        p.add_run(f"{num}\t")
        add_runs(p, text)
        return p

    def cell_text(cell, text, bold=False, size=9, align=WD_ALIGN_PARAGRAPH.LEFT, color=None, italic=False):
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = align
        p.paragraph_format.space_after = Pt(0); p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        add_runs(p, text, size, color, bold, italic)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        return p

    def cell_extra(cell, text, size=7.5):
        pp = cell.add_paragraph(); pp.alignment = cell.paragraphs[0].alignment
        pp.paragraph_format.space_after = Pt(0); pp.paragraph_format.line_spacing = 1.0
        add_runs(pp, text, size, GRAY, italic=True)

    def field(par, instr):
        r = par.add_run()
        for tag, attrs, txt in (("w:fldChar", {"w:fldCharType": "begin"}, None),
                                ("w:instrText", {"xml:space": "preserve"}, instr),
                                ("w:fldChar", {"w:fldCharType": "separate"}, None),
                                ("w:t", {}, "1"),
                                ("w:fldChar", {"w:fldCharType": "end"}, None)):
            e = OxmlElement(tag)
            for k, v in attrs.items():
                e.set(qn(k), v)
            if txt is not None:
                e.text = txt
            r._r.append(e)
        r.font.size = Pt(7.5); r.font.color.rgb = GRAY

    def keep_table(t):
        for row in t.rows[:-1]:
            for c in row.cells:
                for p_ in c.paragraphs:
                    p_.paragraph_format.keep_with_next = True

    def head_row(table, titles, size=8.5):
        for c, t in zip(table.rows[0].cells, titles):
            shade(c, HEAD_HEX); cell_text(c, t, bold=True, size=size, align=WD_ALIGN_PARAGRAPH.CENTER)
        repeat_header(table.rows[0])

    # ── cabeçalho / rodapé ──
    hdr = sec.header
    ht = hdr.add_table(rows=2, cols=3, width=Cm(CONTENT_W))
    borders(ht, "000000", 4); widths(ht, [4.3, 8.9, 3.8]); cell_margins(ht, 40, 40, 90, 90)
    lg = ht.cell(0, 0).merge(ht.cell(1, 0))
    lg.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    logo = os.path.join(BASE_DIR, cfg.get("logo", "static/temazec_logo.png"))
    if os.path.exists(logo):
        lg.paragraphs[0].add_run().add_picture(logo, width=Cm(3.4))
    lg.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    cell_text(ht.cell(0, 1), cfg["cabecalho"]["titulo"], bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.CENTER)
    cell_text(ht.cell(1, 1), cfg["cabecalho"]["subtitulo"], size=8.5, align=WD_ALIGN_PARAGRAPH.CENTER)
    cell_text(ht.cell(0, 2), f"Revisão: {V['revisao']}", size=8)
    cell_text(ht.cell(1, 2), f"Data: {hoje.strftime('%d/%m/%Y')}", size=8)
    hp = hdr.paragraphs[0]; hp._p.getparent().remove(hp._p)
    hp = hdr.add_paragraph(); hp.paragraph_format.space_after = Pt(0); hp.add_run().font.size = Pt(2)

    ftr = sec.footer
    ft = ftr.add_table(rows=1, cols=2, width=Cm(CONTENT_W))
    no_borders(ft); widths(ft, [12.0, 5.0]); cell_margins(ft, 60, 0, 0, 0)
    for c in ft.rows[0].cells:
        tb = OxmlElement("w:tcBorders"); e = OxmlElement("w:top")
        e.set(qn("w:val"), "single"); e.set(qn("w:sz"), "4"); e.set(qn("w:color"), "000000")
        tb.append(e); c._tc.get_or_add_tcPr().append(tb)
    cell_text(ft.cell(0, 0), cfg["rodape"], size=7.5, color=GRAY)
    fp = ft.cell(0, 1).paragraphs[0]; fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    ft.cell(0, 1).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    fp.paragraph_format.space_after = Pt(0)
    for t_, instr in (("Página ", "PAGE"), (" de ", "NUMPAGES")):
        r = fp.add_run(t_); r.font.size = Pt(7.5); r.font.color.rgb = GRAY
        field(fp, instr)
    fp0 = ftr.paragraphs[0]; fp0._p.getparent().remove(fp0._p)
    ftr.add_paragraph().add_run().font.size = Pt(1)

    # ── abertura ──
    data_carta = op.get("data")
    linha_data = (f"{cfg.get('cidade', '')}, {_data_extenso(data_carta)}." if data_carta
                  else f"{cfg.get('cidade', '')}, ____ de ____________________ de {hoje.year}.")
    P(linha_data, align=WD_ALIGN_PARAGRAPH.RIGHT, after=12)
    P("À", align=WD_ALIGN_PARAGRAPH.LEFT, after=0)
    P("**Empresa convidada:** " + (op.get("empresa") or "_" * 46), align=WD_ALIGN_PARAGRAPH.LEFT, after=0)
    P("**A/C:** " + (op.get("ac") or "_" * 46), align=WD_ALIGN_PARAGRAPH.LEFT, after=10)
    P("**Ref.:** " + cfg["referencia"], align=WD_ALIGN_PARAGRAPH.LEFT, after=4)

    # ── 1. identificação ──
    H1(1, "Identificação")
    P(cfg["identificacao"]["intro"], after=8)
    idt = doc.add_table(rows=0, cols=2)
    borders(idt); widths(idt, [4.2, 12.8]); cell_margins(idt)
    for k, v in cfg["identificacao"]["linhas"]:
        row = idt.add_row(); cant_split(row)
        shade(row.cells[0], LIGHT_HEX)
        cell_text(row.cells[0], k, bold=True, size=8.5)
        cell_text(row.cells[1], v, size=9.5)

    # ── 2. objeto ──
    H1(2, "Objeto")
    for t in cfg["objeto"]:
        P(t)

    # ── 3. documentos técnicos (AutoDoc) ──
    dcfg = cfg["documentos"]
    H1(3, "Documentos técnicos para levantamento")
    P(dcfg["intro"], after=3)
    dtab = doc.add_table(rows=1, cols=4)
    borders(dtab); widths(dtab, [4.6, 7.4, 2.6, 2.4]); cell_margins(dtab, 35, 35, 70, 70)
    head_row(dtab, ["Documento", "Descrição", "Revisão", "Data"], size=9)
    tem_pendente = False
    esperados_cfg = dcfg.get("esperados") or {}
    if not grupos:
        avisos.append("Pranchas do AutoDoc ainda não sincronizadas — rode a atualização.")
    for g in grupos:
        row = dtab.add_row(); cant_split(row)
        for c in row.cells: shade(c, LIGHT_HEX)
        m_ = row.cells[0].merge(row.cells[3])
        cell_text(m_, g["grupo"].upper(), bold=True, size=8.5)
        for d in g["pranchas"]:
            row = dtab.add_row(); cant_split(row)
            cell_text(row.cells[0], d["documento"], size=8.5)
            cell_text(row.cells[1], d["descricao"], size=8.5)
            rev_txt = d["rev"] if d.get("aprovada") else f"[[{d['rev']}]]"
            cell_text(row.cells[2], rev_txt, size=8.5, align=WD_ALIGN_PARAGRAPH.CENTER)
            if d.get("pendente"):
                tem_pendente = True
                cell_extra(row.cells[2], f"{d['pendente']} em aprovação", size=6.5)
            cell_text(row.cells[3], d.get("data", ""), size=8.5, align=WD_ALIGN_PARAGRAPH.CENTER)
        esp = esperados_cfg.get(g["grupo"])
        if esp and not any(d["local"] in esp["locais"] or re.fullmatch(r"\d{2}P", d["local"]) for d in g["pranchas"]):
            row = dtab.add_row(); cant_split(row)
            cell_text(row.cells[0], "[[A INCLUIR]]", size=8.5)
            cell_text(row.cells[1], f"[[{esp['texto']}]]", size=8.5)
            cell_text(row.cells[2], "[[—]]", size=8.5, align=WD_ALIGN_PARAGRAPH.CENTER)
            cell_text(row.cells[3], "", size=8.5)
            avisos.append(f"{g['grupo']}: ainda não há no AutoDoc — {esp['texto']}.")
    fixos = dcfg.get("fixos") or []
    if fixos:
        row = dtab.add_row(); cant_split(row)
        for c in row.cells: shade(c, LIGHT_HEX)
        cell_text(row.cells[0].merge(row.cells[3]), "OUTROS DOCUMENTOS", bold=True, size=8.5)
        for a, b, c in fixos:
            row = dtab.add_row(); cant_split(row)
            cell_text(row.cells[0], a, size=8.5); cell_text(row.cells[1], b, size=8.5)
            cell_text(row.cells[2], c, size=8.5, align=WD_ALIGN_PARAGRAPH.CENTER); cell_text(row.cells[3], "", size=8.5)
    P(dcfg.get("nota", ""), size=8.5, italic=True, color=GRAY, before=3, after=0)
    if tem_pendente:
        P("Revisões indicadas como “em aprovação” existem no AutoDoc, mas ainda não foram liberadas; "
          "a proposta deve considerar a revisão aprovada listada.", size=8, italic=True, color=GRAY, after=4)

    # ── 4. matriz de responsabilidades ──
    H1(4, "Matriz de responsabilidades").paragraph_format.page_break_before = True
    P(cfg["matriz"]["intro"], after=3)
    mt = doc.add_table(rows=1, cols=4)
    borders(mt); widths(mt, [1.1, 11.9, 2.0, 2.0]); cell_margins(mt, 35, 35, 70, 70)
    head_row(mt, ["Item", "Descrição", "Contratada", "Contratante"])
    nomes_frente = {f["id"]: f["nome"] for f in frentes_efetivas(pacote, cfg, st)}
    for gi, g in enumerate(matriz_efetiva(pacote, cfg, st), 1):
        row = mt.add_row(); cant_split(row)
        for c in row.cells: shade(c, LIGHT_HEX)
        cell_text(row.cells[0], str(gi), bold=True, size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
        cell_text(row.cells[1].merge(row.cells[3]), g["grupo"].upper(), bold=True, size=9)
        for li, it in enumerate(g["itens"], 1):
            row = mt.add_row(); cant_split(row)
            cell_text(row.cells[0], f"{gi}.{li}", size=8.5, align=WD_ALIGN_PARAGRAPH.CENTER)
            cell_text(row.cells[1], it["texto"], size=8.5)
            if it["frente"] != "geral" and len(frentes) > 1:
                cell_extra(row.cells[1], nomes_frente.get(it["frente"], it["frente"]), size=6.5)
            cell_text(row.cells[2], "X" if it["resp"] == "A" else "", bold=True, size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
            cell_text(row.cells[3], "X" if it["resp"] == "B" else "", bold=True, size=9, align=WD_ALIGN_PARAGRAPH.CENTER)

    # ── 5. quantidades por frente de serviço (preços vão na planilha anexa) ──
    q = cfg["quantidades"]
    V["revisao_base"] = q.get("revisao_base", "")
    H1(5, "Quadro resumo de quantidades")
    P(q["intro"], after=3).paragraph_format.keep_with_next = True
    qt = doc.add_table(rows=1, cols=4)
    borders(qt); widths(qt, [4.6, 8.0, 1.6, 2.8]); cell_margins(qt, 35, 35, 60, 60)
    head_row(qt, ["Frente de serviço", "Serviço", "Un.", "Quantidade"])
    a_levantar = q.get("texto_a_levantar", "a levantar")
    faltando_qtd = []
    if not any(f["servicos"] for f in frentes):
        row = qt.add_row()
        cell_text(row.cells[0].merge(row.cells[3]), "[[Escopo a preencher]]", size=8.5, align=WD_ALIGN_PARAGRAPH.CENTER)
    for f in frentes:
        if not f["servicos"]:
            continue
        primeira_linha = None
        for s in f["servicos"]:
            row = qt.add_row(); cant_split(row)
            if primeira_linha is None:
                primeira_linha = row
            cell_text(row.cells[1], s["servico"], size=8.5)
            cell_text(row.cells[2], s["un"], size=8.5, align=WD_ALIGN_PARAGRAPH.CENTER)
            if s["qtd"] is None:
                faltando_qtd.append(s["servico"])
                cell_text(row.cells[3], f"[[{a_levantar}]]", size=8.5, align=WD_ALIGN_PARAGRAPH.CENTER)
            else:
                cell_text(row.cells[3], _br(float(s["qtd"])), size=8.5, align=WD_ALIGN_PARAGRAPH.RIGHT)
        # frente numa célula só, mesclada na vertical
        cel = primeira_linha.cells[0]
        if len(f["servicos"]) > 1:
            cel = cel.merge(qt.rows[-1].cells[0])
        shade(cel, LIGHT_HEX)
        cell_text(cel, f["nome"], bold=True, size=8.5)
    notas = [_fmt(f["nota"], V) for f in frentes if f.get("nota")]
    for n_ in notas:
        P(n_, size=8, italic=True, color=GRAY, before=2, after=0)
    if faltando_qtd:
        avisos.append(f"{len(faltando_qtd)} serviço(s) sem quantidade (saem como '{a_levantar}').")
    # levantamento feito numa revisão que já não é a vigente -> alerta amarelo
    base_g = next((g for g in grupos if g["grupo"] == q.get("grupo_base")), None)
    if base_g:
        divergentes = [d for d in base_g["pranchas"] if d["rev"] != q.get("revisao_base")]
        if divergentes:
            lista = ", ".join(f"{d['local']} {d['rev']}" for d in divergentes)
            P(f"[[Atenção: quantidades levantadas na {q.get('revisao_base')}; revisão vigente no AutoDoc: {lista}. "
              f"Conferir o levantamento antes do envio.]]", size=8.5, before=2)
            avisos.append(f"Quantidades feitas na {q.get('revisao_base')}, mas há pranchas em outra revisão ({lista}).")
    if q.get("pendencia"):
        P(f"[[{q['pendencia']}]]", size=8.5, before=2, after=4)
    P(cfg["proposta"]["texto"], before=6)

    # ── 6. itens gerais ──
    H1(6, "Itens gerais")
    for i, t in enumerate(cfg["itens_gerais"], 1):
        item(f"6.{i}", t)

    # ── 7. financeiro ──
    fin = cfg["financeiro"]
    H1(7, "Itens financeiros e fiscais")
    H2("7.1  Critérios de medição e pagamento")
    for i, t in enumerate(fin["medicao"], 1):
        item(f"7.1.{i}", t, indent=1.3)
    H2("7.2  Documentos necessários")
    P(fin["documentos"])
    H2("7.3  Retenção técnica")
    P(fin["retencao"])

    # ── 8. prazos ──
    pz_cfg = cfg["prazos"]
    H1(8, "Prazos")
    item("8.1", pz_cfg["entrega_proposta"])
    item("8.2", pz_cfg["mobilizacao"])
    item("8.3", pz_cfg["multa"])

    # ── 9. condições gerais + assinaturas ──
    H1(9, "Condições gerais")
    for t in cfg["condicoes_gerais"]:
        P(t)
    P(cfg["declaracao"], before=6, bold=True).paragraph_format.keep_with_next = True
    sg = doc.add_table(rows=3, cols=3)
    no_borders(sg); widths(sg, [7.5, 2.0, 7.5])
    for col, (a, b) in zip([0, 2], cfg["assinaturas"]):
        cell_text(sg.rows[0].cells[col], "", size=9)
        sg.rows[0].cells[col].paragraphs[0].paragraph_format.space_before = Pt(36)
        cell_text(sg.rows[1].cells[col], a, bold=True, size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
        para_border(sg.rows[1].cells[col].paragraphs[0], "top", "404040", 6, 4)
        cell_text(sg.rows[2].cells[col], b, size=8.5, color=GRAY, align=WD_ALIGN_PARAGRAPH.CENTER)
    for r_ in sg.rows: cant_split(r_)
    ct = cfg["contato"]
    for i, t in enumerate(ct):
        p = P(t, align=WD_ALIGN_PARAGRAPH.LEFT, before=14 if i == 0 else 0, after=0 if i < len(ct) - 1 else 4,
              size=9 if i == len(ct) - 1 else None)
        if i < len(ct) - 1:
            p.paragraph_format.keep_with_next = True
    p.runs[-1].add_break(WD_BREAK.PAGE)

    # ── Anexo I ──
    p = P(align=WD_ALIGN_PARAGRAPH.LEFT, after=2); r = p.add_run("ANEXO I"); r.bold = True; r.font.size = Pt(9.5)
    p = P(align=WD_ALIGN_PARAGRAPH.LEFT, after=8); r = p.add_run(_fmt(cfg["anexo"]["titulo"], V)); r.bold = True; r.font.size = Pt(10.5)
    qr = doc.add_table(rows=0, cols=3)
    borders(qr); widths(qr, [1.0, 6.0, 10.0]); cell_margins(qr, 40, 40, 70, 70)
    for gi, g in enumerate(anexo_efetivo(pacote, cfg, st), 1):
        row = qr.add_row(); cant_split(row)
        for c in row.cells: shade(c, HEAD_HEX)
        cell_text(row.cells[0], str(gi), bold=True, size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
        cell_text(row.cells[1].merge(row.cells[2]), g["grupo"].upper(), bold=True, size=9)
        for li, it in enumerate(g["itens"], 1):
            row = qr.add_row(); cant_split(row)
            cell_text(row.cells[0], f"{gi}.{li}", size=8.5, align=WD_ALIGN_PARAGRAPH.CENTER)
            shade(row.cells[1], LIGHT_HEX)
            cell_text(row.cells[1], it["rotulo"], bold=True, size=8.5)
            v = it.get("valor")
            if it["tipo"] == "texto":
                txt = str(v or "")
            else:
                marcados = set(v or []) if it["tipo"] == "multi" else {v}
                txt = "   ".join(f"{'☒' if o in marcados else '☐'} {o}" for o in it["opcoes"])
                if it.get("complemento"):
                    txt += "  – " + it["complemento"]
            cell_text(row.cells[2], txt, size=8.5)
            if it.get("obs"):
                cell_extra(row.cells[2], "OBS: " + it["obs"])

    # a de documentos cresce com o AutoDoc — não amarrar (jogaria tudo pra outra página)
    for t_ in (idt, qt):
        keep_table(t_)
    cp = doc.core_properties
    cp.title = cfg.get("nome", "Carta Convite")
    cp.author = f"Temazec – Engenharia {cfg.get('obra', '')}"
    cp.subject = _fmt(cfg["referencia"], V)
    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf, avisos


# ─── envio por e-mail ─────────────────────────────────────────────────────
# Sai da caixa de quem preencher o remetente/senha na aba. A senha só fica
# gravada se marcarem "lembrar", e mesmo assim criptografada pelo Windows
# (DPAPI — só este usuário, neste computador, consegue abrir).
EMAIL_CRED_PATH = os.path.join(DATA_DIR, ".carta_email_credenciais.json")
_RE_EMAIL = re.compile(r"[A-Za-z0-9._%+'-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")


def _chave_fernet():
    """Chave pro Linux (sem DPAPI): derivada do segredo do dashboard, que fica
    no servidor (variável DASHBOARD_SECRET_KEY ou dados/synergia/.dashboard_secret)."""
    import base64
    import hashlib
    seg = os.environ.get("DASHBOARD_SECRET_KEY") or os.environ.get("CARTA_EMAIL_KEY") or ""
    if not seg:
        try:
            with open(os.path.join(DATA_DIR, ".dashboard_secret"), "r", encoding="utf-8") as fh:
                seg = fh.read().strip()
        except Exception:
            seg = ""
    if not seg:
        raise OSError("Sem segredo do servidor pra proteger a senha (defina DASHBOARD_SECRET_KEY).")
    return base64.urlsafe_b64encode(hashlib.sha256(("carta-email:" + seg).encode("utf-8")).digest())


def _dpapi(dados, proteger):
    if os.name != "nt":
        # Linux/servidor: DPAPI não existe — Fernet com chave do segredo do servidor.
        # Senhas salvas antes pelo Windows (DPAPI) não abrem aqui: é só salvar de novo.
        from cryptography.fernet import Fernet
        f = Fernet(_chave_fernet())
        return f.encrypt(dados) if proteger else f.decrypt(dados)
    import ctypes
    from ctypes import wintypes

    class BLOB(ctypes.Structure):
        _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]
    buf = ctypes.create_string_buffer(dados, len(dados))
    entrada, saida = BLOB(len(dados), buf), BLOB()
    fn = ctypes.windll.crypt32.CryptProtectData if proteger else ctypes.windll.crypt32.CryptUnprotectData
    if not fn(ctypes.byref(entrada), None, None, None, None, 0x01, ctypes.byref(saida)):
        raise OSError("DPAPI falhou")
    try:
        return ctypes.string_at(saida.pbData, saida.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(saida.pbData)


def _ler_credenciais_email():
    try:
        with open(EMAIL_CRED_PATH, "r", encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return {}


def salvar_senha_email(remetente, senha):
    import base64
    cred = _ler_credenciais_email()
    cred[remetente.lower()] = base64.b64encode(_dpapi(senha.encode("utf-8"), True)).decode()
    with open(EMAIL_CRED_PATH, "w", encoding="utf-8") as fh:
        json.dump(cred, fh)


def esquecer_senha_email(remetente):
    cred = _ler_credenciais_email()
    if cred.pop((remetente or "").lower(), None) is not None:
        with open(EMAIL_CRED_PATH, "w", encoding="utf-8") as fh:
            json.dump(cred, fh)


def senha_email_salva(remetente):
    import base64
    blob = _ler_credenciais_email().get((remetente or "").lower())
    if not blob:
        return None
    try:
        return _dpapi(base64.b64decode(blob), False).decode("utf-8")
    except Exception:
        return None


def smtp_padrao(remetente, cfg_email=None):
    dom = (remetente or "").split("@")[-1].lower()
    if dom in ("gmail.com", "googlemail.com"):
        return "smtp.gmail.com", 587
    if dom in ("outlook.com", "hotmail.com", "live.com", "msn.com"):
        return "smtp-mail.outlook.com", 587
    cfg_email = cfg_email or {}
    return cfg_email.get("smtp_host", "smtp.office365.com"), int(cfg_email.get("smtp_port", 587))


def config_email(pacote, cfg=None, st=None):
    """Padrões do JSON + o que foi salvo na aba (sem senha)."""
    cfg = cfg or carregar_config(pacote)
    st = st if st is not None else estado_pacote(pacote)
    base = dict(cfg.get("email") or {})
    base.update({k: v for k, v in (st.get("email") or {}).items() if v not in (None, "")})
    V = {"obra": cfg.get("obra", ""), **{k: v for k, v in contexto_escopo(pacote, cfg, st).items() if isinstance(v, str)}}
    base["assunto"] = _fmt(base.get("assunto", ""), V)
    base["mensagem"] = _fmt(base.get("mensagem", ""), V)
    base["senha_salva"] = bool(senha_email_salva(base.get("remetente", ""))) if base.get("remetente") else False
    return base


def destinatarios_de(texto):
    vistos, out = set(), []
    for e in _RE_EMAIL.findall(texto or ""):
        if e.lower() not in vistos:
            vistos.add(e.lower())
            out.append(e)
    return out


def _smtp_conectar(host, porta, remetente, senha):
    import smtplib
    import ssl
    ctx = ssl.create_default_context()
    if int(porta) == 465:
        s = smtplib.SMTP_SSL(host, 465, timeout=60, context=ctx)
    else:
        s = smtplib.SMTP(host, int(porta), timeout=60)
        s.ehlo()
        s.starttls(context=ctx)
        s.ehlo()
    s.login(remetente, senha)
    return s


def _erro_smtp_legivel(exc):
    txt = str(exc)
    if "535" in txt or "Authentication unsuccessful" in txt or "Username and Password not accepted" in txt:
        return ("E-mail ou senha recusados pelo servidor. Na conta Microsoft 365 com verificação em duas etapas é "
                "preciso usar uma 'senha de app', e o SMTP autenticado precisa estar liberado para a caixa "
                f"(pedir ao TI). Detalhe: {txt[:200]}")
    return txt[:300]


def enviar_carta(pacote, opcoes, email, senha, anexar_pranchas=True, progresso=None):
    """Gera a carta (PDF, ou .docx se não houver Word), baixa as pranchas
    marcadas e manda um e-mail SEPARADO para cada destinatário (um
    convidado não vê o outro). Anexos acima do limite vão em partes."""
    from email.message import EmailMessage
    from email.utils import formataddr, make_msgid

    def _p(txt):
        _log(txt)
        _atualizar_estado(pacote, envio={**(estado_pacote(pacote).get("envio") or {}), "etapa": txt})
        if progresso:
            progresso(txt)

    cfg = carregar_config(pacote)
    remetente = (email.get("remetente") or "").strip()
    dest = destinatarios_de(email.get("destinatarios"))
    if not _RE_EMAIL.fullmatch(remetente):
        return {"ok": False, "error": "E-mail do remetente inválido."}
    if not dest:
        return {"ok": False, "error": "Informe ao menos um destinatário."}
    senha = senha or senha_email_salva(remetente)
    if not senha:
        return {"ok": False, "error": "Informe a senha do e-mail do remetente."}
    host = (email.get("smtp_host") or "").strip() or smtp_padrao(remetente, cfg.get("email"))[0]
    porta = int(email.get("smtp_port") or smtp_padrao(remetente, cfg.get("email"))[1])

    _atualizar_estado(pacote, envio={"status": "enviando", "inicio": time.strftime("%Y-%m-%d %H:%M:%S"),
                                     "etapa": "Conferindo login do e-mail...", "enviados": [], "erros": []})
    try:
        _smtp_conectar(host, porta, remetente, senha).quit()
    except Exception as exc:
        msg = _erro_smtp_legivel(exc)
        _atualizar_estado(pacote, envio={"status": "erro", "etapa": "", "erros": [msg],
                                         "fim": time.strftime("%Y-%m-%d %H:%M:%S")})
        return {"ok": False, "error": msg}

    # carta
    _p("Gerando a carta convite...")
    buf, avisos = gerar_docx(pacote, opcoes=opcoes)
    os.makedirs(SAIDA_DIR, exist_ok=True)
    docx_path = _gravar(os.path.join(SAIDA_DIR, nome_arquivo(pacote, opcoes)), buf.getvalue())
    carta_path = docx_path
    try:
        pdf_path = os.path.splitext(docx_path)[0] + ".pdf"
        tmp_pdf = os.path.join(tempfile.gettempdir(), f"carta_{int(time.time() * 1000)}.pdf")
        if exportar_pdf(docx_path, tmp_pdf):
            with open(tmp_pdf, "rb") as fh:
                carta_path = _gravar(pdf_path, fh.read())
            os.remove(tmp_pdf)
    except Exception as exc:
        _log(f"PDF da carta não gerado ({exc}); vai o .docx")

    # pranchas
    anexos_pr, erros = [], []
    if anexar_pranchas:
        pranchas = [d for g in grupos_selecionados(pacote, cfg) for d in g["pranchas"]]
        _p(f"Baixando {len(pranchas)} prancha(s) do AutoDoc...")
        locais, erros = baixar_pranchas(pranchas, progresso=_p)
        anexos_pr = [locais[d["arquivo"]] for d in pranchas if d["arquivo"] in locais]
        if erros:
            _atualizar_estado(pacote, envio={**estado_pacote(pacote)["envio"], "erros": list(erros)})

    planilha_path = _gravar(os.path.join(SAIDA_DIR, nome_planilha(pacote)), gerar_planilha_proposta(pacote, opcoes).getvalue())
    limite = float((cfg.get("email") or {}).get("limite_anexos_mb", 18)) * 1024 * 1024
    partes, atual = [], [carta_path, planilha_path]
    tam = sum(os.path.getsize(a) for a in atual)
    for a in anexos_pr:
        t = os.path.getsize(a)
        if atual and tam + t > limite:
            partes.append(atual)
            atual, tam = [], 0
        atual.append(a)
        tam += t
    partes.append(atual)

    V = {"obra": cfg.get("obra", ""), **{k: v for k, v in contexto_escopo(pacote, cfg).items() if isinstance(v, str)}}
    assunto = _fmt(email.get("assunto") or (cfg.get("email") or {}).get("assunto", "Carta convite"), V)
    corpo = _fmt(email.get("mensagem") or (cfg.get("email") or {}).get("mensagem", ""), V)
    nome_rem = (email.get("nome_remetente") or "").strip()
    enviados = []
    try:
        _p("Conectando no servidor de e-mail...")
        smtp = _smtp_conectar(host, porta, remetente, senha)
        try:
            for i, pessoa in enumerate(dest, 1):
                for n, grupo in enumerate(partes, 1):
                    _p(f"Enviando para {pessoa} ({i}/{len(dest)})" + (f" — parte {n}/{len(partes)}" if len(partes) > 1 else ""))
                    msg = EmailMessage()
                    msg["From"] = formataddr((nome_rem, remetente)) if nome_rem else remetente
                    msg["To"] = pessoa
                    msg["Subject"] = assunto + (f" (parte {n}/{len(partes)})" if len(partes) > 1 else "")
                    msg["Message-ID"] = make_msgid()
                    if n == 1:
                        texto = corpo
                        if len(partes) > 1:
                            texto += f"\n\n(Os projetos seguem em {len(partes)} e-mails por causa do tamanho dos anexos.)"
                    else:
                        texto = f"Continuação dos projetos para levantamento – parte {n} de {len(partes)}."
                    msg.set_content(texto)
                    for caminho in grupo:
                        with open(caminho, "rb") as fh:
                            dados = fh.read()
                        ext = os.path.splitext(caminho)[1].lower()
                        sub = {"pdf": "pdf", "docx": "vnd.openxmlformats-officedocument.wordprocessingml.document",
                               "xlsx": "vnd.openxmlformats-officedocument.spreadsheetml.sheet"}.get(ext[1:], "octet-stream")
                        msg.add_attachment(dados, maintype="application", subtype=sub, filename=os.path.basename(caminho))
                    smtp.send_message(msg)
                enviados.append(pessoa)
        finally:
            try:
                smtp.quit()
            except Exception:
                pass
    except Exception as exc:
        erros.append(_erro_smtp_legivel(exc))

    fim = time.strftime("%Y-%m-%d %H:%M:%S")
    status = "ok" if enviados and len(enviados) == len(dest) else ("parcial" if enviados else "erro")
    hist = ([{"data": fim, "remetente": remetente, "destinatarios": enviados, "anexos": 2 + len(anexos_pr),
              "partes": len(partes)}] + (estado_pacote(pacote).get("historico_envios") or []))[:100]
    _atualizar_estado(pacote, envio={"status": status, "etapa": "", "fim": fim, "enviados": enviados,
                                     "erros": erros, "anexos": 2 + len(anexos_pr), "partes": len(partes),
                                     "avisos": avisos}, historico_envios=hist)
    _log(f"E-mail: enviado para {len(enviados)}/{len(dest)} destinatário(s), {1 + len(anexos_pr)} anexo(s).")
    return {"ok": bool(enviados), "enviados": enviados, "erros": erros, "error": "; ".join(erros) if not enviados else ""}


def nome_arquivo(pacote, opcoes=None):
    cfg = carregar_config(pacote)
    base = _fmt(cfg.get("arquivo_saida") or "Carta convite - {titulo}", _vars_nome(pacote, cfg))
    emp = ((opcoes or {}).get("empresa") or "").strip()
    if emp:
        base += " - " + re.sub(r'[\\/:*?"<>|]', "", emp)
    return base + ".docx"


def _exportar_pdf_libreoffice(docx_path, pdf_path):
    """Linux/servidor: converte com o LibreOffice (soffice) em modo headless."""
    import shutil
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return False
    work = tempfile.mkdtemp(prefix="lo_pdf_")
    try:
        r = subprocess.run(
            [soffice, "--headless", "--norestore", f"-env:UserInstallation=file://{work}/profile",
             "--convert-to", "pdf", "--outdir", work, os.path.abspath(docx_path)],
            capture_output=True, text=True, timeout=180)
        gerado = os.path.join(work, os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")
        if r.returncode == 0 and os.path.exists(gerado):
            shutil.move(gerado, pdf_path)
            return True
        return False
    except Exception:
        return False
    finally:
        shutil.rmtree(work, ignore_errors=True)


def exportar_pdf(docx_path, pdf_path):
    """Windows: converte via Word (COM). Linux: LibreOffice. Sem nenhum dos dois, só não gera o PDF."""
    if os.name != "nt":
        return _exportar_pdf_libreoffice(docx_path, pdf_path)
    ps = (
        "$ErrorActionPreference='Stop';"
        "$w=New-Object -ComObject Word.Application;$w.Visible=$false;$w.DisplayAlerts=0;$w.AutomationSecurity=3;"
        "try{$d=$w.Documents.Open($env:CC_IN,$false,$true,$false);$d.Fields.Update()|Out-Null;"
        "$d.ExportAsFixedFormat($env:CC_OUT,17);$d.Close(0)}finally{$w.Quit()}"
    )
    env = dict(os.environ, CC_IN=os.path.abspath(docx_path), CC_OUT=os.path.abspath(pdf_path))
    r = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps],
                       env=env, capture_output=True, text=True, timeout=180)
    return r.returncode == 0 and os.path.exists(pdf_path)


def _gravar(destino, dados):
    """Grava; se o arquivo estiver aberto no Word, salva com sufixo."""
    try:
        with open(destino, "wb") as fh:
            fh.write(dados)
        return destino
    except PermissionError:
        raiz, ext = os.path.splitext(destino)
        alt = f"{raiz} (atualizada {time.strftime('%H%M')}){ext}"
        with open(alt, "wb") as fh:
            fh.write(dados)
        return alt


def rodar_robo(pacote="mdo_bruta", sincronizar=True, saida=None, pdf=True, opcoes=None):
    _log(f"=== Robô Carta Convite ({pacote}) ===")
    if sincronizar:
        res = sincronizar_pranchas(pacote)
        if res.get("ok"):
            if res["mudancas"]:
                _log("Pranchas com revisão nova: " + "; ".join(res["mudancas"]))
            elif res.get("primeira_leitura"):
                _log("Primeira leitura das pranchas registrada.")
            else:
                _log("Nenhuma revisão nova desde a última leitura.")
        else:
            _log("AVISO: " + res.get("error", "falha no AutoDoc") + " — usando a última leitura salva.")
    prazos = ler_prazos(pacote)
    if prazos.get("ok"):
        _log(f"Prazos: {len(prazos['linhas'])} pavimentos ({'cache' if prazos.get('cache') else prazos.get('arquivo')})")
    buf, avisos = gerar_docx(pacote, opcoes=opcoes, prazos=prazos)
    dados = buf.getvalue()
    os.makedirs(SAIDA_DIR, exist_ok=True)
    nome = nome_arquivo(pacote, opcoes)
    principal = _gravar(os.path.join(SAIDA_DIR, nome), dados)
    gerados = [principal]
    planilha = gerar_planilha_proposta(pacote, opcoes).getvalue()
    extras = [_gravar(os.path.join(SAIDA_DIR, nome_planilha(pacote)), planilha)]
    if saida:
        os.makedirs(saida, exist_ok=True)
        gerados.append(_gravar(os.path.join(saida, nome), dados))
        extras.append(_gravar(os.path.join(saida, nome_planilha(pacote)), planilha))
    if pdf:
        for docx_path in list(gerados):
            pdf_path = os.path.splitext(docx_path)[0] + ".pdf"
            try:
                tmp_pdf = os.path.join(tempfile.gettempdir(), f"carta_{int(time.time() * 1000)}.pdf")
                if exportar_pdf(docx_path, tmp_pdf):
                    with open(tmp_pdf, "rb") as fh:
                        gerados.append(_gravar(pdf_path, fh.read()))
                    os.remove(tmp_pdf)
                else:
                    _log("PDF não gerado (Word indisponível?).")
                    break
            except Exception as exc:
                _log(f"PDF não gerado: {exc}")
                break
    for a in avisos:
        _log("PENDÊNCIA: " + a)
    gerados += extras
    for g in gerados:
        _log("Gerado: " + g)
    _atualizar_estado(pacote, ultima_geracao=time.strftime("%Y-%m-%d %H:%M:%S"), ultimos_arquivos=gerados,
                      ultimos_avisos=avisos)
    return {"ok": True, "arquivos": gerados, "avisos": avisos}


def main():
    ap = argparse.ArgumentParser(description="Robô da Carta Convite (AutoDoc + PCP -> .docx/.pdf)")
    ap.add_argument("--pacote", default="mdo_bruta")
    ap.add_argument("--sem-autodoc", action="store_true", help="não loga no AutoDoc; usa a última leitura")
    ap.add_argument("--saida", default=os.path.join(os.path.expanduser("~"), "Downloads") if os.name == "nt" else "",
                    help="pasta extra onde deixar a cópia (padrão no Windows: Downloads; no servidor, nenhuma)")
    ap.add_argument("--sem-pdf", action="store_true")
    a = ap.parse_args()
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    r = rodar_robo(a.pacote, sincronizar=not a.sem_autodoc, saida=a.saida or None, pdf=not a.sem_pdf)
    sys.exit(0 if r.get("ok") else 1)


if __name__ == "__main__":
    main()
