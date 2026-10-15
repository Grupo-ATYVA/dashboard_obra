"""
Dashboard Obra - SPE T2 BUENO  v4
Servidor Flask local — lê o Excel direto da pasta dados ao lado do app

NOTA SOBRE CREDORES SEM NOME (exibidos como 'CREDOR XXXX'):
  Os nomes desses fornecedores existem no SIENGE mas não estão na aba
  'dfin_bancos' do Excel (eles ainda nao tiveram movimento bancario).
  Para resolver: adicione ao SQL do relatorio SIENGE a tabela de cadastro
  de credores com colunas codcredor e nmcredor, salvando como aba 'cad_credores'.
  O codigo abaixo ja usa essa aba automaticamente quando disponivel.

NOTA SOBRE CENTROS DE CUSTO:
  A aba 'fin_cpagar_aprobra' exporta apenas o CC 1016 (OBRA).
  Os demais CCs mostram dados de movimentacao bancaria (dfin_bancos) como
  proxy. Para dados completos de AP por CC, inclua as abas de cada CC no export.
"""
from flask import Flask, render_template, send_from_directory, send_file, request, jsonify, make_response, redirect, session
from werkzeug.utils import secure_filename
from werkzeug.security import generate_password_hash, check_password_hash
import pandas as pd, json, os, re, time, unicodedata, threading, traceback, glob, itertools, tempfile, subprocess, base64, copy, datetime, html
from io import BytesIO
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.request import Request, urlopen
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from functools import wraps
try:
    import docx as _docxlib
except Exception:
    _docxlib = None

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
app = Flask(__name__)
DATA_DIR = os.path.join(BASE_DIR, "dados", "synergia")
UAU_DATA_DIR = os.path.join(BASE_DIR, "dados", "uau")
BUILD_ID = os.environ.get("DASH_BUILD_ID", "dev")


def _ci_path(path):
    """O Windows ignora maiúscula/minúscula no nome do arquivo; o Linux não.
    Se o caminho exato não existe, devolve o arquivo da mesma pasta que só
    difere na caixa (ex.: 'credores.xlsx' x 'Credores.xlsx')."""
    if os.path.exists(path):
        return path
    folder, name = os.path.split(path)
    try:
        for n in os.listdir(folder):
            if n.lower() == name.lower():
                return os.path.join(folder, n)
    except OSError:
        pass
    return path


# Atrás de proxy HTTPS (Caddy/Cloudflare) o Flask precisa confiar nos cabeçalhos
# X-Forwarded-* pra enxergar o esquema (https) e o IP real do visitante.
if os.environ.get("DASHBOARD_BEHIND_PROXY") == "1":
    from werkzeug.middleware.proxy_fix import ProxyFix
    app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1)
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=os.environ.get("DASHBOARD_HTTPS") == "1",
    # a sessão expira após 12 h SEM uso (renova a cada requisição) — antes era cookie de navegador
    PERMANENT_SESSION_LIFETIME=datetime.timedelta(hours=int(os.environ.get("DASHBOARD_SESSION_HOURS", "12") or 12)),
)

# Sites autorizados a EMBUTIR o painel em iframe (ex.: SharePoint). Lista separada por espaço em
# DASHBOARD_FRAME_ANCESTORS, ex.: "https://suaempresa.sharepoint.com". Vazio = ninguém embute.
_FRAME_ANCESTORS = (os.environ.get("DASHBOARD_FRAME_ANCESTORS") or "").strip()


@app.after_request
def _security_headers(resp):
    resp.headers.setdefault("X-Content-Type-Options", "nosniff")
    resp.headers.setdefault("Referrer-Policy", "same-origin")
    resp.headers.setdefault("Permissions-Policy", "camera=(), microphone=(), geolocation=(), payment=()")
    resp.headers.setdefault("Cross-Origin-Opener-Policy", "same-origin")
    # Sem script-src de propósito: o index.html usa scripts/estilos inline em grande volume e
    # uma CSP restritiva o quebraria. Aqui entram só as diretivas que não afetam o funcionamento.
    ancestors = "'self' " + _FRAME_ANCESTORS if _FRAME_ANCESTORS else "'none'"
    resp.headers.setdefault(
        "Content-Security-Policy",
        f"frame-ancestors {ancestors}; base-uri 'self'; object-src 'none'; form-action 'self'",
    )
    if not _FRAME_ANCESTORS:
        resp.headers.setdefault("X-Frame-Options", "DENY")
    if os.environ.get("DASHBOARD_HTTPS") == "1":
        resp.headers.setdefault("Strict-Transport-Security", "max-age=31536000; includeSubDomains")
    # respostas de API e telas de login/admin nunca devem ficar em cache de navegador/proxy
    path = request.path or ""
    if path.startswith(("/api/", "/admin", "/login", "/setup-admin")):
        resp.headers.setdefault("Cache-Control", "no-store")
    return resp


_UNSAFE_METHODS = {"POST", "PUT", "PATCH", "DELETE"}


def _origin_host(value):
    from urllib.parse import urlsplit
    try:
        return (urlsplit(value).netloc or "").lower()
    except Exception:
        return ""


@app.before_request
def _csrf_origin_check():
    """Defesa contra CSRF: requisição que altera dados e vem de um navegador precisa ter Origin
    (ou, na falta, Referer) do próprio painel. Complementa o SameSite=Lax do cookie.
    Clientes sem navegador (curl, scripts do servidor) não enviam esses cabeçalhos e passam."""
    if request.method not in _UNSAFE_METHODS:
        return None
    origin = request.headers.get("Origin") or request.headers.get("Referer") or ""
    if not origin:
        return None
    permitidos = {(request.host or "").lower()}
    for extra in (os.environ.get("DASHBOARD_ALLOWED_ORIGINS") or "").split():
        permitidos.add(_origin_host(extra))
    if _origin_host(origin) in permitidos:
        return None
    if (request.path or "").startswith("/api/"):
        return jsonify({"ok": False, "error": "Origem da requisição não permitida."}), 403
    return make_response("Origem da requisição não permitida.", 403)
ATAS_DIR = os.path.join(DATA_DIR, "Ata de Reunião")
MELIUS_CACHE_PATH = os.path.join(DATA_DIR, "melius_cache.json")
MELIUS_CREDENTIALS_PATH = os.path.join(DATA_DIR, ".melius_credentials.json")
MELIUS_ROMPIMENTOS_CACHE_PATH = os.path.join(DATA_DIR, "melius_rompimentos_cache.json")
MELIUS_ROMPIMENTOS_STATUS_PATH = os.path.join(DATA_DIR, "melius_rompimentos_status.json")
MELIUS_DIARIO_CACHE_PATH = os.path.join(DATA_DIR, "melius_diario_cache.json")
MELIUS_DIARIO_STATUS_PATH = os.path.join(DATA_DIR, "melius_diario_status.json")
MELIUS_BASE_URL = "https://dev.meliusqualidade.com.br/apex/mq/reports"
MELIUS_COMPLETE_URL = "https://dev.inttegrasolutions.com.br/apex/mq/reports/fvs_lista_completa"
_melius_lock = threading.Lock()
AUTH_USERS_PATH = os.path.join(DATA_DIR, "access_users.json")
DASHBOARD_TABS = [
    ("bi", "Painel BI"), ("orc", "Orçamento API"), ("orcrel", "Orçamento Relatório"),
    ("pag", "Pagamentos"), ("ct", "Contratos"),
    ("mapa", "Mapa de Controle"), ("nf", "Notas Fiscais"), ("loc", "Almoxarifado"),
    ("vnd", "Vendas"), ("melius", "Melius Qualidade"), ("execaprop", "Títulos x Apropriação"), ("ifc", "Modelo IFC 3D"),
    ("atas", "Atas de Reunião"), ("minutas", "Minutas de Contrato"), ("cartas", "Cartas Convite"),
    ("taxa", "Taxa de Construção"),
    ("orcsql", "Orçamento SQL (UAU)"), ("uaucj", "CJ03/CJ05 (UAU)"),
    ("uaufin", "Financeiro UAU"), ("uaucomp", "Análise Orçamentária FAR"),
]
ORC_ITEM_STATUS_PATH = os.path.join(DATA_DIR, "orc_item_status.json")
IFC_MEDICAO_SAVES_PATH = os.path.join(DATA_DIR, "ifc_medicao_saves.json")
_ifc_medicao_lock = threading.Lock()
_orc_item_status_lock = threading.Lock()
AUTH_SECRET_PATH = os.path.join(DATA_DIR, ".dashboard_secret")
LOCAL_EXCEL_PATH = os.path.join(DATA_DIR, "bd_sienge_tmz.xlsx")
EXCEL_FILE_NAME = "bd_sienge_tmz.xlsx"
_excel_path_cache = {"path": "", "ts": 0}
REQUIRED_EXCEL_SHEETS = {"fin_cpagar_aprobra", "dfin_bancos"}
ORC_REL_PATH = os.path.join(DATA_DIR, "ORC-EXE-SYN-EXE_R00.xlsx")
NOMINAL_BUDGET_PATH = os.path.join(DATA_DIR, "ORÇAMENTO SIENGE BASICO R00.xlsx")
NOMINAL_BUDGET_TOTAL = 69846583.49
SUPPORT_REPORT_NAME = 'Apoio relatório - SYN - Synergia Bueno.xlsx'
SUPPORT_REPORT_LEGACY_NAME = 'Apoio relatório - SYN - Synergia Bueno (JULHO).xlsx'
SUPPORT_REPORT_PATH = os.path.join(DATA_DIR, SUPPORT_REPORT_NAME)
SUPPORT_REPORT_LEGACY_PATH = os.path.join(DATA_DIR, SUPPORT_REPORT_LEGACY_NAME)
if not os.path.exists(SUPPORT_REPORT_PATH) and os.path.exists(SUPPORT_REPORT_LEGACY_PATH):
    SUPPORT_REPORT_PATH = SUPPORT_REPORT_LEGACY_PATH
MAPA_CONTROLE_NAME = "Relatório de Mapa de Controle.xlsx"
MAPA_CONTROLE_PATH = os.path.join(DATA_DIR, MAPA_CONTROLE_NAME)


def _ensure_secret_key():
    os.makedirs(DATA_DIR, exist_ok=True)
    if os.path.exists(AUTH_SECRET_PATH):
        try:
            return Path(AUTH_SECRET_PATH).read_text(encoding="utf-8").strip()
        except Exception:
            pass
    secret = os.urandom(32).hex()
    Path(AUTH_SECRET_PATH).write_text(secret, encoding="utf-8")
    return secret


app.secret_key = os.environ.get("DASHBOARD_SECRET_KEY") or _ensure_secret_key()


def _melius_decrypt_windows(value):
    """Descriptografa um segredo DPAPI do usuário atual sem registrá-lo em logs."""
    if not value:
        return ""
    if os.name != "nt":
        raise RuntimeError("Credenciais protegidas por DPAPI só abrem no Windows; defina MELIUS_USER/MELIUS_PASSWORD/MELIUS_TOKEN no ambiente.")
    script = (
        "$ErrorActionPreference='Stop';"
        "$s=ConvertTo-SecureString $env:MELIUS_ENCRYPTED;"
        "$p=[Runtime.InteropServices.Marshal]::SecureStringToBSTR($s);"
        "try{[Runtime.InteropServices.Marshal]::PtrToStringBSTR($p)}"
        "finally{[Runtime.InteropServices.Marshal]::ZeroFreeBSTR($p)}"
    )
    env = os.environ.copy()
    env["MELIUS_ENCRYPTED"] = str(value)
    proc = subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-Command", script],
        capture_output=True, text=True, timeout=15, env=env,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    if proc.returncode != 0:
        raise RuntimeError("Não foi possível abrir as credenciais protegidas da Melius.")
    return proc.stdout.rstrip("\r\n")


def _load_melius_credentials():
    env_values = {
        "user": os.environ.get("MELIUS_USER", ""),
        "password": os.environ.get("MELIUS_PASSWORD", ""),
        "token": os.environ.get("MELIUS_TOKEN", ""),
    }
    if all(env_values.values()):
        return env_values
    if not os.path.exists(MELIUS_CREDENTIALS_PATH):
        return env_values
    try:
        encrypted = json.loads(Path(MELIUS_CREDENTIALS_PATH).read_text(encoding="utf-8-sig"))
        return {key: env_values[key] or _melius_decrypt_windows(encrypted.get(key, "")) for key in env_values}
    except Exception:
        return env_values


def _load_melius_site_senha():
    """Senha de LOGIN DO SITE (app.meliusqualidade.com.br), usada pelos robôs
    de rompimentos/diário de obra — é diferente da senha da API pública
    (MELIUS_PASSWORD, usada em Basic Auth). Contar as duas como uma coisa só
    foi o motivo do login dos robôs falhar mesmo com a senha certa da API."""
    return os.environ.get("MELIUS_SITE_SENHA", "").strip()


def _melius_rows(payload):
    if isinstance(payload, list):
        return payload
    if not isinstance(payload, dict):
        return []
    for key in (
        "items", "data", "rows", "result", "results", "obras", "servicos", "servico",
        "locais", "fvs", "registros", "verificacoes",
    ):
        value = payload.get(key)
        if isinstance(value, list):
            return value
        if isinstance(value, dict):
            nested = _melius_rows(value)
            if nested:
                return nested
    return []


def _melius_get(url, params=None):
    credentials = _load_melius_credentials()
    if not all(credentials.values()):
        raise RuntimeError("Credenciais da Melius não configuradas neste computador.")
    if params:
        url += ("&" if "?" in url else "?") + urlencode(params)
    basic = base64.b64encode(
        f"{credentials['user']}:{credentials['password']}".encode("utf-8")
    ).decode("ascii")
    req = Request(url, headers={
        "Accept": "application/json",
        "Authorization": "Basic " + basic,
        "token": credentials["token"],
        "User-Agent": "TEMAZEC-Dashboard/1.0",
    })
    try:
        with urlopen(req, timeout=45) as response:
            raw = response.read().decode("utf-8-sig", errors="replace")
            return json.loads(raw) if raw.strip() else []
    except HTTPError as exc:
        if exc.code in (401, 403):
            raise RuntimeError("A Melius recusou a autenticação. Verifique o usuário, a senha e o token.")
        raise RuntimeError(f"A API Melius respondeu com erro HTTP {exc.code}.")
    except URLError as exc:
        raise RuntimeError("Não foi possível conectar à API Melius.") from exc


def _melius_cache_default():
    return {
        "ok": False, "status": "não_atualizado", "updated_at": "", "error": "",
        "selected_work": "", "works": [], "services": [], "locations": [],
        "fvs_complete": [], "concrete_traceability": [],
        "modules": {
            "fvs": {"available": True, "label": "FVS"},
            "verificacoes": {"available": True, "label": "Verificações"},
            "rastreabilidade_concreto": {"available": False, "label": "Rastreabilidade de concreto", "reason": "Não disponibilizado na versão pública atual da API."},
            "diario_obra": {"available": False, "label": "Diário de obra", "reason": "Não disponibilizado na versão pública atual da API."},
            "fvm": {"available": False, "label": "FVM", "reason": "Não disponibilizado na versão pública atual da API."},
        },
    }


def _load_melius_cache():
    try:
        cached = json.loads(Path(MELIUS_CACHE_PATH).read_text(encoding="utf-8"))
        if isinstance(cached, dict):
            base = _melius_cache_default()
            base.update(cached)
            return base
    except Exception:
        pass
    return _melius_cache_default()


def _save_melius_cache(data):
    os.makedirs(DATA_DIR, exist_ok=True)
    tmp = MELIUS_CACHE_PATH + ".tmp"
    Path(tmp).write_text(json.dumps(data, ensure_ascii=False, default=str), encoding="utf-8")
    os.replace(tmp, MELIUS_CACHE_PATH)


def _melius_work_key(row):
    if not isinstance(row, dict):
        return ""
    for key in ("chave_obra", "obra_chave", "id_obra", "obra_id", "chave", "id"):
        value = row.get(key)
        if value not in (None, ""):
            return str(value)
    return ""


def _melius_text(row, keys):
    if not isinstance(row, dict):
        return ""
    normalized = {str(key).lower().replace("_", ""): value for key, value in row.items()}
    for key in keys:
        value = normalized.get(str(key).lower().replace("_", ""))
        if value not in (None, ""):
            return str(value)
    return ""


def _melius_walk(value):
    """Yield every dictionary in a Melius payload without assuming its nesting."""
    if isinstance(value, dict):
        yield value
        for child in value.values():
            yield from _melius_walk(child)
    elif isinstance(value, list):
        for child in value:
            yield from _melius_walk(child)


def _melius_first(payload, keys):
    """Return the first non-empty value found anywhere in a Melius payload."""
    for node in _melius_walk(payload):
        value = _melius_text(node, keys)
        if value:
            return value
    return ""


def _melius_number(value):
    if value in (None, ""):
        return None
    match = re.search(r"-?\d+(?:[.,]\d+)?", str(value))
    if not match:
        return None
    try:
        return float(match.group(0).replace(",", "."))
    except ValueError:
        return None


def _melius_fck(payload):
    raw = _melius_first(payload, (
        "fck_projeto", "fck", "fck_previsto", "fck_projetado",
        "resistencia_projeto", "resistencia_caracteristica", "classe_concreto",
    ))
    if raw:
        return _melius_number(raw) or raw
    pattern = re.compile(r"\bfck\s*(?:de|projeto|previsto)?\s*[:=-]?\s*(\d{1,3}(?:[.,]\d+)?)", re.I)
    for node in _melius_walk(payload):
        text = " | ".join(str(value) for value in node.values() if isinstance(value, (str, int, float)))
        match = pattern.search(text)
        if match:
            return float(match.group(1).replace(",", "."))
    return ""


def _melius_age_results(detail):
    """Extract concrete test ages and strengths from the complete FVS response."""
    results = {}
    age_pattern = re.compile(r"(?<!\d)(1|2|3|7|14|21|28|56|63|90|91)\s*(?:d|dia|dias)\b", re.I)
    mpa_pattern = re.compile(r"(\d{1,3}(?:[.,]\d{1,3})?)\s*(?:mpa|m\s*pa)\b", re.I)

    def add_result(age, strength=None, status="", date=""):
        if age is None:
            return
        age = int(age)
        current = results.setdefault(age, {
            "idade_dias": age, "resultado_mpa": None, "resultados_mpa": [],
            "status": status or "Nao informado", "data": date,
            "fonte": "Melius - ensaio de concreto",
        })
        numeric = _melius_number(strength)
        if numeric is not None and numeric not in current["resultados_mpa"]:
            current["resultados_mpa"].append(numeric)
            if current["resultado_mpa"] is None:
                current["resultado_mpa"] = numeric
        if status:
            current["status"] = status
        if date:
            current["data"] = date

    for node in _melius_walk(detail):
        # Some payloads return parallel arrays instead of one object per test.
        raw_ages = next((node.get(key) for key in (
            "idades", "idades_dias", "dias_ruptura", "idades_ensaio",
        ) if isinstance(node.get(key), list)), None)
        raw_strengths = next((node.get(key) for key in (
            "resultados", "resultados_mpa", "resistencias", "resistencias_mpa",
        ) if isinstance(node.get(key), list)), None)
        if raw_ages:
            for index, raw_age in enumerate(raw_ages):
                strength = raw_strengths[index] if raw_strengths and index < len(raw_strengths) else None
                add_result(
                    _melius_number(raw_age), strength,
                    _melius_text(node, ("status", "situacao", "estado")) or "Informado",
                    _melius_text(node, ("data_realizacao", "data_resultado", "data_ensaio", "data_ruptura")),
                )
        explicit_age = _melius_number(_melius_text(node, (
            "idade_dias", "idade", "dias", "idade_rompimento", "dias_ruptura", "idade_ensaio",
        )))
        explicit_strength = _melius_text(node, (
            "resistencia_mpa", "resultado_mpa", "resistencia", "tensao", "fc", "mpa", "valor_resultado",
        ))
        if explicit_age is not None:
            add_result(
                explicit_age,
                explicit_strength or _melius_text(node, ("resultado", "valor")),
                _melius_text(node, ("status", "situacao", "estado")),
                _melius_text(node, ("data_realizacao", "data", "data_resultado", "data_ensaio", "data_ruptura")),
            )
        # Some Melius responses expose the age in the field name (for example,
        # resistencia_7_dias) and only the numeric strength in its value.
        for key, raw_value in node.items():
            key_text = str(key).replace("_", " ").replace("-", " ")
            key_age = age_pattern.search(key_text)
            if not key_age or raw_value in (None, ""):
                continue
            numeric = re.search(r"-?\d{1,3}(?:[.,]\d{1,3})?", str(raw_value))
            if not numeric:
                continue
            age = int(key_age.group(1))
            add_result(age, numeric.group(0),
                _melius_text(node, ("status", "situacao", "estado")) or "Informado",
                _melius_text(node, ("data_realizacao", "data", "data_resultado", "data_ensaio", "data_ruptura")))
        text = " | ".join(
            str(value) for value in node.values()
            if isinstance(value, (str, int, float)) and value not in (None, "")
        )
        ages = [int(value) for value in age_pattern.findall(text)]
        if not ages:
            continue
        strengths = [float(value.replace(",", ".")) for value in mpa_pattern.findall(text)]
        status = _melius_text(node, ("status", "situacao", "resultado", "estado"))
        date = _melius_text(node, ("data_realizacao", "data", "data_resultado", "data_ensaio"))
        for index, age in enumerate(ages):
            strength = strengths[index] if index < len(strengths) else (strengths[0] if len(strengths) == 1 else None)
            add_result(age, strength, status, date)
    return [results[key] for key in sorted(results)]


def _melius_concrete_traceability(services, locations, complete=None):
    """Use concrete-related FVS records as the traceability source exposed by the public API."""
    concrete_codes = set()
    service_names = {}
    for row in services:
        code = _melius_text(row, ("servico_cod", "codigo_servico", "cod_servico", "codigo", "id"))
        name = _melius_text(row, ("servico", "nome_servico", "descricao", "nome"))
        if any(term in name.casefold() for term in ("concret", "pilar", "viga", "laje", "bloco de funda")):
            concrete_codes.add(code)
            service_names[code] = name

    details_by_fvs = {}
    for detail in complete or []:
        ids = set()
        for node in _melius_walk(detail):
            value = _melius_text(node, ("fvs_cod", "codigo_fvs", "id_fvs", "fvs", "id"))
            if value:
                ids.add(value)
        for value in ids:
            details_by_fvs.setdefault(value, []).append(detail)

    result = []
    for row in locations:
        code = _melius_text(row, ("servico_cod", "codigo_servico", "cod_servico", "codigo", "id"))
        if code not in concrete_codes:
            continue
        item = dict(row)
        item["servico_cod"] = code
        item["servico"] = service_names.get(code) or _melius_text(row, ("servico", "nome_servico", "descricao"))
        item["local"] = _melius_text(row, ("local", "nome_local", "localizacao", "ambiente"))
        item["fvs_cod"] = _melius_text(row, ("fvs_cod", "codigo_fvs", "fvs", "id_fvs", "id"))
        item["status"] = _melius_text(row, ("status", "situacao", "resultado", "estado")) or "Disponível"
        sources = [row] + details_by_fvs.get(item["fvs_cod"], [])
        item["data_aplicacao"] = _melius_first(sources, (
            "data_aplicacao", "dt_aplicacao", "data_concretagem", "data_execucao", "data",
        ))
        item["fornecedor"] = _melius_first(sources, (
            "fornecedor", "nome_fornecedor", "razao_social", "empresa_concreto", "concreteira",
        ))
        item["fck_projeto"] = _melius_fck(sources)
        item["slump_previsto"] = _melius_first(sources, (
            "slump_previsto", "slump", "espalhamento_previsto", "abatimento_previsto",
        ))
        item["resultados_idade"] = _melius_age_results(sources)
        item["detalhes_disponiveis"] = bool(
            item["fck_projeto"] or item["fornecedor"] or item["slump_previsto"]
            or item["resultados_idade"]
        )
        result.append(item)
    return result


def sync_melius_data(chave_obra=""):
    with _melius_lock:
        previous = _load_melius_cache()
        try:
            works_payload = _melius_get(MELIUS_BASE_URL + "/obra_lista")
            works = _melius_rows(works_payload)
            selected = str(chave_obra or "").strip()
            if not selected:
                selected = _melius_work_key(works[0]) if works else ""
            services = []
            complete = []
            locations = []
            if selected:
                services = _melius_rows(_melius_get(MELIUS_BASE_URL + "/fvs_servicos", {"chave_obra": selected}))
                complete = _melius_rows(_melius_get(MELIUS_COMPLETE_URL, {"chave_obra": selected}))
                # Locais são consultados somente para uma amostra controlada; a hierarquia completa
                # continua sendo a fonte principal para não tornar a atualização lenta.
                seen_codes = []
                for row in services:
                    if not isinstance(row, dict):
                        continue
                    code = next((row.get(k) for k in ("servico_cod", "codigo_servico", "cod_servico", "codigo", "id") if row.get(k) not in (None, "")), None)
                    if code is not None and str(code) not in seen_codes:
                        seen_codes.append(str(code))
                for code in seen_codes[:25]:
                    try:
                        rows = _melius_rows(_melius_get(MELIUS_BASE_URL + "/fvs_servicos_locais", {"servico_cod": code}))
                        for row in rows:
                            if isinstance(row, dict):
                                row.setdefault("servico_cod", code)
                        locations.extend(rows)
                    except Exception:
                        continue
            concrete_traceability = _melius_concrete_traceability(services, locations, complete)
            concrete_with_fck = sum(bool(row.get("fck_projeto")) for row in concrete_traceability)
            concrete_with_results = sum(bool(row.get("resultados_idade")) for row in concrete_traceability)
            data = _melius_cache_default()
            data.update({
                "ok": True, "status": "atualizado", "updated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
                "error": "", "selected_work": selected, "works": works,
                "services": services, "locations": locations, "fvs_complete": complete,
                "concrete_traceability": concrete_traceability,
            })
            data["modules"]["rastreabilidade_concreto"] = {
                "available": bool(concrete_traceability),
                "label": "Rastreabilidade de concreto",
                "total_fichas": len(concrete_traceability),
                "fichas_com_fck": concrete_with_fck,
                "fichas_com_resultados": concrete_with_results,
                "reason": (
                    f"{len(concrete_traceability)} fichas/localizacoes encontradas; "
                    f"{concrete_with_fck} com FCK e {concrete_with_results} com resultados por idade."
                    if concrete_traceability else
                    "Nenhuma FVS de concretagem foi encontrada para a obra."
                ),
            }
            _save_melius_cache(data)
            return data
        except Exception as exc:
            previous["ok"] = bool(previous.get("updated_at"))
            previous["status"] = "erro"
            previous["error"] = str(exc)
            previous["last_attempt_at"] = time.strftime("%Y-%m-%d %H:%M:%S")
            _save_melius_cache(previous)
            return previous


def sync_melius_location(servico_cod):
    """Atualiza os locais de UM serviço específico (fvs_servicos_locais),
    sem re-rodar a sincronização completa. Existe porque sync_melius_data()
    só busca locais para os 25 primeiros códigos de serviço (seen_codes[:25])
    pra não deixar a sincronização completa lenta — serviços além desse
    limite nunca têm seus locais atualizados pelo botão "Atualizar Melius
    completo". Esse endpoint permite atualizar qualquer serviço sob demanda."""
    with _melius_lock:
        data = _load_melius_cache()
        try:
            rows = _melius_rows(_melius_get(MELIUS_BASE_URL + "/fvs_servicos_locais", {"servico_cod": servico_cod}))
            for row in rows:
                if isinstance(row, dict):
                    row.setdefault("servico_cod", servico_cod)
            locations = [r for r in (data.get("locations") or []) if str(r.get("servico_cod", "")) != str(servico_cod)]
            locations.extend(rows)
            data["locations"] = locations
            concrete_traceability = _melius_concrete_traceability(data.get("services") or [], locations, data.get("fvs_complete") or [])
            data["concrete_traceability"] = concrete_traceability
            concrete_with_fck = sum(bool(row.get("fck_projeto")) for row in concrete_traceability)
            concrete_with_results = sum(bool(row.get("resultados_idade")) for row in concrete_traceability)
            data.setdefault("modules", _melius_cache_default()["modules"])
            data["modules"]["rastreabilidade_concreto"] = {
                "available": bool(concrete_traceability),
                "label": "Rastreabilidade de concreto",
                "total_fichas": len(concrete_traceability),
                "fichas_com_fck": concrete_with_fck,
                "fichas_com_resultados": concrete_with_results,
                "reason": (
                    f"{len(concrete_traceability)} fichas/localizacoes encontradas; "
                    f"{concrete_with_fck} com FCK e {concrete_with_results} com resultados por idade."
                    if concrete_traceability else
                    "Nenhuma FVS de concretagem foi encontrada para a obra."
                ),
            }
            data["ok"] = True
            data["status"] = "atualizado"
            data["updated_at"] = time.strftime("%Y-%m-%d %H:%M:%S")
            data["error"] = ""
            _save_melius_cache(data)
            return data
        except Exception as exc:
            data["error"] = str(exc)
            return {**data, "ok": False}


def _load_auth_users():
    if not os.path.exists(AUTH_USERS_PATH):
        return {"users": []}
    try:
        data = json.loads(Path(AUTH_USERS_PATH).read_text(encoding="utf-8"))
        if isinstance(data, dict) and isinstance(data.get("users"), list):
            return data
    except Exception:
        pass
    return {"users": []}


def _save_auth_users(data):
    os.makedirs(DATA_DIR, exist_ok=True)
    Path(AUTH_USERS_PATH).write_text(
        json.dumps(data, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def _find_auth_user(email):
    email = (email or "").strip().lower()
    for user in _load_auth_users().get("users", []):
        if (user.get("email") or "").strip().lower() == email:
            return user
    return None


def _auth_has_users():
    return bool(_load_auth_users().get("users"))


def _pw_fingerprint(user):
    import hashlib
    return hashlib.sha256((user.get("password_hash") or "").encode("utf-8")).hexdigest()[:20]


def _start_session(user):
    """Abre a sessão do usuário. Guarda uma impressão da senha atual: se a senha mudar,
    todas as sessões antigas dessa conta deixam de valer."""
    session.clear()
    session.permanent = True
    session["auth_email"] = (user.get("email") or "").strip().lower()
    session["auth_pv"] = _pw_fingerprint(user)


PASSWORD_MIN_LEN = 12


def _password_problem(password, email=""):
    """Texto do problema da senha, ou '' se estiver aceitável."""
    if len(password) < PASSWORD_MIN_LEN:
        return f"A senha precisa ter pelo menos {PASSWORD_MIN_LEN} caracteres."
    if len(set(password)) < 5:
        return "A senha é muito repetitiva. Use caracteres variados."
    low = password.lower()
    if email and email.split("@")[0] and email.split("@")[0].lower() in low:
        return "A senha não pode conter o seu e-mail."
    comuns = ("123456", "senha", "password", "qwerty", "temazec", "synergia", "admin")
    if any(c in low for c in comuns):
        return "A senha contém uma sequência comum fácil de adivinhar."
    return ""


def _current_auth_user():
    email = session.get("auth_email")
    user = _find_auth_user(email)
    if not user or user.get("blocked") or session.get("auth_pv") != _pw_fingerprint(user):
        session.pop("auth_email", None)
        session.pop("auth_pv", None)
        return None
    return user


def _login_page(error="", setup=False):
    title = "Criar administrador" if setup else "Entrar no Dashboard"
    subtitle = "Primeiro acesso: crie o usuário administrador." if setup else "Use seu e-mail e senha autorizados."
    action = "/setup-admin" if setup else "/login"
    if setup and os.environ.get("DASHBOARD_SETUP_TOKEN", "").strip():
        action += "?" + urlencode({"token": os.environ["DASHBOARD_SETUP_TOKEN"].strip()})
    button = "Criar acesso" if setup else "Entrar"
    name_field = """
        <label>Nome</label>
        <input name="name" autocomplete="name" required>
    """ if setup else ""
    return f"""<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{title} - TEMAZEC</title>
  <style>
    body{{margin:0;min-height:100vh;display:grid;place-items:center;background:#f3f5f8;font-family:Arial,sans-serif;color:#172033}}
    .card{{width:min(420px,calc(100vw - 32px));background:#fff;border:1px solid #d9e0ea;border-radius:12px;box-shadow:0 24px 70px rgba(15,23,42,.12);padding:28px}}
    .brand{{font-weight:900;letter-spacing:.04em;color:#a90025;font-size:28px;line-height:1}}
    .brand small{{display:block;color:#172033;font-size:10px;letter-spacing:.16em;margin-top:4px}}
    h1{{font-size:22px;margin:26px 0 6px}} p{{color:#667085;font-size:14px;margin:0 0 22px}}
    label{{display:block;margin:14px 0 6px;font-size:11px;font-weight:800;letter-spacing:.14em;text-transform:uppercase;color:#667085}}
    input{{width:100%;box-sizing:border-box;border:1px solid #cbd5e1;border-radius:8px;padding:12px;font-size:15px}}
    button{{width:100%;margin-top:20px;border:0;border-radius:8px;background:#a90025;color:#fff;font-weight:800;padding:13px;cursor:pointer}}
    .error{{background:#fff1f2;border:1px solid #fecdd3;color:#b91c1c;border-radius:8px;padding:10px 12px;font-size:13px;margin-bottom:14px}}
  </style>
</head>
<body>
  <form class="card" method="post" action="{action}">
    <div class="brand">TEMAZEC<small>INCORPORADORA</small></div>
    <h1>{title}</h1>
    <p>{subtitle}</p>
    {f'<div class="error">{error}</div>' if error else ''}
    {name_field}
    <label>E-mail</label>
    <input name="email" type="email" autocomplete="username" required>
    <label>Senha</label>
    <input name="password" type="password" autocomplete="current-password" required>
    <button type="submit">{button}</button>
  </form>
</body>
</html>"""


def _require_admin(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        user = _current_auth_user()
        if not user:
            return redirect("/login")
        if user.get("role") != "admin":
            return make_response("Acesso restrito ao administrador.", 403)
        return fn(*args, **kwargs)
    return wrapper


REAUTH_TTL_SECONDS = 600  # confirmacao de senha vale por 10 min


def _require_reauth(fn):
    """Alem de exigir admin, exige que a senha tenha sido reconfirmada
    recentemente nesta sessao — usado em telas sensiveis (cadastro de
    usuarios), mesmo que o usuario ja esteja logado."""
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if session.get("reauth_until", 0) < time.time():
            return redirect("/admin/reauth?next=" + request.path)
        return fn(*args, **kwargs)
    return wrapper


def _reauth_page(error="", next_path="/admin/acessos"):
    safe_next = next_path if str(next_path or "").startswith("/") else "/admin/acessos"
    return f"""<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Confirmar senha - TEMAZEC</title>
  <style>
    body{{margin:0;min-height:100vh;display:grid;place-items:center;background:#f3f5f8;font-family:Arial,sans-serif;color:#172033}}
    .card{{width:min(420px,calc(100vw - 32px));background:#fff;border:1px solid #d9e0ea;border-radius:12px;box-shadow:0 24px 70px rgba(15,23,42,.12);padding:28px}}
    .brand{{font-weight:900;letter-spacing:.04em;color:#a90025;font-size:28px;line-height:1}}
    .brand small{{display:block;color:#172033;font-size:10px;letter-spacing:.16em;margin-top:4px}}
    h1{{font-size:22px;margin:26px 0 6px}} p{{color:#667085;font-size:14px;margin:0 0 22px}}
    label{{display:block;margin:14px 0 6px;font-size:11px;font-weight:800;letter-spacing:.14em;text-transform:uppercase;color:#667085}}
    input{{width:100%;box-sizing:border-box;border:1px solid #cbd5e1;border-radius:8px;padding:12px;font-size:15px}}
    button{{width:100%;margin-top:20px;border:0;border-radius:8px;background:#a90025;color:#fff;font-weight:800;padding:13px;cursor:pointer}}
    .error{{background:#fff1f2;border:1px solid #fecdd3;color:#b91c1c;border-radius:8px;padding:10px 12px;font-size:13px;margin-bottom:14px}}
    a{{display:block;margin-top:14px;text-align:center;color:#667085;font-size:12px;text-decoration:none}}
  </style>
</head>
<body>
  <form class="card" method="post" action="/admin/reauth">
    <div class="brand">TEMAZEC<small>INCORPORADORA</small></div>
    <h1>Confirme sua senha</h1>
    <p>Área sensível (cadastro de usuários) — digite sua senha de novo pra continuar.</p>
    {f'<div class="error">{error}</div>' if error else ''}
    <input type="hidden" name="next" value="{safe_next}">
    <label>Senha</label>
    <input name="password" type="password" autocomplete="current-password" autofocus required>
    <button type="submit">Confirmar</button>
    <a href="/">Voltar ao dashboard</a>
  </form>
</body>
</html>"""


@app.route("/admin/reauth", methods=["GET", "POST"])
def admin_reauth():
    user = _current_auth_user()
    if not user:
        return redirect("/login")
    if user.get("role") != "admin":
        return make_response("Acesso restrito ao administrador.", 403)
    next_path = request.values.get("next") or "/admin/acessos"
    if request.method == "POST":
        password = request.form.get("password") or ""
        if not check_password_hash(user.get("password_hash", ""), password):
            return _reauth_page("Senha incorreta.", next_path)
        session["reauth_until"] = time.time() + REAUTH_TTL_SECONDS
        return redirect(next_path)
    return _reauth_page(next_path=next_path)


PRESENCE_ONLINE_TTL = 90  # segundos sem nenhum sinal = considerado "ausente"
_presence = {}
_presence_lock = threading.Lock()


def _tab_label(code):
    return dict(DASHBOARD_TABS).get(code, code or "—")


def _touch_presence(user, tab=None):
    if not user:
        return
    email = (user.get("email") or "").strip().lower()
    if not email:
        return
    with _presence_lock:
        entry = _presence.setdefault(email, {})
        entry["email"] = email
        entry["name"] = user.get("name") or email
        entry["cargo"] = user.get("cargo") or ""
        entry["empresa"] = user.get("empresa") or ""
        entry["role"] = user.get("role") or "user"
        if tab:
            entry["tab"] = tab
        entry["last_seen"] = time.time()
        entry["ip"] = request.headers.get("X-Forwarded-For", request.remote_addr) or ""


# Autorização no SERVIDOR por aba. Antes, "allowed_tabs" só escondia a aba na tela (CSS) e
# qualquer usuário logado conseguia chamar qualquer /api/* direto. Cada prefixo abaixo exige
# que o usuário tenha pelo menos UMA das abas listadas (admin e usuários sem restrição passam).
# Rotas que não aparecem aqui continuam liberadas a qualquer usuário logado (dados compartilhados
# como /dash_data.js, /api/refresh, presença, whoami). Ajuste esta tabela se uma aba nova surgir.
_TAB_API_PREFIXES = (
    ("/api/melius", {"melius"}),
    ("/api/notas-fiscais", {"nf"}),
    ("/api/export/notas-fiscais", {"nf"}),
    ("/api/almoxarifado", {"loc"}),
    ("/api/export/almoxarifado", {"loc"}),
    ("/api/locacoes", {"loc"}),
    ("/api/atas", {"atas"}),
    ("/api/minutas", {"minutas"}),
    ("/api/cartas", {"cartas"}),
    ("/api/ifc", {"ifc"}),
    ("/api/mapa_controle", {"mapa"}),
    ("/api/export/mapa-controle", {"mapa"}),
    ("/api/uau", {"orcsql", "uaucj", "uaufin", "uaucomp"}),
    ("/api/export/orcamento-sql", {"orcsql"}),
    ("/api/export/orcamento-relatorio", {"orc", "orcrel"}),
    ("/api/orc_item_status", {"orc", "orcrel"}),
)
# Prefixos só para administrador: diagnóstico, túnel e credenciais de integração.
_ADMIN_ONLY_PREFIXES = (
    "/api/debug",
    "/api/tunnel",
    "/api/sienge_sync",
)


def _path_starts(path, prefix):
    return path == prefix or path.startswith(prefix + "/") or path.startswith(prefix + "_") or path.startswith(prefix + "-") or path.startswith(prefix + ".")


def _forbidden_for_user(user, path):
    """True se o usuário logado não pode acessar este caminho (aba não liberada ou área admin)."""
    is_admin = user.get("role") == "admin"
    if is_admin:
        return False
    if any(_path_starts(path, p) for p in _ADMIN_ONLY_PREFIXES):
        return True
    allowed = set(user.get("allowed_tabs") or [])
    if not allowed:  # sem restrição = todas as abas (mesma regra da tela)
        return False
    for prefix, tabs in _TAB_API_PREFIXES:
        if _path_starts(path, prefix) and not (tabs & allowed):
            return True
    return False


@app.before_request
def _dashboard_auth_gate():
    path = request.path or "/"
    public = (
        path.startswith("/static/") or
        path in {"/login", "/logout", "/setup-admin", "/favicon.ico", "/api/status"} or
        path in {"/ifc_viewer.js", "/ifcviz.js"}
    )
    if public:
        return None
    if not _auth_has_users():
        if path.startswith("/api/"):
            return jsonify({"ok": False, "auth_required": True, "setup_required": True}), 401
        return redirect("/setup-admin")
    user = _current_auth_user()
    if user:
        if _forbidden_for_user(user, path):
            if path.startswith("/api/"):
                return jsonify({"ok": False, "error": "Acesso não liberado para o seu perfil."}), 403
            return make_response("Acesso não liberado para o seu perfil.", 403)
        _touch_presence(user)
        return None
    if path.startswith("/api/"):
        return jsonify({"ok": False, "auth_required": True}), 401
    return redirect("/login")


@app.route("/api/presence", methods=["POST"])
def api_presence_ping():
    user = _current_auth_user()
    if not user:
        return jsonify({"ok": False}), 401
    body = request.get_json(silent=True) or {}
    tab = str(body.get("tab") or "").strip()[:40]
    _touch_presence(user, tab=tab)
    return jsonify({"ok": True})


@app.route("/api/presence")
@_require_admin
def api_presence_list():
    now = time.time()
    with _presence_lock:
        items = list(_presence.values())
    out = []
    for e in items:
        last_seen = e.get("last_seen", 0)
        age = now - last_seen
        out.append({
            "email": e.get("email", ""),
            "name": e.get("name", ""),
            "cargo": e.get("cargo", ""),
            "empresa": e.get("empresa", ""),
            "role": e.get("role", "user"),
            "tab": e.get("tab", ""),
            "tab_label": _tab_label(e.get("tab", "")),
            "ip": e.get("ip", ""),
            "age_seconds": round(age),
            "online": age <= PRESENCE_ONLINE_TTL,
        })
    out.sort(key=lambda x: x["age_seconds"])
    return jsonify({"ok": True, "items": out, "online_ttl": PRESENCE_ONLINE_TTL})


@app.route("/admin/presenca")
@_require_admin
def admin_presenca():
    return f"""<!doctype html>
<html lang="pt-BR"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Quem está online - TEMAZEC</title>
<style>
body{{font-family:Arial,sans-serif;margin:0;background:#f3f5f8;color:#172033}}main{{max-width:1100px;margin:32px auto;padding:0 20px}}
.card{{background:#fff;border:1px solid #d9e0ea;border-radius:10px;padding:18px;margin-bottom:16px}}h1{{margin:0 0 14px}}
table{{width:100%;border-collapse:collapse}}th,td{{border-bottom:1px solid #e2e8f0;text-align:left;padding:10px;font-size:14px;vertical-align:top}}
th{{font-size:11px;letter-spacing:.12em;text-transform:uppercase;color:#667085}}
.top{{display:flex;justify-content:space-between;gap:12px;align-items:center}}a{{color:#a90025;text-decoration:none;font-weight:700}}
.dot{{display:inline-block;width:8px;height:8px;border-radius:50%;margin-right:6px}}
.dot-on{{background:#22c55e}} .dot-off{{background:#cbd5e1}}
.muted{{color:#94a3b8;font-size:12px}}
.count{{font-size:12px;color:#667085;font-weight:700}}
</style></head><body><main>
<div class="top"><h1>Quem está online</h1><div><a href="/admin/acessos">Controle de acessos</a> &nbsp; <a href="/">Voltar ao dashboard</a></div></div>
<section class="card">
  <div class="count" id="presence-count">Carregando…</div>
  <table><thead><tr><th>Status</th><th>Nome</th><th>E-mail</th><th>Empresa</th><th>Onde está</th><th>Última atividade</th><th>IP</th></tr></thead>
  <tbody id="presence-body"><tr><td colspan="7" class="muted">Carregando…</td></tr></tbody></table>
</section>
<script>
function pEsc(s){{ return String(s==null?'':s).replace(/[&<>"]/g, c=>({{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}}[c])); }}
function fmtAge(s){{
  if(s<5) return 'agora';
  if(s<60) return s+'s atrás';
  const m=Math.floor(s/60);
  if(m<60) return m+' min atrás';
  const h=Math.floor(m/60);
  if(h<24) return h+'h atrás';
  return Math.floor(h/24)+'d atrás';
}}
async function loadPresence(){{
  try{{
    const r = await fetch('/api/presence',{{cache:'no-store'}});
    const j = await r.json();
    if(!j.ok) return;
    const tb = document.getElementById('presence-body');
    const online = j.items.filter(it=>it.online).length;
    document.getElementById('presence-count').textContent = `${{online}} online de ${{j.items.length}} usuário(s) com atividade recente`;
    if(!j.items.length){{ tb.innerHTML = '<tr><td colspan="7" class="muted">Ninguém logado ainda.</td></tr>'; return; }}
    tb.innerHTML = j.items.map(it => `<tr>
      <td><span class="dot ${{it.online?'dot-on':'dot-off'}}"></span>${{it.online?'Online':'Ausente'}}</td>
      <td>${{pEsc(it.name)}}</td>
      <td>${{pEsc(it.email)}}</td>
      <td>${{pEsc(it.empresa)||'—'}}</td>
      <td>${{pEsc(it.tab_label)||'—'}}</td>
      <td>${{fmtAge(it.age_seconds)}}</td>
      <td class="muted">${{pEsc(it.ip)||'—'}}</td>
    </tr>`).join('');
  }}catch(e){{}}
}}
loadPresence();
setInterval(loadPresence, 8000);
</script>
</main></body></html>"""


@app.route("/setup-admin", methods=["GET", "POST"])
def setup_admin():
    if _auth_has_users():
        return redirect("/login")
    # No servidor público, o primeiro acesso (criar o administrador) exige o token
    # combinado em DASHBOARD_SETUP_TOKEN — senão quem achasse a URL primeiro viraria admin.
    setup_token = os.environ.get("DASHBOARD_SETUP_TOKEN", "").strip()
    if setup_token:
        import hmac
        if not hmac.compare_digest(request.args.get("token", ""), setup_token):
            return "Não encontrado", 404
    if request.method == "POST":
        email = (request.form.get("email") or "").strip().lower()
        password = request.form.get("password") or ""
        name = (request.form.get("name") or email).strip()
        problema = "Informe um e-mail válido." if (not email or "@" not in email) else _password_problem(password, email)
        if problema:
            return _login_page(problema, setup=True)
        novo = {
            "email": email,
            "name": name,
            "password_hash": generate_password_hash(password),
            "role": "admin",
            "blocked": False,
            "created_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        }
        _save_auth_users({"users": [novo]})
        _start_session(novo)
        return redirect("/")
    return _login_page(setup=True)


_login_fails = {}
_login_fails_lock = threading.Lock()
LOGIN_JANELA_SEG = 15 * 60
LOGIN_LIMITE_EMAIL = 5     # falhas por e-mail na janela
LOGIN_LIMITE_IP = 20       # falhas por IP na janela
_DUMMY_HASH = generate_password_hash("sem-usuario")


def _login_espera(keys_limites):
    """Segundos de espera restantes se alguma chave passou do limite (0 = liberado)."""
    now = time.time()
    with _login_fails_lock:
        for key, limite in keys_limites:
            fails = [t for t in _login_fails.get(key, []) if now - t < LOGIN_JANELA_SEG]
            _login_fails[key] = fails
            if len(fails) >= limite:
                return int(LOGIN_JANELA_SEG - (now - fails[0])) + 1
    return 0


def _login_registrar(keys, ok):
    with _login_fails_lock:
        for key in keys:
            if ok:
                _login_fails.pop(key, None)
            else:
                _login_fails.setdefault(key, []).append(time.time())


@app.route("/login", methods=["GET", "POST"])
def login():
    if not _auth_has_users():
        return redirect("/setup-admin")
    if request.method == "POST":
        email = (request.form.get("email") or "").strip().lower()
        password = request.form.get("password") or ""
        k_ip, k_em = f"ip:{request.remote_addr}", f"em:{email}"
        espera = _login_espera([(k_ip, LOGIN_LIMITE_IP), (k_em, LOGIN_LIMITE_EMAIL)])
        if espera:
            return _login_page(f"Muitas tentativas. Tente de novo em {max(1, espera // 60)} min."), 429
        user = _find_auth_user(email)
        # compara sempre um hash (mesmo sem usuário) pra não revelar por tempo se o e-mail existe
        senha_ok = check_password_hash((user or {}).get("password_hash") or _DUMMY_HASH, password)
        if not user or user.get("blocked") or not senha_ok:
            _login_registrar([k_ip, k_em], False)
            return _login_page("E-mail ou senha inválidos, ou usuário bloqueado.")
        _login_registrar([k_ip, k_em], True)
        _start_session(user)
        return redirect("/")
    return _login_page()


@app.route("/logout")
def logout():
    session.pop("auth_email", None)
    return redirect("/login")


@app.route("/api/whoami")
def api_whoami():
    email = (session.get("auth_email") or "").strip().lower()
    if not email:
        return jsonify({"ok": False}), 401
    users = _load_auth_users().get("users", [])
    user = next((u for u in users if str(u.get("email", "")).strip().lower() == email), None)
    return jsonify({
        "ok": True,
        "email": email,
        "name": (user or {}).get("name", "") or email,
        "role": (user or {}).get("role", "user"),
        "cargo": (user or {}).get("cargo", ""),
        "empresa": (user or {}).get("empresa", ""),
        "allowed_tabs": (user or {}).get("allowed_tabs") or [],
        "allowed_cc": (user or {}).get("allowed_cc") or [],
        "allowed_sources": (user or {}).get("allowed_sources") or [],
    })


DASHBOARD_SOURCES = [("synergia", "Synergia"), ("uau", "UAU")]


def _admin_cc_list():
    try:
        data = load_dashboard_cache() or {}
        return data.get("cc_list") or []
    except Exception:
        return []


def _checkbox_group_html(name, options, checked_values, all_label):
    """options: lista de (value, label). checked_values vazio/None = 'todas'
    (nenhum checkbox marcado, mas a lista inteira liberada — ver _parse_checked_list)."""
    checked_values = set(str(v) for v in (checked_values or []))
    items = "\n".join(
        f'<label class="chk"><input type="checkbox" name="{name}" value="{html.escape(str(val))}"'
        f'{" checked" if str(val) in checked_values else ""}> {html.escape(label)}</label>'
        for val, label in options
    )
    return f'<div class="chk-hint">Nenhuma marcada = acesso a {all_label} (padrão)</div><div class="chk-grid">{items}</div>'


def _parse_checked_list(field_name):
    return request.form.getlist(field_name)


def _admin_edit_form_html(action_url, cancel_label, u, tab_options, cc_options, submit_label, extra_hidden="", show_password=False, password_optional=False):
    email = u.get("email", "")
    name = u.get("name", "")
    cargo = u.get("cargo", "")
    empresa = u.get("empresa", "")
    role = u.get("role", "user")
    allowed_tabs = u.get("allowed_tabs") or []
    allowed_cc = u.get("allowed_cc") or []
    allowed_sources = u.get("allowed_sources") or []
    password_field = ('<div class="pw-wrap">'
        f'<input name="password" type="password" placeholder="{"Nova senha (opcional — deixe em branco para não alterar)" if password_optional else "Senha inicial"}" autocomplete="new-password" {"" if password_optional else "required"}>'
        '<button type="button" class="pw-toggle" onclick="const i=this.previousElementSibling;'
        'const show=i.type===\'password\';i.type=show?\'text\':\'password\';this.textContent=show?\'🙈\':\'👁\'">👁</button>'
        '</div>') if show_password else ''
    return f"""
    <form method="post" action="{action_url}" class="perm-form">
      {extra_hidden}
      <div class="form-row">
        <input name="name" placeholder="Nome" value="{html.escape(name)}" required>
        <input name="email" type="email" placeholder="E-mail" value="{html.escape(email)}" required>
        {password_field}
        <input name="cargo" placeholder="Cargo" value="{html.escape(cargo)}">
        <input name="empresa" placeholder="Empresa" value="{html.escape(empresa)}">
        <select name="role">
          <option value="user" {"selected" if role!="admin" else ""}>Usuário</option>
          <option value="admin" {"selected" if role=="admin" else ""}>Administrador</option>
        </select>
      </div>
      <div class="perm-cols">
        <div class="perm-col">
          <b>Abas do dashboard</b>
          {_checkbox_group_html('tabs', tab_options, allowed_tabs, 'todas as abas')}
        </div>
        <div class="perm-col">
          <b>Centros de custo</b>
          {_checkbox_group_html('cc', cc_options, allowed_cc, 'todos os CCs')}
        </div>
        <div class="perm-col" style="flex:0 1 200px">
          <b>Fonte de dados</b>
          {_checkbox_group_html('sources', DASHBOARD_SOURCES, allowed_sources, 'Synergia e UAU')}
        </div>
      </div>
      <button>{submit_label}</button>
    </form>"""


@app.route("/admin/acessos")
@_require_admin
@_require_reauth
def admin_acessos():
    users = _load_auth_users().get("users", [])
    cc_options = [(c.get("cod"), f"{c.get('cod')} — {c.get('nm','')}") for c in _admin_cc_list()]
    tab_options = DASHBOARD_TABS
    current_email = (session.get("auth_email") or "").strip().lower()

    def user_row(u):
        email = u.get("email", "")
        is_self = email.strip().lower() == current_email
        allowed_tabs = u.get("allowed_tabs") or []
        allowed_cc = u.get("allowed_cc") or []
        allowed_sources = u.get("allowed_sources") or []
        tabs_summary = "Todas" if not allowed_tabs else f"{len(allowed_tabs)} de {len(tab_options)}"
        cc_summary = "Todos" if not allowed_cc else f"{len(allowed_cc)} de {len(cc_options)}"
        src_summary = "Ambas" if not allowed_sources else "/".join(s.upper() for s in allowed_sources)
        block_form = f"""
          <form method="post" action="/admin/acessos/{'ativar' if u.get('blocked') else 'bloquear'}" style="display:inline">
            <input type="hidden" name="email" value="{html.escape(email)}">
            <button class="btn-sec" {"disabled title='Não é possível bloquear a própria conta'" if is_self else ""}>{'Ativar' if u.get('blocked') else 'Bloquear'}</button>
          </form>""" if not is_self else '<span class="muted">Sua conta</span>'
        edit_form = _admin_edit_form_html(
            "/admin/acessos/editar", "Cancelar", u, tab_options, cc_options, "Salvar alterações",
            extra_hidden=f'<input type="hidden" name="original_email" value="{html.escape(email)}">'
            + ('<div class="chk-hint" style="margin-bottom:10px">Seu próprio perfil (admin) não pode ser rebaixado por aqui.</div>' if is_self else ''),
            show_password=True, password_optional=True,
        )
        return f"""<tr>
          <td>{html.escape(u.get('name',''))}</td>
          <td>{html.escape(email)}</td>
          <td>{html.escape(u.get('cargo','') or '—')}</td>
          <td>{html.escape(u.get('empresa','') or '—')}</td>
          <td><span class="badge {'badge-admin' if u.get('role')=='admin' else ''}">{html.escape(u.get('role',''))}</span></td>
          <td>{'<span class="badge badge-blocked">Bloqueado</span>' if u.get('blocked') else '<span class="badge badge-ok">Ativo</span>'}</td>
          <td>{tabs_summary}</td>
          <td>{cc_summary}</td>
          <td>{src_summary}</td>
          <td>{block_form}</td>
        </tr>
        <tr class="perm-row"><td colspan="10">
          <details>
            <summary>Editar cadastro</summary>
            {edit_form}
          </details>
        </td></tr>"""

    rows = "\n".join(user_row(u) for u in users)
    new_user_form = _admin_edit_form_html(
        "/admin/acessos/criar", "", {}, tab_options, cc_options, "Criar acesso", show_password=True,
    )
    return f"""<!doctype html>
<html lang="pt-BR"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Acessos - TEMAZEC</title>
<style>
body{{font-family:Arial,sans-serif;margin:0;background:#f3f5f8;color:#172033}}main{{max-width:1240px;margin:32px auto;padding:0 20px}}
.card{{background:#fff;border:1px solid #d9e0ea;border-radius:10px;padding:18px;margin-bottom:16px}}h1{{margin:0 0 14px}}h2{{margin:0 0 14px;font-size:16px}}
input,select{{border:1px solid #cbd5e1;border-radius:8px;padding:10px;font-size:13px}}
button{{border:0;border-radius:8px;background:#a90025;color:white;padding:10px 14px;font-weight:800;cursor:pointer;font-size:13px}}
button.btn-sec{{background:#eef2f6;color:#172033;font-weight:700}}
button:disabled{{opacity:.45;cursor:not-allowed}}
table{{width:100%;border-collapse:collapse}}th,td{{border-bottom:1px solid #e2e8f0;text-align:left;padding:10px;font-size:14px;vertical-align:top}}
th{{font-size:11px;letter-spacing:.12em;text-transform:uppercase;color:#667085}}
.top{{display:flex;justify-content:space-between;gap:12px;align-items:center}}a{{color:#a90025;text-decoration:none;font-weight:700}}
.form-row{{display:flex;gap:10px;flex-wrap:wrap;align-items:center;margin-bottom:14px}}
.form-row input,.form-row select{{flex:1 1 160px}}
.pw-wrap{{position:relative;flex:1 1 160px;display:flex}}
.pw-wrap input{{width:100%;padding-right:34px;flex:1}}
.pw-toggle{{position:absolute;right:4px;top:50%;transform:translateY(-50%);background:none!important;border:0;padding:4px 6px;font-size:14px;cursor:pointer;color:#667085;line-height:1}}
.perm-cols{{display:flex;gap:24px;flex-wrap:wrap;margin-bottom:14px}}
.perm-col{{flex:1 1 280px}}
.perm-col b{{display:block;font-size:12px;color:#667085;text-transform:uppercase;letter-spacing:.08em;margin-bottom:8px}}
.chk-hint{{font-size:11px;color:#94a3b8;margin-bottom:6px}}
.chk-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:4px 12px;max-height:180px;overflow-y:auto;border:1px solid #e2e8f0;border-radius:8px;padding:10px}}
.chk{{display:flex;align-items:center;gap:6px;font-size:12.5px;font-weight:400;cursor:pointer}}
.chk input{{padding:0}}
.badge{{display:inline-block;padding:2px 9px;border-radius:20px;font-size:11px;font-weight:700;background:#eef2f6;color:#475569}}
.badge-admin{{background:#fff1f1;color:#a90025}}
.badge-ok{{background:#ecfdf5;color:#059669}}
.badge-blocked{{background:#fff1f1;color:#b91c1c}}
.muted{{color:#94a3b8;font-size:12px}}
.perm-row td{{padding-top:0;padding-bottom:14px;border-bottom:2px solid #e2e8f0}}
details summary{{cursor:pointer;color:#a90025;font-weight:700;font-size:12.5px;margin-bottom:10px}}
.perm-form{{background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:14px}}
</style></head><body><main>
<div class="top"><h1>Controle de acessos</h1><div><a href="/admin/presenca">Quem está online</a> &nbsp; <a href="/">Voltar ao dashboard</a> &nbsp; <a href="/logout">Sair</a></div></div>
<section class="card">
  <h2>Criar usuário</h2>
  {new_user_form}
</section>
<section class="card"><table><thead><tr><th>Nome</th><th>E-mail</th><th>Cargo</th><th>Empresa</th><th>Perfil</th><th>Status</th><th>Abas</th><th>CCs</th><th>Fonte</th><th>Ação</th></tr></thead><tbody>{rows}</tbody></table></section>
</main></body></html>"""


@app.route("/admin/acessos/criar", methods=["POST"])
@_require_admin
@_require_reauth
def admin_acessos_criar():
    data = _load_auth_users()
    email = (request.form.get("email") or "").strip().lower()
    password = request.form.get("password") or ""
    problema = _password_problem(password, email) if (email and password) else ""
    if problema:
        return make_response(
            f"<p style='font-family:Arial;padding:24px'>{html.escape(problema)} <a href='/admin/acessos'>Voltar</a></p>", 400)
    if email and password and not _find_auth_user(email):
        valid_tabs = {k for k, _ in DASHBOARD_TABS}
        valid_cc = {str(c.get("cod")) for c in _admin_cc_list()}
        valid_sources = {k for k, _ in DASHBOARD_SOURCES}
        data["users"].append({
            "email": email,
            "name": (request.form.get("name") or email).strip(),
            "cargo": (request.form.get("cargo") or "").strip(),
            "empresa": (request.form.get("empresa") or "").strip(),
            "password_hash": generate_password_hash(password),
            "role": "admin" if request.form.get("role") == "admin" else "user",
            "blocked": False,
            "allowed_tabs": [t for t in _parse_checked_list("tabs") if t in valid_tabs],
            "allowed_cc": [c for c in _parse_checked_list("cc") if c in valid_cc],
            "allowed_sources": [s for s in _parse_checked_list("sources") if s in valid_sources],
            "created_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        })
        _save_auth_users(data)
    return redirect("/admin/acessos")


@app.route("/admin/acessos/editar", methods=["POST"])
@_require_admin
@_require_reauth
def admin_acessos_editar():
    data = _load_auth_users()
    original_email = (request.form.get("original_email") or "").strip().lower()
    new_email = (request.form.get("email") or "").strip().lower()
    current_email = (session.get("auth_email") or "").strip().lower()
    is_self = original_email == current_email
    user = next((u for u in data.get("users", []) if (u.get("email") or "").strip().lower() == original_email), None)
    if not user:
        return redirect("/admin/acessos")
    # e-mail novo nao pode colidir com outro usuario ja existente
    if new_email and new_email != original_email:
        if any((u.get("email") or "").strip().lower() == new_email for u in data.get("users", []) if u is not user):
            return make_response(
                f"<p style='font-family:Arial;padding:24px'>O e-mail {html.escape(new_email)} já está em uso por outro usuário. "
                f"<a href='/admin/acessos'>Voltar</a></p>", 400)
        user["email"] = new_email
        if is_self:
            session["auth_email"] = new_email
    user["name"] = (request.form.get("name") or user.get("name", "")).strip()
    user["cargo"] = (request.form.get("cargo") or "").strip()
    user["empresa"] = (request.form.get("empresa") or "").strip()
    if not is_self:
        user["role"] = "admin" if request.form.get("role") == "admin" else "user"
    new_password = request.form.get("password") or ""
    if new_password:
        problema = _password_problem(new_password, user.get("email", ""))
        if problema:
            return make_response(
                f"<p style='font-family:Arial;padding:24px'>{html.escape(problema)} "
                "<a href='/admin/acessos'>Voltar</a></p>", 400)
        user["password_hash"] = generate_password_hash(new_password)
    if is_self:  # o admin que editou a própria conta continua logado após trocar e-mail/senha
        session["auth_email"] = (user.get("email") or "").strip().lower()
        session["auth_pv"] = _pw_fingerprint(user)
    valid_tabs = {k for k, _ in DASHBOARD_TABS}
    valid_cc = {str(c.get("cod")) for c in _admin_cc_list()}
    valid_sources = {k for k, _ in DASHBOARD_SOURCES}
    user["allowed_tabs"] = [t for t in _parse_checked_list("tabs") if t in valid_tabs]
    user["allowed_cc"] = [c for c in _parse_checked_list("cc") if c in valid_cc]
    user["allowed_sources"] = [s for s in _parse_checked_list("sources") if s in valid_sources]
    _save_auth_users(data)
    return redirect("/admin/acessos")


def _set_user_blocked(blocked):
    data = _load_auth_users()
    email = (request.form.get("email") or "").strip().lower()
    current = (session.get("auth_email") or "").strip().lower()
    for user in data.get("users", []):
        if (user.get("email") or "").strip().lower() == email and email != current:
            user["blocked"] = blocked
    _save_auth_users(data)
    return redirect("/admin/acessos")


@app.route("/admin/acessos/bloquear", methods=["POST"])
@_require_admin
@_require_reauth
def admin_acessos_bloquear():
    return _set_user_blocked(True)


@app.route("/admin/acessos/ativar", methods=["POST"])
@_require_admin
@_require_reauth
def admin_acessos_ativar():
    return _set_user_blocked(False)
PPC_PLAN_PATH = os.path.join(DATA_DIR, "PPC JULHO-SETEMBRO_R01.xlsx")
APROP_REL_PATH = os.path.join(DATA_DIR, "Relatório Analítico de Apropriações por Obra.xlsx")
APAGAR_REL_PATH = os.path.join(DATA_DIR, "Relatório de Contas a Pagar.xlsx")
DIST_REL_PATH = os.path.join(DATA_DIR, "Relatório de Distorções das Apropriações por Obra.xlsx")
PAGAS_REL_PATH = os.path.join(DATA_DIR, "Relatório de Contas Pagas.xlsx")
PROJECT_MPP_PATH = os.path.join(DATA_DIR, "CRONOGRAMA_SYNERGIA_BUENO.mpp")
MPP_CRONO_CACHE_PATH = os.path.join(DATA_DIR, "mpp_cronograma_cache.json")
CRONO_HTML_PATH = os.path.join(DATA_DIR, "PLAN_BI_TMZC_SYNERGIA.html")
PA_CRONO_PATH = os.path.join(DATA_DIR, "PA_SYNERGIA_DOCS.xlsm")
PA_MEDICOES_PATH = os.path.join(DATA_DIR, "pa_medicoes.json")
RSO_STATE_PATH = os.path.join(DATA_DIR, "rso_state.json")
RMO_STATE_PATH = os.path.join(DATA_DIR, "rmo_state.json")
RMO_PHOTOS_DIR = os.path.join(DATA_DIR, "rmo_fotos")
RMO_TEMPLATE_PATH = os.path.join(DATA_DIR, "RMO_TEMPLATE.pptx")
CONTRATACOES_PATH = os.path.join(DATA_DIR, "Controle de contratações.xlsx")
SHAREPOINT_STATUS_PATH = os.path.join(DATA_DIR, "sharepoint_sync_status.json")
CNPJ_CACHE_PATH = os.path.join(DATA_DIR, "cnpj_nome_fantasia_cache.json")
CNPJ_FULL_CACHE_PATH = os.path.join(DATA_DIR, "cnpj_full_cache.json")
MINUTA_TEMPLATE_PATH = os.path.join(DATA_DIR, "templates", "minuta_contrato_base.docx")
MINUTA_COUNTER_PATH = os.path.join(DATA_DIR, "minuta_contador.json")
DASH_CACHE_PATH = os.path.join(DATA_DIR, "dashboard_data_cache.json")
DASH_JS_CACHE_PATH = os.path.join(DATA_DIR, "dashboard_data_cache.js")
SINAPI_META_PATH = os.path.join(DATA_DIR, "SINAPI_GO_2026_04_FONTE.json")
SINAPI_TABLE_PATH = os.path.join(DATA_DIR, "SINAPI_GO_2026_04.xlsx")
SIENGE_SYNC_STATUS_PATH = os.path.join(DATA_DIR, "sienge_sync_status.json")
PROTESTOS_PATH = os.path.join(DATA_DIR, "protestos.json")
LOCACOES_STATE_PATH = os.path.join(DATA_DIR, "locacoes_credores.json")


def _load_locacoes_state():
    default = {"incluidos_manual": [], "excluidos_manual": [], "categorias_manual": {}}
    if not os.path.exists(LOCACOES_STATE_PATH):
        return default
    try:
        with open(LOCACOES_STATE_PATH, "r", encoding="utf-8-sig") as fh:
            data = json.load(fh)
        if isinstance(data, dict):
            default["incluidos_manual"] = [str(x) for x in (data.get("incluidos_manual") or [])]
            default["excluidos_manual"] = [str(x) for x in (data.get("excluidos_manual") or [])]
            categorias = data.get("categorias_manual")
            if isinstance(categorias, dict):
                default["categorias_manual"] = {str(k): str(v) for k, v in categorias.items()}
    except Exception:
        pass
    return default


def _save_locacoes_state(state):
    tmp = LOCACOES_STATE_PATH + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(state, fh, ensure_ascii=False, indent=2)
    os.replace(tmp, LOCACOES_STATE_PATH)


_LOCACAO_GRUPO_REGEX = re.compile(r"LOCA", re.IGNORECASE)  # casa "Locação" e "Locação de equipamentos" (desgrins do Sienge)
# Palavras-chave pra sugerir "escoramento metálico" como categoria default de
# um item de locação — tudo que não bater fica em "ferramentas locadas"
# (catch-all). É só o palpite inicial: o usuário pode corrigir item a item,
# e a partir daí a escolha manual (categorias_manual) sempre tem prioridade.
_ESCORAMENTO_REGEX = re.compile(
    r"ESCORA|CIMBRAMENT|PONTALETE|GARFO\s*REGUL|SAPATA\s*REGUL|VIGA\s*MET|"
    r"TRELI[CÇ]A\s*MET|PAINEL\s*(DE\s*)?FORMA|PRANCHA\s*MET|TRAVESSA\s*MET|"
    r"TORRE\s*(DE\s*)?(ESCORA|CIMBRAMENT)|ESCORAS?\s*MET",
    re.IGNORECASE,
)
PROTESTOS_DEFAULT_CNPJ = "45573855000140"
UAU_OBRAS_PATH = os.path.join(DATA_DIR, "uau_obras.json")
UAU_SQL_CONFIG_PATH = os.path.join(DATA_DIR, "uau_sql_config.json")
UAU_EXTERNAL_ENV_PATH = os.path.join(UAU_DATA_DIR, ".env")
_cache = {"data": None, "ts": 0}
CACHE_TTL = 86400  # abertura rápida; o botão Atualizar força o recálculo completo
_sienge_sync = {
    "thread": None,
    "lock": threading.Lock(),
    "code_event": threading.Event(),
    "code": "",
    "confirmation_event": threading.Event(),
    "confirmed": False,
}

def _read_sienge_sync_status():
    default = {
        "state": "idle",
        "message": "Pronto para atualizar os relatórios do Sienge.",
        "running": False,
        "needs_code": False,
        "needs_confirmation": False,
        "needs_password": False,
        "started_at": "",
        "finished_at": "",
        "files": [],
        "error": "",
    }
    try:
        if os.path.exists(SIENGE_SYNC_STATUS_PATH):
            with open(SIENGE_SYNC_STATUS_PATH, "r", encoding="utf-8-sig") as fh:
                saved = json.load(fh)
            if isinstance(saved, dict):
                default.update(saved)
    except Exception:
        pass
    thread = _sienge_sync.get("thread")
    default["running"] = bool(thread and thread.is_alive())
    if default["running"] and default.get("state") == "opening":
        try:
            started = datetime.strptime(str(default.get("started_at") or ""), "%Y-%m-%dT%H:%M:%S")
            if (datetime.now() - started).total_seconds() > 180:
                default.update({
                    "state": "error",
                    "message": "O Chrome automático travou ao abrir. Feche a janela do Sienge se houver e clique em Baixar e atualizar novamente.",
                    "running": False,
                    "needs_code": False,
                    "needs_confirmation": False,
                    "needs_password": False,
                    "error": "Chrome automático travou ao abrir.",
                })
        except Exception:
            pass
    if not default["running"] and default.get("state") in {
        "starting", "opening", "login", "awaiting_code", "awaiting_confirmation", "downloading"
    }:
        default.update({
            "state": "idle",
            "message": "A atualização anterior foi interrompida. Clique em Baixar e atualizar para iniciar novamente.",
            "needs_code": False,
            "needs_confirmation": False,
        })
    default["needs_code"] = default.get("state") == "awaiting_code"
    default["needs_confirmation"] = default.get("state") == "awaiting_confirmation"
    default["needs_password"] = default.get("state") == "password_required"
    return default

def _write_sienge_sync_status(state, message, extra=None):
    current = _read_sienge_sync_status()
    current.update({
        "state": state,
        "message": message,
        "running": state not in {"success", "error", "idle", "password_required"},
        "needs_code": state == "awaiting_code",
        "needs_confirmation": state == "awaiting_confirmation",
        "needs_password": state == "password_required",
    })
    if extra:
        current.update(extra)
    tmp = SIENGE_SYNC_STATUS_PATH + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(current, fh, ensure_ascii=False, indent=2, default=str)
    os.replace(tmp, SIENGE_SYNC_STATUS_PATH)
    return current

def _run_sienge_report_sync(email, password):
    global _cache
    from sienge_sync import run_sync

    _sienge_sync["code"] = ""
    _sienge_sync["code_event"].clear()
    _sienge_sync["confirmed"] = False
    _sienge_sync["confirmation_event"].clear()
    started_at = time.strftime("%Y-%m-%dT%H:%M:%S")
    _write_sienge_sync_status(
        "starting",
        "Iniciando atualização dos relatórios...",
        {"started_at": started_at, "finished_at": "", "files": [], "error": "", "details": ""},
    )

    def status(state, message, extra=None):
        _write_sienge_sync_status(state, message, extra)

    def code_provider():
        if not _sienge_sync["code_event"].wait(timeout=600):
            return ""
        code = str(_sienge_sync.get("code", "") or "").strip()
        _sienge_sync["code"] = ""
        _sienge_sync["code_event"].clear()
        return code

    try:
        result = run_sync(DATA_DIR, email, password, status, code_provider)
        password = ""
        status(
            "recalculating",
            "Arquivos salvos e validados. Recalculando o dashboard...",
            {"files": result.get("files", []), "backup": result.get("backup", "")},
        )
        _cache = {"data": None, "ts": 0}
        _excel_path_cache.update(path="", ts=0)
        data, err = build_data(force=True)
        if err:
            raise RuntimeError("Relatórios atualizados, mas o dashboard não recalculou: " + err)
        _write_sienge_sync_status(
            "success",
            "Relatórios atualizados e dashboard recalculado.",
            {
                "finished_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
                "files": result.get("files", []),
                "backup": result.get("backup", ""),
                "error": "",
            },
        )
    except Exception as exc:
        password = ""
        partial_files = _read_sienge_sync_status().get("files", [])
        error_text = str(exc)
        lowered_error = unicodedata.normalize("NFKD", error_text).encode("ascii", "ignore").decode("ascii").lower()
        credential_error = (
            error_text.startswith(("AUTH_PASSWORD_INVALID:", "AUTH_PASSWORD_REQUIRED:"))
            or ("senha" in lowered_error and "inval" in lowered_error)
            or ("email ou senha" in lowered_error)
            or ("e-mail ou senha" in lowered_error)
            or ("usuario/e-mail" in lowered_error)
            or ("usuario ou senha" in lowered_error)
            or ("auth/login" in lowered_error and "senha" in lowered_error and "entrar" in lowered_error)
        )
        if credential_error:
            clean_error = error_text.split(":", 1)[1].strip()
            if "URL:" in clean_error or len(clean_error) > 180:
                clean_error = "Usuário/e-mail ou senha inválidos. Confira os dados e tente novamente."
            _write_sienge_sync_status(
                "password_required",
                clean_error,
                {
                    "finished_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
                    "error": clean_error,
                    "details": "",
                    "files": partial_files,
                },
            )
            return
        relatorios_existentes = [
            name for name in os.listdir(DATA_DIR)
            if name.lower().startswith("relat") and name.lower().endswith(".xlsx")
        ]
        if partial_files and len(relatorios_existentes) >= 4:
            try:
                _write_sienge_sync_status(
                    "recalculating",
                    "Relatorios disponiveis na pasta. Recalculando o dashboard...",
                    {"files": partial_files, "error": ""},
                )
                _cache = {"data": None, "ts": 0}
                _excel_path_cache.update(path="", ts=0)
                data, err = build_data(force=True)
                if err:
                    raise RuntimeError(err)
                save_dashboard_cache(data)
                _write_sienge_sync_status(
                    "success",
                    "Relatorios salvos na pasta, validados e dashboard recalculado.",
                    {
                        "finished_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
                        "files": partial_files,
                        "error": "",
                        "details": "",
                    },
                )
                return
            except Exception:
                pass
        _write_sienge_sync_status(
            "error",
            (
                f"Atualização parcial: {len(partial_files)} arquivo(s) salvo(s) e validado(s)."
                if partial_files
                else "A atualização não foi concluída. Os arquivos anteriores foram mantidos."
            ),
            {
                "finished_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
                "error": str(exc),
                "details": traceback.format_exc(),
            },
        )

def load_dashboard_cache():
    if not os.path.exists(DASH_CACHE_PATH):
        return None
    try:
        with open(DASH_CACHE_PATH, "r", encoding="utf-8-sig") as fh:
            data = json.load(fh)
        return data if isinstance(data, dict) and data else None
    except Exception:
        return None

def save_dashboard_cache(data):
    try:
        tmp_path = DASH_CACHE_PATH + ".tmp"
        with open(tmp_path, "w", encoding="utf-8") as fh:
            json.dump(data, fh, ensure_ascii=False, default=str, separators=(",", ":"))
        os.replace(tmp_path, DASH_CACHE_PATH)
        js_tmp_path = DASH_JS_CACHE_PATH + ".tmp"
        with open(js_tmp_path, "w", encoding="utf-8") as fh:
            fh.write("const DASH = ")
            json.dump(data, fh, ensure_ascii=False, default=str, separators=(",", ":"))
            fh.write(";")
        os.replace(js_tmp_path, DASH_JS_CACHE_PATH)
    except Exception:
        pass

def load_protestos_state():
    default = {
        "cnpj": PROTESTOS_DEFAULT_CNPJ,
        "consulta": {
            "status": "nao_consultado",
            "message": "Consulta ainda não executada.",
            "executed_at": "",
            "source_url": "https://www.pesquisaprotesto.com.br/",
            "cartorios": [],
        },
        "items": [],
    }
    try:
        if os.path.exists(PROTESTOS_PATH):
            with open(PROTESTOS_PATH, "r", encoding="utf-8-sig") as fh:
                payload = json.load(fh)
            if isinstance(payload, list):
                payload = {"items": payload}
            if isinstance(payload, dict):
                items = payload.get("items", [])
                if isinstance(items, list):
                    by_title = {}
                    for item in items:
                        if isinstance(item, dict) and str(item.get("titulo", "")).strip():
                            item.setdefault("cnpj", payload.get("cnpj") or default["cnpj"])
                            by_title[str(item.get("titulo")).strip()] = item
                    return {
                        "cnpj": str(payload.get("cnpj") or default["cnpj"]).replace(".", "").replace("/", "").replace("-", ""),
                        "consulta": payload.get("consulta") if isinstance(payload.get("consulta"), dict) else default["consulta"],
                        "items": list(by_title.values()),
                        "updated_at": payload.get("updated_at", ""),
                    }
    except Exception:
        pass
    return default

def save_protestos_state(payload):
    items = payload.get("items", []) if isinstance(payload, dict) else []
    cnpj = re.sub(r"\D", "", str(payload.get("cnpj") or PROTESTOS_DEFAULT_CNPJ))
    previous = load_protestos_state()
    clean = []
    for item in items:
        if not isinstance(item, dict):
            continue
        titulo = str(item.get("titulo", "")).strip()
        if not titulo:
            continue
        clean.append({
            "titulo": titulo,
            "cnpj": re.sub(r"\D", "", str(item.get("cnpj") or cnpj)),
            "status": str(item.get("status", "nao_consultado") or "nao_consultado").strip(),
            "cartorio": str(item.get("cartorio", "") or "").strip(),
            "cidade": str(item.get("cidade", "") or "").strip(),
            "uf": str(item.get("uf", "") or "").strip().upper()[:2],
            "consulta_em": str(item.get("consulta_em", "") or "").strip(),
            "observacao": str(item.get("observacao", "") or "").strip(),
        })
    consulta = payload.get("consulta") if isinstance(payload.get("consulta"), dict) else previous.get("consulta", {})
    out = {"cnpj": cnpj, "consulta": consulta, "items": clean, "updated_at": time.strftime("%Y-%m-%dT%H:%M:%S")}
    tmp = PROTESTOS_PATH + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(out, fh, ensure_ascii=False, indent=2)
    os.replace(tmp, PROTESTOS_PATH)
    return out

def consulta_protesto_directd(cnpj_digits):
    """Consulta nacional de protestos via Direct Data (API paga por consumo,
    homologada pelo CENPROT/IEPTB — não depende de hCaptcha).
    Ativa sozinha quando a variável de ambiente DIRECTD_API_TOKEN existir;
    sem o token, retorna None e o chamador cai no fallback público (CENPROT).
    Docs: https://www.directd.com.br/protestos-ieptb
    """
    token = os.environ.get("DIRECTD_API_TOKEN", "").strip()
    if not token:
        return None
    source_url = "https://www.directd.com.br/protestos-ieptb"
    api_url = "https://apiv3.directd.com.br/api/ProtestosOnline"
    executed_at = time.strftime("%Y-%m-%dT%H:%M:%S")
    try:
        req = Request(f"{api_url}?Cnpj={cnpj_digits}&Token={token}", headers={"Accept": "application/json"})
        with urlopen(req, timeout=25) as resp:
            raw = resp.read().decode("utf-8", "ignore")
        payload = json.loads(raw) if raw.strip().startswith(("{", "[")) else {}
        cartorios = []
        for estado in payload.get("protestos") or []:
            if not isinstance(estado, dict):
                continue
            uf = str(estado.get("estado") or estado.get("uf") or "").strip().upper()[:2]
            for c in estado.get("cartorios") or []:
                if not isinstance(c, dict):
                    continue
                nome = c.get("nomeCartorio") or c.get("cartorio") or c.get("nome") or ""
                cidade = c.get("cidade") or c.get("municipio") or ""
                titulos_norm = []
                for t in c.get("titulos") or []:
                    if isinstance(t, dict):
                        numero = (t.get("numeroTitulo") or t.get("titulo") or t.get("protocolo")
                                  or t.get("numero") or "")
                        valor = t.get("valor") or t.get("valorProtesto") or t.get("valorTitulo") or 0
                        data_t = t.get("data") or t.get("dataProtesto") or t.get("dataApresentacao") or ""
                        titulos_norm.append({"numero": str(numero).strip(), "valor": valor, "data": str(data_t or "").strip()})
                if not titulos_norm:
                    # cartório sinalizado sem detalhe de título individual — ainda registra o achado
                    titulos_norm = [{"numero": "", "valor": c.get("valorTotal", 0), "data": ""}]
                cartorios.append({
                    "nome": str(nome or "").strip(),
                    "cidade": str(cidade or "").strip(),
                    "uf": uf,
                    "endereco": "",
                    "telefone": "",
                    "numero_cartorio": str(c.get("numeroCartorio") or "").strip(),
                    "qtd_titulos": len(titulos_norm),
                    "titulos": titulos_norm,
                })
        tem_protesto = bool(payload.get("constamProtestos"))
        status = "consultado" if cartorios else ("sem_protesto" if not tem_protesto else "sem_retorno")
        total = payload.get("numeroTotalProtestos", len(cartorios))
        message = (
            f"Consulta Direct Data (IEPTB/CENPROT) executada. {total} protesto(s) encontrado(s) nacionalmente."
            if tem_protesto else
            "Consulta Direct Data (IEPTB/CENPROT) executada. Nenhum protesto encontrado para este CNPJ."
        )
        return {
            "status": status, "cnpj": cnpj_digits, "message": message,
            "executed_at": executed_at, "source_url": source_url, "api_url": api_url,
            "cartorios": cartorios,
        }
    except HTTPError as exc:
        raw = ""
        try:
            raw = exc.read().decode("utf-8", "ignore")
        except Exception:
            pass
        return {
            "status": "erro", "cnpj": cnpj_digits,
            "message": f"Direct Data retornou HTTP {exc.code}. {raw[:180]}",
            "executed_at": executed_at, "source_url": source_url, "api_url": api_url, "cartorios": [],
        }
    except Exception as exc:
        return {
            "status": "erro", "cnpj": cnpj_digits,
            "message": f"Não foi possível consultar a Direct Data: {exc}",
            "executed_at": executed_at, "source_url": source_url, "api_url": api_url, "cartorios": [],
        }

def consulta_protesto_cnpj(cnpj):
    cnpj_digits = re.sub(r"\D", "", str(cnpj or PROTESTOS_DEFAULT_CNPJ))
    directd = consulta_protesto_directd(cnpj_digits)
    if directd is not None:
        return directd
    source_url = "https://resolve.cenprot.org.br/app"
    api_url = f"https://api.resolve.cenprot.org.br/para-voce/api/protests/v2/research/cenprot/{cnpj_digits}"
    executed_at = time.strftime("%Y-%m-%dT%H:%M:%S")
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) DashboardTemazec/1.0",
        "Accept": "application/json, text/plain, */*",
        "Origin": "https://resolve.cenprot.org.br",
        "Referer": "https://resolve.cenprot.org.br/app",
    }
    try:
        req = Request(api_url, headers=headers)
        with urlopen(req, timeout=20) as resp:
            raw = resp.read().decode("utf-8", "ignore")
            status_code = getattr(resp, "status", 200)
        payload = json.loads(raw) if raw.strip().startswith(("{", "[")) else {}
        cartorios_raw = payload if isinstance(payload, list) else (
            payload.get("cartorios")
            or payload.get("notaries")
            or payload.get("data")
            or payload.get("items")
            or []
        )
        if isinstance(cartorios_raw, dict):
            cartorios_raw = [cartorios_raw]
        cartorios = []
        for c in cartorios_raw if isinstance(cartorios_raw, list) else []:
            if not isinstance(c, dict):
                continue
            nome = c.get("nomeCartorio") or c.get("cartorio") or c.get("nome") or c.get("name") or ""
            cidade = c.get("cidade") or c.get("city") or ""
            uf = c.get("uf") or c.get("estado") or c.get("state") or ""
            titulos_raw = c.get("titulos") or c.get("titles") or []
            titulos_norm = []
            for t in (titulos_raw if isinstance(titulos_raw, list) else []):
                if isinstance(t, dict):
                    numero = (t.get("numeroTitulo") or t.get("numero_titulo") or t.get("titulo")
                              or t.get("numero") or t.get("number") or t.get("protocolo") or "")
                    valor = t.get("valor") or t.get("value") or t.get("valorTitulo") or 0
                    data_t = (t.get("data") or t.get("date") or t.get("dataProtesto")
                              or t.get("dataApresentacao") or "")
                else:
                    numero, valor, data_t = str(t), 0, ""
                numero = str(numero or "").strip()
                if numero:
                    titulos_norm.append({"numero": numero, "valor": valor, "data": str(data_t or "").strip()})
            cartorios.append({
                "nome": str(nome or "").strip(),
                "cidade": str(cidade or "").strip(),
                "uf": str(uf or "").strip().upper()[:2],
                "endereco": str(c.get("endereco") or c.get("address") or "").strip(),
                "telefone": str(c.get("telefone") or c.get("phone") or "").strip(),
                "numero_cartorio": str(c.get("numeroCartorio") or c.get("numero_cartorio") or "").strip(),
                "qtd_titulos": c.get("qtdTitulos") or c.get("qtd_titulos") or len(titulos_norm),
                "titulos": titulos_norm,
            })
        status = "consultado" if cartorios else "sem_retorno"
        message = (
            f"Consulta CENPROT executada. Cartórios retornados: {len(cartorios)}."
            if cartorios else
            "Consulta CENPROT executada, mas nenhum cartório foi retornado para este CNPJ."
        )
        return {
            "status": status,
            "cnpj": cnpj_digits,
            "message": message,
            "executed_at": executed_at,
            "source_url": source_url,
            "api_url": api_url,
            "cartorios": cartorios,
        }
    except HTTPError as exc:
        raw = ""
        try:
            raw = exc.read().decode("utf-8", "ignore")
        except Exception:
            raw = ""
        if exc.code in (401, 403) or "hCaptcha" in raw or "Unauthorized" in raw:
            status = "autenticacao_necessaria"
            message = (
                "A CENPROT/Resolve exige autenticação ou hCaptcha para liberar o nome do cartório. "
                "O dashboard marcou o título, mas o cartório precisa ser confirmado na consulta oficial."
            )
        else:
            status = "erro"
            message = f"API CENPROT retornou HTTP {exc.code}. {raw[:180]}"
    except Exception as exc:
        status = "erro"
        message = f"Não foi possível consultar automaticamente o CENPROT/Resolve: {exc}"
    return {
        "status": status,
        "cnpj": cnpj_digits,
        "message": message,
        "executed_at": executed_at,
        "source_url": source_url,
        "api_url": api_url,
        "cartorios": [],
    }

def dashboard_changed_sources():
    """Retorna somente entradas locais modificadas depois do cache consolidado."""
    if not os.path.exists(DASH_CACHE_PATH):
        return ["cache ausente"]
    cache_mtime = os.path.getmtime(DASH_CACHE_PATH)
    sources = [
        resolve_excel_path(), ORC_REL_PATH, NOMINAL_BUDGET_PATH,
        SUPPORT_REPORT_PATH, PPC_PLAN_PATH, APROP_REL_PATH, APAGAR_REL_PATH,
        DIST_REL_PATH, PAGAS_REL_PATH, PA_CRONO_PATH, PA_MEDICOES_PATH,
        RSO_STATE_PATH, CONTRATACOES_PATH, CREDORES_PATH, SINAPI_TABLE_PATH,
    ]
    changed = []
    for path in dict.fromkeys(p for p in sources if p):
        try:
            if os.path.isfile(path) and os.path.getmtime(path) > cache_mtime + 0.001:
                changed.append(os.path.basename(path))
        except OSError:
            pass
    return changed

def resolve_excel_path():
    """Usa a base local; um caminho configurado explicitamente continua opcional."""
    if _excel_path_cache["path"] and (time.time() - _excel_path_cache["ts"]) < 60:
        return _excel_path_cache["path"]

    configured = os.environ.get("DASH_SHAREPOINT_EXCEL", "").strip()
    try:
        import config
        configured = configured or str(getattr(config, "SP_SYNC_PATH", "") or "").strip()
    except Exception:
        pass
    def is_valid_source(path):
        if not path or not os.path.isfile(path):
            return False
        try:
            from openpyxl import load_workbook
            wb = load_workbook(path, read_only=True, data_only=True)
            names = set(wb.sheetnames)
            wb.close()
            return REQUIRED_EXCEL_SHEETS.issubset(names)
        except Exception:
            return False

    if configured and is_valid_source(configured):
        _excel_path_cache.update(path=configured, ts=time.time())
        return configured

    # O dashboard voltou a operar localmente. Evita varrer todo o OneDrive a
    # cada atualização, que em máquinas com muitos arquivos pode levar minutos.
    if os.path.isfile(LOCAL_EXCEL_PATH):
        _excel_path_cache.update(path=LOCAL_EXCEL_PATH, ts=time.time())
        return LOCAL_EXCEL_PATH

    roots = []
    for key in ("OneDriveCommercial", "OneDrive", "OneDriveConsumer"):
        root = os.environ.get(key, "").strip()
        if root and os.path.isdir(root):
            roots.append(root)
    user_home = os.environ.get("USERPROFILE", "")
    if user_home and os.path.isdir(user_home):
        try:
            roots.extend(
                entry.path for entry in os.scandir(user_home)
                if entry.is_dir() and ("onedrive" in entry.name.lower() or "sharepoint" in entry.name.lower())
            )
        except OSError:
            pass

    found = []
    ignored = {"$recycle.bin", ".git", "node_modules", "__pycache__", "appdata"}
    for root in dict.fromkeys(roots):
        for current, dirs, files in os.walk(root):
            dirs[:] = [d for d in dirs if d.lower() not in ignored and not d.startswith(".")]
            for name in files:
                if name.lower() == EXCEL_FILE_NAME.lower():
                    path = os.path.join(current, name)
                    if os.path.abspath(path).lower() != os.path.abspath(LOCAL_EXCEL_PATH).lower() and is_valid_source(path):
                        found.append(path)
    path = max(found, key=os.path.getmtime) if found else LOCAL_EXCEL_PATH
    _excel_path_cache.update(path=path, ts=time.time())
    return path

def f2(v):
    try: return round(float(v), 2)
    except: return 0.0

def _ascii_upper(s):
    return unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode("ascii").upper()

def _sheet_name_by_ascii(xl, target):
    target = _ascii_upper(target)
    for name in xl.sheet_names:
        if _ascii_upper(name) == target:
            return name
    return None

def _norm_budget_ref(value):
    nums = re.findall(r"\d+", str(value or ""))
    if not nums:
        return ""
    out = []
    for idx, part in enumerate(nums):
        n = int(part)
        out.append(f"{n:02d}" if idx == 0 else f"{n:03d}")
    return ".".join(out)

def _month_key(value):
    dt = pd.to_datetime(value, errors="coerce")
    if pd.isna(dt):
        return ""
    return str(dt.to_period("M"))

def _load_support_physical_sheet(xl, sheet_ascii):
    sheet = _sheet_name_by_ascii(xl, sheet_ascii)
    if not sheet:
        return {}
    raw = xl.parse(sheet, header=None)
    if raw.empty or raw.shape[1] < 8:
        return {}
    header_idx = None
    for idx in raw.index[:15]:
        vals = [_ascii_upper(v) for v in raw.iloc[idx].tolist()]
        if "ITEM" in vals and any("ORCAMENTO" in v for v in vals):
            header_idx = idx
            break
    if header_idx is None:
        header_idx = 4
    date_cols = []
    for col in range(7, raw.shape[1]):
        mk = _month_key(raw.iat[header_idx, col])
        if mk:
            date_cols.append((col, mk))
    rows = {}
    for idx in range(header_idx + 1, len(raw)):
        lvl = pd.to_numeric(pd.Series([raw.iat[idx, 0]]), errors="coerce").iloc[0]
        if pd.isna(lvl) or int(lvl) != 2:
            continue
        cod = _norm_budget_ref(raw.iat[idx, 1])
        if not cod:
            continue
        monthly = {}
        for col, mk in date_cols:
            val = pd.to_numeric(pd.Series([raw.iat[idx, col]]), errors="coerce").fillna(0).iloc[0]
            monthly[mk] = float(val or 0)
        rows[cod] = {
            "cod": cod,
            "item": str(raw.iat[idx, 1] or ""),
            "nm": str(raw.iat[idx, 2] or ""),
            "orcado": f2(raw.iat[idx, 3] if raw.shape[1] > 3 else 0),
            "accum": float(pd.to_numeric(pd.Series([raw.iat[idx, 5] if raw.shape[1] > 5 else 0]), errors="coerce").fillna(0).iloc[0] or 0),
            "monthly": monthly,
        }
    return rows

def _mapa_controle_date(v):
    dt = pd.to_datetime(v, dayfirst=True, errors="coerce")
    return dt.strftime("%d/%m/%Y") if pd.notna(dt) else ""

def build_cred_cnpj_map():
    """codcredor -> CNPJ/CPF normalizado (só dígitos), lido da mesma planilha
    de apoio Credores.xlsx usada por build_cred_map (formato '123 - NOME' na
    1ª coluna, CNPJ/CPF na 3ª). Usado para não cruzar NF/NFSe com título de
    um credor diferente só porque a numeração do documento colidiu."""
    cnpj_map = {}
    if not os.path.exists(CREDORES_PATH):
        return cnpj_map
    try:
        df_cred = pd.read_excel(CREDORES_PATH, sheet_name=0, header=None)
        for _, row in df_cred.iterrows():
            val = row.iloc[0]
            if pd.isna(val):
                continue
            m = re.match(r'^(\d+)\s*-\s*(.*)', str(val).strip())
            if not m:
                continue
            cred_id = int(m.group(1))
            cnpj_raw = row.iloc[2] if len(row) > 2 else None
            if pd.isna(cnpj_raw):
                continue
            cnpj_digits = re.sub(r"\D", "", str(cnpj_raw))
            if cnpj_digits:
                cnpj_map[cred_id] = cnpj_digits
    except Exception:
        pass
    return cnpj_map

def _nf_titulo_candidatos():
    """Como _nf_titulo_lookup, mas preserva TODOS os títulos candidatos por
    número de documento (não só o primeiro encontrado), com o CNPJ do credor
    de cada um. Necessário porque fornecedores diferentes podem usar a mesma
    numeração de NF/NFSe por coincidência (ex.: NFSe nº 54 emitida tanto pelo
    Elvis quanto pela PHD Fundações) — casar só pelo número gera falso
    positivo de 'lançado no Sienge'. Quem consome isso deve confirmar o
    candidato certo comparando o CNPJ do emitente da nota."""
    candidatos = {}
    try:
        excel_path = resolve_excel_path()
        if not os.path.exists(excel_path):
            return candidatos
        xl = pd.ExcelFile(excel_path, engine="openpyxl")
        cnpj_map = build_cred_cnpj_map()
        bancos = xl.parse("dfin_bancos") if "dfin_bancos" in xl.sheet_names else pd.DataFrame(columns=["nmcredor", "codcredor"])
        nome_map = build_cred_map(bancos, xl)

        def _add(key, numtitcp, codcredor):
            if not key:
                return
            cnpj = cnpj_map.get(codcredor, "") if codcredor is not None else ""
            nome = nome_map.get(codcredor, "") if codcredor is not None else ""
            candidatos.setdefault(key, []).append(
                {"numtitcp": numtitcp, "codcredor": codcredor, "cnpj": cnpj, "nome": nome}
            )

        def _add_nome(key, numtitcp, nome_livre):
            """Como _add, mas pra fontes (relatórios manuais) que só trazem o
            nome do credor em texto livre, sem código de credor do Sienge —
            _casar_titulo já sabe desempatar por similaridade de nome quando
            não há CNPJ."""
            if not key or not numtitcp:
                return
            candidatos.setdefault(key, []).append(
                {"numtitcp": numtitcp, "codcredor": None, "cnpj": "", "nome": str(nome_livre or "").strip()}
            )

        _NF_DOC_RE = re.compile(r"^\s*(?:NFE|NFSE)[./]?\s*(\d+)", re.IGNORECASE)

        def _titulo_base(v):
            m = re.match(r"^\s*(\d+)", str(v or ""))
            return m.group(1) if m else ""

        # Relatório de Contas Pagas: linhas de lançamento têm credor (col 0),
        # código do credor (col 3), documento (col 5, ex. 'NFSE.45') e
        # título/parcela (col 8, ex. '137372/1') — cobre pagamentos já
        # baixados que não aparecem em fin_cpagar_aprobra/dfin_cpagar (ex.:
        # relatório exportado antes do bd_sienge_tmz mais recente).
        if os.path.exists(PAGAS_REL_PATH):
            try:
                raw_pagas = pd.read_excel(PAGAS_REL_PATH, sheet_name="Relatório", header=None)
                for _, row in raw_pagas.iterrows():
                    doc = row.get(5)
                    m = _NF_DOC_RE.match(str(doc)) if pd.notna(doc) else None
                    if not m or pd.isna(row.get(8)):
                        continue
                    key = m.group(1).lstrip("0")
                    numtitcp = _titulo_base(row.get(8))
                    if not key or not numtitcp:
                        continue
                    try:
                        codcredor = int(float(row.get(3))) if pd.notna(row.get(3)) else None
                    except Exception:
                        codcredor = None
                    if codcredor is not None:
                        _add(key, numtitcp, codcredor)
                    else:
                        _add_nome(key, numtitcp, row.get(0))
            except Exception:
                pass

        # Relatório de Contas a Pagar: mesmo espírito, mas cobre títulos em
        # aberto (ainda não baixados) — formato 'compacto' (<=20 colunas) ou
        # 'largo' (>20), documento/título em posições diferentes conforme o
        # Sienge exporta.
        if os.path.exists(APAGAR_REL_PATH):
            try:
                raw_apagar = pd.read_excel(APAGAR_REL_PATH, sheet_name="Relatório", header=None)
                compact = raw_apagar.shape[1] <= 20
                doc_idx, tit_idx = (1, 2) if compact else (15, 25)
                for _, row in raw_apagar.iterrows():
                    doc = row.get(doc_idx)
                    m = _NF_DOC_RE.match(str(doc)) if pd.notna(doc) else None
                    if not m or pd.isna(row.get(tit_idx)):
                        continue
                    key = m.group(1).lstrip("0")
                    numtitcp = _titulo_base(row.get(tit_idx))
                    if not key or not numtitcp:
                        continue
                    _add_nome(key, numtitcp, row.get(0))
            except Exception:
                pass

        # Relatório Analítico de Apropriações por Obra: Documento (ex.
        # 'NFSE/45') + Título/Parcela (ex. '137372/1') + Credor/Histórico —
        # cobre casos como MEI/pessoa física apropriados só nesse relatório
        # manual, sem rastro nenhum em fsup_ct/fsup_nf_ped/dfin_cpagar (ex.:
        # Wanderson de Lima Pinto, título 137372).
        if os.path.exists(APROP_REL_PATH):
            try:
                aprop_cand = pd.read_excel(APROP_REL_PATH, sheet_name="Relatório", header=1)
                if "Documento" not in aprop_cand.columns or "Título/Parcela" not in aprop_cand.columns:
                    aprop_cand = _parse_aprop_hierarquico(APROP_REL_PATH)
                if {"Documento", "Título/Parcela"}.issubset(aprop_cand.columns):
                    for _, row in aprop_cand.iterrows():
                        doc = row.get("Documento")
                        m = _NF_DOC_RE.match(str(doc)) if pd.notna(doc) else None
                        if not m or pd.isna(row.get("Título/Parcela")):
                            continue
                        key = m.group(1).lstrip("0")
                        numtitcp = _titulo_base(row.get("Título/Parcela"))
                        if not key or not numtitcp:
                            continue
                        _add_nome(key, numtitcp, row.get("Credor/Histórico"))
            except Exception:
                pass

        if "fin_cpagar_aprobra" in xl.sheet_names:
            tit_doc = xl.parse("fin_cpagar_aprobra")
            if {"nudoc", "numtitcp", "codcredor"}.issubset(tit_doc.columns):
                sub = tit_doc[["nudoc", "numtitcp", "codcredor"]].dropna(subset=["nudoc", "numtitcp"])
                for _, r in sub.iterrows():
                    key = re.sub(r"\D", "", str(r["nudoc"])).lstrip("0")
                    try:
                        codcredor = int(float(r["codcredor"])) if pd.notna(r["codcredor"]) else None
                    except Exception:
                        codcredor = None
                    _add(key, str(int(r["numtitcp"])), codcredor)

        if "fsup_nf_ped" in xl.sheet_names:
            nf_ped = xl.parse("fsup_nf_ped")
            nf_col = next((c for c in [
                "numnf", "numdocfis", "coddocfis", "numdoc", "codnotafis", "nfdoc", "numdocumento",
                "numdocnf", "coddoc", "numdoccp", "nfdocumento", "numeronf", "numero_nf", "numdocumnf",
                "numdocorigem", "coddocorigem", "numnotafiscal", "coddocfiscal",
            ] if c in nf_ped.columns), None)
            cred_col = next((c for c in ["codcredornf", "codcredor"] if c in nf_ped.columns), None)
            if nf_col and "numtitcp" in nf_ped.columns:
                cols = [nf_col, "numtitcp"] + ([cred_col] if cred_col else [])
                sub = nf_ped[cols].dropna(subset=[nf_col, "numtitcp"])
                for _, r in sub.iterrows():
                    key = re.sub(r"\D", "", str(r[nf_col])).lstrip("0")
                    codcredor = None
                    if cred_col and pd.notna(r.get(cred_col)):
                        try:
                            codcredor = int(float(r[cred_col]))
                        except Exception:
                            codcredor = None
                    _add(key, str(int(r["numtitcp"])), codcredor)

        # dfin_cpagar: aba "crua" de contas a pagar, com cobertura mais ampla
        # de centro de custo que fin_cpagar_aprobra (ex.: títulos de despesas
        # administrativas recorrentes tipo consultoria/assessoria, que não
        # têm apropriação de EAP de obra e por isso não aparecem em
        # fin_cpagar_aprobra) — sem ela, notas desses fornecedores ficam
        # mostrando "Não lançada" mesmo já lançadas e pagas no Sienge.
        if "dfin_cpagar" in xl.sheet_names:
            dfc = xl.parse("dfin_cpagar")
            if {"numdocum", "numtitcp", "codcredor"}.issubset(dfc.columns):
                sub = dfc[["numdocum", "numtitcp", "codcredor"]].dropna(subset=["numdocum", "numtitcp"])
                for _, r in sub.iterrows():
                    key = re.sub(r"\D", "", str(r["numdocum"])).lstrip("0")
                    try:
                        codcredor = int(float(r["codcredor"])) if pd.notna(r["codcredor"]) else None
                    except Exception:
                        codcredor = None
                    _add(key, str(int(r["numtitcp"])), codcredor)
    except Exception:
        pass
    return candidatos

_CAU_OBSTIT_RE = re.compile(
    r"medi[cç][aã]o\s*N?[o°]?\.?\s*(\d+).*?contrato\s+(\S+)\s*/\s*(\d+).*?"
    r"documento\(s\)\s*origem:\s*([A-Z0-9/]+).*?"
    r"valor base para c[aá]lculo da cau[cç][aã]o:\s*([\d.,]+)",
    re.IGNORECASE | re.DOTALL,
)

def _load_caucoes_retencao():
    """Retenção CONTRATUAL (caução) descontada da medição do empreiteiro —
    documento tipo 'CAU' em dfin_cpagar. Isso é diferente de imposto retido
    na fonte (ISS/INSS/IRRF/PIS/COFINS/CSLL, tratado em outro lugar): aqui é
    o percentual retido da própria medição do contrato como garantia,
    normalmente liberado só no fechamento/encerramento do contrato. Só
    existe em dfin_cpagar — fin_cpagar_aprobra não tem esse tipo de
    documento. O texto de obstit já vem estruturado pelo Sienge com
    medição/contrato/documento de origem/valor base, então dá pra extrair
    tudo sem depender de nenhum cadastro manual."""
    out = []
    try:
        excel_path = resolve_excel_path()
        if not os.path.exists(excel_path):
            return out
        xl = pd.ExcelFile(excel_path, engine="openpyxl")
        if "dfin_cpagar" not in xl.sheet_names:
            return out
        df = xl.parse("dfin_cpagar")
        if "coddocum" not in df.columns:
            return out
        sub = df[df["coddocum"].astype(str).str.strip().str.upper() == "CAU"]
        if sub.empty:
            return out
        bancos = xl.parse("dfin_bancos") if "dfin_bancos" in xl.sheet_names else pd.DataFrame(columns=["nmcredor", "codcredor"])
        nome_map = build_cred_map(bancos, xl)
        for _, r in sub.iterrows():
            obstit = str(r.get("obstit") or "")
            m = _CAU_OBSTIT_RE.search(obstit)
            try:
                codcredor = int(float(r.get("codcredor"))) if pd.notna(r.get("codcredor")) else None
            except Exception:
                codcredor = None
            valor_caucao = f2(r.get("vlrparcorig"))
            valor_liberado = f2(r.get("vlrliqbaixacp"))
            out.append({
                "numtitcp": str(int(r["numtitcp"])) if pd.notna(r.get("numtitcp")) else "",
                "credor": nome_map.get(codcredor, "") or "—",
                "medicao": m.group(1) if m else "",
                "contrato": f"{m.group(2)}/{m.group(3)}" if m else "",
                "documento_origem": m.group(4) if m else "",
                "valor_base_medicao": f2(m.group(5).replace(".", "").replace(",", ".")) if m else 0,
                "valor_caucao": valor_caucao,
                "valor_liberado": valor_liberado,
                "valor_a_liberar": f2(valor_caucao - valor_liberado),
                "status": str(r.get("statusparc") or "").strip(),
                "dtemissao": _fmt_date(r.get("dtemissao")),
                "obstit": obstit,
            })
        out.sort(key=lambda x: x["valor_a_liberar"], reverse=True)
    except Exception:
        pass
    return out

def _load_orcado_administracao_preliminar():
    """Orçado 'oficial' de mão de obra própria/administração, direto do
    orçamento original do projeto — aba 'ORÇAMENTO PRELIMINAR' da planilha
    de apoio, item '1.4.2 EQUIPE TÉCNICA E ADMINISTRATIVA' (Engenheiro,
    Mestre de Obras, Administrativo da Obra, Técnico de Segurança etc.).
    Essa é a fonte única pedida pelo usuário pra "orçado" de administração,
    usada tanto no gráfico de custo de folha (Painel BI) quanto nas árvores
    de Orçamento API e Orçamento Relatório — antes cada uma vinha de um
    lugar diferente (eng_orc no Sienge, FIN - AS - NOMINAL na planilha de
    apoio, ou nada). A coluna usada é 'Preço Total > Total' (cabeçalho em
    duas linhas na planilha) — não confundir com as colunas seguintes
    (também rotuladas 'Total' pelo Excel por serem mescladas, mas que
    guardam % do orçamento e R$/m², não o valor em reais da linha."""
    out = {"ok": False, "valor": 0.0, "fonte": "ORÇAMENTO PRELIMINAR › 1.4.2 Equipe Técnica e Administrativa"}
    try:
        df = pd.read_excel(SUPPORT_REPORT_PATH, sheet_name="ORÇAMENTO PRELIMINAR", header=[8, 9], engine="openpyxl")
        desc_col = next((c for c in df.columns if str(c[0]).strip() == "Descrição"), None)
        total_col = next((c for c in df.columns if c[0] == "Preço Total" and c[1] == "Total"), None)
        if desc_col is None or total_col is None:
            return out
        desc_norm = df[desc_col].map(_ascii_upper)
        mask = desc_norm.str.contains("EQUIPE TECNICA E ADMINISTRATIVA", na=False)
        rows = df.loc[mask, total_col]
        if not rows.empty:
            valor = f2(pd.to_numeric(rows.iloc[0], errors="coerce"))
            if valor:
                out = {"ok": True, "valor": valor, "fonte": out["fonte"]}
    except Exception:
        pass
    return out


def _previsto_termino_folha(xl, histograma_equipe):
    """Estima o custo previsto ao término da folha própria (mão de obra),
    combinando: 1) custo REAL já pago por competência (títulos FOP/FOLHA do
    Sienge, obra 1016) e 2) para os meses sem dado real ainda, uma projeção
    = efetivo planejado do mês (histograma de equipe, planilha de apoio) x
    custo médio por pessoa observado nos meses onde efetivo real e custo
    real coincidem. Metodologia tipo EAC (Estimate At Completion) de gestão
    de projetos: realizado + tendência observada aplicada ao planejado
    restante. É uma ESTIMATIVA, não um valor contratual/orçado — a
    planilha de apoio (HISTOGRAMA ADM) não traz o custo por mês, só o
    efetivo (nº de pessoas), então o custo precisa ser inferido a partir do
    que já foi realmente pago."""
    out = {"ok": False, "meses": [], "custo_medio_pessoa": 0, "total_real": 0, "total_projetado": 0, "previsto_termino": 0, "meses_com_real": 0}
    try:
        df = xl.parse("dfin_cpagar")
        if not {"coddocum", "codccusto", "numdocum", "vlrliqbaixacp"}.issubset(df.columns):
            return out
        sub = df[(df["coddocum"].astype(str).str.strip().str.upper() == "FOP") & (pd.to_numeric(df["codccusto"], errors="coerce") == 1016)]
        custo_real_por_mes = {}
        qtd_titulos_por_mes = {}
        for _, r in sub.iterrows():
            m = re.search(r"(\d{2})\.(\d{4})", str(r.get("numdocum") or ""))
            if not m:
                continue
            key = f"{m.group(2)}-{m.group(1)}"
            custo_real_por_mes[key] = custo_real_por_mes.get(key, 0.0) + f2(r.get("vlrliqbaixacp"))
            qtd_titulos_por_mes[key] = qtd_titulos_por_mes.get(key, 0) + 1

        # Fechamento de folha de um mês costuma chegar aos poucos no Sienge —
        # o mês mais recente pode estar com poucos títulos lançados ainda
        # (bem abaixo do mês anterior), sinal de que ainda não é o valor
        # final daquele mês. Só verifica o(s) último(s) mês(es) da série
        # (não os primeiros meses de ramp-up, que são legitimamente
        # pequenos) — compara cada mês só com o mês imediatamente anterior,
        # não com o pico da série inteira.
        meses_incompletos = set()
        meses_ordenados = sorted(qtd_titulos_por_mes.keys())
        if len(meses_ordenados) >= 2:
            ultimo, penultimo = meses_ordenados[-1], meses_ordenados[-2]
            if qtd_titulos_por_mes[penultimo] and qtd_titulos_por_mes[ultimo] < qtd_titulos_por_mes[penultimo] * 0.4:
                meses_incompletos.add(ultimo)

        hist_by_mes = {r.get("mes"): r for r in (histograma_equipe or []) if r.get("mes")}

        # Orçado (planejado) oficial de mão de obra própria/administração —
        # vem do orçamento original do projeto (ver
        # _load_orcado_administracao_preliminar), mesma fonte usada nas
        # árvores de Orçamento API/Relatório pro nó 01.004.001.
        orcado_folha = _load_orcado_administracao_preliminar().get("valor", 0.0)

        soma_planejado_headcount = sum(f2((h or {}).get("planejado")) for h in hist_by_mes.values())

        razoes = []
        for mes, custo in custo_real_por_mes.items():
            if mes in meses_incompletos:
                continue
            efetivo_real = f2((hist_by_mes.get(mes) or {}).get("real"))
            if efetivo_real > 0:
                razoes.append(custo / efetivo_real)
        custo_medio_pessoa = f2(sum(razoes) / len(razoes)) if razoes else 0.0

        meses_out = []
        total_real = 0.0
        total_projetado = 0.0
        for mes in sorted(hist_by_mes.keys()):
            h = hist_by_mes[mes]
            # Planejado distribuído pelo mesmo formato da curva de efetivo
            # planejado (histograma) — o orçamento (EAP SIENGE) só tem o
            # total do projeto, não por mês, então distribui proporcional ao
            # ritmo de efetivo planejado pra virar uma curva comparável.
            peso_mes = (f2(h.get("planejado")) / soma_planejado_headcount) if soma_planejado_headcount else 0
            planejado_mes = f2(orcado_folha * peso_mes)
            custo_real = custo_real_por_mes.get(mes) if mes not in meses_incompletos else None
            if custo_real:
                meses_out.append({"mes": mes, "custo": f2(custo_real), "tipo": "real", "efetivo": h.get("real", 0), "planejado": planejado_mes})
                total_real += custo_real
            else:
                projetado = f2((h.get("planejado") or 0) * custo_medio_pessoa)
                meses_out.append({"mes": mes, "custo": projetado, "tipo": "projetado", "efetivo": h.get("planejado", 0), "planejado": planejado_mes})
                total_projetado += projetado

        out.update({
            "ok": bool(razoes),
            "meses": meses_out,
            "custo_medio_pessoa": custo_medio_pessoa,
            "total_real": f2(total_real),
            "total_projetado": f2(total_projetado),
            "previsto_termino": f2(total_real + total_projetado),
            "orcado_folha": orcado_folha,
            "meses_com_real": len(razoes),
        })
    except Exception:
        pass
    return out

def _norm_credor_nome(s):
    """Normaliza nome de credor pra comparação tolerante: maiúsculas, sem
    acento, sem razão social genérica (LTDA/ME/EIRELI/S.A./...), só
    alfanumérico — mesma lógica de notas_fiscais._norm_nome, duplicada aqui
    pra não acoplar os dois módulos por uma função de 5 linhas."""
    s = unicodedata.normalize("NFKD", str(s or "").upper()).encode("ascii", "ignore").decode("ascii")
    s = re.sub(r"[^A-Z0-9 ]", " ", s)
    s = re.sub(r"\b(LTDA|EPP|EIRELI|ME|MEI|S/?A|SA|CIA|COMERCIO|COM|SERVICOS|SERVICO|"
               r"INDUSTRIA|CONSTRUCOES|ENGENHARIA)\b", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def _credores_obra_1016():
    """Fornecedores vinculados ao centro de custo 1016 (a obra), por 3 fontes:
    contrato (fsup_ct) e pedido de compra (fsup_nf_ped) do bd_sienge_tmz —
    que dão o CNPJ completo via codcredor — e o 'Relatório Analítico de
    Apropriações por Obra' (relatório manual, já filtrado pra obra 1016)
    — que só tem o nome do credor em texto livre (às vezes com um prefixo
    numérico de CNPJ/CPF, no caso de pessoa física/MEI). Usado pra sinalizar,
    no Controle de Notas Fiscais, notas de fornecedores que já trabalham pra
    obra mas foram apropriadas em outro centro de custo (ou ainda nem
    lançadas no bd_sienge_tmz) — casos como um MEI com apropriação só nesse
    relatório manual, sem nenhum rastro em fsup_ct/fsup_nf_ped."""
    out = {"cnpjs": set(), "cnpj_roots": set(), "nomes": set()}
    try:
        excel_path = resolve_excel_path()
        if os.path.exists(excel_path):
            xl = pd.ExcelFile(excel_path, engine="openpyxl")
            cnpj_map = build_cred_cnpj_map()

            if "fsup_ct" in xl.sheet_names:
                ct = xl.parse("fsup_ct")
                if {"codccusto", "codcredorct"}.issubset(ct.columns):
                    ct1016 = ct[pd.to_numeric(ct["codccusto"], errors="coerce") == 1016]
                    for cod in ct1016["codcredorct"].dropna().unique():
                        try:
                            cnpj = cnpj_map.get(int(float(cod)))
                        except Exception:
                            cnpj = None
                        if cnpj:
                            out["cnpjs"].add(cnpj)

            if "fsup_nf_ped" in xl.sheet_names:
                nfp = xl.parse("fsup_nf_ped")
                if {"codccusto", "codcredornf"}.issubset(nfp.columns):
                    nfp1016 = nfp[pd.to_numeric(nfp["codccusto"], errors="coerce") == 1016]
                    for cod in nfp1016["codcredornf"].dropna().unique():
                        try:
                            cnpj = cnpj_map.get(int(float(cod)))
                        except Exception:
                            cnpj = None
                        if cnpj:
                            out["cnpjs"].add(cnpj)
    except Exception:
        pass

    try:
        if os.path.exists(APROP_REL_PATH):
            aprop = pd.read_excel(APROP_REL_PATH, sheet_name="Relatório", header=1)
            if "Credor/Histórico" not in aprop.columns:
                aprop = _parse_aprop_hierarquico(APROP_REL_PATH)
            if "Credor/Histórico" in aprop.columns:
                if "Obra" in aprop.columns:
                    aprop = aprop[aprop["Obra"].astype(str).str.strip().str.startswith("1016")]
                for credor in aprop["Credor/Histórico"].dropna().astype(str):
                    m = re.match(r"^\s*([\d.\-/]{6,})\s+(.+)$", credor)
                    if m:
                        raiz = re.sub(r"\D", "", m.group(1))[:8]
                        if len(raiz) == 8:
                            out["cnpj_roots"].add(raiz)
                        nome = _norm_credor_nome(m.group(2))
                    else:
                        nome = _norm_credor_nome(credor)
                    if nome:
                        out["nomes"].add(nome)
    except Exception:
        pass
    return out


def _eh_possivel_custo_obra(cnpj_fornecedor, nome_fornecedor, credores_obra):
    cnpj = re.sub(r"\D", "", str(cnpj_fornecedor or ""))
    if cnpj and cnpj in credores_obra.get("cnpjs", set()):
        return True
    if cnpj and cnpj[:8] in credores_obra.get("cnpj_roots", set()):
        return True
    nome = _norm_credor_nome(nome_fornecedor)
    if nome and nome in credores_obra.get("nomes", set()):
        return True
    return False


def _nf_titulo_ccusto_apropriacao():
    """numtitcp -> {centro_custo, apropriacao}, pra exibir no Controle de
    Notas Fiscais quando a nota já está lançada no Sienge. 'Apropriação' é
    o nível 1 da EAP (macro-serviço/etapa da obra) ao qual o título foi
    apropriado — reaproveita o que build_data() já calcula (cacheado)."""
    out = {}
    try:
        d, err = build_data()
        if err or not d:
            return out
        cc_nm = {str(r.get("cod")): r.get("nm", "") for r in (d.get("tit_cc") or [])}
        n1_nm = {str(r.get("cod")): r.get("nm", "") for r in (d.get("tit_n1") or [])}
        for r in d.get("titulos_pag") or []:
            numtit = r.get("numtit")
            if not numtit:
                continue
            cc_cod = r.get("codccusto") or ""
            n1_cod = r.get("n1_cod") or ""
            out[str(numtit)] = {
                "centro_custo": cc_nm.get(cc_cod, f"CC {cc_cod}") if cc_cod else "",
                "apropriacao": n1_nm.get(n1_cod, "") if n1_cod else "",
            }
    except Exception:
        pass
    return out

def _nf_titulo_lookup():
    """Mapa normalizado 'número da NF/NFSe' -> número do título financeiro
    gerado, para cruzar com o Mapa de Controle e com o Controle de Notas
    Fiscais. Duas fontes, a mais confiável primeiro:
      1) fin_cpagar_aprobra.nudoc — número do documento gravado direto no
         título, cobre QUALQUER lançamento (com ou sem pedido de compra;
         é o que resolve casos como SENAI, que não passa por pedido).
      2) fsup_nf_ped — vínculo por pedido de compra, cobre casos onde o
         nudoc não bate mas o pedido liga a NF ao título."""
    lookup = {}
    try:
        excel_path = resolve_excel_path()
        if not os.path.exists(excel_path):
            return lookup
        xl = pd.ExcelFile(excel_path, engine="openpyxl")

        if "fin_cpagar_aprobra" in xl.sheet_names:
            tit_doc = xl.parse("fin_cpagar_aprobra")
            if "nudoc" in tit_doc.columns and "numtitcp" in tit_doc.columns:
                sub = tit_doc[["nudoc", "numtitcp"]].dropna()
                for _, r in sub.iterrows():
                    key = re.sub(r"\D", "", str(r["nudoc"])).lstrip("0")
                    if not key:
                        continue
                    lookup.setdefault(key, str(int(r["numtitcp"])))

        if "fsup_nf_ped" in xl.sheet_names:
            nf_ped = xl.parse("fsup_nf_ped")
            nf_col = next((c for c in [
                "numnf", "numdocfis", "coddocfis", "numdoc", "codnotafis", "nfdoc", "numdocumento",
                "numdocnf", "coddoc", "numdoccp", "nfdocumento", "numeronf", "numero_nf", "numdocumnf",
                "numdocorigem", "coddocorigem", "numnotafiscal", "coddocfiscal",
            ] if c in nf_ped.columns), None)
            if nf_col and "numtitcp" in nf_ped.columns:
                sub = nf_ped[[nf_col, "numtitcp"]].dropna()
                for _, r in sub.iterrows():
                    key = re.sub(r"\D", "", str(r[nf_col])).lstrip("0")
                    if not key:
                        continue
                    lookup.setdefault(key, str(int(r["numtitcp"])))
    except Exception:
        pass
    return lookup

def _load_mapa_controle():
    """Lê 'Relatório de Mapa de Controle.xlsx': acompanhamento físico
    item a item da cadeia Solicitação de Compra -> Pedido de Compra -> NF/
    entrega na obra. Não tem valores financeiros (é controle de quantidade
    e prazo), por isso os KPIs são de itens/quantidade, não R$."""
    out = {"ok": False, "error": "", "source": "", "rows": [], "summary": {}}
    if not os.path.exists(MAPA_CONTROLE_PATH):
        out["error"] = "Arquivo 'Relatório de Mapa de Controle.xlsx' não encontrado na pasta dados."
        return out
    try:
        xl = pd.ExcelFile(MAPA_CONTROLE_PATH, engine="openpyxl")
        sheet = xl.sheet_names[0]
        raw = xl.parse(sheet, header=None)
        hdr_row = None
        for i in range(min(20, len(raw))):
            vals = [str(v).strip() for v in raw.iloc[i].tolist() if pd.notna(v)]
            if "Item" in vals and any("SC" in v for v in vals):
                hdr_row = i
                break
        if hdr_row is None:
            out["error"] = "Não foi possível localizar o cabeçalho da planilha."
            return out
        df = xl.parse(sheet, header=hdr_row)
        df.columns = [str(c).strip() for c in df.columns]
        df["Item"] = pd.to_numeric(df["Item"], errors="coerce")
        df = df[df["Item"].notna()].copy()

        today = pd.Timestamp.now().normalize()
        nf_titulo_map = _nf_titulo_lookup()

        # apropriação por título (item de orçamento, centro de custo, valores) —
        # vem do cache principal do dashboard (já processado em build_data);
        # o Mapa de Controle em si não tem essa informação (é só quantidade/
        # prazo). Quando a linha não tem "Título Gerado" via NF, cai pro
        # vínculo direto PC→título (pc_titulo_map, fsup_nf_ped) como reforço.
        _dash_cache = load_dashboard_cache() or {}
        _titulo_aprop_map = {}
        for _t in (_dash_cache.get("titulos_pag") or []):
            _titulo_aprop_map.setdefault(_t.get("numtit", ""), []).append(_t)
        _pc_titulo_map = _dash_cache.get("pc_titulo_map") or {}

        def _apropriacoes_da_linha(titulo_gerado, pc_str):
            numtits = []
            if titulo_gerado and titulo_gerado != "—":
                numtits.append(titulo_gerado)
            if pc_str and pc_str != "—":
                numtits.extend(n for n in _pc_titulo_map.get(pc_str, []) if n not in numtits)
            linhas = []
            for numtit in numtits:
                for t in _titulo_aprop_map.get(numtit, []):
                    linhas.append({
                        "numtit": numtit,
                        "apropriacao": t.get("apropriacao") or "",
                        "refitemorc": t.get("refitemorc") or "",
                        "codccusto": t.get("codccusto") or "",
                        "vlr_ap": t.get("vlr_ap") or 0,
                        "vlr_pg": t.get("vlr_pg") or 0,
                        "status": t.get("status") or "",
                    })
            return linhas

        rows = []
        for _, r in df.iterrows():
            qtd_sol = f2(r.get("Qt. solicitada")) if pd.notna(r.get("Qt. solicitada")) else 0.0
            qtd_ent = f2(r.get("Quant. entregue")) if pd.notna(r.get("Quant. entregue")) else 0.0
            pc = r.get("Nº do PC")
            tem_pc = pd.notna(pc)
            nf = r.get("Nº da NF")
            tem_nf = pd.notna(nf) and str(nf).strip() not in ("", "nan")
            titulo_gerado = ""
            if tem_nf:
                nf_key = re.sub(r"\D", "", str(nf)).lstrip("0")
                titulo_gerado = nf_titulo_map.get(nf_key, "")
            prev_entrega_raw = r.get("Previsão de entrega")
            prev_entrega_dt = pd.to_datetime(prev_entrega_raw, dayfirst=True, errors="coerce")
            entregue = qtd_ent > 0 and qtd_ent >= qtd_sol
            if entregue:
                status = "Entregue"
            elif not tem_pc:
                status = "Aguardando PC"
            elif pd.notna(prev_entrega_dt) and prev_entrega_dt < today:
                status = "Atrasado"
            else:
                status = "Aguardando entrega"
            pc_str = str(int(pc)) if tem_pc else "—"
            apropriacoes = _apropriacoes_da_linha(titulo_gerado, pc_str)
            # Prazo de entrada: pedido é considerado aprovado na data de emissão
            # do PC; a partir daí, até 2 dias para dar entrada na obra é normal,
            # acima disso é atraso de entrada (mesmo que ainda não tenha entrado).
            data_pc_dt = pd.to_datetime(r.get("Data emissão do PC"), dayfirst=True, errors="coerce")
            data_entrada_dt = pd.to_datetime(r.get("Data entrada na obra"), dayfirst=True, errors="coerce")
            atraso_entrada_dias = None
            atraso_entrada_status = ""
            if tem_pc and pd.notna(data_pc_dt):
                ref_dt = data_entrada_dt if pd.notna(data_entrada_dt) else today
                atraso_entrada_dias = int((ref_dt - data_pc_dt).days)
                atraso_entrada_status = "Normal" if atraso_entrada_dias <= 2 else "Atraso de entrada"
            rows.append({
                "item": int(r["Item"]),
                "sc": str(int(r["Nº da SC"])) if pd.notna(r.get("Nº da SC")) else "—",
                "data_sc": _mapa_controle_date(r.get("Data da SC")),
                "insumo_cod": str(int(r["Cód. insumo"])) if pd.notna(r.get("Cód. insumo")) else "—",
                "insumo": str(r.get("Descrição do insumo") or "").strip(),
                "qtd_solicitada": qtd_sol,
                "pc": pc_str,
                "data_pc": _mapa_controle_date(r.get("Data emissão do PC")),
                "previsao_entrega": _mapa_controle_date(prev_entrega_raw),
                "qtd_entregue": qtd_ent,
                "saldo": f2(r.get("Saldo")) if pd.notna(r.get("Saldo")) else f2(qtd_sol - qtd_ent),
                "nf": str(nf).strip() if tem_nf else "—",
                "titulo_gerado": titulo_gerado or "—",
                "data_nf": _mapa_controle_date(r.get("Data da NF")),
                "data_entrada_obra": _mapa_controle_date(r.get("Data entrada na obra")),
                "data_vencimento": _mapa_controle_date(r.get("Data vencimento")),
                "status": status,
                "apropriacoes": apropriacoes,
                "atraso_entrada_dias": atraso_entrada_dias,
                "atraso_entrada_status": atraso_entrada_status,
            })

        total = len(rows)
        entregues = sum(1 for x in rows if x["status"] == "Entregue")
        atrasados = sum(1 for x in rows if x["status"] == "Atrasado")
        aguardando_pc = sum(1 for x in rows if x["status"] == "Aguardando PC")
        aguardando_entrega = sum(1 for x in rows if x["status"] == "Aguardando entrega")
        atraso_entrada = sum(1 for x in rows if x["atraso_entrada_status"] == "Atraso de entrada")
        out.update({
            "ok": True,
            "source": MAPA_CONTROLE_NAME,
            "rows": sorted(rows, key=lambda x: (pd.to_numeric(x["sc"], errors="coerce") or 0, x["item"]), reverse=True),
            "summary": {
                "total": total,
                "entregues": entregues,
                "atrasados": atrasados,
                "aguardando_pc": aguardando_pc,
                "aguardando_entrega": aguardando_entrega,
                "atraso_entrada": atraso_entrada,
                "pct_entregue": f2(entregues / total * 100) if total else 0,
            },
        })
    except Exception as exc:
        out["error"] = str(exc)
    return out

def _load_dotacao_fisico_preliminar():
    """Lê a aba 'FÍSICO - EXECUTADO PRELIMINAR' da planilha de apoio: estrutura
    EAP com o valor orçado por item e a curva de distribuição mensal
    planejada (percentual sobre o valor orçado do próprio item). Usada pela
    aba Dotação para saber o que está PLANEJADO por mês; o que já está
    medido/comprometido vem do relatório de apropriações (orc_rel_tree),
    não dessa planilha — a coluna "MEDIDO ACUMULADO" dela é um valor global
    do projeto repetido em toda linha, não um valor real por item."""
    out = {"ok": False, "source": "", "months": [], "rows": {}}
    if not os.path.exists(SUPPORT_REPORT_PATH):
        return out
    try:
        xl = pd.ExcelFile(SUPPORT_REPORT_PATH, engine="openpyxl")
        sheet = _sheet_name_by_ascii(xl, "FISICO - EXECUTADO PRELIMINAR")
        if not sheet:
            out["error"] = "Aba 'FÍSICO - EXECUTADO PRELIMINAR' não encontrada."
            return out
        raw = xl.parse(sheet, header=None)
        header_idx = None
        for idx in raw.index[:20]:
            vals = [_ascii_upper(v) for v in raw.iloc[idx].tolist()]
            if any("ITEMIZACAO" in v for v in vals) and any("PRECO FINAL" in v for v in vals):
                header_idx = idx
                break
        if header_idx is None:
            out["error"] = "Cabeçalho da aba não encontrado."
            return out
        # so aceita colunas cujo cabecalho e um valor de data de verdade —
        # colunas numericas soltas apos os meses (ex.: coluna de checksum)
        # seriam interpretadas erradas pelo pd.to_datetime (viram 1970-01).
        month_cols = []
        for col in range(12, raw.shape[1]):
            header_val = raw.iat[header_idx, col]
            if not isinstance(header_val, (pd.Timestamp, datetime.datetime)):
                continue
            mk = _month_key(header_val)
            if mk:
                month_cols.append((col, mk))
        rows = {}
        for idx in range(header_idx + 1, len(raw)):
            lvl = pd.to_numeric(pd.Series([raw.iat[idx, 0]]), errors="coerce").iloc[0]
            if pd.isna(lvl):
                continue
            cod_raw = raw.iat[idx, 1]
            cod = _norm_rel_code(cod_raw)
            if not cod:
                continue
            orcado = f2(raw.iat[idx, 8] if raw.shape[1] > 8 else 0)
            nm = str(raw.iat[idx, 3] or "").strip()
            monthly_pct = {}
            for col, mk in month_cols:
                v = pd.to_numeric(pd.Series([raw.iat[idx, col]]), errors="coerce").fillna(0).iloc[0]
                if v:
                    monthly_pct[mk] = float(v)
            rows[cod] = {"cod": cod, "cod_raw": str(cod_raw or ""), "nm": nm, "orcado": orcado, "monthly_pct": monthly_pct}
        out.update({
            "ok": True,
            "source": os.path.basename(SUPPORT_REPORT_PATH),
            "months": [mk for _, mk in month_cols],
            "rows": rows,
        })
    except Exception as exc:
        out["error"] = str(exc)
    return out

def _aggregate_physical_curve(items):
    """Agrega um dict {cod: {orcado, monthly:{mk:pct}}} (nivel N2, uma aba
    FISICO - ... inteira) numa unica curva ponderada pelo orcado de cada item
    — % do escopo coberto pela aba, mes a mes e acumulado. Mesma logica de
    'curva do projeto inteiro' ja usada na Dotacao (_load_planejamento_financeiro_curva),
    aqui aplicada ao avanco fisico em vez de desembolso."""
    total = sum(float(v.get("orcado") or 0) for v in items.values())
    if total <= 0:
        return {}
    month_keys = sorted({mk for v in items.values() for mk in (v.get("monthly") or {})})
    monthly = {}
    acc = 0.0
    accum = {}
    for mk in month_keys:
        s = sum(float(v.get("orcado") or 0) * float((v.get("monthly") or {}).get(mk) or 0) for v in items.values())
        pct = s / total
        monthly[mk] = pct
        acc += pct
        accum[mk] = acc
    return {"monthly": monthly, "accum": accum}

def _merge_reprogrammed_curve(physical_months, reprogrammed_items):
    """A aba 'GRÁFICOS' (fonte de physical_months) tem a linha REPROGRAMAÇÃO
    zerada/desatualizada — não é recalculada quando a aba FÍSICO -
    REPROGRAMADO muda. Sobrescreve reprogrammed_month/reprogrammed_accum com
    a curva real, agregada item a item a partir da própria aba REPROGRAMADO
    (a mesma fonte que o usuário aponta como "o planejado" hoje)."""
    curve = _aggregate_physical_curve(reprogrammed_items or {})
    if not curve:
        return physical_months
    months_sorted = sorted(curve["monthly"])
    # A aba so tem dado real preenchido ate uma certa competencia — depois
    # disso todo mundo fica em 0 (mes ainda nao preenchido no arquivo, nao
    # porque o planejamento realmente zera). Sem esse corte, o acumulado
    # reprogramado ficaria travado nesse valor pra sempre dali em diante.
    # Meses sem dado real ficam de fora daqui e caem no fallback pro
    # planned_* que o grafico ja faz (reprogrammed_month||planned_month).
    real_months = [mk for mk in months_sorted if curve["monthly"][mk] > 0.0001]
    cutoff = real_months[-1] if real_months else None
    by_month = {m["mes"]: dict(m) for m in physical_months}
    for mk in months_sorted:
        if cutoff and mk > cutoff:
            continue
        by_month.setdefault(mk, {"mes": mk})
        by_month[mk]["reprogrammed_month"] = curve["monthly"][mk]
        by_month[mk]["reprogrammed_accum"] = curve["accum"][mk]
    return [by_month[mk] for mk in sorted(by_month)]

def _load_support_physical_curves(xl):
    sheet = _sheet_name_by_ascii(xl, "GRAFICOS")
    if not sheet:
        return []
    raw = xl.parse(sheet, header=None)
    if raw.empty or raw.shape[1] < 2:
        return []
    months = {}
    for col in range(1, raw.shape[1]):
        mk = _month_key(raw.iat[0, col])
        if mk:
            months[col] = mk
    label_map = {
        "PLANEJADO MENSAL": "planned_month",
        "EXECUTADO MENSAL": "executed_month",
        "REPROGRAMACAO MENSAL": "reprogrammed_month",
        "PLANEJADO ACUMULADO": "planned_accum",
        "EXECUTADO ACUMULADO": "executed_accum",
        "REPROGRAMACAO ACUMULADA": "reprogrammed_accum",
    }
    by_month = {mk: {"mes": mk} for mk in months.values()}
    for idx in raw.index:
        label = _ascii_upper(raw.iat[idx, 0])
        key = label_map.get(label)
        if not key:
            continue
        for col, mk in months.items():
            val = pd.to_numeric(pd.Series([raw.iat[idx, col]]), errors="coerce").fillna(0).iloc[0]
            by_month[mk][key] = float(val or 0)
    return [by_month[k] for k in sorted(by_month)]

def _load_support_histogram_adm(xl):
    sheet = _sheet_name_by_ascii(xl, "HISTOGRAMA ADM")
    if not sheet:
        return []
    raw = xl.parse(sheet, header=None)
    if raw.empty or raw.shape[1] < 4:
        return []

    header_idx = None
    month_cols = {}
    for idx in raw.index[:8]:
        found = {}
        for col in range(3, raw.shape[1]):
            mk = _month_key(raw.iat[idx, col])
            if mk:
                found[col] = mk
        if len(found) >= 3:
            header_idx = idx
            month_cols = found
            break
    if header_idx is None:
        return []

    by_month = {mk: {"mes": mk, "planejado": 0.0, "real": 0.0, "custo_planejado": 0.0, "custo_real": 0.0} for mk in month_cols.values()}
    for idx in range(header_idx + 1, len(raw)):
        kind = _ascii_upper(raw.iat[idx, 2] if raw.shape[1] > 2 else "")
        if kind not in {"P", "R"}:
            continue
        for col, mk in month_cols.items():
            val = pd.to_numeric(pd.Series([raw.iat[idx, col]]), errors="coerce").fillna(0).iloc[0]
            qty = float(val or 0)
            if not qty:
                continue
            if kind == "P":
                by_month[mk]["planejado"] += qty
            else:
                by_month[mk]["real"] += qty
    return [
        {k: f2(v) if k != "mes" else v for k, v in by_month[mk].items()}
        for mk in sorted(by_month)
        if f2(by_month[mk]["planejado"]) or f2(by_month[mk]["real"])
    ]

def load_ppc_plan():
    out = {"ok": False, "source": "", "updated_at": "", "total_previsto": 0, "total_realizado": 0, "meses": [], "servicos": []}
    if not os.path.exists(PPC_PLAN_PATH):
        return out
    try:
        raw = pd.read_excel(PPC_PLAN_PATH, sheet_name=0, header=None, engine="openpyxl")
        # A linha "Período:" com as datas de início/fim fica na linha 6 da planilha
        # (índice 5), não na linha 3 — usar a linha errada fazia cair no fallback
        # abaixo e desalinhava todos os dias em ~9 dias.
        start = pd.to_datetime(raw.iat[5, 3], errors="coerce") if raw.shape[0] > 5 and raw.shape[1] > 3 else pd.NaT
        end = pd.to_datetime(raw.iat[5, 4], errors="coerce") if raw.shape[0] > 5 and raw.shape[1] > 4 else pd.NaT
        if pd.isna(start):
            start = pd.Timestamp("2026-07-01")
        monthly = {}
        services = []
        for r in range(6, len(raw)):
            marker = _ascii_upper(raw.iat[r, 8] if raw.shape[1] > 8 else "")
            if marker != "PREVISTO":
                continue
            desc = str(raw.iat[r, 5] or "").strip() if raw.shape[1] > 5 else ""
            local = str(raw.iat[r, 4] or "").strip() if raw.shape[1] > 4 else ""
            linha = str(raw.iat[r, 3] or "").strip() if raw.shape[1] > 3 else ""
            equipe = str(raw.iat[r, 6] or "").strip() if raw.shape[1] > 6 else ""
            prev_count = 0
            real_count = 0
            for c in range(9, raw.shape[1]):
                date = start + pd.Timedelta(days=c - 9)
                if pd.notna(end) and date > end:
                    break
                mk = str(date.to_period("M"))
                prev = str(raw.iat[r, c] or "").strip().upper() == "X"
                real = r + 1 < len(raw) and str(raw.iat[r + 1, c] or "").strip().upper() == "X"
                if prev:
                    monthly.setdefault(mk, {"mes": mk, "previsto": 0, "realizado": 0})
                    monthly[mk]["previsto"] += 1
                    prev_count += 1
                if real:
                    monthly.setdefault(mk, {"mes": mk, "previsto": 0, "realizado": 0})
                    monthly[mk]["realizado"] += 1
                    real_count += 1
            if prev_count or real_count:
                services.append({"linha": linha, "local": local, "servico": desc, "equipe": equipe, "previsto": prev_count, "realizado": real_count})
        out.update({
            "ok": True,
            "source": os.path.basename(PPC_PLAN_PATH),
            "updated_at": time.strftime("%d/%m/%Y %H:%M:%S", time.localtime(os.path.getmtime(PPC_PLAN_PATH))),
            "inicio": str(start.date()) if pd.notna(start) else "",
            "fim": str(end.date()) if pd.notna(end) else "",
            "total_previsto": int(sum(v["previsto"] for v in monthly.values())),
            "total_realizado": int(sum(v["realizado"] for v in monthly.values())),
            "meses": [monthly[k] for k in sorted(monthly)],
            "servicos": services[:300],
        })
    except Exception as exc:
        out["error"] = str(exc)
    return out


def load_ppc_week(start_date, end_date):
    """Le a planilha de PPC dia a dia para uma semana especifica (Seg-Sex),
    reproduzindo a tabela 'PLANEJAMENTO SEMANAL' do RSO em papel."""
    out = {"ok": False, "servicos": [], "dias": [], "ppc_pct": 0, "evolucao": []}
    if not os.path.exists(PPC_PLAN_PATH):
        return out
    try:
        raw = pd.read_excel(PPC_PLAN_PATH, sheet_name=0, header=None, engine="openpyxl")
        period_start = pd.to_datetime(raw.iat[5, 3], errors="coerce") if raw.shape[0] > 5 and raw.shape[1] > 3 else pd.NaT
        if pd.isna(period_start):
            return out
        start = pd.to_datetime(start_date, errors="coerce")
        end = pd.to_datetime(end_date, errors="coerce")
        if pd.isna(start) or pd.isna(end):
            return out

        def col_for(date):
            offset = (date.normalize() - period_start.normalize()).days
            col = 9 + offset
            return col if 9 <= col < raw.shape[1] else None

        week_days = [d for d in pd.date_range(start, end, freq="D") if d.weekday() < 5][:5]

        def week_totals(days):
            cols = [col_for(d) for d in days]
            tp = to = 0
            r = 9
            while r + 1 < raw.shape[0]:
                if _ascii_upper(raw.iat[r, 8] if raw.shape[1] > 8 else "") != "PREVISTO":
                    r += 1
                    continue
                for col in cols:
                    if not col:
                        continue
                    if str(raw.iat[r, col] or "").strip().upper() == "X":
                        tp += 1
                        if str(raw.iat[r + 1, col] or "").strip().upper() == "X":
                            to += 1
                r += 2
            return tp, to

        cols = [col_for(d) for d in week_days]
        dia_labels = [{"dia": ["Seg", "Ter", "Qua", "Qui", "Sex"][i], "data": d.strftime("%d/%m")} for i, d in enumerate(week_days)]
        servicos = []
        r = 9
        while r + 1 < raw.shape[0]:
            if _ascii_upper(raw.iat[r, 8] if raw.shape[1] > 8 else "") != "PREVISTO":
                r += 1
                continue
            codigo = raw.iat[r, 2]
            linha = str(raw.iat[r, 3] or "").strip()
            local = str(raw.iat[r, 4] or "").strip()
            desc = str(raw.iat[r, 5] or "").strip()
            equipe = str(raw.iat[r, 6] or "").strip()
            dias = []
            for i, col in enumerate(cols):
                prev = bool(col) and str(raw.iat[r, col] or "").strip().upper() == "X"
                real = bool(col) and str(raw.iat[r + 1, col] or "").strip().upper() == "X"
                dias.append({"dia": dia_labels[i]["dia"] if i < len(dia_labels) else "", "previsto": prev, "realizado": real})
            if any(d["previsto"] or d["realizado"] for d in dias):
                total_prev = sum(1 for d in dias if d["previsto"])
                total_ok = sum(1 for d in dias if d["previsto"] and d["realizado"])
                if total_prev == 0:
                    status = "Não iniciado"
                elif total_ok == total_prev:
                    status = "Concluído"
                elif total_ok > 0:
                    status = "Em andamento"
                else:
                    status = "Atrasado"
                servicos.append({
                    "codigo": codigo if pd.notna(codigo) else "",
                    "linha": linha, "local": local, "servico": desc, "equipe": equipe,
                    "dias": dias, "status": status,
                })
            r += 2

        total_prev, total_ok = week_totals(week_days)
        evolucao = []
        cursor = period_start
        limit = 0
        while cursor <= end and limit < 60:
            w_days = [cursor + pd.Timedelta(days=i) for i in range(5)]
            tp, to = week_totals(w_days)
            if tp:
                evolucao.append({"semana": cursor.strftime("%d/%m"), "pct": round(to / tp * 100, 1)})
            cursor += pd.Timedelta(days=7)
            limit += 1

        out.update({
            "ok": True,
            "servicos": servicos,
            "dias": dia_labels,
            "ppc_pct": round(total_ok / total_prev * 100, 1) if total_prev else 0,
            "evolucao": evolucao[-12:],
        })
    except Exception as exc:
        out["error"] = str(exc)
    return out


def _load_support_budget_n2():
    out = {"ok": False, "source": "", "updated_at": "", "total": 0.0, "n2": {}, "n3": {}, "rows": [], "rows_n3": [], "physical_n2": {}, "physical_months": [], "ppc": load_ppc_plan()}
    if not os.path.exists(SUPPORT_REPORT_PATH):
        return out
    try:
        xl = pd.ExcelFile(SUPPORT_REPORT_PATH, engine="openpyxl")
        sheet = _sheet_name_by_ascii(xl, "FIN - AS - NOMINAL")
        if not sheet:
            out["error"] = "Aba FIN - AS - NOMINAL não encontrada."
            return out
        raw = xl.parse(sheet, header=None)
        rows = []
        n2 = {}
        rows_n3 = []
        n3 = {}
        for idx in raw.index:
            lvl = pd.to_numeric(pd.Series([raw.iat[idx, 0] if raw.shape[1] > 0 else None]), errors="coerce").iloc[0]
            if pd.isna(lvl):
                continue
            cod = _norm_budget_ref(raw.iat[idx, 1])
            val = f2(raw.iat[idx, 3] if raw.shape[1] > 3 else 0)
            # "val <= 0" nao filtra NaN (comparacao com NaN e sempre False em
            # Python) — linhas de subtotal com celula em branco viravam NaN e
            # passavam batido, contaminando a soma de n2/n3 inteira.
            if not cod or pd.isna(val) or val <= 0:
                continue
            nm = str(raw.iat[idx, 2] or "").strip()
            if int(lvl) == 1:
                n2[cod] = f2(n2.get(cod, 0) + val)
                rows.append({"cod": cod, "item": str(raw.iat[idx, 1] or ""), "nm": nm, "orcado": val})
            elif int(lvl) == 2:
                n3[cod] = f2(n3.get(cod, 0) + val)
                rows_n3.append({"cod": cod, "item": str(raw.iat[idx, 1] or ""), "nm": nm, "orcado": val})
        planned = _load_support_physical_sheet(xl, "FISICO - PLANEJADO")
        executed = _load_support_physical_sheet(xl, "FISICO - EXECUTADO")
        reprogrammed = _load_support_physical_sheet(xl, "FISICO - REPROGRAMADO")
        phys = {}
        for cod in sorted(set(planned) | set(executed) | set(reprogrammed)):
            p = planned.get(cod, {})
            e = executed.get(cod, {})
            rp = reprogrammed.get(cod, {})
            phys[cod] = {
                "cod": cod,
                "nm": (p.get("nm") or e.get("nm") or rp.get("nm") or ""),
                "planned_accum": float(p.get("accum") or 0),
                "executed_accum": float(e.get("accum") or 0),
                "reprogrammed_accum": float(rp.get("accum") or 0),
                "planned_monthly": p.get("monthly", {}),
                "executed_monthly": e.get("monthly", {}),
                "reprogrammed_monthly": rp.get("monthly", {}),
            }
        total = f2(sum(n2.values()))
        target_total = f2(NOMINAL_BUDGET_TOTAL)
        scale_factor = (target_total / total) if total > 0 and target_total > 0 else 1
        scaled_n2 = {cod: f2(val * scale_factor) for cod, val in n2.items()}
        scaled_n3 = {cod: f2(val * scale_factor) for cod, val in n3.items()}
        scaled_rows = [
            {**row, "orcado_original": row["orcado"], "orcado": f2(row["orcado"] * scale_factor)}
            for row in rows
        ]
        scaled_rows_n3 = [
            {**row, "orcado_original": row["orcado"], "orcado": f2(row["orcado"] * scale_factor)}
            for row in rows_n3
        ]
        # Se o RESUMO tiver um total nominal diferente, ele é apenas informativo.
        resumo_sheet = _sheet_name_by_ascii(xl, "RESUMO")
        resumo_total = 0.0
        if resumo_sheet:
            resumo = xl.parse(resumo_sheet, header=None)
            try:
                resumo_total = f2(resumo.iat[4, 2])
            except Exception:
                resumo_total = 0.0
        out.update({
            "ok": True,
            "source": os.path.basename(SUPPORT_REPORT_PATH),
            "updated_at": time.strftime("%d/%m/%Y %H:%M:%S", time.localtime(os.path.getmtime(SUPPORT_REPORT_PATH))),
            "total": target_total or total or resumo_total,
            "resumo_total": resumo_total,
            "raw_total": total,
            "scale_factor": scale_factor,
            "n2": scaled_n2,
            "n3": scaled_n3,
            "rows": scaled_rows,
            "rows_n3": scaled_rows_n3,
            "physical_n2": phys,
            "physical_months": _merge_reprogrammed_curve(_load_support_physical_curves(xl), reprogrammed),
            "histograma_equipe": _load_support_histogram_adm(xl),
            "ppc": load_ppc_plan(),
        })
    except Exception as exc:
        out["error"] = str(exc)
    return out

def load_nominal_budget_n2():
    """Lê o orçamento básico nominal (~73M) e devolve valores por nível 2."""
    support = _load_support_budget_n2()
    if support.get("ok"):
        return support
    out = {"ok": False, "source": "", "updated_at": "", "total": 0.0, "n2": {}, "n3": {}, "rows": [], "rows_n3": []}
    candidates = [NOMINAL_BUDGET_PATH] if os.path.exists(NOMINAL_BUDGET_PATH) else []
    candidates += [p for p in glob.glob(os.path.join(DATA_DIR, "*BASICO*.xlsx")) if p not in candidates]
    if not candidates:
        out["error"] = "Arquivo de orçamento nominal não encontrado."
        return out
    path = candidates[0]
    try:
        raw = pd.read_excel(path, sheet_name=0, header=None, engine="openpyxl")
        if raw.shape[1] <= 31:
            out["error"] = "Estrutura do orçamento nominal sem a coluna Elvis / Preço Total."
            return out
        level = pd.to_numeric(raw.iloc[:, 0], errors="coerce")
        value = pd.to_numeric(raw.iloc[:, 31], errors="coerce").fillna(0)
        item = raw.iloc[:, 1].astype(str).str.strip()
        desc = raw.iloc[:, 5].astype(str).str.strip() if raw.shape[1] > 5 else pd.Series([""] * len(raw))
        rows = []
        n2 = {}
        rows_n3 = []
        n3 = {}
        for idx in raw.index[level == 2]:
            parts = re.findall(r"\d+", str(item.iloc[idx]))
            if len(parts) < 2:
                continue
            code = f"{int(parts[0]):02d}.{int(parts[1]):03d}"
            val = f2(value.iloc[idx])
            if val <= 0:
                continue
            n2[code] = f2(n2.get(code, 0) + val)
            rows.append({"cod": code, "item": str(item.iloc[idx]), "nm": str(desc.iloc[idx]), "orcado": val})
        for idx in raw.index[level == 3]:
            parts = re.findall(r"\d+", str(item.iloc[idx]))
            if len(parts) < 3:
                continue
            code = f"{int(parts[0]):02d}.{int(parts[1]):03d}.{int(parts[2]):03d}"
            val = f2(value.iloc[idx])
            if val <= 0:
                continue
            n3[code] = f2(n3.get(code, 0) + val)
            rows_n3.append({"cod": code, "item": str(item.iloc[idx]), "nm": str(desc.iloc[idx]), "orcado": val})
        raw_total = f2(sum(n2.values()))
        factor = NOMINAL_BUDGET_TOTAL / raw_total if raw_total > 0 else 1
        scaled_n2 = {cod: f2(val * factor) for cod, val in n2.items()}
        scaled_n3 = {cod: f2(val * factor) for cod, val in n3.items()}
        scaled_rows = [{**row, "orcado_original": row["orcado"], "orcado": f2(row["orcado"] * factor)} for row in rows]
        scaled_rows_n3 = [{**row, "orcado_original": row["orcado"], "orcado": f2(row["orcado"] * factor)} for row in rows_n3]
        out.update({
            "ok": True,
            "source": os.path.basename(path),
            "updated_at": time.strftime("%d/%m/%Y %H:%M:%S", time.localtime(os.path.getmtime(path))),
            "total": f2(NOMINAL_BUDGET_TOTAL),
            "raw_total": raw_total,
            "scale_factor": factor,
            "n2": scaled_n2,
            "n3": scaled_n3,
            "rows": scaled_rows,
            "rows_n3": scaled_rows_n3,
        })
    except Exception as exc:
        out["error"] = str(exc)
    return out

NWD_DIR = os.path.join(DATA_DIR, "nwd_files")


def list_nwd_files():
    files = []
    seen = set()
    for pasta in (NWD_DIR, DATA_DIR):
        for pattern in ("*.nwd", "*.nwc"):
            for path in glob.glob(os.path.join(pasta, pattern)):
                nome = os.path.basename(path)
                if nome in seen:
                    continue
                seen.add(nome)
                try:
                    files.append({
                        "name": nome,
                        "size": os.path.getsize(path),
                        "updated_at": time.strftime("%d/%m/%Y %H:%M:%S", time.localtime(os.path.getmtime(path))),
                    })
                except Exception:
                    pass
    return sorted(files, key=lambda x: x.get("updated_at", ""), reverse=True)

def load_sharepoint_sync_status():
    default = {
        "executed_at": "",
        "ita_modified_at": "",
        "has_new_version": False,
        "validation_ok": False,
        "status": "aguardando",
        "source_url": "",
        "target_file": LOCAL_EXCEL_PATH,
        "file_updated": False,
        "message": "Aguardando a primeira execução automática às 10h.",
        "error": "",
    }
    if not os.path.exists(SHAREPOINT_STATUS_PATH):
        return default
    try:
        with open(SHAREPOINT_STATUS_PATH, "r", encoding="utf-8-sig") as fh:
            data = json.load(fh)
        default.update(data if isinstance(data, dict) else {})
    except Exception as exc:
        default.update(status="erro", error=str(exc), message="Não foi possível ler o relatório de atualização.")
    return default

def load_rso_state():
    if not os.path.exists(RSO_STATE_PATH):
        return {"ppc_servicos": [], "contratacoes_status": {}}
    try:
        with open(RSO_STATE_PATH, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        data.setdefault("ppc_servicos", [])
        data.setdefault("contratacoes_status", {})
        return data
    except Exception:
        return {"ppc_servicos": [], "contratacoes_status": {}}

def save_rso_state(data):
    data = data if isinstance(data, dict) else {}
    data.setdefault("ppc_servicos", [])
    data.setdefault("contratacoes_status", {})
    with open(RSO_STATE_PATH, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=2)
    return data

def load_rmo_state():
    default = {"sintese_geral": "", "photos": []}
    try:
        if os.path.exists(RMO_STATE_PATH):
            with open(RMO_STATE_PATH, "r", encoding="utf-8-sig") as fh:
                saved = json.load(fh)
            if isinstance(saved, dict):
                default.update(saved)
    except Exception:
        pass
    default["photos"] = [p for p in default.get("photos", []) if isinstance(p, dict)]
    return default


def save_rmo_state(payload):
    state = load_rmo_state()
    if isinstance(payload, dict):
        if "sintese_geral" in payload:
            state["sintese_geral"] = str(payload.get("sintese_geral") or "")
        if "photos" in payload and isinstance(payload.get("photos"), list):
            cleaned = []
            for item in payload.get("photos", []):
                if not isinstance(item, dict):
                    continue
                name = secure_filename(str(item.get("name") or ""))
                if not name:
                    continue
                cleaned.append({
                    "name": name,
                    "url": f"/api/rmo_photo/{name}",
                    "caption": str(item.get("caption") or ""),
                    "uploaded_at": str(item.get("uploaded_at") or ""),
                })
            state["photos"] = cleaned
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(RMO_STATE_PATH, "w", encoding="utf-8") as fh:
        json.dump(state, fh, ensure_ascii=False, indent=2)
    return state


@app.route("/api/rmo_state", methods=["GET", "POST"])
def api_rmo_state():
    if request.method == "POST":
        return jsonify({"ok": True, "state": save_rmo_state(request.get_json(silent=True) or {})})
    return jsonify({"ok": True, "state": load_rmo_state()})


@app.route("/api/rmo_photo", methods=["POST"])
def api_rmo_photo():
    file = request.files.get("photo")
    if not file or not file.filename:
        return jsonify({"ok": False, "error": "Nenhuma foto recebida."}), 400
    ext = os.path.splitext(file.filename)[1].lower()
    if ext not in {".jpg", ".jpeg", ".png", ".webp"}:
        return jsonify({"ok": False, "error": "Envie JPG, PNG ou WEBP."}), 400
    os.makedirs(RMO_PHOTOS_DIR, exist_ok=True)
    safe = secure_filename(file.filename)
    name = f"{int(time.time() * 1000)}_{safe}"
    target = os.path.join(RMO_PHOTOS_DIR, name)
    file.save(target)
    state = load_rmo_state()
    photo = {
        "name": name,
        "url": f"/api/rmo_photo/{name}",
        "caption": "",
        "uploaded_at": time.strftime("%Y-%m-%d %H:%M:%S"),
    }
    state.setdefault("photos", []).append(photo)
    save_rmo_state(state)
    return jsonify({"ok": True, "photo": photo, "state": state})


@app.route("/api/rmo_photo/<path:filename>")
def api_rmo_photo_file(filename):
    return send_from_directory(RMO_PHOTOS_DIR, secure_filename(filename))


def _rmo_month_value(rows, key, month):
    for row in rows or []:
        if str(row.get("mes") or "") == str(month or ""):
            return f2(row.get(key, 0))
    return 0.0


@app.route("/api/export/rmo.pptx")
def api_export_rmo_pptx():
    data, err = build_data()
    if err:
        return jsonify({"ok": False, "error": err}), 500
    try:
        from pptx import Presentation
        from pptx.util import Inches, Pt
        from pptx.dml.color import RGBColor
        from pptx.enum.text import PP_ALIGN
    except Exception as exc:
        return jsonify({
            "ok": False,
            "error": "Para gerar PowerPoint, instale a dependência python-pptx e reinicie o dashboard.",
            "detail": str(exc),
        }), 500

    def add_text(slide, x, y, w, h, text, size=14, bold=False, color=(0, 0, 0), align=None):
        box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = box.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = str(text or "")
        run.font.size = Pt(size)
        run.font.bold = bool(bold)
        run.font.color.rgb = RGBColor(*color)
        if align:
            p.alignment = align
        return box

    def fmt(v):
        try:
            return brl(v)
        except Exception:
            return f"R$ {float(v or 0):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    state = load_rmo_state()
    bi = data.get("bi_report") or {}
    evm = data.get("bi_evm") or {}
    cost = data.get("bi_cost_indicators") or {}
    evaluated = request.args.get("mes") or cost.get("evaluated_month") or ""
    nominal = f2((cost.get("budget") or {}).get("nominal") or data.get("nominal_budget_total") or NOMINAL_BUDGET_TOTAL)
    paid = f2(bi.get("pago") or bi.get("realizado") or 0)
    payable = f2(bi.get("a_pagar") or 0)
    final_cost = f2(paid + payable)
    result = f2(final_cost - nominal)
    evm_rows = evm.get("rows") or cost.get("evm_rows") or []
    vp = _rmo_month_value(evm_rows, "pv", evaluated)
    ac = _rmo_month_value(evm_rows, "ac", evaluated)
    va = _rmo_month_value(evm_rows, "ev", evaluated)

    slide = prs.slides.add_slide(blank)
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = RGBColor(245, 245, 245)
    add_text(slide, .25, .18, 2.7, .45, "TEMAZEC", 28, True, (165, 14, 36))
    add_text(slide, .28, .62, 2.6, .18, "DESENVOLVIMENTO IMOBILIÁRIO", 7, True, (90, 90, 90))
    logo_path = os.path.join(BASE_DIR, "static", "synergia_logo.png")
    if os.path.exists(logo_path):
        slide.shapes.add_picture(logo_path, Inches(4.4), Inches(.12), width=Inches(3.6))
    else:
        add_text(slide, 4.6, .22, 3.3, .45, "SYNERGIA BUENO", 27, True, (81, 116, 172), PP_ALIGN.CENTER)
    add_text(slide, 10.1, .28, 2.8, .32, "RELATÓRIO MENSAL DE OBRA", 15, True, (0, 0, 0), PP_ALIGN.RIGHT)
    add_text(slide, .35, 1.06, 12.5, .36, "SPE T2 BUENO - Gestão Financeira da Obra", 18, True, (0, 0, 0), PP_ALIGN.CENTER)

    metrics = [
        ("Orçamento nominal", fmt(nominal)),
        ("Custo final previsto", fmt(final_cost)),
        ("Resultado", fmt(result)),
        ("Valor planejado no mês", fmt(vp)),
        ("Custo real no mês", fmt(ac)),
        ("Valor agregado no mês", fmt(va)),
    ]
    for i, pair in enumerate(metrics):
        label, value = pair
        x = .45 + (i % 3) * 4.15
        y = 1.7 + (i // 3) * 1.0
        shape = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(3.8), Inches(.74))
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(255, 255, 255)
        shape.line.color.rgb = RGBColor(210, 210, 210)
        add_text(slide, x + .15, y + .08, 3.5, .2, label.upper(), 8, True, (80, 80, 80))
        add_text(slide, x + .15, y + .3, 3.5, .28, value, 15, True, (0, 0, 0))

    add_text(slide, .45, 3.78, 12.2, .26, "SÍNTESE GERAL", 12, True, (0, 0, 0))
    add_text(slide, .45, 4.1, 12.2, 1.5, state.get("sintese_geral") or "Preencha a síntese geral no dashboard antes de exportar.", 12, False, (30, 30, 30))
    add_text(slide, .45, 6.95, 12.3, .18, "Fontes: planilhas locais da pasta dados + controle de contratações.", 7, False, (90, 90, 90))

    rows = data.get("controle_contratacoes") or []
    slide = prs.slides.add_slide(blank)
    add_text(slide, .35, .22, 12.6, .35, "Controle de contratações", 18, True, (0, 0, 0), PP_ALIGN.CENTER)
    table_rows = min(12, len(rows)) + 1
    table = slide.shapes.add_table(table_rows, 6, Inches(.35), Inches(.85), Inches(12.65), Inches(5.85)).table
    headers = ["Item", "Etapa", "Descrição", "Orçado", "Status", "Observações"]
    widths = [.7, 1.4, 4.0, 1.3, 1.2, 4.0]
    for idx, width in enumerate(widths):
        table.columns[idx].width = Inches(width)
    for c, h in enumerate(headers):
        cell = table.cell(0, c)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(220, 220, 220)
    for r, item in enumerate(rows[:12], start=1):
        values = [
            item.get("item", ""),
            item.get("etapa", ""),
            item.get("descricao", ""),
            fmt(item.get("orc", 0)),
            item.get("status", ""),
            item.get("observacoes", ""),
        ]
        for c, val in enumerate(values):
            table.cell(r, c).text = str(val or "")

    photos = [p for p in state.get("photos", []) if p.get("name")]
    if photos:
        slide = prs.slides.add_slide(blank)
        add_text(slide, .35, .22, 12.6, .35, "Registro fotográfico", 18, True, (0, 0, 0), PP_ALIGN.CENTER)
        for i, p in enumerate(photos[:6]):
            path = os.path.join(RMO_PHOTOS_DIR, secure_filename(p.get("name") or ""))
            if not os.path.exists(path):
                continue
            x = .55 + (i % 3) * 4.2
            y = .95 + (i // 3) * 3.0
            try:
                slide.shapes.add_picture(path, Inches(x), Inches(y), width=Inches(3.8), height=Inches(2.35))
            except Exception:
                continue
            add_text(slide, x, y + 2.42, 3.8, .25, p.get("caption") or os.path.basename(path), 8, False, (60, 60, 60), PP_ALIGN.CENTER)

    bio = BytesIO()
    prs.save(bio)
    bio.seek(0)
    name = f"RMO_SYNERGIA_BUENO_{time.strftime('%Y%m%d_%H%M')}.pptx"
    return send_file(bio, as_attachment=True, download_name=name, mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation")

def _convert_pptx_to_pdf_libreoffice(pptx_path, pdf_path):
    """Linux/servidor: converte com o LibreOffice (soffice) em modo headless."""
    import shutil
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return False, "LibreOffice (soffice) não está instalado neste servidor."
    work = tempfile.mkdtemp(prefix="lo_pdf_")
    try:
        proc = subprocess.run(
            [soffice, "--headless", "--norestore", f"-env:UserInstallation=file://{work}/profile",
             "--convert-to", "pdf", "--outdir", work, pptx_path],
            capture_output=True, text=True, timeout=180,
        )
        gerado = os.path.join(work, os.path.splitext(os.path.basename(pptx_path))[0] + ".pdf")
        if proc.returncode != 0 or not os.path.exists(gerado):
            return False, (proc.stderr or proc.stdout or "LibreOffice não retornou o PDF.").strip()
        shutil.move(gerado, pdf_path)
        return True, ""
    except Exception as exc:
        return False, str(exc)
    finally:
        shutil.rmtree(work, ignore_errors=True)


def _convert_pptx_to_pdf_with_powerpoint(pptx_path, pdf_path):
    """Converte o PowerPoint RMO para PDF: PowerPoint no Windows, LibreOffice no Linux."""
    if os.name != "nt":
        return _convert_pptx_to_pdf_libreoffice(pptx_path, pdf_path)
    ps = f"""
$ErrorActionPreference = 'Stop'
$pptx = {json.dumps(pptx_path)}
$pdf = {json.dumps(pdf_path)}
$pp = New-Object -ComObject PowerPoint.Application
$presentation = $null
try {{
  $presentation = $pp.Presentations.Open($pptx, $true, $false, $false)
  $presentation.SaveAs($pdf, 32)
}} finally {{
  if ($presentation -ne $null) {{ $presentation.Close() }}
  $pp.Quit()
}}
"""
    proc = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps],
        capture_output=True,
        text=True,
        timeout=120,
    )
    if proc.returncode != 0 or not os.path.exists(pdf_path):
        detail = (proc.stderr or proc.stdout or "PowerPoint nao retornou o PDF.").strip()
        return False, detail
    return True, ""


@app.route("/api/export/rmo.pdf")
def api_export_rmo_pdf():
    """Gera o RMO em PDF a partir do PowerPoint, evitando a impressao da tela HTML."""
    mes = request.args.get("mes") or ""
    query = f"?mes={mes}" if mes else ""
    with app.test_client() as client:
        pptx_resp = client.get("/api/export/rmo.pptx" + query)
        if pptx_resp.status_code >= 400:
            return pptx_resp
        pptx_bytes = pptx_resp.get_data()
    with tempfile.TemporaryDirectory() as tmp:
        pptx_path = os.path.join(tmp, "rmo_synergia.pptx")
        pdf_path = os.path.join(tmp, "rmo_synergia.pdf")
        with open(pptx_path, "wb") as fh:
            fh.write(pptx_bytes)
        ok, detail = _convert_pptx_to_pdf_with_powerpoint(pptx_path, pdf_path)
        if not ok:
            return jsonify({
                "ok": False,
                "error": "Nao foi possivel converter a apresentacao RMO para PDF.",
                "detail": detail,
                "orientacao": "Use o botao Baixar PowerPoint e salve como PDF pelo PowerPoint enquanto verifico o conversor local.",
            }), 500
        with open(pdf_path, "rb") as fh:
            pdf_bytes = fh.read()
    out = BytesIO(pdf_bytes)
    out.seek(0)
    name = f"RMO_SYNERGIA_BUENO_{time.strftime('%Y%m%d_%H%M')}.pdf"
    return send_file(out, as_attachment=True, download_name=name, mimetype="application/pdf")

def load_controle_contratacoes():
    if not os.path.exists(CONTRATACOES_PATH):
        return []
    try:
        df = pd.read_excel(CONTRATACOES_PATH, sheet_name=0, header=4)
        df = df.dropna(how="all").copy()

        def _txt(col):
            if col not in df.columns:
                return ""
            v = r.get(col)
            return str(v).strip() if pd.notna(v) else ""

        def _num(col):
            if col not in df.columns:
                return 0.0
            v = r.get(col)
            return f2(v) if pd.notna(v) else 0.0

        rows = []
        for _, r in df.iterrows():
            item = r.get("ÍTEM") if "ÍTEM" in df.columns else r.iloc[1] if len(r) > 1 else ""
            desc = r.get("DESCRIÇÃO") if "DESCRIÇÃO" in df.columns else r.iloc[4] if len(r) > 4 else ""
            if pd.isna(item) and pd.isna(desc):
                continue
            rows.append({
                "item": str(item).strip() if pd.notna(item) else "",
                "cod": _txt("COD. ESTR."),
                "etapa": _txt("ETAPA"),
                "descricao": str(desc).strip() if pd.notna(desc) else "",
                "referencia": _txt("REFERÊNCIA ORÇAMENTÁRIA"),
                "orc": _num("ORÇ."),
                "custo": _num("CUSTO ESTIMADO"),
                "pulmao": _num("PULMÃO\nCONTA CORRENTE"),
                "data_aprovacao": _txt("DATA APROVAÇÃO"),
                "dias_aprovacao": _txt("QTE DIAS P/ APROVAÇÃO"),
                "meio": _txt("MEIO DE APROVAÇÃO"),
                "status": _txt("STATUS"),
                "observacoes": _txt("OBSERVAÇÕES"),
            })
        return rows
    except Exception:
        return []

def load_crono_pkgs_from_html():
    if not os.path.exists(CRONO_HTML_PATH):
        return [], []
    try:
        with open(CRONO_HTML_PATH, "r", encoding="utf-8", errors="ignore") as fh:
            html = fh.read()
        m = re.search(r"const\s+PKGS\s*=\s*(\[.*?\]);", html, flags=re.S)
        if not m:
            return [], []
        pkgs = json.loads(m.group(1))
        rows = []
        for i, p in enumerate(pkgs, 1):
            ini = pd.to_datetime(p.get("start_date"), errors="coerce")
            fim = pd.to_datetime(p.get("end_date"), errors="coerce")
            dias = int((fim - ini).days) + 1 if pd.notna(ini) and pd.notna(fim) else int(p.get("duration_weeks") or 0) * 7
            tarefa = " - ".join([str(x) for x in [p.get("zone"), p.get("floor"), p.get("activity")] if x])
            rows.append({
                "id": str(p.get("id") or i),
                "tarefa": tarefa,
                "inicio": str(ini.date()) if pd.notna(ini) else "—",
                "termino": str(fim.date()) if pd.notna(fim) else "—",
                "dias": max(dias, 1),
                "vlr_plan": 0.0,
                "peso": 1.0,
                "pct_real": 0.0,
                "nivel": int(p.get("level") or 0),
                "tipo": "PKG",
                "eap": str(p.get("zone") or "")
            })
        rows.sort(key=lambda x: (x["inicio"], x["nivel"], x["tarefa"]))
        lb = []
        dated = [r for r in rows if r["termino"] != "—"]
        if dated:
            by_month = {}
            for r in dated:
                mes = r["termino"][:7]
                by_month[mes] = by_month.get(mes, 0) + 1
            total = float(len(dated)) or 1.0
            acc = 0
            for mes in sorted(by_month):
                acc += by_month[mes]
                lb.append({"mes": mes, "planejado": round(acc / total * 100, 2), "vlr_plan": 0.0})
        return rows, lb
    except Exception:
        return [], []

def load_orc_saves():
    try:
        with open(ORC_ITEM_STATUS_PATH, "r", encoding="utf-8-sig") as fh:
            data = json.load(fh)
        if isinstance(data, dict) and isinstance(data.get("saves"), list):
            return data["saves"]
    except Exception:
        pass
    return []


def save_orc_saves(saves):
    os.makedirs(DATA_DIR, exist_ok=True)
    tmp = ORC_ITEM_STATUS_PATH + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump({"saves": saves}, fh, ensure_ascii=False, indent=2)
    os.replace(tmp, ORC_ITEM_STATUS_PATH)


def load_ifc_medicao_saves():
    try:
        with open(IFC_MEDICAO_SAVES_PATH, "r", encoding="utf-8-sig") as fh:
            data = json.load(fh)
        if isinstance(data, dict) and isinstance(data.get("saves"), list):
            return data["saves"]
    except Exception:
        pass
    return []


def save_ifc_medicao_saves(saves):
    os.makedirs(DATA_DIR, exist_ok=True)
    tmp = IFC_MEDICAO_SAVES_PATH + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump({"saves": saves}, fh, ensure_ascii=False, indent=2)
    os.replace(tmp, IFC_MEDICAO_SAVES_PATH)


def load_pa_medicoes():
    try:
        with open(PA_MEDICOES_PATH, "r", encoding="utf-8-sig") as fh:
            data = json.load(fh)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}

def save_pa_medicoes(data):
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(PA_MEDICOES_PATH, "w", encoding="utf-8") as fh:
        json.dump(data if isinstance(data, dict) else {}, fh, ensure_ascii=False, indent=2)

def load_pa_cronograma():
    """Extrai a linha base semanal e o histograma de equipe do PA_SYNERGIA_DOCS."""
    if not os.path.exists(PA_CRONO_PATH):
        return [], [], [], {}
    from openpyxl import load_workbook
    # A linha base é uma matriz visual larga; leitura em memória evita milhares
    # de buscas lentas que ocorrem no modo read_only com acesso aleatório.
    wb = load_workbook(PA_CRONO_PATH, read_only=False, data_only=True, keep_vba=True)
    try:
        ws = wb["LINHA BASE-RESUMO"]
        medicoes = load_pa_medicoes()
        dates_start = {c: pd.to_datetime(ws.cell(8, c).value, errors="coerce") for c in range(7, ws.max_column + 1)}
        dates_end = {c: pd.to_datetime(ws.cell(9, c).value, errors="coerce") for c in range(7, ws.max_column + 1)}
        tasks = []
        section = ""
        for r in range(11, ws.max_row + 1):
            sec = _clean_txt(ws.cell(r, 3).value)
            floor = _clean_txt(ws.cell(r, 5).value)
            if sec:
                section = sec
            if not floor:
                continue
            c = 7
            while c <= ws.max_column:
                code = _clean_txt(ws.cell(r, c).value)
                if not code or pd.isna(dates_start.get(c)):
                    c += 1
                    continue
                end_c = c
                while end_c + 1 <= ws.max_column and _clean_txt(ws.cell(r, end_c + 1).value) == code:
                    end_c += 1
                start = dates_start.get(c)
                end = dates_end.get(end_c)
                if pd.notna(start) and pd.notna(end):
                    task_id = f"PA-{r}-{c}-{code}"
                    pct = max(0, min(100, f2(medicoes.get(task_id, 0))))
                    tasks.append({
                        "id": task_id, "tarefa": f"{section} · {floor} · {code}",
                        "atividade": code, "pavimento": floor, "etapa": section,
                        "inicio": str(start.date()), "termino": str(end.date()),
                        "dias": max(1, int((end - start).days) + 1), "vlr_plan": 0,
                        "peso": 1, "pct_real": pct, "nivel": 3, "tipo": "PA_LINHA_BASE", "eap": "",
                    })
                c = end_c + 1

        by_month = {}
        for t in tasks:
            month = t["termino"][:7]
            rec = by_month.setdefault(month, {"planned": 0, "real": 0})
            rec["planned"] += 1
            rec["real"] += t["pct_real"] / 100
        total = float(len(tasks)) or 1
        acc_p = acc_r = 0.0
        linha = []
        for month in sorted(by_month):
            acc_p += by_month[month]["planned"]
            acc_r += by_month[month]["real"]
            linha.append({"mes": month, "planejado": round(acc_p / total * 100, 2),
                          "realizado": round(acc_r / total * 100, 2), "vlr_plan": 0})

        hist_ws = wb["HISTOGRAMA EQUIPE"]
        hist = []
        for c in range(12, hist_ws.max_column + 1):
            dt = pd.to_datetime(hist_ws.cell(12, c).value, errors="coerce")
            if pd.isna(dt):
                continue
            planned = real = planned_cost = real_cost = 0.0
            for r in range(14, 68):
                kind = _clean_txt(hist_ws.cell(r, 10).value).upper()
                qty = f2(hist_ws.cell(r, c).value)
                if kind == "P":
                    planned += qty
                    planned_cost += qty * f2(hist_ws.cell(r, 7).value or hist_ws.cell(r, 4).value)
                elif kind == "R":
                    real += qty
                    real_cost += qty * f2(hist_ws.cell(r, 7).value)
            hist.append({"mes": str(dt.date())[:7], "planejado": f2(planned), "real": f2(real),
                         "custo_planejado": f2(planned_cost), "custo_real": f2(real_cost)})
        meta = {
            "fonte": "PA_SYNERGIA_DOCS.xlsm · LINHA BASE-RESUMO",
            "project_file": os.path.basename(PA_CRONO_PATH), "project_path": PA_CRONO_PATH,
            "tarefas": len(tasks), "inicio": min((t["inicio"] for t in tasks), default="—"),
            "termino": max((t["termino"] for t in tasks), default="—"), "niveis": [3],
            "medicao": "Percentuais salvos no localhost", "updated_at": time.strftime(
                "%d/%m/%Y %H:%M", time.localtime(os.path.getmtime(PA_CRONO_PATH))),
        }
        return tasks, linha, hist, meta
    finally:
        wb.close()


def sync_mpp_cronograma():
    """Lê o cronograma direto do CRONOGRAMA_SYNERGIA_BUENO.mpp via automação
    COM do MS Project (instalado nesta máquina) e salva num cache local —
    não precisa de exportação manual pra Excel/HTML a cada atualização do
    Project. Demora ~1 a 2 minutos (a API COM do Project é lenta por
    tarefa), por isso roda só quando chamado explicitamente (botão
    "Atualizar do Project"), nunca dentro do carregamento normal do
    dashboard."""
    if not os.path.exists(PROJECT_MPP_PATH):
        return {"ok": False, "error": "Arquivo CRONOGRAMA_SYNERGIA_BUENO.mpp não encontrado em dados/."}
    if os.name != "nt":
        return {"ok": False, "error": "A leitura direta do .mpp usa o MS Project (só Windows). No servidor Linux o cronograma fica o da última sincronização; atualize rodando 'Atualizar do Project' no computador com o Project e enviando o mpp_cronograma_cache.json."}
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        return {"ok": False, "error": "pywin32 não está instalado neste ambiente."}

    pythoncom.CoInitialize()
    app = None
    try:
        app = win32com.client.DispatchEx("MSProject.Application")
        app.Visible = False
        app.DisplayAlerts = False
        app.FileOpen(Name=PROJECT_MPP_PATH, ReadOnly=True)
        proj = app.ActiveProject
        tasks_col = proj.Tasks
        n = tasks_col.Count

        tasks = []
        for i in range(1, n + 1):
            t = tasks_col(i)
            if t is None:
                continue
            start, finish = t.Start, t.Finish
            if not start or not finish:
                continue
            ini_d = pd.Timestamp(start).date()
            fim_d = pd.Timestamp(finish).date()
            dias = max(1, (fim_d - ini_d).days + 1)
            tasks.append({
                "id": str(t.UniqueID),
                "tarefa": str(t.Name or ""),
                "inicio": str(ini_d),
                "termino": str(fim_d),
                "dias": dias,
                "vlr_plan": f2(t.Cost or 0),
                "peso": 1,
                "pct_real": f2(t.PercentComplete or 0),
                "nivel": int(t.OutlineLevel or 0),
                "tipo": "RESUMO" if t.Summary else "TAREFA",
                "eap": str(t.WBS or ""),
                "custo_real": f2(t.ActualCost or 0),
            })
        tasks.sort(key=lambda x: (x["inicio"], x["nivel"]))

        # curva S ponderada por custo (só tarefas-folha com custo, senão soma
        # em dobro o que já está nos níveis-resumo acima delas na EAP)
        folhas = [t for t in tasks if t["tipo"] == "TAREFA" and t["vlr_plan"] > 0]
        total_custo = sum(t["vlr_plan"] for t in folhas) or 1.0
        by_month = {}
        for t in folhas:
            mes = t["termino"][:7]
            rec = by_month.setdefault(mes, {"plan": 0.0, "real": 0.0})
            rec["plan"] += t["vlr_plan"]
            rec["real"] += t["custo_real"]
        linha = []
        acc_p = acc_r = 0.0
        for mes in sorted(by_month):
            acc_p += by_month[mes]["plan"]
            acc_r += by_month[mes]["real"]
            linha.append({
                "mes": mes,
                "planejado": round(acc_p / total_custo * 100, 2),
                "realizado": round(acc_r / total_custo * 100, 2),
                "vlr_plan": f2(by_month[mes]["plan"]),
            })

        meta = {
            "fonte": "CRONOGRAMA_SYNERGIA_BUENO.mpp (leitura direta do MS Project)",
            "project_file": os.path.basename(PROJECT_MPP_PATH),
            "project_path": PROJECT_MPP_PATH,
            "tarefas": len(tasks),
            "inicio": min((t["inicio"] for t in tasks), default="—"),
            "termino": max((t["termino"] for t in tasks), default="—"),
            "niveis": sorted(set(t["nivel"] for t in tasks)),
            "updated_at": time.strftime("%d/%m/%Y %H:%M", time.localtime(os.path.getmtime(PROJECT_MPP_PATH))),
        }
        cache = {"tasks": tasks, "linha": linha, "hist": [], "meta": meta,
                 "sync_at": time.strftime("%Y-%m-%d %H:%M:%S")}
        os.makedirs(DATA_DIR, exist_ok=True)
        tmp = MPP_CRONO_CACHE_PATH + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(cache, fh, ensure_ascii=False)
        os.replace(tmp, MPP_CRONO_CACHE_PATH)
        return {"ok": True, "tarefas": len(tasks)}
    except Exception as exc:
        return {"ok": False, "error": f"Falha ao ler o .mpp via MS Project: {exc}"}
    finally:
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


def load_mpp_cronograma_cached():
    """Lê o cache já sincronizado pelo sync_mpp_cronograma() — rápido,
    sem chamar o MS Project. Usado no carregamento normal do dashboard."""
    if not os.path.exists(MPP_CRONO_CACHE_PATH):
        return [], [], [], {}
    try:
        with open(MPP_CRONO_CACHE_PATH, "r", encoding="utf-8") as fh:
            cache = json.load(fh)
        return cache.get("tasks", []), cache.get("linha", []), cache.get("hist", []), cache.get("meta", {})
    except Exception:
        return [], [], [], {}


EXCLUDED_FIN_DOCS = {'PCT', 'PPC'}

def excluded_fin_doc_mask(df):
    mask = pd.Series(False, index=df.index)
    if 'coddocum' in df.columns:
        mask |= df['coddocum'].astype(str).str.upper().str.strip().isin(EXCLUDED_FIN_DOCS)
    for col in ['obstit', 'dsobservacao', 'observacao', 'nmobservacao', 'dsobserv', 'observ', 'obs']:
        if col in df.columns:
            mask |= df[col].astype(str).str.upper().str.contains('PREVIS', na=False)
    return mask

def _resolve_odbc_driver(name):
    """No Windows o UAU usa o driver que estiver configurado (ex.: 'SQL Server');
    no Linux esse nome não existe — cai pro melhor driver Microsoft instalado
    (18, depois 17) pra não depender de editar o .env do servidor."""
    try:
        import pyodbc
        instaladas = list(pyodbc.drivers())
    except Exception:
        return name
    if not instaladas or name in instaladas:
        return name
    for pref in ("ODBC Driver 18 for SQL Server", "ODBC Driver 17 for SQL Server", "ODBC Driver 13 for SQL Server", "SQL Server"):
        if pref in instaladas:
            return pref
    return name


CREDORES_PATH = _ci_path(os.path.join(DATA_DIR, "credores.xlsx"))

def _norm_rel_code(v):
    import re
    raw = str(v or '').split(' - ')[0]
    parts = re.findall(r'\d+', raw)
    if not parts:
        return ''
    return '.'.join(str(int(p)).zfill(2) for p in parts)

_APROP_HIER_LABELS = {
    "Período", "Célula construtiva", "Etapa", "Subetapa", "Obra",
    "Unidade construtiva", "Valores corrigidos por", "Valores apresentados em",
    "Seleção por", "Analítico de Apropriação por Obra",
}

def _parse_aprop_hierarquico(path):
    """Achata o relatório 'Analítico de Apropriações por Obra' quando o Sienge
    o exporta agrupado (Período > Célula construtiva > Etapa > Subetapa >
    Serviço, com linhas de lançamento e "Total ..." entre blocos) em vez de
    tabela plana. Retorna um DataFrame com as mesmas colunas que o formato
    plano teria: Data, Documento, Título/Parcela, Credor/Histórico, Valor,
    Serviço, Observação (sempre vazia nesse formato — o relatório
    hierárquico não traz observação livre por lançamento)."""
    raw = pd.read_excel(path, sheet_name="Relatório", header=None)
    date_re = re.compile(r"^\d{2}/\d{2}/\d{4}$")
    rows_out = []
    cur_servico = ""
    for _, r in raw.iterrows():
        c0 = r.get(0)
        label = str(c0).strip() if pd.notna(c0) else ""
        if not label:
            continue
        if label == "Serviço":
            cur_servico = str(r.get(4) or "").strip()
            continue
        if label in _APROP_HIER_LABELS or label.startswith("Total "):
            continue
        if not date_re.match(label):
            continue
        rows_out.append({
            "Data": label,
            "Documento": r.get(1),
            "Título/Parcela": r.get(4),
            "Credor/Histórico": r.get(7),
            "Valor": r.get(14) if pd.notna(r.get(14)) else r.get(12),
            "Serviço": cur_servico,
            "Observação": "",
        })
    return pd.DataFrame(rows_out, columns=["Data", "Documento", "Título/Parcela", "Credor/Histórico", "Valor", "Serviço", "Observação"])

def _norm_words(v):
    txt = unicodedata.normalize("NFKD", str(v or "")).encode("ascii", "ignore").decode("ascii").upper()
    return set(re.findall(r"[A-Z0-9]{3,}", txt))

def load_sinapi_reference():
    meta = {
        "referencia": "04/2026", "uf": "GO", "fonte": "CAIXA / SINAPI",
        "url": "https://www.caixa.gov.br/poder-publico/modernizacao-gestao/sinapi/Paginas/default.aspx",
        "arquivo": os.path.basename(SINAPI_TABLE_PATH), "disponivel": False, "itens": 0,
    }
    rows = []
    if os.path.exists(SINAPI_META_PATH):
        try:
            with open(SINAPI_META_PATH, "r", encoding="utf-8-sig") as fh:
                meta.update(json.load(fh))
        except Exception:
            pass
    if not os.path.exists(SINAPI_TABLE_PATH):
        return rows, meta
    try:
        src = pd.read_excel(SINAPI_TABLE_PATH)
        cols = {str(c).strip().upper(): c for c in src.columns}
        cod_col = next((cols[c] for c in cols if "COD" in c), None)
        desc_col = next((cols[c] for c in cols if "DESC" in c), None)
        un_col = next((cols[c] for c in cols if "UNID" in c), None)
        price_col = next((cols[c] for c in cols if "PRECO" in c or "CUSTO" in c), None)
        if cod_col is not None and desc_col is not None:
            for _, r in src.iterrows():
                desc = _clean_txt(r.get(desc_col))
                if not desc:
                    continue
                rows.append({
                    "codigo": _clean_txt(r.get(cod_col)), "descricao": desc,
                    "unidade": _clean_txt(r.get(un_col)) if un_col is not None else "",
                    "preco": f2(r.get(price_col)) if price_col is not None else 0,
                    "words": _norm_words(desc),
                })
        meta.update(disponivel=bool(rows), itens=len(rows))
    except Exception as exc:
        meta["erro"] = str(exc)
    return rows, meta

def build_activity_cost_analysis(plan_list):
    """Cruza a EAP do cronograma com o orçamento executivo e a referência SINAPI local."""
    empty_meta = {"atividades": len(plan_list or []), "vinculadas": 0, "sem_vinculo": len(plan_list or []),
                  "custo_mapeado": 0, "sinapi_vinculados": 0}
    if not plan_list or not os.path.exists(ORC_REL_PATH):
        _, sinapi_meta = load_sinapi_reference()
        return [], empty_meta, sinapi_meta
    try:
        orc = pd.read_excel(ORC_REL_PATH, sheet_name=2, header=10)
        orc.columns = [str(c).strip() for c in orc.columns]
        budget = {}
        current = None
        for _, r in orc.iterrows():
            raw_eap = _clean_txt(r.get("ITEMIZAÇÃO"))
            if raw_eap:
                current = _norm_rel_code(raw_eap)
                if len(current.split(".")) == 4:
                    budget[current] = {
                        "eap": current, "codigo": _clean_txt(r.get("CÓDIGO")),
                        "descricao": _clean_txt(r.get("DESCRIÇÃO")),
                        "unidade": _clean_txt(r.get("UNID.")),
                        "quantidade": f2(r.get("QUANT./COEF.")),
                        "custo_orcamento": f2(r.get("PREÇO FINAL (TOT.)(R$)")),
                        "custo_unitario": f2(r.get("PREÇO FINAL (UNIT.)(R$)")),
                        "mat": f2(r.get("PREÇO MAT. (TOT.)(R$)")),
                        "mao_obra": f2(r.get("PREÇO M.O. (TOT.)(R$)")),
                        "outros": f2(r.get("PREÇO OUTROS (TOT.)(R$)")),
                        "componentes": [],
                    }
                continue
            if current in budget and _clean_txt(r.get("CÓDIGO")):
                qty = f2(r.get("QUANT./COEF."))
                unit = f2(r.get("PREÇO FINAL (UNIT.)(R$)"))
                budget[current]["componentes"].append({
                    "codigo": _clean_txt(r.get("CÓDIGO")), "descricao": _clean_txt(r.get("DESCRIÇÃO")),
                    "classe": _clean_txt(r.get("CLASS")), "unidade": _clean_txt(r.get("UNID.")),
                    "quantidade": qty, "unitario": unit, "total": f2(qty * unit),
                })

        sinapi, sinapi_meta = load_sinapi_reference()
        result = []
        for task in plan_list:
            eap = _norm_rel_code(task.get("eap"))
            item = budget.get(eap)
            plan_cost = f2(task.get("vlr_plan"))
            row = dict(task)
            row.update({
                "eap_norm": eap, "orc_codigo": "", "orc_descricao": "", "orc_unidade": "",
                "custo_atividade": plan_cost, "custo_orcamento": 0, "origem_custo": "Cronograma",
                "componentes": 0, "composicao_top": [], "sinapi_codigo": "", "sinapi_descricao": "",
                "sinapi_preco": 0, "sinapi_confianca": 0,
            })
            if item:
                cost = item["custo_orcamento"] or plan_cost
                row.update({
                    "orc_codigo": item["codigo"], "orc_descricao": item["descricao"],
                    "orc_unidade": item["unidade"], "custo_atividade": f2(cost),
                    "custo_orcamento": item["custo_orcamento"],
                    "origem_custo": "Orçamento / EAP" if item["custo_orcamento"] else "Cronograma / EAP",
                    "componentes": len(item["componentes"]),
                    "composicao_top": sorted(item["componentes"], key=lambda x: x["total"], reverse=True)[:5],
                })
                target_words = _norm_words(item["descricao"])
                best = None
                best_score = 0
                for ref in sinapi:
                    union = target_words | ref["words"]
                    score = len(target_words & ref["words"]) / len(union) if union else 0
                    if score > best_score:
                        best, best_score = ref, score
                if best and best_score >= .42:
                    row.update(sinapi_codigo=best["codigo"], sinapi_descricao=best["descricao"],
                               sinapi_preco=best["preco"], sinapi_confianca=round(best_score * 100, 1))
            result.append(row)
        linked = [r for r in result if r["orc_descricao"]]
        sin_linked = [r for r in result if r["sinapi_codigo"]]
        meta = {
            "atividades": len(result), "vinculadas": len(linked), "sem_vinculo": len(result) - len(linked),
            "custo_mapeado": f2(sum(r["custo_atividade"] for r in linked)),
            "sinapi_vinculados": len(sin_linked), "orcamento_atualizado": time.strftime(
                "%d/%m/%Y %H:%M", time.localtime(os.path.getmtime(ORC_REL_PATH))),
        }
        return result, meta, sinapi_meta
    except Exception as exc:
        empty_meta["erro"] = str(exc)
        _, sinapi_meta = load_sinapi_reference()
        return [], empty_meta, sinapi_meta

def _digits_only(v):
    import re
    return ''.join(re.findall(r'\d+', str(v or '')))

def _fmt_date(v):
    d = pd.to_datetime(v, errors="coerce", dayfirst=True)
    return str(d.date()) if pd.notna(d) else "—"

def _clean_txt(v):
    if pd.isna(v):
        return ""
    s = str(v).strip()
    return "" if s.lower() in ("nan", "none", "nat") else s

def _norm_company_name(v):
    s = re.sub(r"\s+\d{11,14}\s*$", "", str(v or "").upper().strip())
    return re.sub(r"[^A-Z0-9]+", "", s)

def load_cnpj_cache():
    try:
        with open(CNPJ_CACHE_PATH, "r", encoding="utf-8-sig") as fh:
            data = json.load(fh)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}

def save_cnpj_cache(data):
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(CNPJ_CACHE_PATH, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=2)

def lookup_cnpj_online(cnpj):
    digits = _digits_only(cnpj)
    if len(digits) != 14:
        return None
    req = Request(f"https://brasilapi.com.br/api/cnpj/v1/{digits}",
                  headers={"User-Agent": "Dashboard-Obra-TEMAZEC/1.0"})
    try:
        with urlopen(req, timeout=8) as response:
            payload = json.loads(response.read().decode("utf-8"))
        return {"cnpj": digits, "nome_fantasia": _clean_txt(payload.get("nome_fantasia")),
                "razao_social": _clean_txt(payload.get("razao_social")),
                "consultado_em": time.strftime("%Y-%m-%d %H:%M:%S")}
    except Exception:
        return None

def load_cnpj_full_cache():
    try:
        with open(CNPJ_FULL_CACHE_PATH, "r", encoding="utf-8-sig") as fh:
            data = json.load(fh)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}

def save_cnpj_full_cache(data):
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(CNPJ_FULL_CACHE_PATH, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=2)

def lookup_cnpj_full(cnpj):
    """Consulta CNPJ na BrasilAPI (dados oficiais da Receita Federal) trazendo
    endereço, telefone, e-mail e quadro societário — usado para pré-preencher
    minutas de contrato. Resultado fica em cache local (dados/cnpj_full_cache.json)."""
    digits = _digits_only(cnpj)
    if len(digits) != 14:
        return None
    cache = load_cnpj_full_cache()
    if digits in cache:
        return cache[digits]
    req = Request(f"https://brasilapi.com.br/api/cnpj/v1/{digits}",
                  headers={"User-Agent": "Dashboard-Obra-TEMAZEC/1.0"})
    try:
        with urlopen(req, timeout=10) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except Exception:
        return None
    endereco_partes = [
        _clean_txt(payload.get("descricao_tipo_de_logradouro")),
        _clean_txt(payload.get("logradouro")),
    ]
    logradouro = " ".join(p for p in endereco_partes if p).strip()
    numero = _clean_txt(payload.get("numero"))
    complemento = _clean_txt(payload.get("complemento"))
    endereco_full = ", ".join(p for p in [logradouro, numero] if p)
    if complemento:
        endereco_full = f"{endereco_full} - {complemento}" if endereco_full else complemento
    qsa = payload.get("qsa") or []
    representante = ""
    for socio in qsa:
        qualif = str(socio.get("qualificacao_socio") or "").upper()
        if "ADMINISTRADOR" in qualif or "SÓCIO-ADMINISTRADOR" in qualif or "DIRETOR" in qualif or "PRESIDENTE" in qualif:
            representante = _clean_txt(socio.get("nome_socio"))
            break
    if not representante and qsa:
        representante = _clean_txt(qsa[0].get("nome_socio"))
    ddd = _clean_txt(payload.get("ddd_telefone_1"))
    result = {
        "cnpj": digits,
        "razao_social": _clean_txt(payload.get("razao_social")),
        "nome_fantasia": _clean_txt(payload.get("nome_fantasia")),
        "endereco": endereco_full,
        "bairro": _clean_txt(payload.get("bairro")),
        "municipio": _clean_txt(payload.get("municipio")),
        "uf": _clean_txt(payload.get("uf")),
        "cep": _clean_txt(payload.get("cep")),
        "telefone": ddd,
        "email": _clean_txt(payload.get("email")),
        "representante_legal": representante,
        "socios": [{"nome": _clean_txt(s.get("nome_socio")), "qualificacao": _clean_txt(s.get("qualificacao_socio"))} for s in qsa],
        "situacao_cadastral": _clean_txt(payload.get("descricao_situacao_cadastral")),
        "consultado_em": time.strftime("%Y-%m-%d %H:%M:%S"),
    }
    cache[digits] = result
    save_cnpj_full_cache(cache)
    return result

def refresh_cnpj_cache(cnpjs):
    cache = load_cnpj_cache()
    pending = sorted({_digits_only(v) for v in cnpjs if len(_digits_only(v)) == 14 and _digits_only(v) not in cache})
    updated = 0
    with ThreadPoolExecutor(max_workers=6) as pool:
        futures = [pool.submit(lookup_cnpj_online, cnpj) for cnpj in pending[:80]]
        for future in as_completed(futures):
            item = future.result()
            if item:
                cache[item["cnpj"]] = item
                updated += 1
    if updated:
        save_cnpj_cache(cache)
    return cache, updated

def build_cnpj_maps(bancos):
    code_to_cnpj, name_to_cnpj = {}, {}
    if os.path.exists(CREDORES_PATH):
        try:
            raw = pd.read_excel(CREDORES_PATH, sheet_name=0, header=None)
            for _, row in raw.iterrows():
                first = _clean_txt(row.iloc[0] if len(row) else "")
                match = re.match(r"^(\d+)\s*-\s*(.+)", first)
                cnpj = _digits_only(row.iloc[2] if len(row) > 2 else "")
                if match and len(cnpj) in (11, 14):
                    code_to_cnpj[int(match.group(1))] = cnpj
                    name_to_cnpj[_norm_company_name(match.group(2))] = cnpj
        except Exception:
            pass
    if bancos is not None and not bancos.empty:
        for _, row in bancos.iterrows():
            try:
                code = int(float(row.get("codcredor")))
                cnpj = _digits_only(row.get("cpfcnpfcredor"))
                name = _clean_txt(row.get("nmcredor"))
                if len(cnpj) in (11, 14):
                    code_to_cnpj.setdefault(code, cnpj)
                    if name:
                        name_to_cnpj.setdefault(_norm_company_name(name), cnpj)
            except Exception:
                pass
    return code_to_cnpj, name_to_cnpj

def build_suprimentos(sol_ped, nf_ped, orc=None, tit=None, cred_map=None):
    if sol_ped is None or sol_ped.empty:
        return {"summary": {}, "solicitacoes": [], "cotacoes": [], "pedidos": []}

    sol = sol_ped.copy()
    nf = nf_ped.copy() if nf_ped is not None else pd.DataFrame()
    if "codobra" in sol.columns:
        sol = sol[pd.to_numeric(sol["codobra"], errors="coerce") == 1016].copy()
    elif "codccusto" in sol.columns:
        sol = sol[pd.to_numeric(sol["codccusto"], errors="coerce") == 1016].copy()
    if not nf.empty:
        if "codobraaprop" in nf.columns:
            nf = nf[pd.to_numeric(nf["codobraaprop"], errors="coerce") == 1016].copy()
        elif "codccusto" in nf.columns:
            nf = nf[pd.to_numeric(nf["codccusto"], errors="coerce") == 1016].copy()

    orc_budget = {}
    orc_name = {}
    if orc is not None and not orc.empty and "refitemorc" in orc.columns:
        ob = orc.copy()
        ob["refitemorc"] = ob["refitemorc"].astype(str).str.strip()
        ob["vlritemorc"] = pd.to_numeric(ob.get("vlritemorc"), errors="coerce").fillna(0)
        ob = ob[ob["refitemorc"].str.len() == 14]
        orc_budget = ob.groupby("refitemorc")["vlritemorc"].sum().to_dict()
        orc_name = ob.groupby("refitemorc")["nmitemorc"].first().to_dict() if "nmitemorc" in ob.columns else {}

    realizado_eap = {}
    if tit is not None and not tit.empty and "refitemorc" in tit.columns:
        tr = tit.copy()
        tr["refitemorc"] = tr["refitemorc"].astype(str).str.strip()
        tr["vlrliqbaixacp"] = pd.to_numeric(tr.get("vlrliqbaixacp"), errors="coerce").fillna(0)
        realizado_eap = tr.groupby("refitemorc")["vlrliqbaixacp"].sum().to_dict()

    nf_ref_by_key = {}
    nf_ref_by_item = {}
    nf_ref_by_id = {}
    if not nf.empty and "refitemorc" in nf.columns:
        nf["refitemorc"] = nf["refitemorc"].astype(str).str.strip()
        nf_valid = nf[nf["refitemorc"].notna() & (nf["refitemorc"] != "") & (nf["refitemorc"].str.lower() != "nan")].copy()
        for c in ["codpedi", "codinsum", "coddet"]:
            if c in nf_valid.columns:
                nf_valid[c] = pd.to_numeric(nf_valid[c], errors="coerce")
        if "idunicocomp" in nf_valid.columns:
            nf_ref_by_id = nf_valid.dropna(subset=["idunicocomp"]).groupby("idunicocomp")["refitemorc"].first().to_dict()
        if all(c in nf_valid.columns for c in ["codpedi", "codinsum", "coddet"]):
            nf_ref_by_key = nf_valid.dropna(subset=["codpedi", "codinsum", "coddet"]).groupby(["codpedi", "codinsum", "coddet"])["refitemorc"].first().to_dict()
        if all(c in nf_valid.columns for c in ["codinsum", "coddet"]):
            nf_ref_by_item = nf_valid.dropna(subset=["codinsum", "coddet"]).groupby(["codinsum", "coddet"])["refitemorc"].first().to_dict()

    for c in ["codccusto", "codsolic", "codpedi", "codinsum", "coddet", "qtsolic", "qtpedi", "qtpendpedi", "qtentregpedi", "vlrinsumpedi"]:
        if c in sol.columns:
            sol[c] = pd.to_numeric(sol[c], errors="coerce")
    for c in ["dtsolic", "dtpedi", "dtautped", "dtpreventsolic"]:
        if c in sol.columns:
            sol[c] = pd.to_datetime(sol[c], errors="coerce")
    sol["vlr_previsto"] = (sol.get("qtpedi", 0).fillna(0) * sol.get("vlrinsumpedi", 0).fillna(0)).round(2)
    sol["vlr_solicitado"] = (sol.get("qtsolic", 0).fillna(0) * sol.get("vlrinsumpedi", 0).fillna(0)).round(2)
    sol["aprovado"] = sol.get("dtautped").notna() & (sol.get("reprovado", "N").astype(str).str.upper() != "S") & (sol.get("statpedi", "").astype(str).str.upper() != "CANCELADO")

    nf_keys = ["codpedi", "codinsum", "coddet"]
    nf_by_item = {}
    if not nf.empty and all(c in nf.columns for c in nf_keys):
        for c in nf_keys:
            nf[c] = pd.to_numeric(nf[c], errors="coerce")
        nf["vlrinsumnf"] = pd.to_numeric(nf.get("vlrinsumnf"), errors="coerce").fillna(0)
        nf["qtentregnf"] = pd.to_numeric(nf.get("qtentregnf"), errors="coerce").fillna(0)
        nf_grp = nf.groupby(nf_keys, dropna=False).agg(
            realizado=("vlrinsumnf", "sum"),
            qtd_realizada=("qtentregnf", "sum"),
            titulos=("numtitcp", lambda x: sorted(set(str(int(v)) for v in x if pd.notna(v)))),
            nfs=("numdocumnf", lambda x: sorted(set(str(v).strip() for v in x if pd.notna(v) and str(v).strip()))),
            codcredores=("codcredornf", lambda x: sorted(set(int(float(v)) for v in x if pd.notna(v)))) if "codcredornf" in nf.columns else ("codpedi", lambda x: []),
        ).reset_index()
        nf_by_item = {
            (f2(r.codpedi), f2(r.codinsum), f2(r.coddet)): {
                "realizado": f2(r.realizado),
                "qtd_realizada": f2(r.qtd_realizada),
                "titulos": r.titulos,
                "nfs": r.nfs,
                "credores": sorted(set((cred_map or {}).get(int(v), f"CREDOR {int(v)}") for v in r.codcredores)),
            }
            for _, r in nf_grp.iterrows()
        }

    cot_rows = []
    for _, r in sol.iterrows():
        key = (f2(r.get("codpedi")), f2(r.get("codinsum")), f2(r.get("coddet")))
        done = nf_by_item.get(key, {})
        ref = nf_ref_by_id.get(_clean_txt(r.get("idunicocomp"))) or nf_ref_by_key.get(key) or nf_ref_by_item.get((f2(r.get("codinsum")), f2(r.get("coddet")))) or ""
        previsto = f2(r.get("vlr_previsto", 0))
        realizado = f2(done.get("realizado", 0))
        saldo = f2(max(previsto - realizado, 0))
        aprovado = bool(r.get("aprovado"))
        cot_rows.append({
            "cc": str(int(r.codccusto)) if pd.notna(r.get("codccusto")) else "",
            "cotacao": str(int(r.codpedi)) if pd.notna(r.get("codpedi")) else "—",
            "solicitacao": str(int(r.codsolic)) if pd.notna(r.get("codsolic")) else "—",
            "data": _fmt_date(r.get("dtpedi")),
            "aprovacao": _fmt_date(r.get("dtautped")),
            "aprovado": aprovado,
            "status": "Aprovada" if aprovado else "Não aprovada",
            "status_pedido": _clean_txt(r.get("statpedi")) or "—",
            "insumo": _clean_txt(r.get("nminsum")) or "—",
            "detalhe": _clean_txt(r.get("nmdetinsum")),
            "eap": ref or "—",
            "item_orc": _clean_txt(orc_name.get(ref, "")) if ref else "",
            "orcado_eap": f2(orc_budget.get(ref, 0)),
            "realizado_eap": f2(realizado_eap.get(ref, 0)),
            "comprometido_eap": 0,
            "saldo_eap": 0,
            "grupo": _clean_txt(r.get("nmgrins")),
            "un": _clean_txt(r.get("codunidmed")) or _clean_txt(r.get("codunidmov")),
            "qtd": f2(r.get("qtpedi", 0)),
            "unit": f2(r.get("vlrinsumpedi", 0)),
            "apropriado": previsto,
            "realizado": realizado,
            "comprometido": saldo if aprovado else 0,
            "impacto": saldo if aprovado else previsto,
            "saldo": saldo,
            "saldo_insumo": 0,
            "comprador": _clean_txt(r.get("codcomprador")),
            "titulos": done.get("titulos", []),
            "nfs": done.get("nfs", []),
            "credores": done.get("credores", []),
        })

    saldo_por_insumo = {}
    for row in cot_rows:
        saldo_por_insumo[row["insumo"]] = f2(saldo_por_insumo.get(row["insumo"], 0) + row["saldo"])
    for row in cot_rows:
        row["saldo_insumo"] = saldo_por_insumo.get(row["insumo"], 0)

    comprometido_por_eap = {}
    for row in cot_rows:
        if row["eap"] and row["eap"] != "—":
            comprometido_por_eap[row["eap"]] = f2(comprometido_por_eap.get(row["eap"], 0) + row["comprometido"])
    for row in cot_rows:
        if row["eap"] and row["eap"] != "—":
            row["comprometido_eap"] = f2(comprometido_por_eap.get(row["eap"], 0))
            row["saldo_eap"] = f2(row["orcado_eap"] - row["realizado_eap"] - row["comprometido_eap"])

    sol_rows = []
    for cod, grp in sol.groupby("codsolic", dropna=False):
        qtd_sol = f2(grp["qtsolic"].sum())
        qtd_at = f2(grp.get("qtatendsolic", pd.Series(dtype=float)).sum())
        val_sol = f2(grp["vlr_solicitado"].sum())
        saldo_qtd = f2(max(qtd_sol - qtd_at, 0))
        saldo_valor = f2(val_sol * (saldo_qtd / qtd_sol)) if qtd_sol else 0
        sol_rows.append({
            "cc": str(int(grp["codccusto"].dropna().iloc[0])) if grp["codccusto"].notna().any() else "",
            "solicitacao": str(int(cod)) if pd.notna(cod) else "—",
            "data": _fmt_date(grp["dtsolic"].min()),
            "previsto": _fmt_date(grp["dtpreventsolic"].min()) if "dtpreventsolic" in grp.columns else "—",
            "status": _clean_txt(grp["statsolic"].dropna().iloc[0]) if grp["statsolic"].notna().any() else "—",
            "solicitante": _clean_txt(grp["codususolic"].dropna().iloc[0]) if grp["codususolic"].notna().any() else "",
            "itens": int(len(grp)),
            "qtd": qtd_sol,
            "valor": val_sol,
            "atendido": qtd_at,
            "saldo": saldo_qtd,
            "saldo_valor": saldo_valor,
        })

    ped_rows = []
    for cod, grp in sol.groupby("codpedi", dropna=False):
        cot_subset = [r for r in cot_rows if r["cotacao"] == (str(int(cod)) if pd.notna(cod) else "—")]
        previsto = f2(sum(r["apropriado"] for r in cot_subset))
        realizado = f2(sum(r["realizado"] for r in cot_subset))
        saldo = f2(max(previsto - realizado, 0))
        aprovado = any(r["aprovado"] for r in cot_subset)
        vinculos = []
        for item in cot_subset:
            vinculos.append({
                "insumo": item.get("insumo", "—"),
                "detalhe": item.get("detalhe", ""),
                "eap": item.get("eap", "—"),
                "item_orc": item.get("item_orc", ""),
                "valor": f2(item.get("apropriado", 0)),
                "realizado": f2(item.get("realizado", 0)),
                "saldo": f2(item.get("saldo", 0)),
                "titulos": item.get("titulos", []),
                "nfs": item.get("nfs", []),
                "credores": item.get("credores", []),
            })
        ped_rows.append({
            "cc": str(int(grp["codccusto"].dropna().iloc[0])) if grp["codccusto"].notna().any() else "",
            "pedido": str(int(cod)) if pd.notna(cod) else "—",
            "data": _fmt_date(grp["dtpedi"].min()),
            "aprovacao": _fmt_date(grp["dtautped"].min()),
            "status": _clean_txt(grp["statpedi"].dropna().iloc[0]) if grp["statpedi"].notna().any() else "—",
            "aprovado": aprovado,
            "comprador": _clean_txt(grp["codcomprador"].dropna().iloc[0]) if grp["codcomprador"].notna().any() else "",
            "itens": int(len(grp)),
            "valor": previsto,
            "realizado": realizado,
            "saldo": saldo,
            "impacto": saldo if aprovado else previsto,
            "nfs": sorted(set(n for r in cot_subset for n in r["nfs"])),
            "titulos": sorted(set(t for r in cot_subset for t in r["titulos"])),
            "credores": sorted(set(c for r in cot_subset for c in r.get("credores", []))),
            "vinculos": vinculos,
        })

    summary = {
        "solicitacoes": len(sol_rows),
        "cotacoes": len(ped_rows),
        "pedidos": len(ped_rows),
        "aprovadas": sum(1 for r in ped_rows if r["aprovado"]),
        "nao_aprovadas": sum(1 for r in ped_rows if not r["aprovado"]),
        "apropriado": f2(sum(r["apropriado"] for r in cot_rows)),
        "realizado": f2(sum(r["realizado"] for r in cot_rows)),
        "comprometido": f2(sum(r["comprometido"] for r in cot_rows)),
        "impacto_pendente": f2(sum(r["impacto"] for r in cot_rows if not r["aprovado"])),
        "saldo": f2(sum(r["saldo"] for r in cot_rows)),
        "cotacoes_com_eap": sum(1 for r in cot_rows if r.get("eap") and r.get("eap") != "—"),
    }
    cotacao_analysis = {}
    for cot, grp_rows in itertools.groupby(
        sorted(cot_rows, key=lambda r: (str(r.get("cotacao", "")), str(r.get("insumo", "")))),
        key=lambda r: str(r.get("cotacao", "") or "—")
    ):
        rows = list(grp_rows)
        if not rows:
            continue
        solicitacoes = sorted(set(str(r.get("solicitacao", "") or "—") for r in rows if str(r.get("solicitacao", "") or "—") != "—"))
        titulos = sorted(set(t for r in rows for t in (r.get("titulos") or [])))
        nfs = sorted(set(n for r in rows for n in (r.get("nfs") or [])))
        credores = sorted(set(c for r in rows for c in (r.get("credores") or [])))
        data_vals = [r.get("data") for r in rows if r.get("data") and r.get("data") != "—"]
        aprov_vals = [r.get("aprovacao") for r in rows if r.get("aprovacao") and r.get("aprovacao") != "—"]
        aprovado = any(bool(r.get("aprovado")) for r in rows)
        itens = []
        seen_refs = set()
        total_orcado_eap = 0.0
        total_realizado_eap = 0.0
        total_comprometido_eap = 0.0
        total_saldo_eap = 0.0
        for r in rows:
            ref = str(r.get("eap", "") or "").strip()
            unique_ref = ref and ref != "—" and ref not in seen_refs
            if unique_ref:
                seen_refs.add(ref)
                orc_ref = f2(r.get("orcado_eap", 0))
                real_ref = f2(r.get("realizado_eap", 0))
                comp_ref = f2(r.get("comprometido_eap", 0))
                saldo_ref = f2(r.get("saldo_eap", 0))
                total_orcado_eap = f2(total_orcado_eap + orc_ref)
                total_realizado_eap = f2(total_realizado_eap + real_ref)
                total_comprometido_eap = f2(total_comprometido_eap + comp_ref)
                total_saldo_eap = f2(total_saldo_eap + saldo_ref)
            else:
                orc_ref = real_ref = comp_ref = saldo_ref = 0.0
            itens.append({
                "solicitacao": r.get("solicitacao", "—"),
                "insumo": r.get("insumo", "—"),
                "detalhe": r.get("detalhe", ""),
                "grupo": r.get("grupo", ""),
                "un": r.get("un", ""),
                "qtd": f2(r.get("qtd", 0)),
                "unit": f2(r.get("unit", 0)),
                "valor_cotado": f2(r.get("apropriado", 0)),
                "realizado": f2(r.get("realizado", 0)),
                "comprometido": f2(r.get("comprometido", 0)),
                "impacto": f2(r.get("impacto", 0)),
                "saldo": f2(r.get("saldo", 0)),
                "eap": ref or "—",
                "item_orc": r.get("item_orc", ""),
                "orcado_eap": orc_ref,
                "realizado_eap": real_ref,
                "comprometido_eap": comp_ref,
                "saldo_eap": saldo_ref,
                "titulos": r.get("titulos", []),
                "nfs": r.get("nfs", []),
                "credores": r.get("credores", []),
            })
        valor_cotado = f2(sum(r.get("apropriado", 0) for r in rows))
        realizado = f2(sum(r.get("realizado", 0) for r in rows))
        comprometido = f2(sum(r.get("comprometido", 0) for r in rows))
        saldo = f2(sum(r.get("saldo", 0) for r in rows))
        cotacao_analysis[cot] = {
            "cotacao": cot,
            "solicitacoes": solicitacoes,
            "data": min(data_vals) if data_vals else "—",
            "aprovacao": min(aprov_vals) if aprov_vals else "—",
            "status": "Aprovada" if aprovado else "Não aprovada",
            "aprovado": aprovado,
            "status_pedido": _clean_txt(rows[0].get("status_pedido", "")) or "—",
            "comprador": _clean_txt(rows[0].get("comprador", "")) or "—",
            "cc": rows[0].get("cc", ""),
            "valor_cotado": valor_cotado,
            "realizado": realizado,
            "comprometido": comprometido,
            "saldo": saldo,
            "impacto": f2(sum(r.get("impacto", 0) for r in rows)),
            "total_orcado_eap": total_orcado_eap,
            "total_realizado_eap": total_realizado_eap,
            "total_comprometido_eap": total_comprometido_eap,
            "total_saldo_eap": total_saldo_eap,
            "titulos": titulos,
            "nfs": nfs,
            "credores": credores,
            "itens": itens,
        }
    return {
        "summary": summary,
        "solicitacoes": sorted(sol_rows, key=lambda r: r["data"], reverse=True),
        "cotacoes": sorted(cot_rows, key=lambda r: (r["data"], r["cotacao"]), reverse=True),
        "pedidos": sorted(ped_rows, key=lambda r: (r["data"], r["pedido"]), reverse=True),
        "cotacao_analysis": cotacao_analysis,
    }

TAX_DOC_PREFIXES = {
    "ISS", "INSS", "IRRF", "CSLL", "PIS", "COFINS", "DARF", "GPS",
    "PCC", "RF", "GI", "DUAM", "GFIP", "GRF", "DAS", "ICMS", "IPI",
    "RET", "RETENCAO", "RETENÇÃO",
}
# Estes credores representam custos de incorporação. No dashboard,
# títulos fora do CC 1011 podem ser exibidos em uma reapropriação gerencial.
# Impostos/retenções vinculados a estes credores seguem a mesma origem.
INCORPORATION_CREDITOR_CODES = set()
INCORPORATION_CREDITOR_NAMES = {
    "ALTCES ARQUITETURA",
    "STUDIO E PROJETOS JATOBA",
    "PAR GESTAO",
    "HIRATA",
    "RAIZ ENGENHARIA",
    "GUIOTTICI",
    "DOMO ENGENHARIA",
    "BSC ENGENHEIROS ASSOCIADOS",
    "ET ENGENHARIA E PROJETOS",
    "AR VERTI",
    "MARCUS BASILIO PROJETOS",
    "EMPORIO",
    "PROASSP",
    "SUPPORT",
    "ALEX CARV",
    "TOR ENGENHARIA",
    "MITECH",
    "GABRIEL",
    "PLANEJAR",
    "VILIP",
    "CASSIA",
}
PAYROLL_CONSUMO_CODE = "01.004.002.010"
PAYROLL_CONSUMO_PARENT_CODE = "01.004.002"
PAYROLL_CONSUMO_PARENT_NAME = "01.004.002-MAO DE OBRA PROPRIA"
PAYROLL_CONSUMO_NAME = "Encargos de m.o. propria (Impostos da Folha, 13 salario, Ferias)"
TAX_SUBJECT_STOPWORDS = {
    "ISS", "INSS", "IRRF", "CSLL", "PIS", "COFINS", "DARF", "GPS",
    "PCC", "RF", "GI", "DUAM", "GFIP", "GRF", "DAS", "ICMS", "IPI",
    "RET", "RETENCAO", "IMPOSTO", "TRIBUTO", "GUIA", "MULTA", "FGTS",
}

def _tax_norm(value):
    text = unicodedata.normalize("NFKD", str(value or "").upper())
    return re.sub(r"[^A-Z0-9]", "", "".join(c for c in text if not unicodedata.combining(c)))

def _fold_text(*values):
    text = " ".join(str(value or "") for value in values)
    text = unicodedata.normalize("NFKD", text.upper())
    return "".join(c for c in text if not unicodedata.combining(c))

def _is_synergia_payroll_charge(document="", observation="", creditor="", extra=""):
    """Encargos/impostos de folha propria do Synergia seguem o titulo original da folha."""
    folded = _fold_text(document, observation, creditor, extra)
    compact = re.sub(r"[^A-Z0-9]", "", folded)

    # Regra solicitada: todos os titulos com DARF, INSS, FGTS, GFIP, GRRF, GPS, IRRF
    # ou sindicato devem apontar para o insumo de encargos de mao de obra propria.
    if any(token in compact for token in (
        "DARF", "INSS", "FGTS", "GFIP", "GRF", "GRRF", "GPS", "IRRF",
        "SINDICATO", "SINDICAT", "SINDTRAB", "SINDICAL",
    )):
        return True
    if re.search(r"\bSIND\b", folded):
        return True
    if "PREVIDENCIA" in compact and "SOCIAL" in compact:
        return True

    return False

def _is_payroll_correctly_allocated(value):
    return _norm_rel_code(value) == _norm_rel_code(PAYROLL_CONSUMO_CODE)

def _payroll_period_candidates(document="", observation="", reference_date=""):
    text = _fold_text(document, observation)
    candidates = []

    def add(month, year):
        try:
            month_i = int(str(month).strip())
        except Exception:
            return
        if month_i < 1 or month_i > 12:
            return
        year_s = str(year or "").strip()
        if len(year_s) == 2:
            year_s = "20" + year_s
        if not re.fullmatch(r"20\d{2}", year_s):
            return
        key = f"FOLHA:{month_i:02d}.{year_s}"
        if key not in candidates:
            candidates.append(key)

    for match in re.finditer(r"(?<!\d)(0?[1-9]|1[0-2])\s*[./-]\s*(20\d{2}|\d{2})(?!\d)", text):
        add(match.group(1), match.group(2))

    ref_text = str(reference_date or "").strip()
    ref_date = pd.to_datetime(
        reference_date,
        errors="coerce",
        dayfirst=not bool(re.match(r"^\d{4}-\d{2}-\d{2}", ref_text)),
    )
    ref_year = str(ref_date.year) if pd.notna(ref_date) else ""
    if ref_year:
        for match in re.finditer(
            r"\b(?:GPS|GFIP|FGTS|GRF|GRRF|DARF|INSS|IRRF|SINDICATO|SINDICAL|SIND|BOL)[A-Z0-9 ./-]{0,18}[/ .-](0?[1-9]|1[0-2])\b",
            text,
        ):
            add(match.group(1), ref_year)
        if not candidates and _is_synergia_payroll_charge(document, observation):
            add(ref_date.month, ref_year)

    return candidates

def _closest_prior_payroll_key(target_key, origin_map):
    """Quando a competencia exata (ex.: FOLHA:08.2026) ainda nao foi apropriada
    na planilha fonte (fechamento de folha do mes atrasado), cai pra folha
    mais recente ja apropriada ANTES dela, como estimativa — a distribuicao
    proporcional por EAP tende a ser parecida mes a mes pro mesmo time de
    obra, entao serve de base ate a apropriacao real da competencia certa
    chegar (quando chegar, o match exato volta a ter prioridade sozinho)."""
    m = re.match(r"^FOLHA:(\d{2})\.(\d{4})$", target_key or "")
    if not m:
        return None
    target = int(m.group(2)) * 12 + int(m.group(1))
    best_key, best_val = None, -1
    for key in (origin_map or {}):
        km = re.match(r"^FOLHA:(\d{2})\.(\d{4})$", key)
        if not km:
            continue
        val = int(km.group(2)) * 12 + int(km.group(1))
        if val < target and val > best_val:
            best_val, best_key = val, key
    return best_key

def _payroll_origin_for_tax(document="", observation="", origin_map=None, reference_date=""):
    candidates = _payroll_period_candidates(document, observation, reference_date)
    for key in candidates:
        origin, rule = _choose_tax_origin(key, "", "", origin_map, reference_date)
        if not origin:
            continue
        origins = []
        for item in origin.get("origins", []) or [origin]:
            row = dict(item)
            row["funcionario_cod"] = row.get("cod", "")
            row["funcionario_servico"] = row.get("servico", "")
            # Regra correta: o imposto/encargo segue a EAP do titulo original
            # de folha de cada funcionario, proporcional ao valor da folha.
            origins.append(row)
        return {
            "key": key,
            "origin": origin,
            "origins": origins,
            "rule": rule,
            "estimado": False,
        }
    # Nenhuma competencia exata bateu (folha do mes ainda nao apropriada na
    # fonte) — estima usando a distribuicao proporcional da folha apropriada
    # mais recente antes da competencia pretendida, deixando claro na regra
    # que e uma estimativa (nao a apropriacao real dessa competencia).
    if candidates:
        wanted = candidates[0]
        fallback_key = _closest_prior_payroll_key(wanted, origin_map)
        if fallback_key:
            origin, rule = _choose_tax_origin(fallback_key, "", "", origin_map, reference_date)
            if origin:
                origins = []
                for item in origin.get("origins", []) or [origin]:
                    row = dict(item)
                    row["funcionario_cod"] = row.get("cod", "")
                    row["funcionario_servico"] = row.get("servico", "")
                    origins.append(row)
                return {
                    "key": fallback_key,
                    "origin": origin,
                    "origins": origins,
                    "rule": (
                        f"ESTIMADO — folha da competência {wanted.split(':', 1)[1]} ainda não apropriada na fonte; "
                        f"rateio usa a distribuição da folha de {fallback_key.split(':', 1)[1]} (mês mais recente já apropriado) "
                        f"como aproximação, proporcional ao valor de cada funcionário."
                    ),
                    "estimado": True,
                }
    return None

def _payroll_consumo_node():
    return {
        "cod": PAYROLL_CONSUMO_CODE,
        "nm": PAYROLL_CONSUMO_NAME,
        "un": "VB",
        "qt": 0,
        "orcado": 0,
        "realizado": 0,
        "a_pagar": 0,
        "saldo": 0,
        "titulos": [],
    }

def _find_tree_node(nodes, code):
    target = _norm_rel_code(code)
    for node in nodes or []:
        if _norm_rel_code(node.get("cod")) == target:
            return node
        found = _find_tree_node(node.get("children", []), target)
        if found:
            return found
    return None

def _ensure_payroll_consumo_leaf(nodes):
    parent = _find_tree_node(nodes, PAYROLL_CONSUMO_PARENT_CODE)
    if not parent:
        return None
    children = parent.setdefault("children", [])
    leaf = _find_tree_node(children, PAYROLL_CONSUMO_CODE)
    if not leaf:
        leaf = _payroll_consumo_node()
        children.append(leaf)
    return leaf

def _append_payroll_consumo_orc4(orc4):
    if orc4 is None or orc4.empty or PAYROLL_CONSUMO_CODE in set(orc4.get("refitemorc", pd.Series(dtype=str)).astype(str)):
        return orc4
    siblings = orc4[orc4["refitemorc"].astype(str).str.startswith(PAYROLL_CONSUMO_PARENT_CODE + ".")].copy()
    if siblings.empty:
        return orc4
    row = siblings.iloc[0].copy()
    row["refitemorc"] = PAYROLL_CONSUMO_CODE
    row["nmitemorc"] = PAYROLL_CONSUMO_NAME
    row["codunid"] = "VB"
    row["qtitemorc"] = 0
    row["vlritemorc"] = 0
    row["n1"] = PAYROLL_CONSUMO_CODE[:2]
    row["n2"] = PAYROLL_CONSUMO_CODE[:6]
    row["n3"] = PAYROLL_CONSUMO_PARENT_CODE
    if "refitemorc3n" in row.index:
        row["refitemorc3n"] = PAYROLL_CONSUMO_PARENT_NAME
    return pd.concat([orc4, row.to_frame().T], ignore_index=True)

def _is_incorporation_creditor(value):
    creditor = _tax_norm(value)
    return any(
        creditor == _tax_norm(name)
        or creditor in _tax_norm(name)
        or _tax_norm(name) in creditor
        for name in INCORPORATION_CREDITOR_NAMES
    )

def _tax_words(value):
    text = unicodedata.normalize("NFKD", str(value or "").upper())
    clean = "".join(c for c in text if not unicodedata.combining(c))
    return re.findall(r"[A-Z0-9]+", clean)

def _tax_subject_tokens(document):
    return [
        word for word in _tax_words(document)
        if len(word) >= 4 and word not in TAX_SUBJECT_STOPWORDS and not word.isdigit()
    ]

def _tax_doc_prefix(value):
    text = str(value or "").upper().strip()
    match = re.match(r"^([A-ZÇÃÕ]+)", text)
    return match.group(1) if match else ""

def _is_tax_document(document="", observation=""):
    obs_norm = _tax_norm(observation)
    return (
        _tax_doc_prefix(document) in TAX_DOC_PREFIXES
        or "RETENCAO" in obs_norm
        or any(term in obs_norm for term in ("IMPOSTO", "TRIBUTO", "CSRF"))
    )

def _tax_original_creditor(observation):
    text = " ".join(str(observation or "").replace("_x000D_", " ").split())
    match = re.search(r"CREDOR\s+(.+?)\s+REFERENTE", text, flags=re.IGNORECASE)
    return match.group(1).strip() if match else ""

def _tax_nf_candidates(document="", observation=""):
    text = str(observation or "").upper()
    found = re.findall(r"(?:NFS?E?|NFE|NF)[/ .-]*0*([0-9]+)", text)
    candidates = []
    for value in found:
        value = value.lstrip("0") or "0"
        if value not in candidates:
            candidates.append(value)
    # Relatórios antigos trazem a NF original no documento do imposto: DARF.634, DUAM.326 etc.
    if _is_tax_document(document, observation):
        match = re.match(r"^[A-ZÇÃÕ]+[/.]\s*0*([0-9]+)\s*$", str(document or "").upper().strip())
        if match:
            value = match.group(1).lstrip("0") or "0"
            if value not in candidates:
                candidates.append(value)
        for value in _payroll_period_candidates(document, observation):
            if value not in candidates:
                candidates.append(value)
    return candidates

def _regular_nf_number(document):
    text = str(document or "").upper().strip()
    match = re.search(r"(?:NFS?E?|NFE|NF|NFA|ME)[/ .-]*0*([0-9]+)", text)
    if match:
        return match.group(1).lstrip("0") or "0"
    return ""

def _build_local_nf_origin_map():
    origins = {}
    if not os.path.exists(APROP_REL_PATH):
        return origins
    try:
        df = pd.read_excel(APROP_REL_PATH, sheet_name="Relatório", header=1)
        required = {"Documento", "Credor/Histórico", "Serviço"}
        if not required.issubset(df.columns):
            return origins
        for _, row in df.iterrows():
            document = str(row.get("Documento", "") or "")
            observation = str(row.get("Observação", "") or "")
            if _is_tax_document(document, observation):
                continue
            nf_num = _regular_nf_number(document)
            code = _norm_rel_code(row.get("Serviço"))
            payroll = re.search(r"FOP[/ .-]*FOLHA\s+(0[1-9]|1[0-2])\.(20\d{2})", document.upper())
            payroll_key = f"FOLHA:{payroll.group(1)}.{payroll.group(2)}" if payroll else ""
            if not code:
                continue
            rec = {
                "nf_num": nf_num or payroll_key,
                "cod": code,
                "servico": str(row.get("Serviço", "") or code),
                "credor": str(row.get("Credor/Histórico", "") or ""),
                "credor_norm": _tax_norm(row.get("Credor/Histórico", "")),
                "documento": document,
                "titulo": str(row.get("Título/Parcela", "") or ""),
                "data": row.get("Data"),
                "valor": f2(row.get("Valor", 0)),
            }
            identity = (rec["cod"], rec["credor_norm"], rec["documento"], rec["titulo"])
            text_bucket = origins.setdefault("__TEXT__", [])
            if not any((r["cod"], r["credor_norm"], r["documento"], r["titulo"]) == identity for r in text_bucket):
                text_bucket.append(rec)
            if nf_num:
                bucket = origins.setdefault(nf_num, [])
                if not any((r["cod"], r["credor_norm"], r["documento"], r["titulo"]) == identity for r in bucket):
                    bucket.append(rec)
            if payroll_key:
                payroll_bucket = origins.setdefault(payroll_key, [])
                if not any((r["cod"], r["credor_norm"], r["documento"], r["titulo"]) == identity for r in payroll_bucket):
                    payroll_bucket.append(rec)
    except Exception:
        return {}
    return origins

def _tax_origin_group(options, rule, allow_multiple_titles=False):
    if not options:
        return None, rule
    if len(options) == 1:
        return options[0], rule
    codes = {option["cod"] for option in options}
    titles = {option.get("titulo", "") for option in options if option.get("titulo")}
    if len(titles) == 1 or allow_multiple_titles:
        total = sum(abs(f2(option.get("valor", 0))) for option in options)
        origins = []
        for option in options:
            item = dict(option)
            item["share"] = abs(f2(option.get("valor", 0))) / total if total else 1 / len(options)
            origins.append(item)
        first = origins[0]
        return {
            "cod": f"Rateio em {len(codes)} EAPs",
            "servico": "; ".join(sorted({option["servico"] for option in origins})),
            "credor": first.get("credor", ""),
            "credor_norm": first.get("credor_norm", ""),
            "documento": first.get("documento", ""),
            "titulo": first.get("titulo", "") if len(titles) == 1 else f"{len(titles)} títulos de folha",
            "origins": origins,
        }, f"{rule} + rateio do título original"
    if len(codes) == 1:
        return options[0], rule
    return None, rule

def _choose_tax_origin(nf_num, creditor="", current_ref="", origin_map=None, reference_date=""):
    options = (origin_map or {}).get(str(nf_num or ""), []) or []
    if not options:
        return None, "NF original não encontrada"
    if str(nf_num).startswith("FOLHA:"):
        return _tax_origin_group(options, "Competência da folha", allow_multiple_titles=True)
    cred_norm = _tax_norm(creditor)
    exact = [
        option for option in options
        if cred_norm and option.get("credor_norm")
        and (
            cred_norm == option["credor_norm"]
            or cred_norm in option["credor_norm"]
            or option["credor_norm"] in cred_norm
        )
    ]
    if exact:
        selected, rule = _tax_origin_group(exact, "NF + credor")
        if selected:
            return selected, rule
    ref_date = pd.to_datetime(reference_date, errors="coerce", dayfirst=True)
    if pd.notna(ref_date) and len(options) > 1:
        dated = []
        for option in options:
            option_date = pd.to_datetime(option.get("data"), errors="coerce", dayfirst=True)
            if pd.notna(option_date):
                dated.append((abs((option_date - ref_date).days), option))
        dated.sort(key=lambda item: item[0])
        if dated and dated[0][0] <= 120:
            closest_days = dated[0][0]
            closest = [option for days, option in dated if days == closest_days]
            selected, rule = _tax_origin_group(closest, "NF + data/competência")
            if selected and (len(dated) == len(closest) or closest_days + 14 < dated[len(closest)][0]):
                return selected, rule
    current_ref = _norm_rel_code(current_ref)
    same_ref = [option for option in options if current_ref and option["cod"] == current_ref]
    if len(same_ref) == 1:
        return same_ref[0], "NF + EAP atual"
    selected, rule = _tax_origin_group(options, "NF única")
    if selected:
        return selected, rule
    identities = {(option["cod"], option["credor_norm"]) for option in options}
    return None, f"NF encontrada em {len(identities)} credores/EAPs; credor original necessário"

def _choose_tax_subject_origin(document, origin_map=None, reference_date=""):
    tokens = _tax_subject_tokens(document)
    if not tokens:
        return None, ""
    options = []
    ref_date = pd.to_datetime(reference_date, errors="coerce", dayfirst=True)
    for option in (origin_map or {}).get("__TEXT__", []) or []:
        doc_words = set(_tax_words(option.get("documento", "")))
        if not all(token in doc_words for token in tokens):
            continue
        option_date = pd.to_datetime(option.get("data"), errors="coerce", dayfirst=True)
        distance = abs((option_date - ref_date).days) if pd.notna(option_date) and pd.notna(ref_date) else 999999
        if distance <= 120:
            options.append((distance, option))
    if not options:
        return None, ""
    closest_days = min(distance for distance, _ in options)
    closest = [option for distance, option in options if distance == closest_days]
    return _tax_origin_group(closest, "Descrição do imposto + data", allow_multiple_titles=False)

def _build_tax_link(document="", observation="", current_ref="", origin_map=None, reference_date=""):
    if _is_synergia_payroll_charge(document, observation):
        payroll_origin = _payroll_origin_for_tax(document, observation, origin_map, reference_date)
        competence = payroll_origin["key"].split(":", 1)[1] if payroll_origin else ""
        origins = payroll_origin["origins"] if payroll_origin else []
        estimado = bool(payroll_origin and payroll_origin.get("estimado"))
        return {
            "status": "VINCULADO",
            "tipo_imposto": "FOLHA_PROPRIA",
            "nf_num": payroll_origin["key"] if payroll_origin else "FOLHA",
            "link_label": (f"Folha {competence} (estimado)" if estimado else f"Folha {competence}") if competence else "Encargos de m.o. propria",
            "cred_orig": f"{len(origins)} funcionário(s)" if origins else "",
            "ref_orig": origins[0].get("cod", "") if len(origins) == 1 else f"Rateio em {len({o.get('cod','') for o in origins if o.get('cod','')})} EAPs" if origins else "",
            "nm_orig": origins[0].get("servico", "") if len(origins) == 1 else "Rateio proporcional pela folha" if origins else "",
            "titulo_orig": (
                payroll_origin["origin"].get("titulo", "") if payroll_origin else "Regra de folha propria"
            ),
            "regra": (
                payroll_origin["rule"] if payroll_origin and payroll_origin.get("rule") else
                "DARF, INSS, FGTS, GFIP, GRF, GRRF, GPS, IRRF e sindicato seguem a apropriação do título original da folha; competência da folha não localizada."
            ),
            "origins": origins,
            "estimado": estimado,
        }
    if not _is_tax_document(document, observation):
        return None
    original_creditor = _tax_original_creditor(observation)
    candidates = _tax_nf_candidates(document, observation)
    for nf_num in candidates:
        origin, rule = _choose_tax_origin(nf_num, original_creditor, current_ref, origin_map, reference_date)
        if origin:
            return {
                "status": "VINCULADO",
                "tipo_imposto": _tax_doc_prefix(document) or "IMPOSTO",
                "nf_num": nf_num,
                "link_label": f"Competência {nf_num.split(':', 1)[1]}" if nf_num.startswith("FOLHA:") else f"NF {nf_num}",
                "cred_orig": origin.get("credor") or original_creditor,
                "ref_orig": origin.get("cod", ""),
                "nm_orig": origin.get("servico", ""),
                "titulo_orig": origin.get("titulo", ""),
                "regra": rule,
                "origins": origin.get("origins", []),
            }
    subject_origin, subject_rule = _choose_tax_subject_origin(document, origin_map, reference_date)
    if subject_origin:
        origin_document = str(subject_origin.get("documento", "") or "").split("/", 1)[-1]
        return {
            "status": "VINCULADO",
            "tipo_imposto": _tax_doc_prefix(document) or "IMPOSTO",
            "nf_num": "?",
            "link_label": f"Origem {origin_document}",
            "cred_orig": subject_origin.get("credor", ""),
            "ref_orig": subject_origin.get("cod", ""),
            "nm_orig": subject_origin.get("servico", ""),
            "titulo_orig": subject_origin.get("titulo", ""),
            "regra": subject_rule,
            "origins": subject_origin.get("origins", []),
        }
    return {
        "status": "NÃO VINCULADO",
        "tipo_imposto": _tax_doc_prefix(document) or "IMPOSTO",
        "nf_num": candidates[0] if candidates else "?",
        "link_label": (
            f"Competência {candidates[0].split(':', 1)[1]}"
            if candidates and candidates[0].startswith("FOLHA:") else
            (f"NF {candidates[0]}" if candidates else "Sem NF/competência")
        ),
        "cred_orig": original_creditor,
        "ref_orig": "",
        "nm_orig": "",
        "titulo_orig": "",
        "regra": (
            _choose_tax_origin(candidates[0], original_creditor, current_ref, origin_map, reference_date)[1]
            if candidates else "NF original não identificada no histórico/documento"
        ),
    }

def build_distorcoes_relatorio_local(tit):
    if not os.path.exists(DIST_REL_PATH):
        return []
    raw = pd.read_excel(DIST_REL_PATH, sheet_name="Relatório", header=None)
    tit_base = tit.copy()
    tit_base["numtitcp_n"] = pd.to_numeric(tit_base.get("numtitcp"), errors="coerce")
    look = {}
    for num, grp in tit_base.dropna(subset=["numtitcp_n"]).groupby("numtitcp_n"):
        g = grp.iloc[0]
        look[str(int(num))] = {
            "credor": str(g.get("nmcredor", "—")),
            "dtvenc": _fmt_date(g.get("dtvenc")),
            "dtpag": _fmt_date(g.get("dtbaixacp")),
            "coddocum": str(g.get("coddocum", "") or ""),
            "nudoc": str(g.get("nudoc", "") or ""),
            "obstit": str(g.get("_observ", g.get("obstit", "")) or ""),
            "refitemorc": str(g.get("refitemorc", "") or ""),
        }

    if os.path.exists(PAGAS_REL_PATH):
        raw_pg = pd.read_excel(PAGAS_REL_PATH, sheet_name="Relatório", header=None)
        data_baixa = "—"
        for _, row in raw_pg.iterrows():
            first = str(row.get(0)).strip() if pd.notna(row.get(0)) else ""
            if first == "Data da baixa" and pd.notna(row.get(3)):
                data_baixa = _fmt_date(row.get(3))
                continue
            if not first or first in ("Credor", "Total do dia", "Total do mês", "Total geral"):
                continue
            lanc = row.get(8)
            if pd.isna(lanc):
                continue
            numtit = str(lanc).split("/")[0].strip()
            if not numtit:
                continue
            doc = str(row.get(5) or "")
            info = look.setdefault(numtit, {})
            if not info.get("credor") or info.get("credor") == "—":
                info["credor"] = first
            if not info.get("dtvenc") or info.get("dtvenc") == "—":
                info["dtvenc"] = _fmt_date(row.get(12))
            if not info.get("dtpag") or info.get("dtpag") == "—":
                info["dtpag"] = data_baixa
            if not info.get("coddocum"):
                info["coddocum"] = doc.split(".", 1)[0].strip() if "." in doc else doc
            if not info.get("nudoc"):
                info["nudoc"] = doc.split(".", 1)[1].strip() if "." in doc else ""

    def _doc_parts(doc, fallback_tipo="", fallback_nf=""):
        doc_s = str(doc or "").strip()
        if "/" in doc_s:
            tipo, nf = doc_s.split("/", 1)
        elif "." in doc_s:
            tipo, nf = doc_s.split(".", 1)
        else:
            tipo, nf = (fallback_tipo or doc_s), fallback_nf
        return doc_s, str(tipo or "—").strip().upper(), str(nf or "—").strip()

    origin_map = _build_local_nf_origin_map()
    # O próprio relatório de distorções pode conter a NF original ainda sem EAP.
    # Nesse caso vinculamos os impostos ao título original, mas mantemos "Sem EAP".
    for _, raw_row in raw.iterrows():
        original_document = str(raw_row.get(0) or "")
        if _is_tax_document(original_document):
            continue
        nf_num = _regular_nf_number(original_document)
        raw_title = raw_row.get(2)
        if not nf_num or pd.isna(raw_title):
            continue
        title_text = str(int(raw_title)) if isinstance(raw_title, (int, float)) else str(raw_title).strip()
        title_info = look.get(title_text.split("/")[0], {})
        rec = {
            "nf_num": nf_num,
            "cod": "Sem EAP",
            "servico": "Título original também está sem apropriação",
            "credor": title_info.get("credor", ""),
            "credor_norm": _tax_norm(title_info.get("credor", "")),
            "documento": original_document,
            "titulo": title_text,
            "data": raw_row.get(3),
            "valor": f2(raw_row.get(13)),
        }
        bucket = origin_map.setdefault(nf_num, [])
        if not any(item.get("titulo") == title_text and item.get("documento") == original_document for item in bucket):
            bucket.append(rec)

    # Layout atual do relatório: cabeçalho na linha 2, com Documento/Título/Serviço.
    try:
        hdr = [str(v).strip() for v in raw.iloc[1].tolist()]
        if (
            "Título/Parcela" in hdr
            and "Documento" in hdr
            and "Credor/Histórico" in hdr
            and "Valor do documento" in hdr
            and "Valor" in hdr
        ):
            df = pd.read_excel(DIST_REL_PATH, sheet_name="Relatório", header=1)
            grouped = {}
            for _, row in df.iterrows():
                numtit_raw = row.get("Título/Parcela")
                valor_doc = pd.to_numeric(row.get("Valor do documento"), errors="coerce")
                valor_aprop = pd.to_numeric(row.get("Valor"), errors="coerce")
                if pd.isna(numtit_raw) or pd.isna(valor_doc):
                    continue
                servico = str(row.get("Serviço", "") or "").strip()
                doc_s, tipo, nf = _doc_parts(row.get("Documento"))
                doc_up = doc_s.upper()
                obs_up = str(row.get("Observação", "") or "").upper()
                if doc_up.startswith(("PPC", "PCT")) or "PREVIS" in doc_up or "PREVIS" in obs_up:
                    continue
                numtit_s = str(numtit_raw).strip()
                num_base = numtit_s.split("/")[0].strip()
                info = look.get(num_base, {})
                tax_link = _build_tax_link(
                    doc_s,
                    info.get("obstit", ""),
                    info.get("refitemorc", ""),
                    origin_map,
                    row.get("Data"),
                )
                rec = grouped.setdefault(numtit_s, {
                    "numtit": numtit_s,
                    "nf": nf or info.get("nudoc", "—") or "—",
                    "documento": doc_s,
                    "tipo": tipo or info.get("coddocum", "—") or "—",
                    "vencimento": info.get("dtvenc", "—"),
                    "pagamento": _fmt_date(row.get("Data da baixa")) or info.get("dtpag", "—"),
                    "credor": str(row.get("Credor/Histórico", "") or info.get("credor", "—")),
                    "emissao": _fmt_date(row.get("Data")),
                    "valor_documento": 0.0,
                    "valor_apropriado_assinado": 0.0,
                    "servicos": set(),
                    "tax_link": tax_link,
                })
                rec["valor_documento"] = max(rec["valor_documento"], abs(float(valor_doc)))
                if pd.notna(valor_aprop):
                    # Estornos aparecem como linhas negativas e devem abater a soma.
                    rec["valor_apropriado_assinado"] += float(valor_aprop)
                if servico:
                    rec["servicos"].add(servico)

            rows = []
            for rec in grouped.values():
                valor_documento = f2(rec.pop("valor_documento", 0))
                valor_apropriado = f2(abs(rec.pop("valor_apropriado_assinado", 0)))
                saldo = f2(max(valor_documento - valor_apropriado, 0))
                # Relatório de distorções: somente títulos sem apropriação ou abaixo de 100%.
                if saldo <= 0:
                    continue
                servicos = sorted(rec.pop("servicos", set()))
                rec.update({
                    "origem": "Sem apropriação" if valor_apropriado <= 0.01 else (servicos[0] if servicos else "Sem EAP informada"),
                    "valor_documento": valor_documento,
                    "valor_apropriado": valor_apropriado,
                    "saldo_apropriar": saldo,
                    "pct_apropriado": f2((valor_apropriado / valor_documento * 100) if valor_documento else 0),
                    "valor": saldo,
                })
                rows.append(rec)
            rows.sort(key=lambda r: (r["pagamento"] == "—", r["pagamento"], r["credor"], r["numtit"]))
            return rows
    except Exception:
        pass

    rows = []
    for _, row in raw.iterrows():
        doc = row.get(0)
        numtit = row.get(2)
        origem = row.get(5)
        pct_obra = pd.to_numeric(row.get(10), errors="coerce")
        valor = pd.to_numeric(row.get(13), errors="coerce")
        if pd.isna(doc) or pd.isna(numtit) or pd.isna(valor):
            continue
        # O total "não apropriado para obra" inclui tanto títulos zerados
        # quanto parcialmente apropriados. Só ficam fora os que chegaram a 100%.
        if pd.notna(pct_obra) and float(pct_obra) >= 100:
            continue
        pct_obra_f = max(0.0, min(float(pct_obra) if pd.notna(pct_obra) else 0.0, 100.0))
        saldo_apropriar = f2(abs(valor))
        fator_pendente = 1 - (pct_obra_f / 100)
        valor_documento = f2(saldo_apropriar / fator_pendente) if fator_pendente > 0 else saldo_apropriar
        valor_apropriado = f2(max(valor_documento - saldo_apropriar, 0))
        numtit_s = str(int(numtit)) if isinstance(numtit, (int, float)) and pd.notna(numtit) else str(numtit).strip()
        info = look.get(numtit_s, {})
        doc_s, tipo, nf = _doc_parts(doc, info.get("coddocum") or str(origem or "—"), info.get("nudoc", ""))
        tax_link = _build_tax_link(
            doc_s,
            info.get("obstit", ""),
            info.get("refitemorc", ""),
            origin_map,
            row.get(3),
        )
        rows.append({
            "numtit": numtit_s,
            "nf": nf or "—",
            "documento": doc_s,
            "tipo": tipo or "—",
            "origem": "Sem apropriação" if pct_obra_f <= 0.01 else str(origem or "Apropriação parcial"),
            "valor_documento": valor_documento,
            "valor_apropriado": valor_apropriado,
            "saldo_apropriar": saldo_apropriar,
            "pct_apropriado": f2(pct_obra_f),
            "valor": saldo_apropriar,
            "vencimento": info.get("dtvenc", "—"),
            "pagamento": info.get("dtpag", "—"),
            "credor": info.get("credor", "—"),
            "emissao": _fmt_date(row.get(3)),
            "tax_link": tax_link,
        })
    rows.sort(key=lambda r: (r["pagamento"] == "—", r["pagamento"], r["credor"], r["numtit"]))
    return rows

def load_incc_di_series():
    """Lê os números-índice da planilha de apoio e completa o último dado oficial."""
    series = {}
    try:
        from openpyxl import load_workbook
        wb = load_workbook(SUPPORT_REPORT_PATH, read_only=True, data_only=True)
        if "INCC" in wb.sheetnames:
            ws = wb["INCC"]
            for row in ws.iter_rows(min_row=3, min_col=1, max_col=3, values_only=True):
                dt = pd.to_datetime(row[0], errors="coerce")
                value = pd.to_numeric(row[2], errors="coerce")
                if pd.notna(dt) and pd.notna(value) and float(value) > 0:
                    series[dt.strftime("%Y-%m")] = float(value)
        wb.close()
    except Exception:
        pass
    # FGV/IBRE: INCC-DI de junho/2026 = +0,78%. O nível parte de maio/2026.
    if "2026-05" in series and "2026-06" not in series:
        series["2026-06"] = round(series["2026-05"] * 1.0078, 3)
    return series

def build_orcamento_relatorio(api_tree=None, title_context=None):
    if not os.path.exists(ORC_REL_PATH):
        return [], {"ok": False, "msg": f"Arquivo nao encontrado: {ORC_REL_PATH}"}, []
    title_context = title_context or {}

    def title_key(value):
        value = str(value or "").strip()
        match = re.search(r"(\d+)(?:\s*[/.-]\s*(\d+))?", value)
        if not match:
            return ""
        return f"{match.group(1)}/{match.group(2) or '1'}"

    def title_base(value):
        return title_key(value).split("/")[0]

    def report_money(value):
        if pd.isna(value):
            return 0.0
        if isinstance(value, (int, float)):
            return f2(value)
        text = re.sub(r"[^0-9,.\-]", "", str(value).strip())
        if "," in text:
            text = text.replace(".", "").replace(",", ".")
        try:
            return f2(float(text))
        except Exception:
            return 0.0

    def excluded_report_doc(value):
        doc = str(value or "").upper().strip()
        return doc.startswith(("PPC", "PCT")) or "PREVIS" in doc

    # Status financeiro do relatório manual: somente Contas Pagas comprova realizado.
    paid_by_title = {}
    paid_date_by_title = {}
    paid_info_by_title = {}
    aprop_info_by_title = {}
    title_aprop_totals = {}
    if os.path.exists(PAGAS_REL_PATH):
        raw_paid = pd.read_excel(PAGAS_REL_PATH, sheet_name="Relatório", header=None)
        baixa = "—"
        for _, row in raw_paid.iterrows():
            first = str(row.get(0)).strip() if pd.notna(row.get(0)) else ""
            if first == "Data da baixa":
                baixa = _fmt_date(row.get(3))
                continue
            if pd.isna(row.get(8)) or excluded_report_doc(row.get(5)):
                continue
            key = title_key(row.get(8))
            if not key or not title_base(key).isdigit():
                continue
            # A coluna "Líquido" (23) só vem preenchida quando há
            # acréscimo/desconto na baixa; na maioria das linhas do relatório
            # ela fica em branco e o valor pago real está em "Valor baixa"
            # (14) — sem esse fallback, quase todas as linhas eram
            # descartadas e o realizado somava ~0.
            liquid = report_money(row.get(23)) or report_money(row.get(14))
            if liquid == 0:
                continue
            paid_by_title[key] = f2(paid_by_title.get(key, 0) + liquid)
            paid_date_by_title[key] = baixa
            info = paid_info_by_title.setdefault(key, {
                "credor": first,
                "documento": str(row.get(5) or "—"),
                "data": baixa,
            })
            info["credor"] = info.get("credor") or first
            info["documento"] = info.get("documento") or str(row.get(5) or "—")
            info["data"] = baixa

    payable_by_title = {}
    if os.path.exists(APAGAR_REL_PATH):
        raw_payable = pd.read_excel(APAGAR_REL_PATH, sheet_name="Relatório", header=None)
        compact_payable = raw_payable.shape[1] <= 20
        for _, row in raw_payable.iterrows():
            title_value = row.get(2) if compact_payable else row.get(25)
            document_value = row.get(1) if compact_payable else row.get(15)
            total_value = row.get(9) if compact_payable else row.get(62)
            if pd.isna(title_value) or excluded_report_doc(document_value):
                continue
            key = title_key(title_value)
            if not key or not title_base(key).isdigit():
                continue
            total = report_money(total_value)
            if total == 0:
                continue
            payable_by_title[key] = f2(payable_by_title.get(key, 0) + total)

    orc = pd.read_excel(ORC_REL_PATH, sheet_name="Orçamento", header=10)
    orc.columns = [str(c).strip() for c in orc.columns]
    orc["ITEMIZAÇÃO"] = orc["ITEMIZAÇÃO"].astype(str).str.strip()

    rel_code_set = set()
    leaf_mask = orc["ITEMIZAÇÃO"].str.match(r"^\d+\.\d+\.\d+\.\d+$", na=False)
    for _, r in orc[leaf_mask].iterrows():
        code = _norm_rel_code(r.get("ITEMIZAÇÃO"))
        if not code:
            continue
        rel_code_set.add(code)

    api_title_by_num = {}
    api_realizado_by_num = {}
    def collect_api_titles(nodes):
        for node in nodes or []:
            node_code = _norm_rel_code(node.get("cod"))
            for t in node.get("titulos", []) or []:
                num = str(t.get("numtit", "")).split("/")[0].strip()
                if num:
                    api_t = dict(t)
                    api_t["ref"] = node_code
                    api_t["item"] = node.get("nm", "")
                    api_title_by_num[num] = api_t
                    if str(t.get("status", "")).upper() != "A PAGAR" and f2(t.get("vlr_pg", 0)) != 0:
                        rec = api_realizado_by_num.setdefault(num, {
                            "numtit": num, "credor": str(t.get("credor", "")), "valor": 0,
                            "eap": node_code, "item": node.get("nm", ""), "data": str(t.get("dt_baixa", "—")),
                        })
                        rec["valor"] = f2(rec["valor"] + t.get("vlr_pg", 0))
            collect_api_titles(node.get("children", []))
    collect_api_titles(api_tree or [])

    def rel_ret_from_api(api_t, cod, item_nm=""):
        """Na aba Relatório, a apropriação exibida deve seguir a planilha local."""
        if not api_t or not (api_t.get("ret_alert") or api_t.get("ret_aprop")):
            return None, None
        item_txt = str(item_nm or "")
        m_ref = re.search(r"\d+(?:\.\d+){3}", item_txt)
        ref = m_ref.group(0) if m_ref else str(cod or "")
        return None, {
            "tipo": "APROP_RELATORIO",
            "ref_orig": ref,
            "nm_orig": str(item_nm or api_t.get("item") or ref),
            "nf_num": (api_t.get("ret_alert") or api_t.get("ret_aprop") or {}).get("nf_num", "?"),
            "cred_orig": (api_t.get("ret_alert") or api_t.get("ret_aprop") or {}).get("cred_orig", ""),
        }

    realizado_by_code = {}
    a_pagar_by_code = {}
    titulos_by_code = {}
    rel_realizado_by_num = {}
    aprop_title_keys = set()
    multi_aprop_titulos = []
    # nf_original_local e usado mais abaixo mesmo quando o relatorio de
    # Apropriacoes nao existe/nao pode ser lido (ex.: logo apos um sync do
    # Sienge, antes desse relatorio especifico terminar de baixar) — precisa
    # estar sempre definido, nao so dentro do bloco condicional abaixo.
    nf_original_local = _build_local_nf_origin_map()
    if os.path.exists(APROP_REL_PATH):
        aprop = pd.read_excel(APROP_REL_PATH, sheet_name="Relatório", header=1)
        if "Serviço" not in aprop.columns or "Valor" not in aprop.columns:
            # O Sienge as vezes exporta esse relatorio no formato hierarquico
            # (agrupado por Periodo > Celula construtiva > Etapa > Subetapa >
            # Servico, com "Total ..." entre blocos) em vez de tabela plana —
            # sem isso o relatorio inteiro ficava sem "Serviço"/"Valor" e o
            # Realizado saia zerado silenciosamente. Achata para o mesmo
            # formato de colunas que o resto do codigo abaixo espera.
            aprop = _parse_aprop_hierarquico(APROP_REL_PATH)
        if "Serviço" in aprop.columns and "Valor" in aprop.columns:
            aprop = aprop.copy()
            aprop["_cod"] = aprop["Serviço"].apply(_norm_rel_code)
            aprop["Valor"] = pd.to_numeric(aprop["Valor"], errors="coerce").fillna(0)
            doc_col = aprop["Documento"].astype(str).str.upper() if "Documento" in aprop.columns else ""
            obs_col = aprop["Observação"].astype(str).str.upper() if "Observação" in aprop.columns else ""
            if hasattr(doc_col, "str"):
                aprop = aprop[~(doc_col.str.startswith("PPC") | doc_col.str.startswith("PCT") | doc_col.str.contains("PREVIS", na=False) | obs_col.str.contains("PREVIS", na=False))].copy()
            aprop["_is_payroll_consumo"] = aprop.apply(
                lambda row: _is_synergia_payroll_charge(
                    row.get("Documento", ""),
                    row.get("Observação", ""),
                    row.get("Credor/Histórico", ""),
                    row.get("Título/Parcela", ""),
                ),
                axis=1,
            )

            aprop["_cod_lancado"] = aprop["_cod"]
            def rel_tax_link(row):
                num_base = title_base(row.get("Título/Parcela", ""))
                api_t = api_title_by_num.get(num_base, {})
                api_ret = api_t.get("ret_alert") or api_t.get("ret_aprop") or {}
                observation = str(row.get("Observação", "") or "")
                if api_ret.get("cred_orig") or api_ret.get("nf_num"):
                    observation += (
                        f" Retenção do credor {api_ret.get('cred_orig', '')}"
                        f" referente ao(s) doc(s) NFSE/{api_ret.get('nf_num', '')}."
                    )
                return _build_tax_link(
                    row.get("Documento", ""),
                    observation,
                    row.get("_cod_lancado", ""),
                    nf_original_local,
                    row.get("Data"),
                )
            aprop["_tax_link"] = aprop.apply(
                rel_tax_link,
                axis=1,
            )
            expanded_rows = []
            linked_tax_groups = {}
            for _, tax_row in aprop.iterrows():
                tax_link = tax_row.get("_tax_link") if isinstance(tax_row.get("_tax_link"), dict) else None
                if tax_link and tax_link.get("status") == "VINCULADO":
                    # Encargos de folha também passam por aqui para manter o rateio
                    # proporcional por funcionário, em vez de virar uma linha única.
                    key = title_key(tax_row.get("Título/Parcela", ""))
                    linked_tax_groups.setdefault(key, []).append(tax_row)
                else:
                    expanded_rows.append(tax_row.to_dict())
            for tax_rows in linked_tax_groups.values():
                base_row = tax_rows[0]
                tax_link = base_row["_tax_link"]
                total_tax = sum(float(row.get("Valor", 0)) for row in tax_rows)
                origins = tax_link.get("origins", []) or [{
                    "cod": tax_link.get("ref_orig", ""),
                    "servico": tax_link.get("nm_orig", ""),
                    "credor": tax_link.get("cred_orig", ""),
                    "titulo": tax_link.get("titulo_orig", ""),
                    "share": 1,
                }]
                remaining = total_tax
                for index, origin in enumerate(origins):
                    item = base_row.to_dict()
                    allocated = remaining if index == len(origins) - 1 else total_tax * float(origin.get("share", 0))
                    remaining -= allocated
                    item["Valor"] = allocated
                    item["_tax_link"] = {
                        **tax_link,
                        "ref_orig": origin.get("cod", ""),
                        "nm_orig": origin.get("servico", ""),
                        "cred_orig": origin.get("credor", tax_link.get("cred_orig", "")),
                        "titulo_orig": origin.get("titulo", tax_link.get("titulo_orig", "")),
                        "origins": [],
                    }
                    expanded_rows.append(item)
            aprop = pd.DataFrame(expanded_rows)
            if "_is_payroll_consumo" not in aprop.columns:
                aprop["_is_payroll_consumo"] = False
            aprop["_is_ret"] = aprop["_tax_link"].apply(lambda value: isinstance(value, dict))
            aprop["_ret_nf_num"] = aprop["_tax_link"].apply(
                lambda value: value.get("nf_num", "") if isinstance(value, dict) else ""
            )
            aprop["_orig_info"] = aprop["_tax_link"].apply(
                lambda value: {
                    "cod": value.get("ref_orig", ""),
                    "servico": value.get("nm_orig", ""),
                    "credor": value.get("cred_orig", ""),
                    "documento": f"NF/{value.get('nf_num', '')}",
                    "titulo": value.get("titulo_orig", ""),
                } if isinstance(value, dict) and value.get("status") == "VINCULADO" else None
            )
            aprop["_cod"] = aprop.apply(
                lambda r: _norm_rel_code(r.get("_tax_link", {}).get("ref_orig", ""))
                if bool(r.get("_is_payroll_consumo")) and isinstance(r.get("_tax_link"), dict) and r.get("_tax_link", {}).get("ref_orig")
                else (r["_orig_info"]["cod"] if isinstance(r.get("_orig_info"), dict) else r["_cod"]),
                axis=1,
            )
            aprop["_title_key"] = aprop["Título/Parcela"].apply(title_key)
            aprop_title_keys = set(aprop["_title_key"].dropna().astype(str))
            title_aprop_totals = aprop.groupby("_title_key")["Valor"].sum().to_dict()

            # Títulos lançados em MAIS DE UMA apropriação (EAP) — depois de toda
            # a realocação de retenção/folha acima, "_cod" já é o EAP real de
            # cada linha. Um mesmo título/parcela com mais de um "_cod"
            # distinto significa que o valor foi dividido entre itens de
            # orçamento diferentes (ex.: uma NF que cobre dois serviços).
            for _key, _grp in aprop.groupby("_title_key"):
                if not _key:
                    continue
                _codes_presentes = sorted({str(c) for c in _grp["_cod"] if str(c or "").strip()})
                if len(_codes_presentes) <= 1:
                    continue
                _first = _grp.iloc[0]
                _breakdown = []
                for _cod, _cgrp in _grp.groupby("_cod"):
                    if not str(_cod or "").strip():
                        continue
                    _breakdown.append({
                        "eap": str(_cod),
                        "servico": str(_cgrp.iloc[0].get("Serviço", "") or ""),
                        "valor": f2(_cgrp["Valor"].sum()),
                    })
                _breakdown.sort(key=lambda x: -abs(x["valor"]))
                multi_aprop_titulos.append({
                    "titulo": str(_key),
                    "credor": str(_first.get("Credor/Histórico", "—")),
                    "documento": str(_first.get("Documento", "—")),
                    "data_baixa": _fmt_date(_first.get("Data da baixa")),
                    "valor_total": f2(_grp["Valor"].sum()),
                    "n_apropriacoes": len(_breakdown),
                    "apropriacoes": _breakdown,
                })
            multi_aprop_titulos.sort(key=lambda x: -abs(x["valor_total"]))
            for key, group in aprop.groupby("_title_key"):
                first_row = group.iloc[0]
                aprop_info_by_title[str(key)] = {
                    "credor": str(first_row.get("Credor/Histórico", "—")),
                    "documento": str(first_row.get("Documento", "—")),
                    "data": _fmt_date(first_row.get("Data da baixa")),
                    "eap": str(first_row.get("_cod_lancado", "") or ""),
                    "servico": str(first_row.get("Serviço", "") or ""),
                }

            def allocate_financial_column(column, source_for_key):
                aprop[column] = 0.0
                for key, indexes in aprop.groupby("_title_key").groups.items():
                    indexes = list(indexes)
                    total_aprop = float(title_aprop_totals.get(key, 0))
                    source_total = f2(source_for_key(key))
                    if not key or source_total == 0 or total_aprop == 0:
                        continue
                    remaining = source_total
                    for position, index in enumerate(indexes):
                        allocated = remaining if position == len(indexes) - 1 else f2(
                            source_total * (float(aprop.at[index, "Valor"]) / total_aprop)
                        )
                        remaining = f2(remaining - allocated)
                        aprop.at[index, column] = allocated

            allocate_financial_column("_paid_alloc", lambda key: paid_by_title.get(key, 0))
            allocate_financial_column(
                "_payable_alloc",
                lambda key: 0 if key in paid_by_title else payable_by_title.get(key, title_aprop_totals.get(key, 0)),
            )
            realizado_by_code = aprop.groupby("_cod")["_paid_alloc"].sum().to_dict()
            a_pagar_by_code = aprop.groupby("_cod")["_payable_alloc"].sum().to_dict()
            for cod, grp in aprop.groupby("_cod"):
                rows = []
                for _, row in grp.iterrows():
                    numtit = str(row.get("Título/Parcela", "—"))
                    title_full = title_key(numtit)
                    num_base = title_base(numtit)
                    api_t = api_title_by_num.get(num_base, {})
                    title_ctx = title_context.get(num_base, {})
                    report_creditor = str(row.get("Credor/Histórico", "—"))
                    cc_real = str(title_ctx.get("cc_real", "") or "1016")
                    orig_info = row.get("_orig_info") if isinstance(row.get("_orig_info"), dict) else None
                    local_item = orig_info.get("servico") if orig_info else row.get("Serviço", "")
                    tax_link = row.get("_tax_link") if isinstance(row.get("_tax_link"), dict) else None
                    payroll_consumo = bool(row.get("_is_payroll_consumo"))
                    incorporation_reap = bool(
                        cc_real != "1011"
                        and (
                            title_ctx.get("incorporacao_reap")
                            or _is_incorporation_creditor(report_creditor)
                            or (
                                tax_link
                                and tax_link.get("status") == "VINCULADO"
                                and _is_incorporation_creditor(tax_link.get("cred_orig", ""))
                            )
                        )
                    )
                    if payroll_consumo:
                        ret_alert = None
                        link = tax_link if isinstance(tax_link, dict) else {}
                        ret_aprop = {
                            "tipo": "FOLHA_PROPRIA",
                            "status": link.get("status", "VINCULADO"),
                            "ref_orig": link.get("ref_orig") or cod,
                            "nm_orig": link.get("nm_orig") or local_item,
                            "nf_num": link.get("nf_num") or str(row.get("Documento", "") or "FOLHA"),
                            "cred_orig": link.get("cred_orig") or report_creditor,
                            "titulo_orig": link.get("titulo_orig") or "Regra de folha propria",
                            "regra": link.get("regra") or "DARF, INSS, FGTS, GFIP, GRF, GRRF, GPS, IRRF e sindicato seguem a apropriação do título original da folha; rateio proporcional por funcionário.",
                            "tipo_imposto": "FOLHA_PROPRIA",
                            "link_label": link.get("link_label") or f"Folha -> {link.get('ref_orig') or cod}",
                            "origins": link.get("origins", []),
                        }
                        local_item = f"{ret_aprop.get('ref_orig', cod)} - {ret_aprop.get('nm_orig', local_item)}"
                    elif tax_link and tax_link.get("status") == "VINCULADO":
                        ret_alert = None
                        ret_aprop = {
                            "tipo": "APROP_RELATORIO",
                            "ref_orig": tax_link.get("ref_orig", cod),
                            "nm_orig": tax_link.get("nm_orig", local_item),
                            "nf_num": tax_link.get("nf_num", "?"),
                            "cred_orig": tax_link.get("cred_orig", ""),
                            "titulo_orig": tax_link.get("titulo_orig", ""),
                            "regra": tax_link.get("regra", ""),
                            "tipo_imposto": tax_link.get("tipo_imposto", "IMPOSTO"),
                            "link_label": tax_link.get("link_label", f"NF {tax_link.get('nf_num', '?')}"),
                        }
                    elif tax_link:
                        ret_aprop = None
                        ret_alert = {
                            "tipo": "SEM_VINCULO",
                            "ref_orig": "Sem vínculo",
                            "nm_orig": tax_link.get("regra", "Origem não identificada"),
                            "nf_num": tax_link.get("nf_num", "?"),
                            "cred_orig": tax_link.get("cred_orig", ""),
                            "tipo_imposto": tax_link.get("tipo_imposto", "IMPOSTO"),
                            "link_label": tax_link.get("link_label", f"NF {tax_link.get('nf_num', '?')}"),
                        }
                    else:
                        ret_alert, ret_aprop = rel_ret_from_api(api_t, cod, local_item)
                    is_paid = title_full in paid_by_title
                    if is_paid:
                        rel_rec = rel_realizado_by_num.setdefault(num_base, {
                            "numtit": numtit,
                            "credor": str(row.get("Credor/Histórico", "—")),
                            "valor": 0,
                            "eap": cod,
                            "item": str(local_item or "—"),
                            "data": paid_date_by_title.get(title_full, "—"),
                            "documento": str(row.get("Documento", "—")),
                        })
                        rel_rec["valor"] = f2(rel_rec["valor"] + row.get("_paid_alloc", 0))
                    rows.append({
                        "numtit": numtit,
                        "credor": report_creditor,
                        "vlr_ap": f2(row.get("Valor", 0)),
                        "vlr_pg": f2(row.get("_paid_alloc", 0)),
                        "vlr_a_pagar": f2(row.get("_payable_alloc", 0)),
                        "status": "PAGO" if is_paid else "A PAGAR",
                        "dt_baixa": paid_date_by_title.get(title_full, "—") if is_paid else "—",
                        "dt_venc": str(api_t.get("dt_venc", "—")),
                        "tpbaixa": "",
                        "observ": str(row.get("Observação", "") or api_t.get("observ", "") or ""),
                    "ref_lancada": str(row.get("_cod_lancado", "") or ""),
                        "ref_original_nf": str(cod or ""),
                        "pedidos": api_t.get("pedidos", []),
                        "contratos": api_t.get("contratos", []),
                        "nfs": api_t.get("nfs", []) or ([row.get("_ret_nf_num")] if row.get("_is_ret") and row.get("_ret_nf_num") else []),
                        "ret_alert": ret_alert,
                        "ret_aprop": ret_aprop,
                        "payroll_consumo": payroll_consumo,
                        "cc_real": cc_real,
                        "codcredor": title_ctx.get("codcredor"),
                        "incorporacao_reap": incorporation_reap,
                    })
                merged_rows = {}
                for item in rows:
                    merge_key = (
                        str(item.get("numtit", "")),
                        str(item.get("credor", "")),
                        str(item.get("status", "")),
                        str(item.get("dt_baixa", "")),
                        str(item.get("dt_venc", "")),
                        str(item.get("ref_lancada", "")),
                        str(item.get("ref_original_nf", "")),
                        str(item.get("cc_real", "")),
                        bool(item.get("incorporacao_reap")),
                    )
                    current = merged_rows.get(merge_key)
                    if not current:
                        current = {**item}
                        current["pedidos"] = list(item.get("pedidos") or [])
                        current["contratos"] = list(item.get("contratos") or [])
                        current["nfs"] = list(item.get("nfs") or [])
                        merged_rows[merge_key] = current
                        continue
                    current["vlr_ap"] = f2(current.get("vlr_ap", 0) + item.get("vlr_ap", 0))
                    current["vlr_pg"] = f2(current.get("vlr_pg", 0) + item.get("vlr_pg", 0))
                    current["vlr_a_pagar"] = f2(current.get("vlr_a_pagar", 0) + item.get("vlr_a_pagar", 0))
                    for field in ("pedidos", "contratos", "nfs"):
                        seen = {str(v) for v in current.get(field, [])}
                        for value in item.get(field, []) or []:
                            if str(value) not in seen:
                                current.setdefault(field, []).append(value)
                                seen.add(str(value))
                    if not current.get("ret_alert") and item.get("ret_alert"):
                        current["ret_alert"] = item.get("ret_alert")
                    if not current.get("ret_aprop") and item.get("ret_aprop"):
                        current["ret_aprop"] = item.get("ret_aprop")
                titulos_by_code[cod] = [
                    item for item in merged_rows.values()
                    if abs(f2(item.get("vlr_ap", 0))) > 0.01
                    or abs(f2(item.get("vlr_pg", 0))) > 0.01
                    or abs(f2(item.get("vlr_a_pagar", 0))) > 0.01
                ]

    divergencias_realizado = []
    for num, rel in rel_realizado_by_num.items():
        api = api_realizado_by_num.get(num)
        if not api:
            tipo = "No relatório, não está na API"
        else:
            dif_val = f2(rel.get("valor", 0) - api.get("valor", 0))
            if abs(dif_val) <= 0.01:
                continue
            tipo = "Valor diferente"
        api = api or {}
        divergencias_realizado.append({
            "tipo": tipo,
            "numtit": rel.get("numtit", num),
            "credor_rel": rel.get("credor", "—"),
            "credor_api": api.get("credor", "—"),
            "eap_rel": rel.get("eap", "—"),
            "eap_api": api.get("eap", "—"),
            "valor_rel": f2(rel.get("valor", 0)),
            "valor_api": f2(api.get("valor", 0)),
            "diferenca": f2(rel.get("valor", 0) - api.get("valor", 0)),
            "data_rel": rel.get("data", "—"),
            "data_api": api.get("data", "—"),
            "documento_rel": rel.get("documento", "—"),
        })
    for num, api in api_realizado_by_num.items():
        if num in rel_realizado_by_num:
            continue
        divergencias_realizado.append({
            "tipo": "Na API, não está no relatório",
            "numtit": num,
            "credor_rel": "—",
            "credor_api": api.get("credor", "—"),
            "eap_rel": "—",
            "eap_api": api.get("eap", "—"),
            "valor_rel": 0,
            "valor_api": f2(api.get("valor", 0)),
            "diferenca": f2(0 - api.get("valor", 0)),
            "data_rel": "—",
            "data_api": api.get("data", "—"),
            "documento_rel": "—",
        })
    divergencias_realizado.sort(key=lambda r: (r["tipo"], str(r["numtit"])))

    if os.path.exists(APAGAR_REL_PATH):
        raw = pd.read_excel(APAGAR_REL_PATH, sheet_name="Relatório", header=None)
        compact_payable = raw.shape[1] <= 20
        venc = "—"
        for _, row in raw.iterrows():
            vals = [row.get(i) for i in range(len(row))]
            first = str(vals[0]).strip() if pd.notna(vals[0]) else ""
            if first == "Data de vencimento":
                venc_val = next((v for v in vals[1:] if pd.notna(v)), "—")
                venc = str(venc_val)
                continue
            if not first or first in ("Credor", "Total do dia", "Total do mês a vencer", "Total geral"):
                continue
            if first.startswith("Obs:") or first.startswith("Empresa") or first.startswith("Obra") or first.startswith("Período"):
                continue
            title_index, document_index, total_index = ((2, 1, 9) if compact_payable else (25, 15, 62))
            if len(vals) <= total_index or pd.isna(vals[title_index]):
                continue
            doc = str(vals[document_index] or "").upper()
            if doc.startswith("PPC") or doc.startswith("PCT") or "PREVIS" in doc:
                continue
            numtit = title_key(vals[title_index])
            num_base = title_base(vals[title_index])
            # Já classificado e distribuído pela EAP no relatório de apropriações.
            if numtit in aprop_title_keys:
                continue
            api_t = api_title_by_num.get(num_base, {})
            if not api_t:
                continue
            title_ctx = title_context.get(num_base, {})
            cc_real = str(title_ctx.get("cc_real", "") or "1016")
            ret = api_t.get("ret_alert") or {}
            is_payroll_consumo = _is_synergia_payroll_charge(
                vals[document_index],
                api_t.get("observ", ""),
                first,
                vals[title_index],
            )
            payroll_link_for_code = None
            if is_payroll_consumo:
                payroll_link_for_code = _build_tax_link(vals[document_index], api_t.get("observ", ""), "", nf_original_local, venc) or {}
            if is_payroll_consumo:
                _pl_origins_for_code = (payroll_link_for_code or {}).get("origins", []) or []
                code = _norm_rel_code(_pl_origins_for_code[0].get("cod", "")) if _pl_origins_for_code else _norm_rel_code((payroll_link_for_code or {}).get("ref_orig", ""))
            else:
                code = _norm_rel_code(ret.get("ref_orig") or api_t.get("ref") or "")
            if not code:
                # api_t nao guarda ref; acha pelo no contendo esse titulo
                def find_code(nodes):
                    for node in nodes or []:
                        if any(str(t.get("numtit", "")).split("/")[0].strip() == num_base for t in node.get("titulos", []) or []):
                            return _norm_rel_code(node.get("cod"))
                        found = find_code(node.get("children", []))
                        if found:
                            return found
                    return ""
                code = find_code(api_tree or [])
            if not code:
                continue
            if is_payroll_consumo:
                ret_alert = None
                payroll_link = _build_tax_link(
                    vals[document_index],
                    api_t.get("observ", ""),
                    code,
                    nf_original_local,
                    venc,
                ) or payroll_link_for_code or {}
                ret_aprop = {
                    "tipo": "FOLHA_PROPRIA",
                    "status": payroll_link.get("status", "VINCULADO"),
                    "ref_orig": payroll_link.get("ref_orig") or code,
                    "nm_orig": payroll_link.get("nm_orig") or api_t.get("item", ""),
                    "nf_num": payroll_link.get("nf_num") or str(vals[document_index] or "FOLHA"),
                    "cred_orig": payroll_link.get("cred_orig") or first,
                    "titulo_orig": payroll_link.get("titulo_orig") or "Regra de folha propria",
                    "regra": payroll_link.get("regra") or "DARF, INSS, FGTS, GFIP, GRF, GRRF, GPS, IRRF e sindicato seguem a apropriação do título original da folha; rateio proporcional por funcionário.",
                    "tipo_imposto": "FOLHA_PROPRIA",
                    "link_label": payroll_link.get("link_label") or "Encargos de m.o. propria",
                    "origins": payroll_link.get("origins", []),
                }
            else:
                ret_alert, ret_aprop = rel_ret_from_api(api_t, code, api_t.get("item", ""))
            incorporation_reap = bool(
                cc_real != "1011"
                and (
                    title_ctx.get("incorporacao_reap")
                    or _is_incorporation_creditor(first)
                    or (
                        ret_aprop
                        and _is_incorporation_creditor(ret_aprop.get("cred_orig", ""))
                    )
                )
            )
            total = report_money(vals[total_index])
            allocations = [(code, total, None)]
            if is_payroll_consumo:
                origins = (ret_aprop or {}).get("origins", []) or []
                if origins:
                    allocations = []
                    remaining = total
                    for pos, origin in enumerate(origins):
                        origin_code = _norm_rel_code(origin.get("cod", ""))
                        if not origin_code:
                            continue
                        allocated = remaining if pos == len(origins) - 1 else f2(total * float(origin.get("share", 0) or 0))
                        remaining = f2(remaining - allocated)
                        origin_aprop = dict(ret_aprop or {})
                        origin_aprop.update({
                            "ref_orig": origin.get("cod", origin_code),
                            "nm_orig": origin.get("servico", ""),
                            "cred_orig": origin.get("credor", first),
                            "titulo_orig": origin.get("titulo", origin_aprop.get("titulo_orig", "")),
                            "origins": [],
                        })
                        allocations.append((origin_code, allocated, origin_aprop))
            for alloc_code, alloc_total, alloc_aprop in allocations:
                if not alloc_code:
                    continue
                a_pagar_by_code[alloc_code] = f2(a_pagar_by_code.get(alloc_code, 0) + alloc_total)
                titulos_by_code.setdefault(alloc_code, []).append({
                    "numtit": str(vals[title_index]),
                    "credor": first,
                    "vlr_ap": alloc_total,
                    "vlr_pg": 0,
                    "vlr_a_pagar": alloc_total,
                    "status": "A PAGAR",
                    "dt_baixa": venc,
                    "dt_venc": venc,
                    "tpbaixa": "",
                    "observ": str(vals[document_index] or ""),
                    "pedidos": api_t.get("pedidos", []),
                    "contratos": api_t.get("contratos", []),
                    "nfs": api_t.get("nfs", []),
                    "ret_alert": ret_alert,
                    "ret_aprop": alloc_aprop or ret_aprop,
                    "payroll_consumo": is_payroll_consumo,
                    "cc_real": cc_real,
                    "codcredor": title_ctx.get("codcredor"),
                    "incorporacao_reap": incorporation_reap,
                })

    roots = json.loads(json.dumps(api_tree or []))
    # Encargos de folha seguem a EAP original da folha; nao cria folha artificial.
    matched = 0
    unmatched = 0

    def roll(node):
        nonlocal matched, unmatched
        if node.get("children"):
            for ch in node["children"]:
                roll(ch)
            node["orcado"] = f2(sum(ch.get("orcado", 0) for ch in node["children"]))
            node["realizado"] = f2(sum(ch.get("realizado", 0) for ch in node["children"]))
            node["a_pagar"] = f2(sum(ch.get("a_pagar", 0) for ch in node["children"]))
            node["saldo"] = f2(node["orcado"] - node["realizado"] - node["a_pagar"])
        else:
            code = _norm_rel_code(node.get("cod"))
            has_cost = code in rel_code_set
            matched += 1 if has_cost else 0
            unmatched += 0 if has_cost else 1
            orcado = f2(node.get("orcado", 0))
            realizado = f2(realizado_by_code.get(code, 0))
            a_pagar = f2(a_pagar_by_code.get(code, 0))
            node["orcado"] = orcado
            node["realizado"] = realizado
            node["a_pagar"] = a_pagar
            node["saldo"] = f2(orcado - realizado - a_pagar)
            node["titulos"] = titulos_by_code.get(code, [])
        return node

    for n in roots:
        roll(n)

    # O orçado do Relatório é o mesmo da API (estrutura e valor 100% vindos
    # do bd_sienge_tmz/eng_orc, via o "roots" copiado da API acima) — sem
    # nenhum reescalonamento pra um nominal externo. Os mecanismos antigos de
    # recalibragem (por N2 via "FIN - AS - NOMINAL" e, mais recente, por N1
    # via ORÇAMENTO PRELIMINAR) foram removidos a pedido do usuário: a
    # estrutura de EAP e o custo devem ser exatamente os do Sienge, iguais
    # nas duas telas (API e Relatório) — só o realizado/a pagar continuam
    # vindo dos relatórios manuais (essa é a diferença real entre as telas).
    nominal_budget = {"ok": False}
    nominal_n2 = {}
    nominal_applied = 0
    nominal_missing = 0

    incc_series = load_incc_di_series()
    current_prev_period = (pd.Timestamp.now().to_period("M") - 1).strftime("%Y-%m")
    available_refs = sorted(key for key, value in incc_series.items() if value > 0 and key <= current_prev_period)
    current_incc_ref = available_refs[-1] if available_refs else ""
    current_incc = float(incc_series.get(current_incc_ref, 0) or 0)

    def previous_month_key(value):
        dt = pd.to_datetime(value, errors="coerce", dayfirst=True)
        if pd.isna(dt):
            return ""
        return (dt.to_period("M") - 1).strftime("%Y-%m")

    def incc_for_reference(reference):
        candidates = sorted(key for key, value in incc_series.items() if value > 0 and key <= reference)
        if not candidates:
            return "", 0.0
        key = candidates[-1]
        return key, float(incc_series[key])

    analytic_total = f2(sum(float(value or 0) for value in title_aprop_totals.values()))
    paid_matching_total = f2(sum(float(paid_by_title.get(key, 0) or 0) for key in title_aprop_totals))
    paid_report_total = f2(sum(float(value or 0) for value in paid_by_title.values()))
    paid_outside_analytic = f2(sum(
        float(value or 0) for key, value in paid_by_title.items() if key not in title_aprop_totals
    ))
    paid_incc_units = 0.0
    paid_incc_adjusted = 0.0
    reconciliation = []
    all_financial_titles = sorted(set(title_aprop_totals) | set(paid_by_title))
    for key in all_financial_titles:
        analytic_value = f2(title_aprop_totals.get(key, 0))
        paid_value = f2(paid_by_title.get(key, 0))
        difference = f2(paid_value - analytic_value)
        paid_info = paid_info_by_title.get(key, {})
        aprop_info = aprop_info_by_title.get(key, {})
        paid_date = paid_info.get("data", "—")
        incc_ref, payment_incc = incc_for_reference(previous_month_key(paid_date))
        adjusted = (
            f2(paid_value / payment_incc * current_incc)
            if paid_value and payment_incc and current_incc else paid_value
        )
        if key in title_aprop_totals and paid_value:
            if payment_incc:
                paid_incc_units += paid_value / payment_incc
            paid_incc_adjusted += adjusted
        if key not in title_aprop_totals:
            status = "Pago não consta no analítico"
        elif paid_value == 0 and analytic_value:
            status = "Apropriado, ainda não pago"
        elif difference > 0.01:
            status = "Pagamento maior que apropriação"
        elif difference < -0.01:
            status = "Apropriação maior que pagamento"
        else:
            status = "Conciliado"
        if status != "Conciliado":
            reconciliation.append({
                "titulo": key,
                "credor": aprop_info.get("credor") or paid_info.get("credor") or "—",
                "documento": aprop_info.get("documento") or paid_info.get("documento") or "—",
                "eap": aprop_info.get("eap", ""),
                "servico": aprop_info.get("servico", ""),
                "data_pagamento": paid_date,
                "valor_analitico": analytic_value,
                "valor_pago": paid_value,
                "diferenca": difference,
                "status": status,
                "considerado_dashboard": key in title_aprop_totals,
                "incc_ref": incc_ref,
                "incc_indice": round(payment_incc, 3),
                "pago_reajustado_incc": adjusted,
            })
    reconciliation.sort(key=lambda row: (
        0 if row["considerado_dashboard"] else 1,
        -abs(row["diferenca"]),
        row["titulo"],
    ))
    payroll_titles = titulos_by_code.get(_norm_rel_code(PAYROLL_CONSUMO_CODE), []) or []
    payroll_realizado = f2(sum(float(t.get("vlr_pg", 0) or 0) for t in payroll_titles))
    payroll_a_pagar = f2(sum(float(t.get("vlr_a_pagar", 0) or 0) for t in payroll_titles))
    payroll_apropriado = f2(sum(float(t.get("vlr_ap", 0) or 0) for t in payroll_titles))

    meta = {
        "ok": True,
        "fonte_orcamento": os.path.basename(LOCAL_EXCEL_PATH),
        "fonte_eap": os.path.basename(LOCAL_EXCEL_PATH),
        "orcamento_base_unica": "sienge",
        "orcamento_nominal_total": f2(nominal_budget.get("total", 0)) if nominal_budget.get("ok") else 0,
        "orcamento_nominal_raw_total": f2(nominal_budget.get("raw_total", 0)) if nominal_budget.get("ok") else 0,
        "orcamento_nominal_scale_factor": round(float(nominal_budget.get("scale_factor", 1) or 1), 8),
        "orcamento_nominal_n2_aplicados": nominal_applied,
        "orcamento_nominal_n2_zerados": nominal_missing,
        "fonte_realizado": os.path.basename(PAGAS_REL_PATH) if os.path.exists(PAGAS_REL_PATH) else "",
        "fonte_apropriacao": os.path.basename(APROP_REL_PATH) if os.path.exists(APROP_REL_PATH) else "",
        "fonte_a_pagar": os.path.basename(APAGAR_REL_PATH) if os.path.exists(APAGAR_REL_PATH) else "",
        "total_orcado": f2(sum(n.get("orcado", 0) for n in roots)),
        "total_realizado": f2(sum(n.get("realizado", 0) for n in roots)),
        "total_a_pagar": f2(sum(n.get("a_pagar", 0) for n in roots)),
        "incc_base_ref": "2025-01",
        "incc_base_indice": round(float(incc_series.get("2025-01", 0) or 0), 3),
        "incc_atual_ref": current_incc_ref,
        "incc_atual_indice": round(current_incc, 3),
        "incc_series": {k: round(v, 3) for k, v in incc_series.items()},
        "folha_propria_consumo_code": PAYROLL_CONSUMO_CODE,
        "folha_propria_consumo_nome": PAYROLL_CONSUMO_NAME,
        "folha_propria_consumo_regra": "DARF, INSS, FGTS, GFIP, GRF, GRRF, GPS, IRRF e sindicato seguem o título original da folha",
        "folha_propria_consumo_titulos": len(payroll_titles),
        "folha_propria_consumo_apropriado": payroll_apropriado,
        "folha_propria_consumo_realizado": payroll_realizado,
        "folha_propria_consumo_a_pagar": payroll_a_pagar,
        "eap": "Mesma EAP do Orçamento API",
        "itens_com_custo": matched,
        "itens_sem_custo": unmatched,
        "qtd_divergencias_realizado": len(divergencias_realizado),
        "conciliacao_pago": {
            "total_analitico": analytic_total,
            "total_pago_dashboard": paid_matching_total,
            "total_pago_relatorio": paid_report_total,
            "pago_fora_analitico": paid_outside_analytic,
            "diferenca_dashboard_analitico": f2(paid_matching_total - analytic_total),
            "qtd_diferencas": len(reconciliation),
            "incc_referencia": current_incc_ref,
            "incc_indice_referencia": round(current_incc, 3),
            "incc_fonte": "Planilha de apoio; jun/2026 atualizado pela variação oficial INCC-DI FGV/IBRE de 0,78%",
            "pago_em_incc": round(paid_incc_units, 6),
            "pago_reajustado_incc": f2(paid_incc_adjusted),
            "linhas": reconciliation,
        },
    }
    return roots, meta, divergencias_realizado, multi_aprop_titulos

def build_cred_map(bancos, xl):
    import re
    cred_map = {}
    # 1) planilha de apoio credores.xlsx (fonte principal)
    if os.path.exists(CREDORES_PATH):
        try:
            # tenta ler formato padrão e formatos alternativos
            try:
                df_cred = pd.read_excel(CREDORES_PATH, sheet_name=0)
                cols = {str(c).strip().lower(): c for c in df_cred.columns}

                cod_col = next((cols[k] for k in cols
                                if 'cod' in k and 'cred' in k), None)
                nome_col = next((cols[k] for k in cols
                                 if ('credor' in k or 'nome' in k or 'nm' in k)), None)

                if cod_col and nome_col:
                    for _, row in df_cred.dropna(subset=[cod_col]).iterrows():
                        try:
                            cred_id = int(float(row[cod_col]))
                            cred_nm = str(row[nome_col]).strip()
                            if cred_nm and cred_nm.lower() != 'nan':
                                cred_map.setdefault(cred_id, cred_nm)
                        except:
                            pass
                else:
                    raise Exception('colunas nao detectadas')
            except Exception:
                # fallback: formato "123 - NOME"
                df_cred = pd.read_excel(CREDORES_PATH, sheet_name=0, header=None)
                for _, row in df_cred.iterrows():
                    val = row.iloc[0]
                    if pd.isna(val): 
                        continue
                    m = re.match(r'^(\d+)\s*-\s*(.*)', str(val).strip())
                    if m:
                        cred_id = int(m.group(1))
                        cred_nm = m.group(2).strip()
                        if cred_id not in cred_map:
                            cred_map[cred_id] = cred_nm
        except Exception:
            pass
    # 2) aba cad_credores no proprio Excel (fallback)
    if 'cad_credores' in xl.sheet_names:
        try:
            cad = xl.parse('cad_credores')
            if 'codcredor' in cad.columns and 'nmcredor' in cad.columns:
                for _, row in cad.dropna(subset=['nmcredor','codcredor']).iterrows():
                    try:
                        k = int(float(row['codcredor']))
                        if k not in cred_map:
                            cred_map[k] = str(row['nmcredor'])
                    except: pass
        except: pass
    # 3) dfin_bancos (último recurso)
    for _, row in bancos.dropna(subset=['nmcredor','codcredor']).iterrows():
        try:
            k = int(float(row['codcredor']))
            if k not in cred_map:
                cred_map[k] = str(row['nmcredor'])
        except: pass
    return cred_map

def build_data(force=False):
    global _cache
    if not force and _cache["data"] and (time.time() - _cache["ts"]) < CACHE_TTL:
        return _cache["data"], None
    if not force:
        disk_cache = load_dashboard_cache()
        if disk_cache:
            _cache = {"data": disk_cache, "ts": time.time()}
            return disk_cache, None
    excel_path = resolve_excel_path()
    if not os.path.exists(excel_path):
        return None, f"Arquivo nao encontrado: {excel_path}"
    try:
        xl = pd.ExcelFile(excel_path, engine="openpyxl")
        tit     = xl.parse('fin_cpagar_aprobra')
        bancos  = xl.parse('dfin_bancos')
        acr     = xl.parse('ffin_cpagar_acrtit')
        orc     = xl.parse('eng_orc')
        vnd     = xl.parse('dcom_vendas')
        cti     = xl.parse('fsup_ctitensct')
        cti_med = xl.parse('fsup_ctitens')
        # fsup_ct = cabecalho do contrato (1 linha por nuct, com codcredorct,
        # codccusto, statct/sitautct reais) — nao tem coditemct/itens, por isso
        # so serve de fonte auxiliar pros fallbacks abaixo, nao substitui cti.
        try:
            ct_hdr = xl.parse('fsup_ct')
            ct_hdr['_nuct_n'] = pd.to_numeric(ct_hdr['nuct'], errors='coerce')
            ct_hdr = ct_hdr.dropna(subset=['_nuct_n']).drop_duplicates('_nuct_n')
            _ct_hdr_credor = pd.to_numeric(ct_hdr.set_index('_nuct_n')['codcredorct'], errors='coerce').dropna().astype(int).to_dict()
            _ct_hdr_cc = pd.to_numeric(ct_hdr.set_index('_nuct_n')['codccusto'], errors='coerce').dropna().astype(int).to_dict()
            _ct_hdr_status = ct_hdr.set_index('_nuct_n')['statct'].dropna().astype(str).to_dict()
            _ct_hdr_sitaut = ct_hdr.set_index('_nuct_n')['sitautct'].dropna().astype(str).to_dict()
        except Exception as _ex_ct_hdr:
            print(f"  [CONTRATOS] fsup_ct (cabecalho) indisponivel: {_ex_ct_hdr}")
            _ct_hdr_credor, _ct_hdr_cc, _ct_hdr_status, _ct_hdr_sitaut = {}, {}, {}, {}

        # ── mapa credores (todas as fontes) ──────────────────────────────────
        cred_map = build_cred_map(bancos, xl)
        _, cred_name_to_cnpj = build_cnpj_maps(bancos)
        cnpj_cache = load_cnpj_cache()

        # ── preparar dfin_bancos ─────────────────────────────────────────────
        bancos['codccusto_n'] = pd.to_numeric(bancos['codccusto'], errors='coerce')
        bancos['vlrmov']      = pd.to_numeric(bancos['vlrmov'], errors='coerce').fillna(0)
        bancos['dtlancto']    = pd.to_datetime(bancos['dtlancto'], errors='coerce')
        # mapa CC -> nome real do SIENGE
        _cc_df = (bancos[['codccusto_n','nmccusto']].dropna()
            .drop_duplicates('codccusto_n').sort_values('codccusto_n'))
        _cc_nm_map = {int(r.codccusto_n): str(r.nmccusto) for _,r in _cc_df.iterrows()}

        # ── titulos AP ───────────────────────────────────────────────────────
        # Remove títulos PCT (Previsão Financeira) e PPC — campo coddocum
        tit = tit.copy()
        _tit_excluidos = pd.DataFrame()
        if 'coddocum' in tit.columns:
            _mask_fin_prev = excluded_fin_doc_mask(tit)
            _tit_excluidos = tit.loc[_mask_fin_prev].copy()
            _por_tipo = tit.loc[_mask_fin_prev, 'coddocum'].astype(str).str.upper().str.strip().value_counts().to_dict()
            tit = tit.loc[~_mask_fin_prev].copy()
            print(f"  [FILTRO PPC/PCT] {_por_tipo} titulos de previsao financeira removidos via coddocum, {len(tit)} restantes.")
        tit['codcredor']  = pd.to_numeric(tit['codcredor'], errors='coerce').fillna(0).astype(int)
        tit['nmcredor']   = tit['codcredor'].map(cred_map).fillna('SEM CREDOR CADASTRADO')
        tit['dtbaixacp']  = pd.to_datetime(tit['dtbaixacp'], errors='coerce')
        tit['STATUS_PAG'] = tit['dtbaixacp'].apply(lambda x: 'PAGO' if pd.notna(x) else 'A PAGAR')
        # campo observação (nome pode variar conforme versão do Sienge)
        _obs_col = next((c for c in [
            'dsobservacao','observacao','nmobservacao','dsobserv','observ','obs',
            'histlancto','nmhistorico','dshist','dshistorico','nmhist',
            'descricao','nmdescricao','dsdescricao','complemento'
        ] if c in tit.columns), None)
        if _obs_col:
            tit['_observ'] = tit[_obs_col].fillna('').astype(str).str.strip()
        else:
            # tenta qualquer coluna cujo nome contenha 'obs' ou 'hist'
            _obs_fallback = [c for c in tit.columns if ('obs' in c.lower() or 'hist' in c.lower()) and c not in ('codccusto',)]
            tit['_observ'] = tit[_obs_fallback[0]].fillna('').astype(str).str.strip() if _obs_fallback else ''
        print(f"  [OBSERV] coluna usada: {_obs_col or (_obs_fallback[0] if _obs_fallback else 'NENHUMA')} | colunas tit: {list(tit.columns)}")
        tit['refitemorc'] = tit['refitemorc'].astype(str).str.strip()
        tit['_refitemorc_lancada'] = tit['refitemorc']
        _payroll_doc_cols = [c for c in ["coddocum", "docnuct", "nudoc", "numtitcp"] if c in tit.columns]
        tit["_payroll_consumo"] = tit.apply(
            lambda row: _is_synergia_payroll_charge(
                " ".join(str(row.get(c, "") or "") for c in _payroll_doc_cols),
                row.get("_observ", ""),
                row.get("nmcredor", ""),
                row.get("nmtitcp", ""),
            ),
            axis=1,
        )
        _payroll_count = int(tit["_payroll_consumo"].sum())
        _payroll_origin_map = _build_local_nf_origin_map()
        if _payroll_count:
            print(f"  [FOLHA] {_payroll_count} titulos de encargos seguem a EAP original da folha, com rateio por funcionario.")
            _expanded_tit_rows = []
            _split_count = 0
            _split_value = 0.0
            for _, _tax_row in tit.iterrows():
                if not bool(_tax_row.get("_payroll_consumo", False)):
                    _expanded_tit_rows.append(_tax_row.copy())
                    continue
                _payroll_doc_text = ' '.join(str(_tax_row.get(c, '') or '') for c in list(_payroll_doc_cols) + ['nmtitcp', 'nmcredor'])
                _payroll_ref_date = next((
                    _tax_row.get(c) for c in ['dtbaixacp', 'dtvenc', 'dtemissao', 'dtlancto', 'dtcad']
                    if c in _tax_row.index and pd.notna(_tax_row.get(c))
                ), None)
                _payroll_link = _build_tax_link(
                    _payroll_doc_text,
                    str(_tax_row.get('_observ', '') or ''),
                    str(_tax_row.get('_refitemorc_lancada', _tax_row.get('refitemorc', '')) or ''),
                    _payroll_origin_map,
                    _payroll_ref_date,
                ) or {}
                _origins = _payroll_link.get('origins', []) if isinstance(_payroll_link, dict) else []
                _origins = [o for o in _origins if _norm_rel_code(o.get('cod', ''))]
                if not _origins:
                    _expanded_tit_rows.append(_tax_row.copy())
                    continue

                _vlr_ap_total = f2(_tax_row.get('vlrapropobra', 0))
                _vlr_pg_total = f2(_tax_row.get('vlrliqbaixacp', 0))
                _rem_ap = _vlr_ap_total
                _rem_pg = _vlr_pg_total
                for _pos, _origin in enumerate(_origins):
                    _share = float(_origin.get('share', 0) or 0)
                    _alloc_ap = _rem_ap if _pos == len(_origins) - 1 else f2(_vlr_ap_total * _share)
                    _alloc_pg = _rem_pg if _pos == len(_origins) - 1 else f2(_vlr_pg_total * _share)
                    _rem_ap = f2(_rem_ap - _alloc_ap)
                    _rem_pg = f2(_rem_pg - _alloc_pg)
                    _new_row = _tax_row.copy()
                    _origin_cod = _norm_rel_code(_origin.get('cod', ''))
                    _new_row['refitemorc'] = _origin_cod
                    _new_row['vlrapropobra'] = _alloc_ap
                    _new_row['vlrliqbaixacp'] = _alloc_pg
                    _new_row['_folha_rateio'] = True
                    _new_row['_folha_origin_ref'] = _origin_cod
                    _new_row['_folha_origin_servico'] = _origin.get('servico', '')
                    _new_row['_folha_origin_credor'] = _origin.get('credor', '')
                    _new_row['_folha_origin_titulo'] = _origin.get('titulo', '')
                    _new_row['_folha_origin_documento'] = _origin.get('documento', '')
                    _new_row['_folha_origin_share'] = _share
                    _expanded_tit_rows.append(_new_row)
                    _split_count += 1
                    _split_value = f2(_split_value + _alloc_ap)
            tit = pd.DataFrame(_expanded_tit_rows).reset_index(drop=True)
            print(f"  [FOLHA] rateio aplicado em {_split_count} linhas | valor rateado R$ {_split_value:,.2f}")
        tit['n1_cod'] = tit['refitemorc'].str[:2]
        tit['n2_cod'] = tit['refitemorc'].str[:6]
        tit['n3_cod'] = tit['refitemorc'].str[:10]
        tit['codccusto']  = pd.to_numeric(tit['codccusto'], errors='coerce')
        for c in ['vlrapropobra','vlrliqbaixacp']:
            tit[c] = pd.to_numeric(tit[c], errors='coerce').fillna(0)
        acr_grp = acr.groupby('numtitcp').agg(vlrjur=('vlrjurtitbx','sum'),vlrmulta=('vlrmultatitbx','sum')).reset_index()
        tit = tit.merge(acr_grp, on='numtitcp', how='left')
        tit['vlrjur']   = tit['vlrjur'].fillna(0)
        tit['vlrmulta'] = tit['vlrmulta'].fillna(0)

        # CCs com dados reais de AP
        _cc_com_ap = set(tit['codccusto'].dropna().astype(int).unique())
        cc_list_out = [
            {'cod': int(r.codccusto_n), 'nm': str(r.nmccusto),
             'tem_ap': int(r.codccusto_n) in _cc_com_ap}
            for _,r in _cc_df.iterrows()
        ]

        # ── extras: nf_ped / pedidos / insumos (precisa ser antes da árvore ORC) ─
        import re as _re
        def _norm(s): return _re.sub(r'[^A-Z0-9]','',str(s).upper())
        def _nf_norm(v):
            s = str(v or '').strip()
            if s.endswith('.0'):
                s = s[:-2]
            d = _re.sub(r'\D', '', s)
            return d.lstrip('0') or d

        def _nf_keys(v):
            raw = str(v or '').strip()
            keys = []
            for k in (raw, _nf_norm(raw)):
                if k and k not in ('nan', 'None', '') and k not in keys:
                    keys.append(k)
            return keys

        _nf_ref_map_ped = {}  # nf_num(str) → refitemorc  — construído aqui
        nf_ped = pd.DataFrame()
        sol_ped = pd.DataFrame()
        try:
            nf_ped  = xl.parse('fsup_nf_ped')
            sol_ped = xl.parse('fsup_sol_ped')
            print(f"  [NF_PED cols] {list(nf_ped.columns)}")
            nf_ped['numtitcp'] = pd.to_numeric(nf_ped['numtitcp'], errors='coerce')
            if not _tit_excluidos.empty:
                _numtit_excluidos = set(pd.to_numeric(_tit_excluidos['numtitcp'], errors='coerce').dropna().astype(int))
                _antes_nf_ped = len(nf_ped)
                nf_ped = nf_ped[~nf_ped['numtitcp'].isin(_numtit_excluidos)].copy()
                print(f"  [FILTRO NF/PED PPC/PCT] {_antes_nf_ped - len(nf_ped)} vinculos de contrato/pedido removidos.")
            nf_ped['codpedi']  = pd.to_numeric(nf_ped['codpedi'],  errors='coerce')
            sol_ped['codpedi'] = pd.to_numeric(sol_ped['codpedi'], errors='coerce')
            sol_nm = sol_ped[['codpedi','nminsum']].dropna(subset=['nminsum']).groupby('codpedi')['nminsum'].first().to_dict()
            # detecta coluna do número da NF — lista extensa com variações do Sienge
            _nf_col = next((c for c in [
                'numnf','numdocfis','coddocfis','numdoc','codnotafis','nfdoc','numdocumento',
                'numdocnf','coddoc','numdoccp','nfdocumento','numeronf','numero_nf',
                'numdocorigem','coddocorigem','numnotafiscal','coddocfiscal',
            ] if c in nf_ped.columns), None)
            # fallback: qualquer coluna com 'nf' ou 'doc' ou 'nota' no nome
            if not _nf_col:
                _nf_col = next((c for c in nf_ped.columns
                    if any(k in c.lower() for k in ['numnf','numdoc','notafis','nfdoc','docfis'])
                ), None)
            print(f"  [NF_PED] coluna NF detectada: '{_nf_col}'")
            _aggs = {
                'codpedi_list': ('codpedi', lambda x: sorted(set(int(v) for v in x if pd.notna(v)))),
                'docnuct_list': ('docnuct', lambda x: sorted(set(str(v) for v in x if pd.notna(v) and str(v) not in ('nan','None','')))),
            }
            if _nf_col:
                _aggs['numnf_list'] = (_nf_col, lambda x: sorted(set(str(v).strip() for v in x if pd.notna(v) and str(v).strip() not in ('nan','None',''))))
            nf_grp = nf_ped.groupby('numtitcp').agg(**_aggs).reset_index()
            tit_extra = {}
            for _,row in nf_grp.iterrows():
                peds = row['codpedi_list']
                ins  = list({sol_nm.get(p,'') for p in peds if sol_nm.get(p)})
                _nfs_this = list(row['numnf_list']) if 'numnf_list' in row.index else []
                tit_extra[int(row.numtitcp)] = {
                    'pedidos': [str(p) for p in peds],
                    'contratos': row['docnuct_list'],
                    'nfs': _nfs_this,
                    'insumos': [i for i in ins if i],
                }
            # ── monta _nf_ref_map_ped: nf_num → refitemorc via nf_ped + tit ──
            if _nf_col:
                _nf_tit = nf_ped[['numtitcp', _nf_col]].dropna(subset=[_nf_col, 'numtitcp']).copy()
                _nf_tit[_nf_col] = _nf_tit[_nf_col].astype(str).str.strip()
                _nf_tit['numtitcp'] = pd.to_numeric(_nf_tit['numtitcp'], errors='coerce')
                _nf_tit = _nf_tit.dropna(subset=['numtitcp'])
                _nf_tit['refitemorc'] = _nf_tit['numtitcp'].map(_tit_ref_map_tmp := tit.set_index('numtitcp')['refitemorc'].to_dict())
                _nf_tit = _nf_tit.dropna(subset=['refitemorc'])
                _nf_ref_map_ped = {}
                for _, _rnf in _nf_tit.iterrows():
                    _ref_nf = str(_rnf['refitemorc']).strip()
                    for _knf in _nf_keys(_rnf[_nf_col]):
                        _nf_ref_map_ped.setdefault(_knf, _ref_nf)
                print(f"  [NF_REF_MAP_PED] {len(_nf_ref_map_ped)} NFs mapeadas via nf_ped")
            # fallback docnuct→credor via NFs (contratos sem título CP vinculado)
            try:
                _nf_c = nf_ped[['docnuct','numtitcp']].dropna(subset=['numtitcp']).copy()
                _nf_c['numtitcp'] = pd.to_numeric(_nf_c['numtitcp'], errors='coerce')
                _nf_c = _nf_c.merge(tit[['numtitcp','nmcredor']].drop_duplicates('numtitcp'), on='numtitcp', how='left')
                _docnuct_cred_nf = _nf_c.dropna(subset=['nmcredor']).groupby('docnuct')['nmcredor'].first().to_dict()
            except Exception:
                _docnuct_cred_nf = {}
        except Exception as _ex_nfped:
            print(f"  [NF_PED ERRO] {_ex_nfped}")
            tit_extra = {}
            _docnuct_cred_nf = {}
        suprimentos = build_suprimentos(sol_ped, nf_ped, orc, tit, cred_map)

        # ── mapa numtit → refitemorc e nf_num → refitemorc ──────────────────────
        _tit_ref_map = tit.set_index('numtitcp')['refitemorc'].to_dict()

        # ── _nf_ref_map: nf_num(str) → refitemorc ───────────────────────────────
        _nf_ref_map = {}
        _nf_ref_map.update(_nf_ref_map_ped)  # Estratégia 1: via nf_ped

        # Estratégia 2: coluna de NF direto em tit
        _nf_num_col_tit = next((c for c in [
            'numnf','numdocfis','coddocfis','numdoc','codnotafis','nfdoc',
            'numdocnf','coddoc','numdoccp','numeronf',
        ] if c in tit.columns), None)
        if not _nf_num_col_tit:
            _nf_num_col_tit = next((c for c in tit.columns
                if any(k in c.lower() for k in ['numnf','numdoc','notafis','nfdoc','docfis'])
            ), None)
        if _nf_num_col_tit:
            for _, _r in tit[['numtitcp',_nf_num_col_tit,'refitemorc']].dropna(subset=[_nf_num_col_tit]).iterrows():
                _k = str(_r[_nf_num_col_tit]).strip()
                if _k and _k not in ('nan','None',''):
                    for _knf in _nf_keys(_k):
                        _nf_ref_map.setdefault(_knf, str(_r['refitemorc']).strip())

        # Estratégia 3: docnuct em tit
        if 'docnuct' in tit.columns:
            for _, _r in tit[['docnuct','refitemorc']].dropna(subset=['docnuct']).iterrows():
                _k = str(_r['docnuct']).strip()
                if _k and _k not in ('nan','None',''):
                    for _knf in _nf_keys(_k):
                        _nf_ref_map.setdefault(_knf, str(_r['refitemorc']).strip())

        # Estratégia 4: parse observ dos títulos NF/NFSE de origem
        _coddocum_nf_set = {'ME','NFSE','NF','NFE'}
        if 'coddocum' in tit.columns:
            _tit_nf_src = tit[tit['coddocum'].astype(str).str.upper().str.strip().isin(_coddocum_nf_set)].copy()
        else:
            _tit_nf_src = tit.copy()
        for _, _r in _tit_nf_src.iterrows():
            _obs = str(_r.get('_observ','') or '')
            _hits = _re.findall(r'(?:NFS?E?[/ ]+)([0-9]+)', _obs.upper())
            for _h in _hits:
                for _knf in _nf_keys(_h):
                    _nf_ref_map.setdefault(_knf, str(_r['refitemorc']).strip())

        # ── _credor_nf_ref: (nmcredor_norm, nf_num) → refitemorc ─────────────
        # Chave PRINCIPAL para retenção: o credor na observação da retenção +
        # número da NF → busca no título de origem daquele credor
        _tit_nf_src2 = _tit_nf_src.copy()
        _tit_nf_src2['_cred_norm'] = _tit_nf_src2['nmcredor'].apply(_norm)
        # docnuct como número da NF
        _credor_nf_ref = {}
        if 'docnuct' in tit.columns:
            _tmp = _tit_nf_src2[['_cred_norm','docnuct','refitemorc']].dropna(subset=['docnuct'])
            for _, _r in _tmp.iterrows():
                for _knf in _nf_keys(_r['docnuct']):
                    _k2 = (str(_r['_cred_norm']), _knf)
                    _credor_nf_ref.setdefault(_k2, str(_r['refitemorc']).strip())
        if _nf_num_col_tit:
            _tmp2 = _tit_nf_src2[['_cred_norm',_nf_num_col_tit,'refitemorc']].dropna(subset=[_nf_num_col_tit])
            for _, _r in _tmp2.iterrows():
                for _knf in _nf_keys(_r[_nf_num_col_tit]):
                    _k2 = (str(_r['_cred_norm']), _knf)
                    _credor_nf_ref.setdefault(_k2, str(_r['refitemorc']).strip())

        # ── _credor_ref_map: nmcredor_norm → list of (refitemorc, nmitemorc) ──
        # Fallback final: se só há 1 refitemorc para o credor, usamos ele
        _credor_refs = {}
        for _, _r in _tit_nf_src2.iterrows():
            _c = str(_r['_cred_norm'])
            _credor_refs.setdefault(_c, set()).add(str(_r['refitemorc']).strip())

        print(f"  [NF_REF_MAP] total: {len(_nf_ref_map)} | credor+NF: {len(_credor_nf_ref)} | colunas tit: {list(tit.columns)}")

        # ── _credor_refs_all: nmcredor_norm → set(refitemorc) ─────────────────
        # Construído de TODOS os títulos (não só NF/NFSE), para que ISS/INSS
        # possa localizar a apropriação do credor de origem via nome
        _credor_refs_all = {}
        for _, _r in tit.iterrows():
            _c = _norm(str(_r.get('nmcredor','') or ''))
            _ref = str(_r['refitemorc']).strip()
            if _c and _ref and _ref not in ('nan','None',''):
                _credor_refs_all.setdefault(_c, set()).add(_ref)

        # ── _numtit_ref: numtitcp → refitemorc (lookup rápido) ────────────────
        _numtit_ref = tit.set_index('numtitcp')['refitemorc'].to_dict()

        print(f"  [CREDOR_REFS] {len(_credor_refs_all)} credores mapeados")
        _payroll_origin_map = _build_local_nf_origin_map()
        _tit_nm_map = {}

        # ── orcamento ────────────────────────────────────────────────────────
        orc['refitemorc'] = orc['refitemorc'].astype(str).str.strip()
        orc['vlritemorc'] = pd.to_numeric(orc['vlritemorc'], errors='coerce').fillna(0)
        orc['qtitemorc']  = pd.to_numeric(orc['qtitemorc'],  errors='coerce').fillna(0)
        orc4 = orc[orc['refitemorc'].str.len()==14].copy()
        orc4['n1'] = orc4['refitemorc'].str[:2]
        orc4['n2'] = orc4['refitemorc'].str[:6]
        orc4['n3'] = orc4['refitemorc'].str[:10]
        # Encargos de folha seguem a EAP original da folha; nao cria insumo artificial.
        rr = tit.groupby('refitemorc').agg(
            realizado=('vlrliqbaixacp','sum'),
            a_pagar=('vlrapropobra', lambda x: x[tit.loc[x.index,'STATUS_PAG']=='A PAGAR'].sum())
        ).reset_index()
        orc4 = orc4.merge(rr, on='refitemorc', how='left')
        orc4['realizado'] = orc4['realizado'].fillna(0)
        orc4['a_pagar']   = orc4['a_pagar'].fillna(0)
        orc4['saldo']     = orc4['vlritemorc'] - orc4['realizado']
        # mapas de nomes dos itens orçamento (para exibir no alerta de retenção)
        _orc_nm_map = orc4.set_index('refitemorc')['nmitemorc'].to_dict()

        n1_nm = orc4.groupby('n1')['refitemorc1n'].first().to_dict()
        n2_nm = orc4.groupby('n2')['refitemorc2n'].first().to_dict()
        n3_nm = orc4.groupby('n3')['refitemorc3n'].first().to_dict()
        n1g = orc4.groupby('n1').agg(o=('vlritemorc','sum'),r=('realizado','sum'),ap=('a_pagar','sum'),s=('saldo','sum')).reset_index()
        n2g = orc4.groupby(['n1','n2']).agg(o=('vlritemorc','sum'),r=('realizado','sum'),ap=('a_pagar','sum'),s=('saldo','sum')).reset_index()
        n3g = orc4.groupby(['n1','n2','n3']).agg(o=('vlritemorc','sum'),r=('realizado','sum'),ap=('a_pagar','sum'),s=('saldo','sum')).reset_index()
        tree = []
        for _,r1 in n1g.iterrows():
            nd1={'cod':str(r1.n1),'nm':str(n1_nm.get(r1.n1,'')),'orcado':f2(r1.o),'realizado':f2(r1.r),'a_pagar':f2(r1.ap),'saldo':f2(r1.s),'children':[]}
            for _,r2 in n2g[n2g.n1==r1.n1].iterrows():
                nd2={'cod':str(r2.n2),'nm':str(n2_nm.get(r2.n2,'')),'orcado':f2(r2.o),'realizado':f2(r2.r),'a_pagar':f2(r2.ap),'saldo':f2(r2.s),'children':[]}
                for _,r3 in n3g[n3g.n2==r2.n2].iterrows():
                    nd3={'cod':str(r3.n3),'nm':str(n3_nm.get(r3.n3,'')),'orcado':f2(r3.o),'realizado':f2(r3.r),'a_pagar':f2(r3.ap),'saldo':f2(r3.s),'children':[]}
                    for _,r4 in orc4[orc4.n3==r3.n3].iterrows():
                        # titulos apropriados neste item (N5)
                        titulos_n4 = tit[tit['refitemorc']==r4.refitemorc].copy()
                        tits_list = []
                        for _,rt in titulos_n4.sort_values('dtbaixacp',na_position='last').iterrows():
                            _pago = f2(rt['vlrliqbaixacp']) if pd.notna(rt['dtbaixacp']) else 0.0
                            _ex = tit_extra.get(int(rt['numtitcp']),{})
                            # ── retenção: check se apropriação bate com NF de origem ─────
                            _ret_alert = None
                            _ret_aprop = None
                            _observ_str = str(rt.get('_observ','') or '')
                            _coddocum   = str(rt.get('coddocum','') or '').upper().strip()
                            _payroll_consumo = bool(rt.get('_payroll_consumo', False))
                            _is_retencao = (
                                any(k in _observ_str.upper() for k in ['RETENÇÃO','RETENCAO','INSS','ISS','IRRF','CSLL','PIS','COFINS'])
                                or _coddocum in TAX_DOC_PREFIXES
                            )
                            if _is_retencao and not _payroll_consumo:
                                _obs_up = _observ_str.upper()
                                # extrai números de NF da observação
                                _nf_matches = _re.findall(r'(?:NFS?E?[/ ]*)([0-9]+)', _obs_up)
                                if not _nf_matches:
                                    _nf_matches = _re.findall(r'DOC(?:S?)[./ ]*([0-9]+)', _obs_up)
                                # também usa as NFs vinculadas no Sienge (nf_ped), que aparecem como tags verdes
                                _nf_candidates = []
                                for _nf_src in list(_nf_matches) + list(_ex.get('nfs', []) or []):
                                    for _knf in _nf_keys(_nf_src):
                                        if _knf and _knf not in _nf_candidates:
                                            _nf_candidates.append(_knf)
                                # extrai credor mencionado: "do credor CIPLAN ... referente"
                                _cred_obs_match = _re.search(r'CREDOR[ ]+(.+?)[ ]+REFERENTE', _obs_up)
                                _cred_obs_norm  = _norm(_cred_obs_match.group(1)) if _cred_obs_match else ''
                                _curr_ref = str(rt['refitemorc']).strip()
                                _orig_ref = None

                                # Estratégia A: nf_num → refitemorc direto
                                for _nf_num_str in _nf_candidates[:8]:
                                    _orig_ref = _nf_ref_map.get(_nf_num_str)
                                    if not _orig_ref and _cred_obs_norm:
                                        _orig_ref = _credor_nf_ref.get((_cred_obs_norm, _nf_num_str))
                                    if _orig_ref:
                                        break

                                # Estratégia B: credor mencionado tem refs únicas → usa a delas
                                if not _orig_ref and _cred_obs_norm:
                                    # busca parcial: credor observado pode ser substring da chave
                                    _cand_keys = [k for k in _credor_refs_all
                                                  if _cred_obs_norm and (_cred_obs_norm in k or k in _cred_obs_norm)]
                                    _cand_refs = set()
                                    for _ck in _cand_keys:
                                        _cand_refs |= _credor_refs_all.get(_ck, set())
                                    if len(_cand_refs) == 1:
                                        _orig_ref = next(iter(_cand_refs))
                                    elif len(_cand_refs) > 1:
                                        # múltiplas refs — tenta coincidir com CC atual
                                        # se todas são da mesma N2, pode ser consenso
                                        _n2s = set(r[:6] for r in _cand_refs if len(r) >= 6)
                                        if len(_n2s) == 1:
                                            # pega a ref mais frequente
                                            _ref_counts = {}
                                            for _ck in _cand_keys:
                                                for _r2 in _credor_refs_all.get(_ck, set()):
                                                    _ref_counts[_r2] = _ref_counts.get(_r2, 0) + 1
                                            _orig_ref = max(_ref_counts, key=_ref_counts.get)

                                # Estratégia C: credor do próprio título de retenção
                                # (ex: INSS onde o credor é Previdência — busca via nmcredor do tit origem)
                                if not _orig_ref:
                                    _own_cred_norm = _norm(str(rt.get('nmcredor','') or ''))
                                    # não é o mesmo credor do imposto — busca pelo credor da observação
                                    pass  # já coberto pela B

                                if _orig_ref:
                                    _ret_aprop = {
                                        'tipo': 'APROP_OK' if _orig_ref == _curr_ref else 'APROP_DIVERGENTE',
                                        'ref_orig': str(_orig_ref),
                                        'nm_orig': _orc_nm_map.get(_orig_ref, _orig_ref),
                                        'nf_num': _nf_candidates[0] if _nf_candidates else '?',
                                        'cred_orig': _cred_obs_match.group(1).title() if _cred_obs_match else '',
                                    }
                                elif _curr_ref:
                                    _ret_aprop = {
                                        'tipo': 'APROP_ATUAL',
                                        'ref_orig': str(_curr_ref),
                                        'nm_orig': _orc_nm_map.get(_curr_ref, _curr_ref),
                                        'nf_num': _nf_candidates[0] if _nf_candidates else str(rt.get('nudoc','') or '?'),
                                        'cred_orig': _cred_obs_match.group(1).title() if _cred_obs_match else '',
                                    }
                                if _ret_aprop and _orig_ref != _curr_ref:
                                    _ret_alert = {
                                        'tipo': 'APROP_DIVERGENTE',
                                        'ref_orig': str(_orig_ref),
                                        'nm_orig': _orc_nm_map.get(_orig_ref, _orig_ref),
                                        'nf_num': _nf_candidates[0] if _nf_candidates else '?',
                                        'cred_orig': _cred_obs_match.group(1).title() if _cred_obs_match else '',
                                    }
                                # se _orig_ref == _curr_ref: apropriação correta, sem alerta
                            if _payroll_consumo:
                                _ret_alert = None
                                _payroll_doc_text = ' '.join(str(rt.get(c, '') or '') for c in list(_payroll_doc_cols) + ['nmtitcp', 'nmcredor'])
                                _payroll_ref_date = next((rt.get(c) for c in ['dtbaixacp', 'dtvenc', 'dtemissao', 'dtlancto', 'dtcad'] if c in rt.index and pd.notna(rt.get(c))), None)
                                if bool(rt.get('_folha_rateio', False)):
                                    _payroll_origins = []
                                    _payroll_link = {
                                        'status': 'VINCULADO',
                                        'ref_orig': _norm_rel_code(rt.get('_folha_origin_ref', rt.get('refitemorc', ''))),
                                        'nm_orig': str(rt.get('_folha_origin_servico', '') or _orc_nm_map.get(str(rt.get('refitemorc', '')).strip(), '')),
                                        'nf_num': 'FOLHA',
                                        'cred_orig': str(rt.get('_folha_origin_credor', '') or rt.get('nmcredor', '') or ''),
                                        'titulo_orig': str(rt.get('_folha_origin_titulo', '') or 'Título original da folha'),
                                        'origins': [],
                                    }
                                else:
                                    _payroll_link = _build_tax_link(
                                        _payroll_doc_text,
                                        str(rt.get('_observ', '') or ''),
                                        str(rt.get('_refitemorc_lancada', rt.get('refitemorc', '')) or ''),
                                        _payroll_origin_map,
                                        _payroll_ref_date,
                                    ) or {}
                                    _payroll_origins = _payroll_link.get('origins', []) if isinstance(_payroll_link, dict) else []
                                _curr_payroll_ref = str(rt.get('refitemorc', '') or '').strip()
                                _ret_aprop = {
                                    'tipo': 'FOLHA_PROPRIA',
                                    'status': _payroll_link.get('status', 'VINCULADO') if isinstance(_payroll_link, dict) else 'VINCULADO',
                                    'ref_orig': _payroll_link.get('ref_orig', _curr_payroll_ref) if isinstance(_payroll_link, dict) else _curr_payroll_ref,
                                    'nm_orig': _payroll_link.get('nm_orig', _orc_nm_map.get(_curr_payroll_ref, _curr_payroll_ref)) if isinstance(_payroll_link, dict) else _orc_nm_map.get(_curr_payroll_ref, _curr_payroll_ref),
                                    'nf_num': _payroll_link.get('nf_num') if isinstance(_payroll_link, dict) else str(rt.get('coddocum','') or rt.get('nudoc','') or 'FOLHA'),
                                    'cred_orig': f"{len(_payroll_origins)} funcionário(s)" if _payroll_origins else (_payroll_link.get('cred_orig') if isinstance(_payroll_link, dict) else str(rt.get('nmcredor','') or '')),
                                    'titulo_orig': _payroll_link.get('titulo_orig', 'Título original da folha') if isinstance(_payroll_link, dict) else 'Título original da folha',
                                    'regra': 'DARF, INSS, FGTS, GFIP, GRF, GRRF, GPS, IRRF e sindicato seguem o título original da folha da competência, rateado proporcionalmente por funcionário.',
                                    'tipo_imposto': 'FOLHA_PROPRIA',
                                    'link_label': _payroll_link.get('link_label', f"Folha -> {_curr_payroll_ref}") if isinstance(_payroll_link, dict) else f"Folha -> {_curr_payroll_ref}",
                                    'origins': _payroll_origins,
                                }
                            # NFs da observação (fallback para ISS/INSS sem nf_ped)
                            _obs_nfs = _re.findall(r'(?:NFS?E?[/ ]+)([0-9]+)', str(rt.get('_observ','') or '').upper())
                            _nfs_merged = list(dict.fromkeys(_ex.get('nfs',[]) + _obs_nfs))  # dedup mantendo ordem
                            _dtvenc = pd.to_datetime(rt.get('dtvenc'), errors='coerce')
                            _cc_real = int(rt.get('codccusto')) if pd.notna(rt.get('codccusto')) else None
                            _codcredor_real = int(rt.get('codcredor')) if pd.notna(rt.get('codcredor')) else None
                            _inc_reap = bool(
                                _cc_real != 1011
                                and (
                                    _codcredor_real in INCORPORATION_CREDITOR_CODES
                                    or _is_incorporation_creditor(rt.get('nmcredor', ''))
                                    or (
                                        _ret_aprop
                                        and _is_incorporation_creditor(_ret_aprop.get('cred_orig', ''))
                                    )
                                )
                            )
                            tits_list.append({
                                'numtit': str(rt['numtitcp']),
                                'credor': str(rt['nmcredor']),
                                'vlr_ap': f2(rt['vlrapropobra']),
                                'vlr_pg': _pago,
                                'vlr_imposto': f2(rt.get('vlrimpbaixacp', 0)),
                                'coddocum': _coddocum,
                                'nudoc': str(rt.get('nudoc','') or ''),
                                'is_retencao': _is_retencao,
                                'status': str(rt['STATUS_PAG']),
                                'dt_baixa': str(rt['dtbaixacp'].date()) if pd.notna(rt['dtbaixacp']) else '—',
                                'dt_venc': str(_dtvenc.date()) if pd.notna(_dtvenc) else '—',
                                'tpbaixa': str(rt.get('tpbaixa','') or ''),
                                'observ': str(rt.get('_observ','') or ''),
                                'ref_lancada': str(rt.get('_refitemorc_lancada', rt.get('refitemorc', '')) or ''),
                                'pedidos': _ex.get('pedidos',[]),
                                'contratos': _ex.get('contratos',[]),
                                'nfs': _nfs_merged,
                                'ret_alert': _ret_alert,
                                'ret_aprop': _ret_aprop,
                                'payroll_consumo': _payroll_consumo,
                                'cc_real': str(_cc_real) if _cc_real is not None else '',
                                'codcredor': _codcredor_real,
                                'incorporacao_reap': _inc_reap,
                            })
                        _ap_n4 = f2(titulos_n4[titulos_n4['STATUS_PAG']=='A PAGAR']['vlrapropobra'].sum())
                        nd3['children'].append({'cod':str(r4.refitemorc),'nm':str(r4.nmitemorc),'un':str(r4.codunid),'qt':f2(r4.qtitemorc),'orcado':f2(r4.vlritemorc),'realizado':f2(r4.realizado),'a_pagar':_ap_n4,'saldo':f2(r4.saldo),'titulos':tits_list})
                    nd2['children'].append(nd3)
                nd1['children'].append(nd2)
            tree.append(nd1)

        report_title_context = {}
        for numtit_ctx, grp_ctx in tit.groupby("numtitcp"):
            ctx_row = grp_ctx.iloc[0]
            try:
                ctx_key = str(int(float(numtit_ctx)))
            except Exception:
                ctx_key = str(numtit_ctx).split("/")[0].strip()
            cc_real = int(ctx_row["codccusto"]) if pd.notna(ctx_row.get("codccusto")) else None
            codcredor = int(ctx_row["codcredor"]) if pd.notna(ctx_row.get("codcredor")) else None
            report_title_context[ctx_key] = {
                "cc_real": str(cc_real) if cc_real is not None else "",
                "codcredor": codcredor,
                "payroll_consumo": bool(ctx_row.get("_payroll_consumo", False)),
                "credor": str(ctx_row.get("nmcredor", "") or ""),
                "incorporacao_reap": bool(
                    cc_real != 1011
                    and (
                        codcredor in INCORPORATION_CREDITOR_CODES
                        or _is_incorporation_creditor(ctx_row.get("nmcredor", ""))
                    )
                ),
            }

        orc_rel_tree, orc_rel_meta, orc_rel_divergencias, titulos_multi_apropriacao = build_orcamento_relatorio(
            tree,
            report_title_context,
        )
        orc_rel_meta["incorporacao_reap_cc"] = "INC-REAP"
        orc_rel_meta["incorporacao_reap_nome"] = "Incorporação - reapropriação"
        orc_rel_meta["incorporacao_reap_credores"] = sorted(INCORPORATION_CREDITOR_NAMES)

        # ── fluxo ────────────────────────────────────────────────────────────
        pg = tit[tit['dtbaixacp'].notna()].copy()
        pg['mes'] = pg['dtbaixacp'].dt.to_period('M').astype(str)
        fl = pg.groupby('mes')['vlrliqbaixacp'].sum().reset_index().sort_values('mes')
        fluxo_list = [{'mes':str(r.mes),'v':f2(r.vlrliqbaixacp)} for _,r in fl.iterrows()]
        fluxo_per_cc = {}
        for cc in tit['codccusto'].dropna().unique():
            sub = tit[(tit['dtbaixacp'].notna())&(tit['codccusto']==cc)].copy()
            if sub.empty: continue
            sub['mes'] = sub['dtbaixacp'].dt.to_period('M').astype(str)
            fg = sub.groupby('mes')['vlrliqbaixacp'].sum().reset_index().sort_values('mes')
            fluxo_per_cc[str(int(cc))] = [{'mes':str(r.mes),'v':f2(r.vlrliqbaixacp)} for _,r in fg.iterrows()]
        # Demais CCs: pagamentos do banco
        bpag = bancos[(bancos['nmoper']=='Pagamento')&(bancos['vlrmov']<0)&bancos['dtlancto'].notna()].copy()
        bpag['mes'] = bpag['dtlancto'].dt.to_period('M').astype(str)
        bpag['vlr'] = bpag['vlrmov'].abs()
        for cc in bancos['codccusto_n'].dropna().unique():
            cc_key = str(int(cc))
            if cc_key in fluxo_per_cc: continue
            sub = bpag[bpag['codccusto_n']==cc]
            if sub.empty: continue
            fg = sub.groupby('mes')['vlr'].sum().reset_index().sort_values('mes')
            fluxo_per_cc[cc_key] = [{'mes':str(r.mes),'v':f2(r.vlr),'banco':True} for _,r in fg.iterrows()]

        # ── taxa de administração da obra (construtora), mês a mês ─────────────
        _admin_cod = '01.004.001'
        _admin_tit = tit[tit['n3_cod'] == _admin_cod].copy()
        _admin_pg = _admin_tit[_admin_tit['dtbaixacp'].notna()].copy()
        _admin_pg['mes'] = _admin_pg['dtbaixacp'].dt.to_period('M').astype(str)
        _admin_mensal = _admin_pg.groupby('mes')['vlrliqbaixacp'].sum().reset_index().sort_values('mes')
        taxa_construtora = {
            'cod': _admin_cod,
            'nome': 'Administração da Obra (taxa da construtora)',
            'vlr_ap': f2(_admin_tit['vlrapropobra'].sum()),
            'vlr_pg': f2(_admin_tit['vlrliqbaixacp'].sum()),
            'qt': int(len(_admin_tit)),
            'meses': [{'mes': str(r.mes), 'v': f2(r.vlrliqbaixacp)} for _, r in _admin_mensal.iterrows()],
        }

        # ── apuração independente da taxa de construção (10% do custo incorrido,
        # regime de caixa) ───────────────────────────────────────────────────
        # "taxa_construtora" acima só mostra o que JÁ foi lançado no Sienge sob o
        # código orçamentário da administração — não confirma se o valor lançado
        # está certo. Esta apuração recalcula do zero, mês a mês, pela mesma
        # metodologia da própria Gama Dias (conferida contra a "Medição de Obra"
        # oficial): soma o valor líquido de TODOS os títulos pagos no mês
        # (regime de caixa — data de baixa, não de apropriação) e aplica 10%,
        # com 5% de retenção técnica sobre a taxa. Guarda o detalhe por título
        # pra permitir auditoria — é o "quais títulos estão inclusos" que dá pra
        # comparar linha a linha com a medição oficial.
        _TAXA_PCT = 0.10
        _RETENCAO_TAXA_PCT = 0.05
        _tx_pagos = tit[(tit['STATUS_PAG'] == 'PAGO') & tit['dtbaixacp'].notna()].copy()
        _tx_pagos['mes'] = _tx_pagos['dtbaixacp'].dt.to_period('M').astype(str)
        _doc_col_tx = next((c for c in ['coddocum', 'nudoc'] if c in _tx_pagos.columns), None)
        apuracao_taxa = []
        for mes, grp in _tx_pagos.groupby('mes'):
            base = f2(grp['vlrliqbaixacp'].sum())
            taxa_bruta = f2(base * _TAXA_PCT)
            retencao = f2(taxa_bruta * _RETENCAO_TAXA_PCT)
            # a base tem uma linha por rubrica orçamentária apropriada do título
            # (um mesmo título pago costuma vir dividido em várias linhas) — pra
            # exibição/auditoria agrupa de volta por numtitcp, senão a lista
            # mostra ~10x mais "títulos" do que a Gama Dias reconhece como tal
            # (validado contra a Medição nº 2: 151 títulos únicos aqui pra 153 na
            # medição oficial de julho/2026, mesmo valor batendo ao centavo).
            _por_titulo = {}
            for r in grp.itertuples():
                key = str(int(r.numtitcp)) if pd.notna(r.numtitcp) else f"s/num-{r.Index}"
                _dtvenc = pd.to_datetime(getattr(r, 'dtvenc', None), errors='coerce')
                item = _por_titulo.setdefault(key, {
                    'numtit': key, 'credor': str(r.nmcredor),
                    'dt_pag': str(r.dtbaixacp.date()) if pd.notna(r.dtbaixacp) else '—',
                    'dt_venc': str(_dtvenc.date()) if pd.notna(_dtvenc) else '—',
                    'vlr_ap': 0.0, 'vlr_pg': 0.0, 'vlrjur': 0.0, 'vlrmulta': 0.0,
                    'status': 'PAGO',
                    'documento': str(getattr(r, _doc_col_tx, '') or '') if _doc_col_tx else '',
                })
                item['vlr_ap'] += float(r.vlrapropobra or 0)
                item['vlr_pg'] += float(r.vlrliqbaixacp or 0)
                item['vlrjur'] += float(getattr(r, 'vlrjur', 0) or 0)
                item['vlrmulta'] += float(getattr(r, 'vlrmulta', 0) or 0)
            titulos_mes = [{**t, 'vlr_ap': f2(t['vlr_ap']), 'vlr_pg': f2(t['vlr_pg']), 'vlrjur': f2(t['vlrjur']), 'vlrmulta': f2(t['vlrmulta'])} for t in _por_titulo.values()]
            apuracao_taxa.append({
                'mes': str(mes),
                'base': base,
                'qt_titulos': len(titulos_mes),
                'taxa_bruta': taxa_bruta,
                'retencao': retencao,
                'liberado': f2(taxa_bruta - retencao),
                'titulos': sorted(titulos_mes, key=lambda t: -t['vlr_pg']),
            })
        apuracao_taxa.sort(key=lambda x: x['mes'])

        # ── taxa de gestão financeira/carteira — sócio operador (centro de custo
        # 1019, fora da OBRA) ────────────────────────────────────────────────
        # Diferente da taxa da Gama Dias acima (10% do custo incorrido em Obra):
        # esta é um valor FIXO mensal ("Gestão Financeira" + "Gestão de
        # Carteira") pago ao sócio operador — Temazec Desenvolvimento até
        # dez/2025, sucedida por RGD Incorporação e Participação a partir de
        # jan/2026 (mesma rubrica, credor diferente, valores corrigidos pelo
        # INCC). Achado a pedido do usuário: "recebimento em outro centro de
        # custo" do pagamento de taxa — cruzado aqui pra aparecer junto.
        try:
            _cp_raw = xl.parse('dfin_cpagar')
            _tx_gestao_raw = _cp_raw[_cp_raw['codccusto'] == 1019].copy()
            _CREDOR_NOMES_1019 = {989: 'Temazec Desenvolvimento', 6250: 'RGD Incorporação e Participação'}

            def _tipo_taxa_gestao(obstit):
                t = str(obstit or '').upper()
                if 'CARTEIRA' in t:
                    return 'Gestão de Carteira'
                if 'FINANCEIRA' in t:
                    return 'Gestão Financeira'
                return 'Outra'

            _tx_gestao_raw['credor_nome'] = _tx_gestao_raw['codcredor'].map(_CREDOR_NOMES_1019)
            _tx_gestao_raw = _tx_gestao_raw[_tx_gestao_raw['credor_nome'].notna()].copy()
            _tx_gestao_raw['tipo'] = _tx_gestao_raw['obstit'].apply(_tipo_taxa_gestao)
            _tx_gestao_raw['dtvenc'] = pd.to_datetime(_tx_gestao_raw['dtvenc'], errors='coerce')
            _tx_gestao_raw['dtbaixacp'] = pd.to_datetime(_tx_gestao_raw['dtbaixacp'], errors='coerce')
            _tx_gestao_raw['mes'] = _tx_gestao_raw['dtvenc'].dt.to_period('M').astype(str)
            taxa_gestao_temazec = []
            for mes, grp in _tx_gestao_raw.groupby('mes'):
                valor_total = f2(grp['vlrparcorig'].sum())
                valor_pago = f2(grp.loc[grp['dtbaixacp'].notna(), 'vlrliqbaixacp'].sum())
                titulos_mes = [{
                    'numtit': str(r.numdocum), 'credor': r.credor_nome, 'tipo': r.tipo,
                    'dt_venc': str(r.dtvenc.date()) if pd.notna(r.dtvenc) else '—',
                    'dt_pag': str(r.dtbaixacp.date()) if pd.notna(r.dtbaixacp) else '—',
                    'vlr_ap': f2(r.vlrparcorig),
                    'vlr_pg': f2(r.vlrliqbaixacp) if pd.notna(r.dtbaixacp) else 0.0,
                    'vlrjur': 0.0, 'vlrmulta': 0.0,
                    'status': 'PAGO' if pd.notna(r.dtbaixacp) else 'A PAGAR',
                    'documento': str(r.numdocum),
                } for r in grp.itertuples()]
                taxa_gestao_temazec.append({
                    'mes': str(mes), 'valor_total': valor_total, 'valor_pago': valor_pago,
                    'qt_titulos': int(len(grp)),
                    'titulos': sorted(titulos_mes, key=lambda t: t['tipo']),
                })
            taxa_gestao_temazec.sort(key=lambda x: x['mes'])
        except Exception as _exc_txg:
            print(f"  [TAXA_GESTAO] falha ao montar taxa de gestão (cc 1019): {_exc_txg}")
            taxa_gestao_temazec = []

        # ── credores ─────────────────────────────────────────────────────────
        bc = tit.groupby('nmcredor').agg(vlr_ap=('vlrapropobra','sum'),vlr_pg=('vlrliqbaixacp','sum'),qt=('numtitcp','count'),vlrjur=('vlrjur','sum')).reset_index().sort_values('vlr_ap',ascending=False)
        by_cred_list = [{'nmcredor':r.nmcredor,'vlr_ap':f2(r.vlr_ap),'vlr_pg':f2(r.vlr_pg),'qt':int(r.qt),'vlrjur':f2(r.vlrjur)} for _,r in bc.iterrows()]

        # ── aging / vencidos (títulos em aberto por prazo de atraso) ───────────
        _hoje = pd.Timestamp.now().normalize()
        _ap_aberto = tit[tit['STATUS_PAG'] == 'A PAGAR'].copy()
        _ap_aberto['dtvenc_dt'] = pd.to_datetime(_ap_aberto['dtvenc'], errors='coerce') if 'dtvenc' in _ap_aberto.columns else pd.NaT
        _ap_aberto['dias_atraso'] = (_hoje - _ap_aberto['dtvenc_dt']).dt.days
        _venc_geral = _ap_aberto[_ap_aberto['dias_atraso'] > 0]
        _a_vencer_geral = _ap_aberto[(_ap_aberto['dias_atraso'] <= 0) & (_ap_aberto['dias_atraso'] >= -7)]

        def _aging_bucket_label(days):
            if days <= 30: return '1-30'
            if days <= 60: return '31-60'
            if days <= 90: return '61-90'
            return '90+'

        def _build_aging(sub_venc, sub_a_vencer):
            buckets = {k: {'qt': 0, 'valor': 0.0} for k in ('1-30', '31-60', '61-90', '90+')}
            for _, r in sub_venc.iterrows():
                b = _aging_bucket_label(r['dias_atraso'])
                buckets[b]['qt'] += 1
                buckets[b]['valor'] = f2(buckets[b]['valor'] + float(r['vlrapropobra'] or 0))
            bucket_labels = {'1-30': 'Até 30 dias', '31-60': '31–60 dias', '61-90': '61–90 dias', '90+': 'Mais de 90 dias'}
            return {
                'vencido_qt': int(len(sub_venc)),
                'vencido_valor': f2(sub_venc['vlrapropobra'].sum()),
                'dias_atraso_medio': round(float(sub_venc['dias_atraso'].mean()), 1) if len(sub_venc) else 0,
                'buckets': [{'key': k, 'label': bucket_labels[k], 'qt': v['qt'], 'valor': v['valor']} for k, v in buckets.items()],
                'a_vencer_7d_qt': int(len(sub_a_vencer)),
                'a_vencer_7d_valor': f2(sub_a_vencer['vlrapropobra'].sum()),
            }

        aging_geral = _build_aging(_venc_geral, _a_vencer_geral)
        aging_per_cc = {}
        for cc in tit['codccusto'].dropna().unique():
            cc_key = str(int(cc))
            aging_per_cc[cc_key] = _build_aging(
                _venc_geral[_venc_geral['codccusto'] == cc],
                _a_vencer_geral[_a_vencer_geral['codccusto'] == cc],
            )
        aging_titulos = [{
            'numtit': str(r['numtitcp']),
            'credor': str(r['nmcredor']),
            'valor': f2(r['vlrapropobra']),
            'dtvenc': str(r['dtvenc_dt'].date()) if pd.notna(r['dtvenc_dt']) else '—',
            'dias_atraso': int(r['dias_atraso']),
            'bucket': _aging_bucket_label(r['dias_atraso']),
            'codccusto': str(int(r['codccusto'])) if pd.notna(r['codccusto']) else '',
        } for _, r in _venc_geral.sort_values('dias_atraso', ascending=False).iterrows()]

        # ── títulos detalhados (drill-down por credor / nível / CC na aba Pagamentos) ─
        _tp = tit.copy()
        _tp['dtvencparsed'] = pd.to_datetime(_tp['dtvenc'], errors='coerce') if 'dtvenc' in _tp.columns else pd.NaT
        # nome real do item de orçamento (ex.: "Corte, dobra e montagem de aço")
        # pelo código refitemorc — nmuc só dá o nome da obra/CC, não do item.
        _nmitemorc_map = orc.dropna(subset=['refitemorc']).drop_duplicates('refitemorc').set_index('refitemorc')['nmitemorc'].to_dict()
        titulos_pag = [{
            'numtit': str(r.numtitcp),
            'credor': str(r.nmcredor),
            'vlr_ap': f2(r.vlrapropobra),
            'vlr_pg': f2(r.vlrliqbaixacp),
            'status': str(r.STATUS_PAG),
            'dt_venc': str(r.dtvencparsed.date()) if pd.notna(r.dtvencparsed) else '—',
            'dt_pag': str(r.dtbaixacp.date()) if pd.notna(r.dtbaixacp) else '—',
            'documento': str(getattr(r, 'coddocum', '') or ''),
            'vlrjur': f2(getattr(r, 'vlrjur', 0)),
            'vlrmulta': f2(getattr(r, 'vlrmulta', 0)),
            'codccusto': str(int(r.codccusto)) if pd.notna(r.codccusto) else '',
            'n1_cod': str(r.n1_cod or ''),
            'n2_cod': str(r.n2_cod or ''),
            'n3_cod': str(r.n3_cod or ''),
            'apropriacao': _nmitemorc_map.get(str(getattr(r, 'refitemorc', '') or '').strip(), '') or str(getattr(r, 'nmuc', '') or '').strip(),
            'refitemorc': str(getattr(r, 'refitemorc', '') or '').strip(),
        } for r in _tp.itertuples()]

        # ── busca por Pedido de Compra (PC) → título vinculado + apropriação ──
        # Ligação direta PC→título (fsup_nf_ped: codpedi→numtitcp) — mesma
        # fonte que alimenta o Mapa de Controle (lá é derivado PC→NF→título;
        # aqui é direto), cruzada com a apropriação de titulos_pag pra trazer
        # centro de custo/item orçamento sem precisar reabrir o relatório.
        pc_titulo_map = {}
        try:
            _pcmap = nf_ped[nf_ped['codpedi'].notna() & nf_ped['numtitcp'].notna()][['codpedi', 'numtitcp']].copy()
            _pcmap['codpedi'] = _pcmap['codpedi'].astype(int).astype(str)
            _pcmap['numtitcp'] = _pcmap['numtitcp'].astype(int).astype(str)
            for pc, grp in _pcmap.groupby('codpedi'):
                pc_titulo_map[pc] = sorted(set(grp['numtitcp'].tolist()), key=lambda x: int(x))
        except Exception as _exc_pcmap:
            print(f"  [PC_TITULO_MAP] falha ao montar mapa PC→título: {_exc_pcmap}")

        # ── controle de locações por equipamento (insumo) ───────────────────────
        # Classifica pelo grupo de insumo do próprio Sienge (desgrins = "Locação" /
        # "Locação de equipamentos"), cruzando fsup_nf_ped (item da NF: quantidade,
        # valor, data) com feng_itensobra (nome do item) e com tit (pago ou não).
        _loc_state = _load_locacoes_state()
        _loc_incluidos = set(_loc_state.get('incluidos_manual', []))
        _loc_excluidos = set(_loc_state.get('excluidos_manual', []))

        def _norm_item_nome(s):
            # maiusculas, sem acento, espaços colapsados — pra "master locações"
            # e "MASTER LOCACOES" (ou com espaço duplo) casarem com o mesmo item,
            # em vez de exigir bater byte a byte com o nome exato do Sienge.
            return re.sub(r'\s+', ' ', _ascii_upper(s)).strip()

        _loc_incluidos_norm = {_norm_item_nome(n) for n in _loc_incluidos}
        _loc_excluidos_norm = {_norm_item_nome(n) for n in _loc_excluidos}
        _loc_categorias = _loc_state.get('categorias_manual', {})
        _loc_categorias_norm = {_norm_item_nome(k): v for k, v in _loc_categorias.items()}

        def _categoria_loc(nome):
            nome_norm = _norm_item_nome(nome)
            escolha = _loc_categorias_norm.get(nome_norm)
            if escolha in ('escoramento', 'ferramentas'):
                return escolha
            return 'escoramento' if _ESCORAMENTO_REGEX.search(nome_norm) else 'ferramentas'

        locacoes = {'vlr_pg': 0.0, 'vlr_total': 0.0, 'qt_lancamentos': 0, 'itens': []}
        try:
            _itens_obra = xl.parse('feng_itensobra')
            _itens_obra['coditemunico'] = _itens_obra['coditemunico'].astype(str)
            _itens_map = _itens_obra.drop_duplicates('coditemunico').set_index('coditemunico')[['descitem', 'desgrins']].to_dict('index')

            _nf_loc = nf_ped.copy() if isinstance(nf_ped, pd.DataFrame) and not nf_ped.empty else pd.DataFrame()
            if not _nf_loc.empty and 'coditemunico' in _nf_loc.columns:
                _nf_loc['coditemunico'] = _nf_loc['coditemunico'].astype(str)
                _nf_loc['descitem'] = _nf_loc['coditemunico'].map(lambda k: (_itens_map.get(k) or {}).get('descitem'))
                _nf_loc['desgrins'] = _nf_loc['coditemunico'].map(lambda k: (_itens_map.get(k) or {}).get('desgrins'))

                def _item_e_locacao(row):
                    nome = str(row.get('descitem') or f"Insumo {row.get('codinsum','?')}").strip()
                    nome_norm = _norm_item_nome(nome)
                    if nome_norm in _loc_excluidos_norm:
                        return False
                    if nome_norm in _loc_incluidos_norm:
                        return True
                    return bool(_LOCACAO_GRUPO_REGEX.search(str(row.get('desgrins') or '')))

                _mask_loc = _nf_loc.apply(_item_e_locacao, axis=1)
                _nf_loc = _nf_loc[_mask_loc].copy()
                _nf_loc['numtitcp'] = pd.to_numeric(_nf_loc['numtitcp'], errors='coerce')
                _tit_slim = tit[['numtitcp', 'nmcredor', 'dtbaixacp', 'STATUS_PAG']].drop_duplicates('numtitcp')
                _nf_loc = _nf_loc.merge(_tit_slim, on='numtitcp', how='left')
                _nf_loc['dtemissnf'] = pd.to_datetime(_nf_loc['dtemissnf'], errors='coerce')
                _nf_loc['mes'] = _nf_loc['dtemissnf'].dt.to_period('M').astype(str)
                _nf_loc['vlrinsumnf'] = pd.to_numeric(_nf_loc['vlrinsumnf'], errors='coerce').fillna(0)
                _nf_loc['qtentregnf'] = pd.to_numeric(_nf_loc['qtentregnf'], errors='coerce').fillna(0)
                _nf_loc['descitem'] = _nf_loc['descitem'].fillna(_nf_loc['codinsum'].apply(lambda c: f"Insumo {c}"))

                locacoes['vlr_total'] = f2(_nf_loc['vlrinsumnf'].sum())
                locacoes['vlr_pg'] = f2(_nf_loc.loc[_nf_loc['STATUS_PAG'] == 'PAGO', 'vlrinsumnf'].sum())
                locacoes['qt_lancamentos'] = int(len(_nf_loc))

                # agrupa por (credor, item) — não só por item — porque um mesmo
                # equipamento (ex.: "Betoneira") é frequentemente alugado de
                # fornecedores diferentes em obras distintas do período, e
                # antes disso ficava escondido atrás de um "credor moda" só.
                _nf_loc['nmcredor'] = _nf_loc['nmcredor'].fillna('Sem fornecedor identificado')
                itens_out = []
                _doc_col_loc = next((c for c in ['numdocumnf', 'numdocfis', 'coddocfis'] if c in _nf_loc.columns), None)
                for (credor, nome), grp in _nf_loc.groupby(['nmcredor', 'descitem']):
                    grp = grp.sort_values('dtemissnf')
                    meses_grp = grp.groupby('mes').agg(qtd=('qtentregnf', 'sum'), valor=('vlrinsumnf', 'sum')).reset_index()
                    with_qtd = grp[grp['qtentregnf'] > 0]
                    valor_unit_medio = f2((with_qtd['vlrinsumnf'] / with_qtd['qtentregnf']).mean()) if len(with_qtd) else 0.0
                    ultima_qtd = grp.iloc[-1]['qtentregnf'] if len(grp) else 0
                    dt_min = grp['dtemissnf'].min()
                    dt_max = grp['dtemissnf'].max()
                    titulos_item = [{
                        'numtit': str(int(r.numtitcp)) if pd.notna(r.numtitcp) else '—',
                        'credor': str(credor),
                        'vlr_ap': f2(r.vlrinsumnf),
                        'vlr_pg': f2(r.vlrinsumnf) if str(r.STATUS_PAG) == 'PAGO' else 0.0,
                        'vlrjur': 0.0, 'vlrmulta': 0.0,
                        'dt_venc': str(r.dtemissnf.date()) if pd.notna(r.dtemissnf) else '—',
                        'dt_pag': str(r.dtbaixacp.date()) if pd.notna(r.dtbaixacp) else '—',
                        'status': str(r.STATUS_PAG) if pd.notna(r.STATUS_PAG) else '—',
                        'documento': str(r[_doc_col_loc]) if _doc_col_loc and pd.notna(r.get(_doc_col_loc)) else '—',
                    } for _, r in grp.iterrows()]
                    itens_out.append({
                        'nome': str(nome),
                        'credor': str(credor),
                        'qtd_atual': f2(ultima_qtd),
                        'valor_unitario_medio': valor_unit_medio,
                        'valor_total': f2(grp['vlrinsumnf'].sum()),
                        'valor_pago': f2(grp.loc[grp['STATUS_PAG'] == 'PAGO', 'vlrinsumnf'].sum()),
                        'qt_lancamentos': int(len(grp)),
                        'primeira_cobranca': str(dt_min.date()) if pd.notna(dt_min) else '',
                        'ultima_cobranca': str(dt_max.date()) if pd.notna(dt_max) else '',
                        'origem': 'manual' if str(nome) in _loc_incluidos else 'auto',
                        'categoria': _categoria_loc(nome),
                        'meses': [{'mes': str(r.mes), 'qtd': f2(r.qtd), 'valor': f2(r.valor)} for _, r in meses_grp.iterrows()],
                        'titulos': titulos_item,
                    })
                itens_out.sort(key=lambda x: -x['valor_total'])
                locacoes['itens'] = itens_out
            locacoes['excluidos_manual'] = sorted(_loc_excluidos)
        except Exception as _exc_loc:
            print(f"  [LOCACOES] falha ao montar controle de locações: {_exc_loc}")
            locacoes['excluidos_manual'] = sorted(_loc_excluidos)

        tit_por_cc = {}
        for cc in tit['codccusto'].dropna().unique():
            sub = tit[tit['codccusto']==cc]
            byc = sub.groupby('nmcredor').agg(vlr_ap=('vlrapropobra','sum'),vlr_pg=('vlrliqbaixacp','sum'),qt=('numtitcp','count'),vlrjur=('vlrjur','sum')).reset_index().sort_values('vlr_ap',ascending=False)
            tit_por_cc[str(int(cc))] = [{'nmcredor':r.nmcredor,'vlr_ap':f2(r.vlr_ap),'vlr_pg':f2(r.vlr_pg),'qt':int(r.qt),'vlrjur':f2(r.vlrjur)} for _,r in byc.iterrows()]
        for cc in bancos['codccusto_n'].dropna().unique():
            cc_key = str(int(cc))
            if cc_key in tit_por_cc: continue
            sub = bancos[(bancos['codccusto_n']==cc)&(bancos['nmoper']=='Pagamento')].copy()
            if sub.empty: continue
            sub['codcredor_int'] = pd.to_numeric(sub['codcredor'], errors='coerce').fillna(0).astype(int)
            sub['nmcredor_c'] = sub['codcredor_int'].map(cred_map).fillna(
                sub['nmcredor'].fillna('CREDOR ' + sub['codcredor'].astype(str)))
            sub['vlr'] = sub['vlrmov'].abs()
            byc = sub.groupby('nmcredor_c').agg(vlr_pg=('vlr','sum'),qt=('vlr','count')).reset_index().sort_values('vlr_pg',ascending=False)
            tit_por_cc[cc_key] = [{'nmcredor':r.nmcredor_c,'vlr_ap':f2(r.vlr_pg),'vlr_pg':f2(r.vlr_pg),'qt':int(r.qt),'vlrjur':0,'banco':True} for _,r in byc.iterrows()]

        def make_tbl(gcol, nm_map):
            g = tit.groupby(gcol).agg(vlr_ap=('vlrapropobra','sum'),vlr_pg=('vlrliqbaixacp','sum'),qt=('numtitcp','count')).reset_index()
            g['nm'] = g[gcol].map(nm_map).fillna('—')
            return [{'cod':str(r[gcol]),'nm':r.nm,'vlr_ap':f2(r.vlr_ap),'vlr_pg':f2(r.vlr_pg),'qt':int(r.qt)} for _,r in g.iterrows()]

        tit_cc_q = tit.groupby(['codccusto']).agg(vlr_ap=('vlrapropobra','sum'),vlr_pg=('vlrliqbaixacp','sum'),qt=('numtitcp','count')).reset_index()
        tit_cc_q['nm'] = tit_cc_q['codccusto'].apply(lambda x: _cc_nm_map.get(int(x),'CC '+str(int(x))) if pd.notna(x) else '—')
        tit_cc_list = [{'cod':int(r.codccusto) if pd.notna(r.codccusto) else 0,'nm':r.nm,'vlr_ap':f2(r.vlr_ap),'vlr_pg':f2(r.vlr_pg),'qt':int(r.qt)} for _,r in tit_cc_q.iterrows()]
        jur_q = tit[tit['vlrjur']>0].groupby(['codccusto','nmcredor']).agg(vlrjur=('vlrjur','sum'),vlrmulta=('vlrmulta','sum'),qt=('numtitcp','count')).reset_index()
        jur_q['nmcc'] = jur_q['codccusto'].apply(lambda x: _cc_nm_map.get(int(x),'CC '+str(int(x))) if pd.notna(x) else '—')
        jur_cc_list = [{'codcc':str(int(r.codccusto)) if pd.notna(r.codccusto) else '','nmcc':r.nmcc,'nmcredor':r.nmcredor,'vlrjur':f2(r.vlrjur),'vlrmulta':f2(r.vlrmulta),'qt':int(r.qt)} for _,r in jur_q.iterrows()]
        tit_cc3 = tit.groupby(['codccusto','n3_cod']).agg(vlr_ap=('vlrapropobra','sum'),vlr_pg=('vlrliqbaixacp','sum')).reset_index()
        tit_cc3['nmcc'] = tit_cc3['codccusto'].apply(lambda x: _cc_nm_map.get(int(x),'CC '+str(int(x))) if pd.notna(x) else '—')
        tit_cc3['nm_n3'] = tit_cc3['n3_cod'].map(n3_nm)
        tit_cc3_list = [{'nmcc':r.nmcc,'n3_cod':r.n3_cod,'nm_n3':str(r.nm_n3),'vlr_ap':f2(r.vlr_ap),'vlr_pg':f2(r.vlr_pg)} for _,r in tit_cc3.iterrows()]

        # ── contratos ────────────────────────────────────────────────────────
        for c in ['vlrtotitem','vlrtotct','vlrmedidoct','vlrpendct','vlrliberct']:
            cti[c] = pd.to_numeric(cti[c], errors='coerce').fillna(0)
        cti['codccusto'] = pd.to_numeric(cti['codccusto'], errors='coerce')
        cti['nmcc'] = cti['codccusto'].apply(lambda x: _cc_nm_map.get(int(x),'CC '+str(int(x))) if pd.notna(x) else '—')
        cti_med = cti_med.copy()
        if not _tit_excluidos.empty and 'numtitcp' in cti_med.columns:
            _numtit_excluidos = set(pd.to_numeric(_tit_excluidos['numtitcp'], errors='coerce').dropna().astype(int))
            cti_med['_numtitcp_n'] = pd.to_numeric(cti_med['numtitcp'], errors='coerce')
            _antes_med = len(cti_med)
            cti_med = cti_med[~cti_med['_numtitcp_n'].isin(_numtit_excluidos)].copy()
            cti_med.drop(columns=['_numtitcp_n'], inplace=True, errors='ignore')
            print(f"  [FILTRO CONTRATOS PPC/PCT] {_antes_med - len(cti_med)} medicoes vinculadas a titulos PPC/PCT removidas.")
        cti_med2 = cti_med[['nuct','numtitcp']].dropna(subset=['numtitcp']).copy()
        cti_med2['numtitcp'] = pd.to_numeric(cti_med2['numtitcp'], errors='coerce')
        def _valid_contract_creditor_name(v):
            s = str(v or '').strip()
            if not s or s.lower() in {'nan', 'none', 'nat', 'null'}:
                return None
            if all((ch.isspace() or ch == '-' or ord(ch) in (0x2014, 0xfffd)) for ch in s):
                return None
            if s.upper() == 'SEM CREDOR CADASTRADO':
                return None
            return s

        def _contract_creditors_from_manual_reports():
            title_candidates = {}
            contract_candidates = {}
            report_paths = []
            for _pat in ('*Contas Pagas*.xlsx', '*Contas a Pagar*.xlsx'):
                report_paths.extend(glob.glob(os.path.join(DATA_DIR, _pat)))

            _bad_creditor_labels = {
                'CREDOR', 'TOTAL', 'TOTAIS', 'OBSERVACAO', 'OBSERVAÇÃO',
                'RELATÓRIO DE CONTAS PAGAS', 'RELATORIO DE CONTAS PAGAS',
                'RELATÓRIO DE CONTAS A PAGAR', 'RELATORIO DE CONTAS A PAGAR',
            }
            _title_re = re.compile(r'\b(\d{4,})\s*/\s*\d+\b')
            _subst_re = re.compile(r'Substitui\w*\s+(\d{4,})', re.I)
            _contract_re = re.compile(r'\b(?:PCT|CTS)\s*[./-]?\s*(\d{1,6})\b', re.I)

            def _add_candidate(bucket, key, name):
                name = _valid_contract_creditor_name(name)
                if not name or key is None:
                    return
                bucket.setdefault(int(key), set()).add(name)

            for _path in sorted(set(report_paths)):
                try:
                    _df_rep = pd.read_excel(_path, header=None, dtype=object)
                except Exception:
                    continue
                for _, _row in _df_rep.iterrows():
                    _vals = []
                    for _v in _row.tolist():
                        if pd.isna(_v):
                            continue
                        _s = str(_v).strip()
                        if _s:
                            _vals.append(_s)
                    if not _vals:
                        continue
                    _cred = _valid_contract_creditor_name(_vals[0])
                    if not _cred or _cred.upper() in _bad_creditor_labels:
                        continue
                    if re.fullmatch(r'[\d\s./-]+', _cred):
                        continue

                    _txt = ' | '.join(_vals)
                    _title_nums = set()
                    for _m in _title_re.finditer(_txt):
                        _title_nums.add(int(_m.group(1)))
                    for _m in _subst_re.finditer(_txt):
                        _title_nums.add(int(_m.group(1)))
                    for _num in _title_nums:
                        _add_candidate(title_candidates, _num, _cred)

                    for _m in _contract_re.finditer(_txt):
                        _add_candidate(contract_candidates, int(_m.group(1)), _cred)

            _title_map = {k: next(iter(v)) for k, v in title_candidates.items() if len(v) == 1}
            _contract_map = {k: next(iter(v)) for k, v in contract_candidates.items() if len(v) == 1}
            if _title_map or _contract_map:
                print(f"  [CONTRATOS] {len(_title_map)} titulos e {len(_contract_map)} contratos com credor recuperado nos relatorios manuais.")
            return _title_map, _contract_map

        _numtit_report_cred, _nuct_report_cred = _contract_creditors_from_manual_reports()
        cti_med2 = cti_med2.merge(tit[['numtitcp','nmcredor']].drop_duplicates('numtitcp'), on='numtitcp', how='left')
        cti_med2['nmcredor'] = cti_med2['nmcredor'].apply(_valid_contract_creditor_name)
        cti_med2['_numtit_key'] = pd.to_numeric(cti_med2['numtitcp'], errors='coerce')
        cti_med2['_nm_report'] = cti_med2['_numtit_key'].apply(lambda x: _numtit_report_cred.get(int(x)) if pd.notna(x) else None)
        cti_med2['nmcredor'] = cti_med2['nmcredor'].fillna(cti_med2['_nm_report'])
        nuct_cred = cti_med2.dropna(subset=['nmcredor']).groupby('nuct')['nmcredor'].first().to_dict()
        _nuct_cred_obs = {}
        try:
            _obs_col_ct = '_observ' if '_observ' in tit.columns else ('obstit' if 'obstit' in tit.columns else None)
            if _obs_col_ct and 'numtitcp' in tit.columns and 'nmcredor' in tit.columns:
                _tmp_obs_ct = tit[[ _obs_col_ct, 'nmcredor' ]].dropna(subset=[_obs_col_ct]).copy()
                _tmp_obs_ct['nmcredor'] = _tmp_obs_ct['nmcredor'].apply(_valid_contract_creditor_name)
                _obs_candidates = {}
                for _, _robs in _tmp_obs_ct.dropna(subset=['nmcredor']).iterrows():
                    for _m_ct in re.finditer(r'\bCTS\s*/\s*(\d+)\b', str(_robs[_obs_col_ct] or ''), flags=re.I):
                        _obs_candidates.setdefault(int(_m_ct.group(1)), set()).add(str(_robs['nmcredor']).strip())
                _nuct_cred_obs = {k: next(iter(v)) for k, v in _obs_candidates.items() if len(v) == 1}
                if _nuct_cred_obs:
                    print(f"  [CONTRATOS] {len(_nuct_cred_obs)} credores recuperados pelo histórico do título.")
        except Exception as _ex_ct_cred_obs:
            print(f"  [CONTRATOS] fallback credor por histórico ignorado: {_ex_ct_cred_obs}")
        cti_med_ap = cti_med[['docnuct','coditemct','numtitcp','vlrmeditem']].dropna(subset=['numtitcp']).copy()
        cti_med_ap['numtitcp'] = pd.to_numeric(cti_med_ap['numtitcp'], errors='coerce')
        cti_med_ap['vlrmeditem'] = pd.to_numeric(cti_med_ap['vlrmeditem'], errors='coerce').fillna(0)
        cti_med_ap = cti_med_ap.merge(
            tit[['numtitcp','codccusto','STATUS_PAG','vlrapropobra']].drop_duplicates('numtitcp'),
            on='numtitcp', how='left'
        )
        _doc_cc_map_ct = cti_med_ap.dropna(subset=['codccusto']).groupby('docnuct')['codccusto'].first().to_dict()
        cti['codccusto'] = (cti['codccusto']
            .fillna(pd.to_numeric(cti['nuct'], errors='coerce').map(_ct_hdr_cc))
            .fillna(cti['docnuct'].map(_doc_cc_map_ct)))
        cti['nmcc'] = cti['codccusto'].apply(lambda x: _cc_nm_map.get(int(x),'CC '+str(int(x))) if pd.notna(x) else '—')
        cti_med_ap_aberto = cti_med_ap[cti_med_ap['STATUS_PAG']=='A PAGAR'].copy()
        ct_ap_aberto = (cti_med_ap_aberto.drop_duplicates(['docnuct','numtitcp'])
            .groupby('docnuct')['vlrapropobra'].sum().to_dict())
        item_ap_aberto = (cti_med_ap_aberto.groupby(['docnuct','coditemct'])['vlrmeditem'].sum().to_dict())
        # Fallback 1: map codcredor from cti itself via cred_map
        if 'codcredor' in cti.columns:
            cti['_cod_c'] = pd.to_numeric(cti['codcredor'], errors='coerce').fillna(0).astype(int)
            cti['_nm_direct'] = cti['_cod_c'].map(cred_map).apply(_valid_contract_creditor_name)
        else:
            cti['_nm_direct'] = None
        # Fallback 2: map nmcredor directly from cti if column exists
        if 'nmcredor' in cti.columns:
            cti['_nm_cti'] = cti['nmcredor'].apply(_valid_contract_creditor_name)
        else:
            cti['_nm_cti'] = None
        # também tenta docnuct → credor via NF/pedido (contratos incorporação sem título AP)
        _docnuct_cred_nf_ct = {}
        try:
            if '_docnuct_cred_nf' in dir():
                _docnuct_cred_nf_ct = _docnuct_cred_nf
        except Exception:
            pass
        cti['_nm_nf'] = cti['docnuct'].map(_docnuct_cred_nf_ct).apply(_valid_contract_creditor_name)
        # Fallback 3: algumas exportacoes do Sienge nao trazem o credor no
        # cadastro do contrato. Quando o insumo do contrato aparece em NF/pedido
        # com um unico fornecedor possivel, usa esse fornecedor sem chutar casos
        # ambiguos.
        _nuct_cred_by_insumo_nf = {}
        try:
            if 'coditemsol' in cti.columns and 'codinsum' in nf_ped.columns and 'codcredornf' in nf_ped.columns:
                _nf_ins = nf_ped[['codinsum', 'codcredornf']].dropna(subset=['codinsum', 'codcredornf']).copy()
                _nf_ins['codinsum_n'] = pd.to_numeric(_nf_ins['codinsum'], errors='coerce')
                _nf_ins['codcredor_n'] = pd.to_numeric(_nf_ins['codcredornf'], errors='coerce')
                _nf_ins['nmcredor_nf'] = _nf_ins['codcredor_n'].map(cred_map).apply(_valid_contract_creditor_name)
                _nf_ins = _nf_ins.dropna(subset=['codinsum_n', 'nmcredor_nf'])
                _ins_cred_candidates = (_nf_ins.groupby('codinsum_n')['nmcredor_nf']
                    .agg(lambda s: sorted(set(str(v).strip() for v in s if str(v).strip() and str(v).strip() != 'nan')))
                    .to_dict())
                _ins_cred_unique = {int(k): v[0] for k, v in _ins_cred_candidates.items() if len(v) == 1}

                _ct_ins = cti[['nuct', 'coditemsol']].dropna(subset=['nuct', 'coditemsol']).copy()
                _ct_ins['coditemsol_n'] = pd.to_numeric(_ct_ins['coditemsol'], errors='coerce')
                _ct_ins = _ct_ins.dropna(subset=['coditemsol_n'])
                for _nuct, _grp_ins in _ct_ins.groupby('nuct'):
                    _names = sorted(set(
                        _ins_cred_unique.get(int(_v))
                        for _v in _grp_ins['coditemsol_n'].dropna().astype(int)
                        if int(_v) in _ins_cred_unique
                    ))
                    _names = [n for n in _names if n]
                    if len(_names) == 1:
                        _nuct_cred_by_insumo_nf[_nuct] = _names[0]
                if _nuct_cred_by_insumo_nf:
                    print(f"  [CONTRATOS] {len(_nuct_cred_by_insumo_nf)} credores recuperados por NF/pedido.")
        except Exception as _ex_ct_cred_nf:
            print(f"  [CONTRATOS] fallback credor por NF/pedido ignorado: {_ex_ct_cred_nf}")
        cti['_nm_insumo_nf'] = cti['nuct'].map(_nuct_cred_by_insumo_nf)
        # fsup_ct traz codcredorct direto no cabecalho do contrato (sem
        # depender de titulo/medicao existir) — fonte mais confiavel que
        # existe hoje, por isso entra primeiro na cadeia de fallback.
        cti['_nm_hdr'] = pd.to_numeric(cti['nuct'], errors='coerce').map(_ct_hdr_credor).map(cred_map).apply(_valid_contract_creditor_name)
        cti['nmcredor_final'] = cti['_nm_hdr']
        cti['nmcredor_final'] = (cti['nmcredor_final']
            .fillna(cti['nuct'].map(nuct_cred))
            .fillna(cti['nuct'].map(_nuct_cred_obs))
            .fillna(cti['nuct'].map(_nuct_report_cred))
            .fillna(cti.get('_nm_direct'))
            .fillna(cti.get('_nm_cti'))
            .fillna(cti['_nm_nf'])
            .fillna(cti['_nm_insumo_nf'])
            .fillna('Credor não localizado'))
        cti['nmcredor'] = cti['nmcredor_final']
        _orc_by_coditem = {}
        _orc_by_ref = {}
        _nominal_budget = {}
        _nominal_n2 = {}
        _nominal_n3 = {}
        try:
            _nominal_budget = load_nominal_budget_n2()
            _nominal_n2 = _nominal_budget.get('n2', {}) if isinstance(_nominal_budget, dict) else {}
            _nominal_n3 = _nominal_budget.get('n3', {}) if isinstance(_nominal_budget, dict) else {}
            _orc_cols = [c for c in ['coditemorc','refitemorc','nmitemorc','refitemorc2n','vlritemorc','qtitemorc','codunid'] if c in orc4.columns]
            for _, _or in orc4[_orc_cols].dropna(subset=['coditemorc']).iterrows():
                try:
                    _k = str(int(float(_or['coditemorc'])))
                except Exception:
                    _k = str(_or['coditemorc']).strip()
                _ref = str(_or.get('refitemorc', '') or '').strip()
                _n2 = _ref[:6] if _ref else ''
                _n3 = _ref[:10] if _ref else ''
                _orc_by_coditem[_k] = {
                    'ref': _ref,
                    'n2': _n2,
                    'n3': _n3,
                    'nm': str(_or.get('nmitemorc', '') or ''),
                    'n2_nm': str(_or.get('refitemorc2n', '') or ''),
                    'orc_api_item': f2(_or.get('vlritemorc', 0)),
                    'orc_nominal_n2': f2(_nominal_n2.get(_n2, 0)),
                    'orc_nominal_n3': f2(_nominal_n3.get(_n3, 0)),
                    'un_orc': str(_or.get('codunid', '') or ''),
                }
                if _ref:
                    _orc_by_ref[_ref] = _orc_by_coditem[_k]
            _orc_by_nm = {}
            for _info in _orc_by_ref.values():
                _nm_key = _ascii_upper(_info.get('nm', ''))
                if _nm_key and _nm_key not in _orc_by_nm:
                    _orc_by_nm[_nm_key] = _info
        except Exception:
            _orc_by_coditem = {}
            _orc_by_ref = {}
            _orc_by_nm = {}
        # groupby() descarta silenciosamente linhas cuja chave e NaN — contratos
        # sem nenhum titulo/medicao vinculado ainda (ex: CTS/1073-1075) ficam
        # com codccusto NaN (nao ha titulo pra descobrir o centro de custo) e
        # desapareciam da lista inteira. codccusto=0 mantem o contrato visivel
        # (nmcc ja fica '-' pra esses, sinalizando centro de custo desconhecido).
        cti['codccusto'] = cti['codccusto'].fillna(0)
        ct_grp = cti.groupby(['docnuct','nuct','codccusto','nmcc','nmcredor']).agg(
            vlrtotct=('vlrtotitem','sum'),vlrmedidoct=('vlrmedidoct','sum'),
            vlrpendct=('vlrpendct','sum'),vlrliberct=('vlrliberct','sum'),itens=('coditemct','count')
        ).reset_index()
        ct_grp['vlrapct'] = ct_grp['docnuct'].map(ct_ap_aberto).fillna(0)
        ct_grp['pct'] = ct_grp.apply(lambda r: round(f2(r.vlrmedidoct)/f2(r.vlrtotct)*100,1) if f2(r.vlrtotct)>0 else 0.0, axis=1)
        contratos_list = [{'docnuct':str(r.docnuct),'nuct':int(r.nuct),'nmcc':str(r.nmcc),'nmcredor':str(r.nmcredor),
            'cnpj':cred_name_to_cnpj.get(_norm_company_name(r.nmcredor), ''),
            'nome_fantasia':(cnpj_cache.get(cred_name_to_cnpj.get(_norm_company_name(r.nmcredor), ''), {}) or {}).get('nome_fantasia', ''),
            'vlrtotct':f2(r.vlrtotct),'vlrmedidoct':f2(r.vlrmedidoct),'vlrpendct':f2(r.vlrapct),'vlrliberct':f2(r.vlrliberct),
            'pct':float(r.pct),'itens':int(r.itens),'codccusto':int(r.codccusto) if pd.notna(r.codccusto) else 0,
            'status_sienge':_ct_hdr_status.get(int(r.nuct), ''),'aprovacao_sienge':_ct_hdr_sitaut.get(int(r.nuct), '')
            } for _,r in ct_grp.sort_values('vlrtotct',ascending=False).iterrows()]
        itens_ct = {}
        _minuta_unid_map = {'16':'m','17':'m²','18':'m³','19':'mês','34':'un','35':'vb','88':'m³','89':'mês'}
        def _minuta_unid_label(v):
            raw = str(v or '').strip()
            if raw.endswith('.0'):
                raw = raw[:-2]
            return _minuta_unid_map.get(raw, 'un')
        for docnuct, grp in cti.groupby('docnuct'):
            itens_ct[str(docnuct)] = [{'cod':str(r.coditemct),'nm':str(r.nmitemct),
                'vlrtot':f2(r.vlrtotitem),'vlrmed':f2(r.vlrmedidoct),'vlrpend':f2(item_ap_aberto.get((r.docnuct, r.coditemct), 0)),
                'pct':round(f2(r.vlrmedidoct)/f2(r.vlrtotitem)*100,1) if f2(r.vlrtotitem)>0 else 0.0,
                'qt':f2(r.get('qtitemct',0)),'un':_minuta_unid_label(r.get('codunidmeditemct','')),'pr_unit':f2(r.get('pritemunit',0))
            } for _,r in grp.sort_values('coditemct').iterrows()]
        contract_analysis = {}
        try:
            def _dt_out(v):
                d = pd.to_datetime(v, errors='coerce')
                return str(d.date()) if pd.notna(d) else '—'
            def _coditem_key(v):
                try:
                    return str(int(float(v)))
                except Exception:
                    return str(v or '').strip()
            def _contract_unit_label(v, fallback=''):
                raw = str(v or '').strip()
                if raw.endswith('.0'):
                    raw = raw[:-2]
                unit_map = {
                    '16': 'm',
                    '17': 'm²',
                    '18': 'm³',
                    '19': 'mes',
                    '34': 'un',
                    '35': 'vb',
                    '88': 'm³',
                    '89': 'mes',
                }
                if raw.lower() in {'', 'nan', 'none', 'null'}:
                    raw = ''
                fb = str(fallback or '').strip()
                if fb.endswith('.0'):
                    fb = fb[:-2]
                return unit_map.get(raw) or unit_map.get(fb) or (fb if fb and not fb.isdigit() else raw)
            _cti_med_det = cti_med.copy()
            for _c in ['qtitemct','pritemunit','vlrtotitem','qtmeditem','qtsditem','vlrmeditem','vlrsditem','numtitcp','numed','coditemct','coditemorc']:
                if _c in _cti_med_det.columns:
                    _cti_med_det[_c] = pd.to_numeric(_cti_med_det[_c], errors='coerce')
            _tit_contract_cols = [c for c in ['numtitcp','STATUS_PAG','dtbaixacp','dtvenc','vlrapropobra','vlrliqbaixacp','coddocum','docnuct','nmcredor','refitemorc'] if c in tit.columns]
            _tit_contract = tit[_tit_contract_cols].drop_duplicates('numtitcp').copy() if 'numtitcp' in tit.columns else pd.DataFrame()
            if not _tit_contract.empty:
                _tit_contract['numtitcp'] = pd.to_numeric(_tit_contract['numtitcp'], errors='coerce')
            _med_with_tit = _cti_med_det.merge(_tit_contract, on='numtitcp', how='left', suffixes=('', '_tit')) if not _tit_contract.empty else _cti_med_det.copy()
            _item_refs_lookup = {}
            _ap_by_ref = {}
            _contract_ref_values = {}
            _contract_ref_open = {}
            if 'refitemorc' in tit.columns and 'vlrapropobra' in tit.columns:
                _ap_ref_src = tit[['refitemorc','vlrapropobra']].dropna(subset=['refitemorc']).copy()
                _ap_ref_src['refitemorc'] = _ap_ref_src['refitemorc'].astype(str).str.strip()
                _ap_ref_src = _ap_ref_src[_ap_ref_src['refitemorc'].str.len() == 14]
                _ap_ref_src['vlrapropobra'] = pd.to_numeric(_ap_ref_src['vlrapropobra'], errors='coerce').fillna(0)
                _ap_by_ref = {str(k): f2(v) for k, v in _ap_ref_src.groupby('refitemorc')['vlrapropobra'].sum().to_dict().items()}
            if 'numtitcp' in tit.columns and 'refitemorc' in tit.columns:
                _tit_ref_multi = tit[['numtitcp','refitemorc','vlrapropobra']].dropna(subset=['numtitcp','refitemorc']).copy()
                _tit_ref_multi['numtitcp'] = pd.to_numeric(_tit_ref_multi['numtitcp'], errors='coerce')
                _tit_ref_multi['refitemorc'] = _tit_ref_multi['refitemorc'].astype(str).str.strip()
                _tit_ref_multi['vlrapropobra'] = pd.to_numeric(_tit_ref_multi['vlrapropobra'], errors='coerce').fillna(0)
                _tit_ref_multi = _tit_ref_multi[_tit_ref_multi['refitemorc'].str.len() == 14]
                _tit_ref_multi = _tit_ref_multi.groupby(['numtitcp','refitemorc'], as_index=False)['vlrapropobra'].sum()
                _med_ref_src = _cti_med_det.dropna(subset=['docnuct','coditemct','numtitcp']).copy()
                if not _med_ref_src.empty:
                    _med_ref_src['numtitcp'] = pd.to_numeric(_med_ref_src['numtitcp'], errors='coerce')
                    _med_ref_src['vlrmeditem'] = pd.to_numeric(_med_ref_src.get('vlrmeditem', 0), errors='coerce').fillna(0)
                    _med_ref_src['qtmeditem'] = pd.to_numeric(_med_ref_src.get('qtmeditem', 0), errors='coerce').fillna(0)
                    _med_ref_src = _med_ref_src[(_med_ref_src['vlrmeditem'].abs() > 0.005) | (_med_ref_src['qtmeditem'].abs() > 0.005)].copy()
                _assigned_refs = []
                for _numtit, _mr in (_med_ref_src.groupby('numtitcp') if not _med_ref_src.empty else []):
                    _refs_for_title = _tit_ref_multi[_tit_ref_multi['numtitcp'] == _numtit].copy()
                    if _refs_for_title.empty:
                        continue
                    _ref_records = _refs_for_title.to_dict('records')
                    if len(_ref_records) == 1:
                        _ref = _ref_records[0]['refitemorc']
                        for _, _mrow in _mr.iterrows():
                            _assigned_refs.append({
                                'docnuct': str(_mrow.get('docnuct')),
                                'coditemct': _mrow.get('coditemct'),
                                'numtitcp': _numtit,
                                'refitemorc': _ref,
                                'valor_medido': abs(f2(_mrow.get('vlrmeditem', 0))),
                            })
                    else:
                        _available = list(_ref_records)
                        _med_rows = sorted(_mr.to_dict('records'), key=lambda x: abs(f2(x.get('vlrmeditem', 0))), reverse=True)
                        for _mrow in _med_rows:
                            if not _available:
                                break
                            _mval = abs(f2(_mrow.get('vlrmeditem', 0)))
                            _best_idx = min(range(len(_available)), key=lambda idx: abs(abs(f2(_available[idx].get('vlrapropobra', 0))) - _mval))
                            _best = _available.pop(_best_idx)
                            _assigned_refs.append({
                                'docnuct': str(_mrow.get('docnuct')),
                                'coditemct': _mrow.get('coditemct'),
                                'numtitcp': _numtit,
                                'refitemorc': _best.get('refitemorc'),
                                'valor_medido': _mval,
                            })
                if _assigned_refs:
                    _assigned_df = pd.DataFrame(_assigned_refs)
                    if not _tit_contract.empty and 'numtitcp' in _tit_contract.columns:
                        _st_cols = [c for c in ['numtitcp','STATUS_PAG'] if c in _tit_contract.columns]
                        if _st_cols:
                            _assigned_df = _assigned_df.merge(_tit_contract[_st_cols].drop_duplicates('numtitcp'), on='numtitcp', how='left')
                    _assigned_df['valor_medido'] = pd.to_numeric(_assigned_df.get('valor_medido', 0), errors='coerce').fillna(0)
                    for (_d, _ref), _rg in _assigned_df.groupby(['docnuct','refitemorc']):
                        _contract_ref_values[(str(_d), str(_ref))] = f2(_rg['valor_medido'].sum())
                        if 'STATUS_PAG' in _rg.columns:
                            _open = _rg[_rg['STATUS_PAG'].astype(str).str.upper().ne('PAGO')]['valor_medido'].sum()
                        else:
                            _open = 0
                        _contract_ref_open[(str(_d), str(_ref))] = f2(_open)
                    for (_d, _i), _rg in _assigned_df.groupby(['docnuct','coditemct']):
                        _refs = []
                        for _ref, _rr in _rg.groupby('refitemorc'):
                            _val = f2(_rr.get('valor_medido', pd.Series(dtype=float)).sum())
                            _info = _orc_by_ref.get(_ref, {})
                            _n3 = str(_ref)[:10]
                            _refs.append({
                                'ref': _ref,
                                'nm': _info.get('nm', ''),
                                'n2': str(_ref)[:6],
                                'n3': _n3,
                                'n2_nm': _info.get('n2_nm', ''),
                                'orc_nominal_n3': f2(_nominal_n3.get(_n3, 0)),
                                'valor_medido': abs(_val),
                            })
                        _refs.sort(key=lambda x: x.get('valor_medido', 0), reverse=True)
                        _item_refs_lookup[(str(_d), _i)] = _refs
            for docnuct, grp in cti.groupby('docnuct'):
                _doc = str(docnuct)
                _meta = ct_grp[ct_grp['docnuct'].astype(str) == _doc]
                _meta_row = _meta.iloc[0] if not _meta.empty else grp.iloc[0]
                _items = []
                _seen_contract_refs = set()
                _total_orcado_apropriacao = 0.0
                _total_consumido_apropriacao = 0.0
                _total_a_pagar_apropriacao = 0.0
                _total_saldo_apropriacao = 0.0
                for _, r in grp.sort_values('coditemct').iterrows():
                    _cod_key = _coditem_key(r.get('coditemorc', ''))
                    _orc_info = _orc_by_coditem.get(_cod_key, {})
                    _linked_refs = _item_refs_lookup.get((_doc, r.get('coditemct')), [])
                    if _linked_refs:
                        _ref = _linked_refs[0].get('ref', '')
                        _orc_info = _orc_by_ref.get(_ref, _orc_info)
                    else:
                        _ref = _orc_info.get('ref', '')
                    if not _ref:
                        _svc_norm = _ascii_upper(r.get('nmitemct', ''))
                        _nm_info = _orc_by_nm.get(_svc_norm, {})
                        if _nm_info:
                            _orc_info = _nm_info
                            _ref = _orc_info.get('ref', '')
                    if not _ref:
                        _svc_norm = _ascii_upper(r.get('nmitemct', ''))
                        if 'BOTA FORA' in _svc_norm and 'BLOCO' in _svc_norm:
                            _ref = '02.003.002.003'
                            _orc_info = _orc_by_ref.get(_ref, _orc_info)
                    if not _ref:
                        # o item do contrato TEM um código de apropriação lançado
                        # no Sienge (coditemorc na própria minuta) — só não bate com
                        # nenhum item do orçamento nominal (numeração própria da
                        # minuta, diferente do id interno do eng_orc) nem tem
                        # título/medição pra confirmar valores ainda. Mostrar como
                        # "apropriação pendente" seria enganoso: o item JÁ está
                        # apropriado, só falta o lançamento financeiro (autorização).
                        _cod_bruto = str(r.get('coditemorc', '') or '').strip()
                        if _cod_bruto and _cod_bruto.lower() not in ('nan', 'none', '0', '0.0'):
                            _ref = _cod_bruto
                    _n2 = _orc_info.get('n2', '')
                    _n3 = _orc_info.get('n3', (_ref[:10] if _ref else ''))
                    _vlrtot = f2(r.get('vlrtotitem', 0))
                    _vlrmed = f2(r.get('vlrmedidoct', 0))
                    _saldo = max(f2(_vlrtot - _vlrmed), 0)
                    _qtd_contrato = f2(r.get('qtitemct', 0))
                    _qtd_medida = f2(_med_with_tit[(_med_with_tit['docnuct'].astype(str)==_doc) & (_med_with_tit['coditemct']==r.get('coditemct'))]['qtmeditem'].sum()) if 'qtmeditem' in _med_with_tit.columns else 0
                    _qtd_saldo = max(f2(_qtd_contrato - _qtd_medida), 0)
                    _is_empty_item = (
                        abs(_qtd_contrato) < 0.005 and
                        abs(f2(r.get('pritemunit', 0))) < 0.005 and
                        abs(_vlrtot) < 0.005 and
                        abs(_vlrmed) < 0.005 and
                        abs(_saldo) < 0.005
                    )
                    _refs_to_calc = []
                    if not _is_empty_item:
                        if _linked_refs:
                            _refs_to_calc = [str(_lr.get('ref', '') or '').strip() for _lr in _linked_refs if str(_lr.get('ref', '') or '').strip()]
                        elif _ref:
                            _refs_to_calc = [_ref]
                    _orcado_apropriacao = 0.0
                    _consumido_apropriacao = 0.0
                    _a_pagar_apropriacao = 0.0
                    _saldo_apropriacao = 0.0
                    for _calc_ref in dict.fromkeys(_refs_to_calc):
                        if len(_calc_ref) != 14 or _calc_ref in _seen_contract_refs:
                            continue
                        _seen_contract_refs.add(_calc_ref)
                        _calc_info = _orc_by_ref.get(_calc_ref, {})
                        _orc_ref = f2(_calc_info.get('orc_api_item', 0))
                        _cons_ref = f2(_contract_ref_values.get((_doc, _calc_ref), _ap_by_ref.get(_calc_ref, 0)))
                        _open_ref = f2(_contract_ref_open.get((_doc, _calc_ref), 0))
                        _saldo_ref = f2(_orc_ref - _cons_ref)
                        _orcado_apropriacao = f2(_orcado_apropriacao + _orc_ref)
                        _consumido_apropriacao = f2(_consumido_apropriacao + _cons_ref)
                        _a_pagar_apropriacao = f2(_a_pagar_apropriacao + _open_ref)
                        _saldo_apropriacao = f2(_saldo_apropriacao + _saldo_ref)
                    _total_orcado_apropriacao = f2(_total_orcado_apropriacao + _orcado_apropriacao)
                    _total_consumido_apropriacao = f2(_total_consumido_apropriacao + _consumido_apropriacao)
                    _total_a_pagar_apropriacao = f2(_total_a_pagar_apropriacao + _a_pagar_apropriacao)
                    _total_saldo_apropriacao = f2(_total_saldo_apropriacao + _saldo_apropriacao)
                    _un = _contract_unit_label(r.get('codunidmeditemct', ''), _orc_info.get('un_orc', ''))
                    _items.append({
                        'cod': str(r.get('coditemct', '') or ''),
                        'servico': str(r.get('nmitemct', '') or ''),
                        'detalhe': str(r.get('nmdetitemct', '') or '') if pd.notna(r.get('nmdetitemct', '')) else '',
                        'un': _un,
                        'preco_unit': f2(r.get('pritemunit', 0)),
                        'qtd_contrato': _qtd_contrato,
                        'valor_contrato': _vlrtot,
                        'qtd_medida': _qtd_medida,
                        'valor_medido': _vlrmed,
                        'qtd_saldo': _qtd_saldo,
                        'valor_saldo': _saldo,
                        'is_empty_item': bool(_is_empty_item),
                        'orcado_apropriacao': _orcado_apropriacao,
                        'consumido_apropriacao': _consumido_apropriacao,
                        'a_pagar_apropriacao': _a_pagar_apropriacao,
                        'saldo_apropriacao': _saldo_apropriacao,
                        'uc': str(r.get('coducct', '') or ''),
                        'apropriacao': _ref,
                        'apropriacao_nm': _orc_info.get('nm', ''),
                        'apropriacoes_medidas': _linked_refs[:6],
                        'n2': _n2,
                        'n3': _n3,
                        'n2_nm': _orc_info.get('n2_nm', ''),
                        'orcado_nominal_n2': f2(_orc_info.get('orc_nominal_n2', 0)),
                        'orcado_nominal_n3': _orcado_apropriacao,
                    })
                _med_grp = _med_with_tit[_med_with_tit['docnuct'].astype(str) == _doc].copy()
                _medicoes = []
                if not _med_grp.empty:
                    _med_grp = _med_grp[(_med_grp.get('vlrmeditem', 0).fillna(0).abs() > 0.005) | (_med_grp.get('qtmeditem', 0).fillna(0).abs() > 0.005)].copy()
                    for _, mr in _med_grp.sort_values(['dtmed','numed','coditemct']).iterrows():
                        _numtit = mr.get('numtitcp', None)
                        _status = str(mr.get('STATUS_PAG', '') or '').strip()
                        _medicoes.append({
                            'medicao': str(int(mr['numed'])) if pd.notna(mr.get('numed', None)) else '—',
                            'data_medicao': _dt_out(mr.get('dtmed')),
                            'vencimento': _dt_out(mr.get('dtvenctomed')),
                            'titulo': str(int(_numtit)) if pd.notna(_numtit) else '—',
                            'status_titulo': _status or '—',
                            'item': str(mr.get('coditemct', '') or ''),
                            'servico': str(mr.get('nmitemct', '') or ''),
                            'qtd': f2(mr.get('qtmeditem', 0)),
                            'valor': f2(mr.get('vlrmeditem', 0)),
                            'usuario': str(mr.get('usurespmed', '') or ''),
                            'situacao_medicao': str(mr.get('statmed', '') or ''),
                            'dt_baixa': _dt_out(mr.get('dtbaixacp')),
                        })
                _previsoes = []
                if not _med_grp.empty and 'numtitcp' in _med_grp.columns:
                    _pay = _med_grp.dropna(subset=['numtitcp']).copy()
                    if not _pay.empty:
                        _pay['numtitcp'] = pd.to_numeric(_pay['numtitcp'], errors='coerce')
                        _pay_q = _pay.groupby('numtitcp', dropna=True).agg(
                            valor=('vlrmeditem','sum'),
                            data=('dtvenctomed','first'),
                            status=('STATUS_PAG','first'),
                            baixa=('dtbaixacp','first'),
                        ).reset_index()
                        for _, pr in _pay_q.sort_values('data').iterrows():
                            _st = 'Paga' if str(pr.get('status','')).upper() == 'PAGO' or pd.notna(pd.to_datetime(pr.get('baixa'), errors='coerce')) else 'Em aberto'
                            _previsoes.append({
                                'titulo': str(int(pr['numtitcp'])) if pd.notna(pr.get('numtitcp')) else '—',
                                'data': _dt_out(pr.get('data')),
                                'valor': f2(pr.get('valor', 0)),
                                'situacao': _st,
                            })
                contract_analysis[_doc] = {
                    'docnuct': _doc,
                    'contrato': _doc,
                    'nuct': int(_meta_row.get('nuct', 0)) if pd.notna(_meta_row.get('nuct', None)) else 0,
                    'fornecedor': str(_meta_row.get('nmcredor', '') or '—'),
                    'centro_custo': str(_meta_row.get('nmcc', '') or '—'),
                    'codccusto': int(_meta_row.get('codccusto', 0)) if pd.notna(_meta_row.get('codccusto', None)) else 0,
                    'valor_total': f2(_meta_row.get('vlrtotct', 0)),
                    'valor_medido': f2(_meta_row.get('vlrmedidoct', 0)),
                    'saldo_contrato': max(f2(_meta_row.get('vlrtotct', 0) - _meta_row.get('vlrmedidoct', 0)), 0),
                    'a_pagar_lancado': f2(ct_ap_aberto.get(docnuct, 0)),
                    'orcamento_nominal_obra': f2((_nominal_budget or {}).get('total', 0)) if isinstance(_nominal_budget, dict) else 0,
                    'total_orcado_apropriacao': _total_orcado_apropriacao,
                    'total_consumido_apropriacao': _total_consumido_apropriacao,
                    'total_a_pagar_apropriacao': _total_a_pagar_apropriacao,
                    'total_saldo_apropriacao': _total_saldo_apropriacao,
                    'itens': _items,
                    'medicoes': _medicoes,
                    'previsoes': _previsoes,
                }
        except Exception as _contract_exc:
            print(f"  [CONTRATOS ANALISE] falha ao montar painel: {_contract_exc}")
            contract_analysis = {}

        # ── vendas ───────────────────────────────────────────────────────────
        for c in ['vlrtotal','vlrprvenda','areapriv']:
            vnd[c] = pd.to_numeric(vnd[c], errors='coerce').fillna(0)
        vnd['dtct']      = pd.to_datetime(vnd['dtct'], errors='coerce')
        vnd['dtcancela'] = pd.to_datetime(vnd['dtcancela'], errors='coerce')
        ativos = vnd[vnd['flsituacao']=='Emitido'].copy()
        ativos['mes'] = ativos['dtct'].dt.to_period('M').astype(str)
        vnd_mes = ativos.groupby('mes').agg(qt=('codct','count'),vlr=('vlrtotal','sum')).reset_index().sort_values('mes')
        vnd_mes_list = [{'mes':str(r.mes),'qt':int(r.qt),'vlr':f2(r.vlr)} for _,r in vnd_mes.iterrows()]
        vnd_status_list = [{'status':str(r.flsituacao),'qt':int(r.qt)} for _,r in vnd.groupby('flsituacao').size().reset_index(name='qt').iterrows()]
        vnd_table = [{'codct':str(r.codct),'nuunid':str(int(r.nuunid)) if pd.notna(r.nuunid) else '—',
            'flsituacao':str(r.flsituacao),
            'dtct':str(r.dtct.date()) if pd.notna(r.dtct) else '—',
            'dtcancela':str(r.dtcancela.date()) if pd.notna(r.dtcancela) else '—',
            'area':f2(r.areapriv),'vlrprvenda':f2(r.vlrprvenda),'vlrtotal':f2(r.vlrtotal),
            'tpvenda':str(r.tpvenda),'nmcatvenda':str(r.nmcatvenda)
        } for _,r in vnd.sort_values('dtct').iterrows()]

        # ── summary ──────────────────────────────────────────────────────────
        vlr_orcado = float(orc4['vlritemorc'].sum())
        def smry_ap(sub):
            ap   = float(sub['vlrapropobra'].sum())
            pg   = float(sub['vlrliqbaixacp'].sum())
            comp = float(sub[sub['STATUS_PAG']=='A PAGAR']['vlrapropobra'].sum())
            return {'vlr_orcado':round(vlr_orcado,2),'vlr_desembolsado':round(pg,2),
                    'vlr_comprometido':round(comp,2),'vlr_apropriado':round(ap,2),
                    'vlr_saldo':round(vlr_orcado-ap,2),
                    'pct_realizado':round(ap/vlr_orcado*100,2) if vlr_orcado>0 else 0,
                    'fonte':'AP'}
        def smry_banco(cc_code):
            sub = bancos[bancos['codccusto_n']==cc_code]
            pag_out = float(sub[sub['vlrmov']<0]['vlrmov'].abs().sum())
            rec_in  = float(sub[sub['vlrmov']>0]['vlrmov'].sum())
            return {'vlr_orcado':0,'vlr_desembolsado':round(pag_out,2),
                    'vlr_comprometido':0,'vlr_apropriado':round(rec_in,2),
                    'vlr_saldo':round(rec_in-pag_out,2),'pct_realizado':0,
                    'fonte':'BANCO'}

        summary_all = smry_ap(tit)
        summary_all.update({
            'vgv':round(float(ativos['vlrtotal'].sum()),2),
            'unidades':int(len(ativos)),
            'unidades_canceladas':int(len(vnd[vnd['flsituacao']=='Cancelado'])),
            'vgv_cancelado':round(float(vnd[vnd['flsituacao']=='Cancelado']['vlrtotal'].sum()),2),
        })
        summary_per_cc = {}
        for cc in tit['codccusto'].dropna().unique():
            summary_per_cc[str(int(cc))] = smry_ap(tit[tit['codccusto']==cc])
        for cc in bancos['codccusto_n'].dropna().unique():
            cc_key = str(int(cc))
            if cc_key not in summary_per_cc:
                summary_per_cc[cc_key] = smry_banco(cc)


        # ── tit_extra já foi computado antes da árvore ORC ──────────────────

        # ── cronograma ───────────────────────────────────────────────────────
        try:
            plan = xl.parse('feng_plan')
            plan = plan[pd.to_numeric(plan.get('codobra'), errors='coerce') == 1016].copy()
            plan['dtinicio']  = pd.to_datetime(plan['dtinicio'],  errors='coerce')
            plan['dttermino'] = pd.to_datetime(plan['dttermino'], errors='coerce')
            plan['mesplan']   = pd.to_datetime(plan.get('mesplan', pd.Series(dtype='datetime64[ns]')), errors='coerce')
            plan['vlrplan']   = pd.to_numeric(plan['vlrplan'], errors='coerce').fillna(0)
            plan['pesotar']   = pd.to_numeric(plan.get('pesotar', pd.Series(dtype=float)), errors='coerce').fillna(0)
            plan['peplan']    = pd.to_numeric(plan.get('peplan', pd.Series(dtype=float)), errors='coerce').fillna(0)
            plan['diasdura']  = pd.to_numeric(plan.get('diasdura', pd.Series(dtype=float)), errors='coerce').fillna(0)
            pct_col = next((c for c in ['pctconcl','pctrealar','pctreal','pctconc'] if c in plan.columns), None)
            plan_list = []
            cost_plan_list = []
            cost_grp = plan.groupby(['nutarefa','nmtarefa','refitemorc'], dropna=False).agg(
                inicio=('dtinicio','min'), termino=('dttermino','max'), dias=('diasdura','max'),
                vlr_plan=('vlrplan','sum'), peso=('pesotar','sum'),
                nivel=('codnivel','max'), tipo=('tptarefa','first'),
            ).reset_index()
            for _, r in cost_grp.iterrows():
                cost_plan_list.append({
                    'id': str(int(r.nutarefa)) if pd.notna(r.nutarefa) else '',
                    'tarefa': str(r.nmtarefa), 'eap': str(r.refitemorc) if pd.notna(r.refitemorc) else '',
                    'inicio': str(r.inicio.date()) if pd.notna(r.inicio) else '—',
                    'termino': str(r.termino.date()) if pd.notna(r.termino) else '—',
                    'dias': int(r.dias) if pd.notna(r.dias) else 0, 'vlr_plan': f2(r.vlr_plan),
                    'peso': f2(r.peso), 'pct_real': 0.0,
                    'nivel': int(r.nivel) if pd.notna(r.nivel) else 0,
                    'tipo': str(r.tipo) if pd.notna(r.tipo) else '',
                })
            plan_dates = plan[plan['dtinicio'].notna() & plan['dttermino'].notna()].copy()
            if not plan_dates.empty:
                task_grp = plan_dates.groupby(['nutarefa','nmtarefa'], dropna=False).agg(
                    inicio=('dtinicio','min'), termino=('dttermino','max'), dias=('diasdura','max'),
                    vlr_plan=('vlrplan','sum'), peso=('pesotar','sum'),
                    pct_real=(pct_col, 'max') if pct_col else ('pesotar', lambda x: 0),
                    nivel=('codnivel','max'), tipo=('tptarefa','first'), eap=('refitemorc','first'),
                ).reset_index()
                for _,r in task_grp.iterrows():
                    plan_list.append({'id':str(int(r.nutarefa)) if pd.notna(r.nutarefa) else '',
                        'tarefa':str(r.nmtarefa),
                        'inicio':str(r.inicio.date()) if pd.notna(r.inicio) else '—',
                        'termino':str(r.termino.date()) if pd.notna(r.termino) else '—',
                        'dias':int(r.dias) if pd.notna(r.dias) else 0,
                        'vlr_plan':f2(r.vlr_plan),'peso':f2(r.peso),
                        'pct_real':round(float(r.pct_real),2) if pd.notna(r.pct_real) else 0.0,
                        'nivel':int(r.nivel) if pd.notna(r.nivel) else 0,
                        'tipo':str(r.tipo) if pd.notna(r.tipo) else '',
                        'eap':str(r.eap) if pd.notna(r.eap) else ''})
            plan_list.sort(key=lambda x: (x['inicio'], x['termino'], x['tarefa']))
            lb_src = plan[plan['mesplan'].notna()].copy()
            linha_balanco = []
            if not lb_src.empty:
                lb_grp = lb_src.groupby(lb_src['mesplan'].dt.to_period('M')).agg(peplan=('peplan','sum'), vlrplan=('vlrplan','sum')).reset_index()
                total_pe = float(lb_grp['peplan'].sum()) or 100.0
                acc = 0.0
                for _,r in lb_grp.sort_values('mesplan').iterrows():
                    acc += float(r.peplan or 0) / total_pe * 100
                    linha_balanco.append({'mes':str(r.mesplan), 'planejado':round(acc,2), 'vlr_plan':f2(r.vlrplan)})
            if len(plan_list) < 20:
                pkg_list, pkg_lb = load_crono_pkgs_from_html()
                if pkg_list:
                    plan_list = pkg_list
                    linha_balanco = pkg_lb
                    crono_fonte = 'PLAN_BI_TMZC_SYNERGIA.html extraído do Project'
                else:
                    crono_fonte = 'feng_plan extraído do Project'
            else:
                crono_fonte = 'feng_plan extraído do Project'
            crono_meta = {
                'fonte': crono_fonte,
                'project_file': os.path.basename(PROJECT_MPP_PATH) if os.path.exists(PROJECT_MPP_PATH) else '',
                'project_path': PROJECT_MPP_PATH if os.path.exists(PROJECT_MPP_PATH) else '',
                'tarefas': len(plan_list),
                'inicio': min((r['inicio'] for r in plan_list if r['inicio'] != '—'), default='—'),
                'termino': max((r['termino'] for r in plan_list if r['termino'] != '—'), default='—'),
                'niveis': sorted(set(r.get('nivel', 0) for r in plan_list)),
            }
        except Exception:
            plan_list = []
            cost_plan_list = []
            linha_balanco = []
            crono_meta = {'fonte':'', 'project_file':'', 'tarefas':0, 'inicio':'—', 'termino':'—', 'niveis':[]}

        pa_tasks, pa_linha, histograma_equipe, pa_meta = load_pa_cronograma()
        if pa_tasks:
            plan_list = pa_tasks
            cost_plan_list = pa_tasks
            linha_balanco = pa_linha
            crono_meta = pa_meta
        else:
            histograma_equipe = []
        nominal_budget = load_nominal_budget_n2()
        support_histograma = nominal_budget.get("histograma_equipe") if isinstance(nominal_budget, dict) else []
        if support_histograma:
            histograma_equipe = support_histograma

        try:
            previsto_termino_folha = _previsto_termino_folha(xl, histograma_equipe)
        except Exception:
            previsto_termino_folha = {"ok": False, "meses": []}

        # cronograma lido direto do .mpp (via "Atualizar do Project" na aba
        # Planejamento) é a fonte mais atual quando disponível — sobrepõe
        # tudo acima, exceto o histograma de equipe (o .mpp não tem alocação
        # de recursos lançada, então mantém o que já foi resolvido).
        mpp_tasks, mpp_linha, _mpp_hist, mpp_meta = load_mpp_cronograma_cached()
        if mpp_tasks:
            plan_list = mpp_tasks
            cost_plan_list = mpp_tasks
            linha_balanco = mpp_linha
            crono_meta = mpp_meta

        activity_costs, activity_cost_meta, sinapi_meta = build_activity_cost_analysis(cost_plan_list or plan_list)

        # ── avanço financeiro mensal ─────────────────────────────────────────
        try:
            pg2 = tit[tit['dtbaixacp'].notna()].copy()
            pg2['mes_bx'] = pg2['dtbaixacp'].dt.to_period('M').astype(str)
            av_fin = pg2.groupby('mes_bx')['vlrliqbaixacp'].sum().reset_index()
            av_fin_list = [{'mes':str(r.mes_bx),'vlr':f2(r.vlrliqbaixacp)} for _,r in av_fin.sort_values('mes_bx').iterrows()]
        except Exception:
            av_fin_list = []

        # ── distorcoes ───────────────────────────────────────────────────────
        try:
            DIST_DATE = pd.Timestamp('2025-12-03')
            CUTOFF_DATE = pd.Timestamp('2026-01-13')
            paid_since = tit[tit['dtbaixacp'] >= DIST_DATE].copy()
            # Pagamentos ANTES de 13/01/2026 (para tabela separada)
            paid_before_cutoff = tit[(tit['dtbaixacp'].notna()) & (tit['dtbaixacp'] < CUTOFF_DATE)].copy()
            paid_since_list = []
            for _,r in paid_since.sort_values('dtbaixacp').iterrows():
                ex = tit_extra.get(int(r['numtitcp']),{})
                paid_since_list.append({
                    'numtit':str(r['numtitcp']),'credor':str(r['nmcredor']),
                    'ref':str(r['refitemorc']),'vlr_ap':f2(r['vlrapropobra']),
                    'vlr_pg':f2(r['vlrliqbaixacp']),
                    'dt_baixa':str(r['dtbaixacp'].date()),
                    'obs':str(r.get('obstit','') or ''),
                    'pedidos':ex.get('pedidos',[]),'insumos':ex.get('insumos',[]),
                    'cc':str(int(r['codccusto'])) if pd.notna(r['codccusto']) else '—',
                    'tpbaixa':str(r.get('tpbaixa','') or ''),
                })
            # Lista pagamentos antes de 13/01/2026
            paid_before_list = []
            for _,r in paid_before_cutoff.sort_values('dtbaixacp').iterrows():
                ex = tit_extra.get(int(r['numtitcp']),{})
                paid_before_list.append({
                    'numtit':str(r['numtitcp']),'credor':str(r['nmcredor']),
                    'ref':str(r['refitemorc']),'vlr_ap':f2(r['vlrapropobra']),
                    'vlr_pg':f2(r['vlrliqbaixacp']),
                    'dt_baixa':str(r['dtbaixacp'].date()),
                    'tpbaixa':str(r.get('tpbaixa','') or ''),
                    'cc':str(int(r['codccusto'])) if pd.notna(r['codccusto']) else '—',
                    'pedidos':ex.get('pedidos',[]),
                })
        except Exception:
            paid_since_list = []
            paid_before_list = []

        try:
            # Itens com alguma apropriação
            refs_com_tit = set(tit['refitemorc'].astype(str).str.strip().unique())
            sem_aprop = []
            for _,r4 in orc4.iterrows():
                ref = str(r4.refitemorc).strip()
                if f2(r4.vlritemorc) > 0:
                    aprop = f2(r4.realizado)
                    saldo = f2(r4.saldo)
                    # Inclui todos os itens sem apropriação OU com saldo remanescente
                    if ref not in refs_com_tit:
                        sem_aprop.append({'ref':ref,'nm':str(r4.nmitemorc),'un':str(r4.codunid),
                            'orcado':f2(r4.vlritemorc),'aprop':aprop,'saldo':saldo,
                            'n1':str(r4.get('refitemorc1n','')) if pd.notna(r4.get('refitemorc1n','')) else '',
                            'n3':str(r4.get('refitemorc3n','')) if pd.notna(r4.get('refitemorc3n','')) else '',
                        })
            sem_aprop.sort(key=lambda x: x['orcado'], reverse=True)
        except Exception:
            sem_aprop = []

        try:
            dist_rel_sem_aprop = build_distorcoes_relatorio_local(tit)
        except Exception:
            dist_rel_sem_aprop = []

        try:
            caucoes_retencao = _load_caucoes_retencao()
        except Exception:
            caucoes_retencao = []

        try:
            orc4_refs = set(orc4['refitemorc'].tolist())
            nao_map_list = []
            for _,r in tit[~tit['refitemorc'].isin(orc4_refs)].iterrows():
                nao_map_list.append({
                    'numtit':str(r['numtitcp']),'credor':str(r['nmcredor']),
                    'ref':str(r['refitemorc']),'vlr_ap':f2(r['vlrapropobra']),
                    'vlr_pg':f2(r['vlrliqbaixacp']),'status':str(r['STATUS_PAG']),
                    'dt_baixa':str(r['dtbaixacp'].date()) if pd.notna(r['dtbaixacp']) else '—',
                })
        except Exception:
            nao_map_list = []

        data = {
            'summary':summary_all,'summary_per_cc':summary_per_cc,'cc_list':cc_list_out,
            'fluxo':fluxo_list,'fluxo_per_cc':fluxo_per_cc,'orc_tree':tree,
            'orc_rel_tree':orc_rel_tree,'orc_rel_meta':orc_rel_meta,
            'titulos_multi_apropriacao':titulos_multi_apropriacao,
            'orc_nominal_budget':nominal_budget,
            'nwd_files':list_nwd_files(),
            'orc_rel_divergencias':orc_rel_divergencias,
            'by_cred':by_cred_list,'tit_por_cc':tit_por_cc,
            'aging_geral':aging_geral,'aging_per_cc':aging_per_cc,'aging_titulos':aging_titulos,
            'titulos_pag':titulos_pag,'pc_titulo_map':pc_titulo_map,'taxa_construtora':taxa_construtora,'apuracao_taxa':apuracao_taxa,'taxa_gestao_temazec':taxa_gestao_temazec,'locacoes':locacoes,
            'tit_n1':make_tbl('n1_cod',n1_nm),'tit_n2':make_tbl('n2_cod',n2_nm),'tit_n3':make_tbl('n3_cod',n3_nm),
            'tit_cc':tit_cc_list,'jur_cc':jur_cc_list,'tit_cc_n3':tit_cc3_list,
            'contratos':contratos_list,'itens_ct':itens_ct,'contract_analysis':contract_analysis,
            'vnd_mes':vnd_mes_list,'vnd_status':vnd_status_list,'vnd_table':vnd_table,
            'plan_cronograma':plan_list,'linha_balanco':linha_balanco,'crono_meta':crono_meta,'av_fin_mes':av_fin_list,
            'histograma_equipe':histograma_equipe,
            'previsto_termino_folha':previsto_termino_folha,
            'sobreconsumo_concreto':_load_sobreconsumo_concreto(),
            'activity_costs':activity_costs,'activity_cost_meta':activity_cost_meta,'sinapi_meta':sinapi_meta,
            'suprimentos':suprimentos,
            'controle_contratacoes':load_controle_contratacoes(),
            'rso_state':load_rso_state(),
            'sharepoint_sync_status':load_sharepoint_sync_status(),
            'distorcoes_pagos':paid_since_list,'distorcoes_sem_aprop':sem_aprop,
            'distorcoes_relatorio_sem_aprop':dist_rel_sem_aprop,
            'caucoes_retencao':caucoes_retencao,
            'distorcoes_nao_mapeados':nao_map_list,'pagamentos_antes_cutoff':paid_before_list,
            'protestos':load_protestos_state(),
        }
        data["excel_source"] = excel_path
        data["excel_source_kind"] = "sharepoint_sync" if os.path.abspath(excel_path).lower() != os.path.abspath(LOCAL_EXCEL_PATH).lower() else "local_backup"
        data["excel_updated_at"] = time.strftime("%d/%m/%Y %H:%M:%S", time.localtime(os.path.getmtime(excel_path)))
        _cache = {"data": data, "ts": time.time()}
        save_dashboard_cache(data)
        return data, None
    except Exception as e:
        import traceback; return None, traceback.format_exc()

def load_uau_obras():
    """Carrega obras UAU para o seletor superior sem bloquear o dashboard."""
    fallback_message = "Lista UAU ainda nao configurada. Use uau_obras.json ou uau_sql_config.json em dados."

    def _normalize_uau_rows(rows):
        obras = []
        for row in rows if isinstance(rows, list) else []:
            if not isinstance(row, dict):
                continue
            empresa = str(row.get("empresa") or row.get("cod_empresa") or row.get("empresa_codigo") or "").strip()
            obra = str(row.get("obra") or row.get("cod_obra") or row.get("obra_codigo") or row.get("codigo") or row.get("cod") or row.get("id") or "").strip()
            codigo = str(row.get("codigo") or row.get("cod") or row.get("id") or "").strip()
            if empresa and obra and "|" not in codigo:
                codigo = f"{empresa}|{obra}"
            codigo = codigo or obra
            nome = str(row.get("nome") or row.get("obra_nome") or row.get("nome_obra") or row.get("descricao_obra") or row.get("descricao") or row.get("desc") or row.get("label") or "").strip()
            empresa_nome = str(row.get("empresa_nome") or row.get("nome_empresa") or row.get("descricao_empresa") or row.get("empresa_descricao") or "").strip()
            if codigo or nome:
                obras.append({
                    "codigo": codigo,
                    "nome": nome,
                    "empresa": empresa,
                    "empresa_nome": empresa_nome,
                    "obra": obra,
                    "obra_nome": nome,
                })
        return obras

    def _read_env_file(path):
        data = {}
        if not os.path.exists(path):
            return data
        with open(path, "r", encoding="utf-8-sig") as fh:
            for line in fh:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                key, value = line.split("=", 1)
                data[key.strip()] = value.strip().strip('"').strip("'")
        return data

    def _query_uau_sql(cfg, source_name):
        query = str(cfg.get("query_obras") or cfg.get("query") or "").strip()
        if not query:
            query = """
                SELECT DISTINCT
                    o.Empresa_obr AS empresa,
                    COALESCE(NULLIF(LTRIM(RTRIM(e.NomeFantasia_emp)), ''), NULLIF(LTRIM(RTRIM(e.Desc_emp)), '')) AS empresa_nome,
                    o.Cod_Obr AS codigo,
                    o.Descr_Obr AS nome,
                    ISNULL(o.AtInat_obr, 0) AS inativa
                FROM Obras o
                LEFT JOIN Empresas e
                  ON e.Codigo_emp = o.Empresa_obr
                INNER JOIN ObrUsr ON o.Empresa_obr = Emp_uo AND o.Cod_Obr = Obr_uo AND Usr_uo = 'FCRUZ'
                WHERE o.Empresa_obr IS NOT NULL
                  AND ISNULL(LTRIM(RTRIM(o.Cod_Obr)), '') <> ''
                ORDER BY ISNULL(o.AtInat_obr, 0), o.Empresa_obr, o.Cod_Obr
            """
        conn_str = str(cfg.get("connection_string") or "").strip()
        if not conn_str:
            server = str(cfg.get("server") or cfg.get("UAU_DB_HOST") or "").strip()
            port = str(cfg.get("port") or cfg.get("UAU_DB_PORT") or "").strip()
            database = str(cfg.get("database") or cfg.get("UAU_DB_NAME") or "").strip()
            driver = str(cfg.get("driver") or cfg.get("UAU_ODBC_DRIVER") or "ODBC Driver 17 for SQL Server").strip()
            driver = _resolve_odbc_driver(driver)
            user = str(cfg.get("user") or cfg.get("username") or cfg.get("UAU_DB_USER") or "").strip()
            password = str(cfg.get("password") or cfg.get("UAU_DB_PASSWORD") or "").strip()
            trusted = bool(cfg.get("trusted_connection", False))
            sql_server = f"{server},{port}" if server and port else server
            parts = []
            if driver:
                parts.append(f"DRIVER={{{driver}}}")
            if sql_server:
                parts.append(f"SERVER={sql_server}")
            if database:
                parts.append(f"DATABASE={database}")
            if user:
                parts.append(f"UID={user}")
                parts.append(f"PWD={password}")
            elif trusted:
                parts.append("Trusted_Connection=yes")
            parts.extend(["TrustServerCertificate=yes", "Encrypt=no", "Connection Timeout=8"])
            conn_str = ";".join(parts)
        if not conn_str:
            return {"ok": False, "obras": [], "message": f"{source_name} sem dados de conexao."}
        import pyodbc
        obras = []
        with pyodbc.connect(conn_str, timeout=8) as conn:
            cur = conn.cursor()
            cur.execute("SET NOCOUNT ON;\n" + query)
            cols = [c[0].lower() for c in cur.description]
            for row in cur.fetchall():
                item = dict(zip(cols, row))
                empresa = str(item.get("empresa") or "").strip()
                codigo = str(item.get("codigo") or item.get("cod") or item.get("obra") or item.get("id") or "").strip()
                nome = str(item.get("nome") or item.get("descricao") or item.get("desc") or item.get("label") or "").strip()
                empresa_nome = str(item.get("empresa_nome") or item.get("nome_empresa") or item.get("descricao_empresa") or item.get("empresa_descricao") or "").strip()
                inativa = str(item.get("inativa") or "").strip()
                value = f"{empresa}|{codigo}" if empresa and codigo else (codigo or nome)
                label = f"{empresa}-{codigo} - {nome}" if empresa and codigo and nome else (nome or value)
                if inativa not in ("", "0", "0.0", "False", "false"):
                    label += " (inativa)"
                if value or label:
                    obras.append({
                        "codigo": value,
                        "nome": label,
                        "empresa": empresa,
                        "empresa_nome": empresa_nome,
                        "obra": codigo,
                        "obra_nome": nome,
                    })
        return {"ok": True, "obras": obras, "message": f"{len(obras)} obra(s) UAU carregada(s) de {source_name}."}

    try:
        if os.path.exists(UAU_OBRAS_PATH):
            with open(UAU_OBRAS_PATH, "r", encoding="utf-8") as f:
                raw = json.load(f)
            rows = raw.get("obras", raw) if isinstance(raw, dict) else raw
            obras = _normalize_uau_rows(rows)
            return {"ok": True, "obras": obras, "message": f"{len(obras)} obra(s) UAU carregada(s) do arquivo local."}
    except Exception as exc:
        return {"ok": False, "obras": [], "message": f"Falha ao ler uau_obras.json: {exc}"}

    try:
        if os.path.exists(UAU_SQL_CONFIG_PATH):
            with open(UAU_SQL_CONFIG_PATH, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            query = str(cfg.get("query_obras") or cfg.get("query") or "").strip()
            conn_str = str(cfg.get("connection_string") or "").strip()
            if not query:
                return {"ok": False, "obras": [], "message": "uau_sql_config.json sem query_obras."}
            if not conn_str:
                server = str(cfg.get("server") or "").strip()
                database = str(cfg.get("database") or "").strip()
                driver = str(cfg.get("driver") or "ODBC Driver 17 for SQL Server").strip()
                driver = _resolve_odbc_driver(driver)
                trusted = bool(cfg.get("trusted_connection", True))
                if server and database and trusted:
                    conn_str = f"DRIVER={{{driver}}};SERVER={server};DATABASE={database};Trusted_Connection=yes;"
            if not conn_str:
                return {"ok": False, "obras": [], "message": "uau_sql_config.json sem dados de conexao."}
            cfg["connection_string"] = conn_str
            return _query_uau_sql(cfg, "uau_sql_config.json")
        if os.path.exists(UAU_EXTERNAL_ENV_PATH):
            env_cfg = _read_env_file(UAU_EXTERNAL_ENV_PATH)
            if env_cfg:
                return _query_uau_sql(env_cfg, "dashboard_uau")
    except ImportError:
        return {"ok": False, "obras": [], "message": "pyodbc nao instalado. Instale o driver para consultar o SQL do UAU."}
    except Exception as exc:
        return {"ok": False, "obras": [], "message": f"Falha ao consultar SQL do UAU: {exc}"}

    return {"ok": True, "obras": [], "message": fallback_message}


def _uau_read_env_file(path=UAU_EXTERNAL_ENV_PATH):
    data = {}
    if not os.path.exists(path):
        return data
    with open(path, "r", encoding="utf-8-sig") as fh:
        for line in fh:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            data[key.strip()] = value.strip().strip('"').strip("'")
    return data


def _uau_conn_str(cfg):
    server = str(cfg.get("UAU_DB_HOST") or cfg.get("server") or "").strip()
    port = str(cfg.get("UAU_DB_PORT") or cfg.get("port") or "").strip()
    database = str(cfg.get("UAU_DB_NAME") or cfg.get("database") or "").strip()
    driver = str(cfg.get("UAU_ODBC_DRIVER") or cfg.get("driver") or "ODBC Driver 17 for SQL Server").strip()
    driver = _resolve_odbc_driver(driver)
    user = str(cfg.get("UAU_DB_USER") or cfg.get("user") or cfg.get("username") or "").strip()
    password = str(cfg.get("UAU_DB_PASSWORD") or cfg.get("password") or "").strip()
    server_part = f"{server},{port}" if server and port else server
    parts = [
        f"DRIVER={{{driver}}}",
        f"SERVER={server_part}",
        f"DATABASE={database}",
        f"UID={user}",
        f"PWD={password}",
        "TrustServerCertificate=yes",
        "Encrypt=no",
        "Connection Timeout=12",
    ]
    return ";".join(p for p in parts if not p.endswith("=") and "{}" not in p)


def _uau_rows_to_dicts(cursor):
    cols = [c[0] for c in cursor.description]
    return [dict(zip(cols, row)) for row in cursor.fetchall()]


def _uau_json_safe(value):
    if hasattr(value, "isoformat"):
        return value.isoformat()
    return value


def _uau_nat(value):
    return [int(p) if p.isdigit() else p.lower() for p in re.split(r"(\d+)", str(value or ""))]


def _uau_split_eap(code):
    return [p for p in str(code or "").strip().split(".") if p]


def _uau_child(parent, code, name):
    for child in parent["children"]:
        if child["cod"] == code:
            if name and (not child.get("nm") or child["nm"] == code):
                child["nm"] = name
            return child
    child = {
        "cod": code,
        "nm": name or code,
        "orcado": 0.0,
        "realizado": 0.0,
        "a_pagar": 0.0,
        "saldo": 0.0,
        "children": [],
        "titulos": [],
    }
    parent["children"].append(child)
    return child


UAU_CJ0305_PLANILHA_PATH = os.path.join(UAU_DATA_DIR, "Orçamento CJ 03 CJ 05.xlsx")
_uau_cj0305_cache = {"mtime": None, "mapa": {}}


def _uau_orcado_cj0305_por_cc():
    """A EAP do UAU (PlanTotal) não guarda valor orçado, só item/quantidade
    — pros 3 centros de custo do CJ03/CJ05 (116-11602 Conjunto 3, 116-11603
    Conjunto 5, 116-11611 Indiretos), o orçado real vem da planilha
    'Orçamento CJ 03 CJ 05.xlsx' (mesma usada no dashboard_uau): direto
    (nível 2, coluna BC) 50/50 entre os dois conjuntos; indireto (nível 2,
    coluna M — BC não tem valor mapeado pra indireto) 100% no 116-11611.
    Devolve {"116-11602": {"02.01": valor, ...}, ...}. Cacheado pelo mtime
    do arquivo."""
    if not os.path.exists(UAU_CJ0305_PLANILHA_PATH):
        return {}
    mtime = os.path.getmtime(UAU_CJ0305_PLANILHA_PATH)
    if _uau_cj0305_cache["mtime"] == mtime:
        return _uau_cj0305_cache["mapa"]
    mapa = {"116-11602": {}, "116-11603": {}, "116-11611": {}}
    try:
        import openpyxl
        wb = openpyxl.load_workbook(UAU_CJ0305_PLANILHA_PATH, read_only=True, data_only=True)
        ws = wb[wb.sheetnames[0]]
        for i in range(6, ws.max_row + 1):
            nivel = ws.cell(row=i, column=5).value  # coluna E
            if nivel != 2:
                continue
            item = ws.cell(row=i, column=6).value  # coluna F
            if not item:
                continue
            item = str(item).strip()
            n1 = item.split(".")[0].zfill(2)
            total_m = ws.cell(row=i, column=13).value  # coluna M
            custo_bc = ws.cell(row=i, column=55).value  # coluna BC
            if n1 == "02" and custo_bc not in (None, ""):
                metade = f2(float(custo_bc) / 2)
                mapa["116-11602"][item] = metade
                mapa["116-11603"][item] = metade
            elif n1 == "01" and total_m not in (None, ""):
                mapa["116-11611"][item] = f2(float(total_m))
        wb.close()
    except Exception:
        return {}
    _uau_cj0305_cache.update({"mtime": mtime, "mapa": mapa})
    return mapa


def _norm_far_code(v):
    """Normaliza código de item FAR CJ03/CJ05 pra comparar entre as duas
    planilhas — uma usa '01'/'02' (zero à esquerda), a outra usa '1'/'2'
    (sem zero). Ex.: '02.01' e '2.01' viram os dois '02.01'."""
    parts = [p for p in re.split(r"[.\s]+", str(v or "").strip()) if p]
    try:
        return ".".join(str(int(p)).zfill(2) for p in parts)
    except ValueError:
        return str(v or "").strip()


def _uau_orcado_cj0305_bruto():
    """Como _uau_orcado_cj0305_por_cc, mas devolve o valor CHEIO por item
    (sem dividir 50/50 entre os conjuntos) — usado pra comparar item a item
    com a segunda planilha de orçamento (ANÁLISE DE CUSTO). Chave já
    normalizada via _norm_far_code."""
    if not os.path.exists(UAU_CJ0305_PLANILHA_PATH):
        return {}
    out = {}
    try:
        import openpyxl
        wb = openpyxl.load_workbook(UAU_CJ0305_PLANILHA_PATH, read_only=True, data_only=True)
        ws = wb[wb.sheetnames[0]]
        for i in range(6, ws.max_row + 1):
            nivel = ws.cell(row=i, column=5).value
            if nivel != 2:
                continue
            item = ws.cell(row=i, column=6).value
            if not item:
                continue
            item = str(item).strip()
            n1 = item.split(".")[0].zfill(2)
            descricao = ws.cell(row=i, column=54).value or ws.cell(row=i, column=10).value
            total_m = ws.cell(row=i, column=13).value
            custo_bc = ws.cell(row=i, column=55).value
            valor = custo_bc if (n1 == "02" and custo_bc not in (None, "")) else total_m
            if valor in (None, ""):
                continue
            out[_norm_far_code(item)] = {"cod": item, "nm": str(descricao or item).strip(), "valor": f2(float(valor))}
        wb.close()
    except Exception:
        return {}
    # sintetiza o nível 1 (essa planilha só tem valor no nível 2) somando os
    # filhos — sem isso "01"/"02" ficariam sem total pra comparar.
    n1_nomes = {"01": "CUSTOS INDIRETOS", "02": "CUSTO DIRETO DE OBRA"}
    n1_totais = {}
    for chave, item in out.items():
        n1 = chave.split(".")[0]
        n1_totais[n1] = f2(n1_totais.get(n1, 0.0) + item["valor"])
    for n1, total in n1_totais.items():
        out[n1] = {"cod": n1, "nm": n1_nomes.get(n1, n1), "valor": total}
    return out


def _uau_analise_planilha_path():
    """Acha a planilha 'ANÁLISE_CUSTO_FAR_CJ03_CJ05...' mais recente em
    dados/uau — o nome muda com a data (ex. '_2026.09'), então busca por
    padrão em vez de nome fixo."""
    import glob
    candidatos = [
        p for p in glob.glob(os.path.join(UAU_DATA_DIR, "ANALISE_CUSTO_FAR_CJ03_CJ05*.xlsx"))
        if not os.path.basename(p).startswith("~$")
    ]
    if not candidatos:
        return None
    return max(candidatos, key=os.path.getmtime)


def _uau_orcado_analise_custo():
    """Lê a aba 'ANÁLISE DE CUSTO' da planilha de análise (a segunda fonte
    de orçamento do FAR CJ03/CJ05) — coluna B 'Cód. Estruturado', F
    'Descrição', W 'CUSTO AO TÉRMINO2' (comprometido + saldo de contratos +
    a contratar — o custo projetado real, confirmado com o usuário: bate
    R$57.634.698,28 no total, W7+W283 na planilha). Só nível 1 e 2 (mesma
    granularidade da outra planilha, pra comparar item a item). Chave
    normalizada via _norm_far_code."""
    path = _uau_analise_planilha_path()
    if not path:
        return {}, None
    out = {}
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        nome_aba = next((n for n in wb.sheetnames if _ascii_upper(n).replace(" ", "") == "ANALISEDECUSTO"), None)
        if not nome_aba:
            return {}, os.path.basename(path)
        ws = wb[nome_aba]
        for row in ws.iter_rows(min_row=7, values_only=True):
            cod = row[1] if len(row) > 1 else None  # B: Cód. Estruturado
            if cod in (None, ""):
                continue
            cod_str = str(cod).strip()
            if cod_str.count(".") > 1:
                continue  # só nível 1 ("1") e nível 2 ("1.01") — mesma granularidade da outra planilha
            descricao = row[5] if len(row) > 5 else None  # F: Descrição
            custo_termino = row[22] if len(row) > 22 else None  # W: CUSTO AO TÉRMINO2
            if custo_termino in (None, ""):
                continue
            try:
                valor = f2(float(custo_termino))
            except (TypeError, ValueError):
                continue
            out[_norm_far_code(cod_str)] = {"cod": cod_str, "nm": str(descricao or cod_str).strip(), "valor": valor}
        wb.close()
    except Exception:
        return {}, os.path.basename(path)
    return out, os.path.basename(path)


def _uau_comparar_orcamentos():
    """Compara, item a item (nível 1 e 2), o orçado da planilha original
    ('Orçamento CJ 03 CJ 05.xlsx') com o da planilha de análise de custo
    mais recente ('ANÁLISE_CUSTO_FAR_CJ03_CJ05...') — mostra as duas fontes
    lado a lado e a diferença, sem decidir qual está certa (isso fica pro
    usuário avaliar)."""
    base = _uau_orcado_cj0305_bruto()
    analise, analise_arquivo = _uau_orcado_analise_custo()
    codigos = sorted(set(base.keys()) | set(analise.keys()))
    linhas = []
    for cod in codigos:
        item_base = base.get(cod)
        item_analise = analise.get(cod)
        v1 = item_base.get("valor", 0.0) if item_base else 0.0
        v2 = item_analise.get("valor", 0.0) if item_analise else 0.0
        nome = (item_base or {}).get("nm") or (item_analise or {}).get("nm") or cod
        diferenca = f2(v2 - v1)
        diferenca_pct = f2((diferenca / v1 * 100) if v1 else (100.0 if v2 else 0.0))
        linhas.append({
            "cod": cod,
            "nm": nome,
            "orcamento_original": v1,
            "analise_custo": v2,
            "diferenca": diferenca,
            "diferenca_pct": diferenca_pct,
            "nivel": cod.count(".") + 1,
        })
    # soma pelo nível 2 (não nível 1): a planilha original só tem valor no
    # nível 2 (o nível 1 dela fica sempre 0 aqui), então somar pelo nível 1
    # zeraria o total original — nível 2 tem dado completo nas duas fontes.
    total_original = f2(sum(l["orcamento_original"] for l in linhas if l["nivel"] == 2))
    total_analise = f2(sum(l["analise_custo"] for l in linhas if l["nivel"] == 2))
    return {
        "ok": True,
        "linhas": linhas,
        "resumo": {
            "total_orcamento_original": total_original,
            "total_analise_custo": total_analise,
            "diferenca_total": f2(total_analise - total_original),
        },
        "fonte_original": os.path.basename(UAU_CJ0305_PLANILHA_PATH),
        "fonte_analise": analise_arquivo or "não encontrada",
    }


def load_uau_dashboard_data(empresa="", obra="", dt_ini="", dt_fim=""):
    cfg = _uau_read_env_file()
    if not cfg:
        return {"ok": False, "error": "Configuração do UAU não encontrada.", "tree": [], "financeiro": []}
    import pyodbc
    empresa = str(empresa or "").strip()
    # obra pode vir com varios centros de custo selecionados, separados por
    # virgula (seletor virou multi-select) — cada item ainda pode estar no
    # formato "empresa|obra" (valor original das options).
    obras = []
    empresas_from_obra = set()
    for item in str(obra or "").split(","):
        item = item.strip()
        if not item:
            continue
        if "|" in item:
            emp_part, obra_part = item.split("|", 1)
            emp_part, obra_part = emp_part.strip(), obra_part.strip()
            if emp_part:
                empresas_from_obra.add(emp_part)
            if obra_part:
                obras.append(obra_part)
        else:
            obras.append(item)
    if not empresa and len(empresas_from_obra) == 1:
        empresa = next(iter(empresas_from_obra))
    dt_ini = re.sub(r"\D", "", str(dt_ini or "")) or "20000101"
    dt_fim = re.sub(r"\D", "", str(dt_fim or "")) or "21000101"
    where_plan = ["ISNULL(p.Item_plt, '') <> ''", "ISNULL(p.AtInat_plt, 0) = 0"]
    where_tit = ["d.DtPgto_Des >= CONVERT(datetime, ?, 112)", "d.DtPgto_Des < DATEADD(day, 1, CONVERT(datetime, ?, 112))"]
    plan_params = []
    tit_params = [dt_ini, dt_fim]
    if empresa:
        where_plan.append("p.Empresa_plt = ?")
        where_tit.append("d.Empresa_Des = ?")
        plan_params.append(empresa)
        tit_params.append(empresa)
    if obras:
        placeholders = ",".join(["?"] * len(obras))
        where_plan.append(f"p.Obra_plt IN ({placeholders})")
        where_tit.append(f"d.Obra_Des IN ({placeholders})")
        plan_params.extend(obras)
        tit_params.extend(obras)

    plan_sql = f"""
        SELECT
            p.Empresa_plt AS empresa,
            p.Obra_plt AS obra,
            o.Descr_Obr AS obra_nome,
            p.Item_plt AS eap,
            p.Serv_plt AS servico,
            p.Descr_plt AS descricao,
            p.Qtde_plt AS quantidade,
            p.Unid_plt AS unidade
        FROM PlanTotal p
        LEFT JOIN Obras o
            ON o.Empresa_obr = p.Empresa_plt
           AND o.Cod_Obr = p.Obra_plt
        INNER JOIN ObrUsr ON p.Empresa_plt = Emp_uo AND p.Obra_plt = Obr_uo AND Usr_uo = 'FCRUZ'
        WHERE {" AND ".join(where_plan)}
        ORDER BY p.Obra_plt, p.Item_plt, p.Serv_plt
    """
    tit_sql = f"""
        SELECT
            d.Empresa_Des AS empresa,
            d.Obra_Des AS obra,
            d.ItemPl_Des AS eap,
            d.DescCompPl_Des AS servico_nome,
            d.DescInsPl_Des AS insumo_nome,
            d.DescItemProc_Des AS item_processo,
            d.NumOc_Des AS pedido_oc,
            d.ItemProc_Des AS item_pedido,
            d.UnidItemProc_Des AS unidade_item,
            d.QtdeItem_Des AS qtd_item,
            d.ValorUnitItem_Des AS valor_unit_item,
            d.VlDocFisc_Des AS valor_doc_fiscal,
            CAST(d.NumProc_Des AS VARCHAR) + '/' + CAST(d.NumParc_Des AS VARCHAR) AS numtit,
            CAST(d.DocFiscal_Des AS VARCHAR) AS doc_fiscal,
            d.DtPgto_Des AS data_movimento,
            d.TotalLiq_Des AS valor,
            d.StatusParc_Des AS status_parcela,
            d.NominalProc_Des AS credor,
            d.Banco_Des AS banco,
            d.ContaCorr_Des AS conta,
            d.ContratoServ_Des AS contrato_serv,
            d.CodMedicao_Des AS cod_medicao,
            d.ContratoPl_Des AS contrato_pl
        FROM VwDesembolso d
        INNER JOIN ObrUsr ON d.Empresa_Des = Emp_uo AND d.Obra_Des = Obr_uo AND Usr_uo = 'FCRUZ'
        WHERE {" AND ".join(where_tit)}
        ORDER BY d.DtPgto_Des DESC, d.Empresa_Des, d.Obra_Des, d.NumProc_Des, d.NumParc_Des
    """
    with pyodbc.connect(_uau_conn_str(cfg), timeout=12) as conn:
        cur = conn.cursor()
        cur.execute("SET DATEFORMAT mdy; SET NOCOUNT ON;\n" + plan_sql, *plan_params)
        plan_rows = _uau_rows_to_dicts(cur)
        cur.execute("SET DATEFORMAT mdy; SET NOCOUNT ON;\n" + tit_sql, *tit_params)
        tit_rows = _uau_rows_to_dicts(cur)

        sup_errors = []

        def _uau_period_where(alias, col, params):
            params.extend([dt_ini, dt_fim])
            return [f"{alias}.{col} >= CONVERT(datetime, ?, 112)", f"{alias}.{col} < DATEADD(day, 1, CONVERT(datetime, ?, 112))"]

        sol_rows = []
        try:
            sol_params = []
            where_sol = _uau_period_where("s", "DataCad_sap", sol_params)
            if empresa:
                where_sol.append("s.Empresa_sap = ?")
                sol_params.append(empresa)
            if obras:
                placeholders = ",".join(["?"] * len(obras))
                where_sol.append(f"s.Obra_sap IN ({placeholders})")
                sol_params.extend(obras)
            sol_sql = f"""
                SELECT
                    s.Empresa_sap AS empresa, s.Obra_sap AS obra, s.Num_sap AS solicitacao,
                    s.StatusAprovacao_sap AS status_code, s.UsrCad_sap AS solicitante,
                    s.DataCad_sap AS data, s.DataAprovacao_sap AS aprovado_em,
                    s.NumCot_sap AS cotacao, s.NumProc_sap AS processo,
                    i.Ins_isap AS insumo, i.Item_isap AS eap, i.Comp_isap AS complemento,
                    i.TotalPlanejado_isap AS total_planejado, i.SaldoPlanejado_isap AS saldo_planejado,
                    i.ValorSolicitado_isap AS valor_solicitado, i.ValorAprovado_isap AS valor_aprovado,
                    i.QdadeSolicitada_isap AS qtd_solicitada, i.QdadeAprovada_isap AS qtd_aprovada,
                    g.Descr_ins AS insumo_nome, g.Unid_ins AS unidade
                FROM SolicitacaoAprovacaoPlanejamento s
                LEFT JOIN ItensSolicitacaoAprovacaoPlanejamento i
                    ON i.Empresa_isap = s.Empresa_sap AND i.Obra_isap = s.Obra_sap AND i.NumSap_isap = s.Num_sap
                LEFT JOIN InsumosGeral g ON g.Cod_ins = i.Ins_isap
                INNER JOIN ObrUsr ON s.Empresa_sap = Emp_uo AND s.Obra_sap = Obr_uo AND Usr_uo = 'FCRUZ'
                WHERE {" AND ".join(where_sol)}
                ORDER BY s.DataCad_sap DESC, s.Num_sap DESC
            """
            cur.execute("SET DATEFORMAT mdy; SET NOCOUNT ON;\n" + sol_sql, *sol_params)
            sol_rows = _uau_rows_to_dicts(cur)
        except Exception as exc:
            sup_errors.append(f"solicitacoes: {exc}")

        cot_rows = []
        try:
            cot_params = []
            where_cot = _uau_period_where("c", "Data_cot", cot_params)
            if empresa:
                where_cot.append("c.Empresa_cot = ?")
                cot_params.append(empresa)
            if obras:
                ph1 = ",".join(["?"] * len(obras))
                ph2 = ",".join(["?"] * len(obras))
                where_cot.append(f"(EXISTS (SELECT 1 FROM OrdemCompra oc2 WHERE oc2.Empresa_Ocp = c.Empresa_cot AND oc2.NumCot_Ocp = c.Num_cot AND oc2.Obra_Ocp IN ({ph1})) OR EXISTS (SELECT 1 FROM SolicitacaoAprovacaoPlanejamento s2 WHERE s2.Empresa_sap = c.Empresa_cot AND s2.NumCot_sap = c.Num_cot AND s2.Obra_sap IN ({ph2})))")
                cot_params.extend(obras)
                cot_params.extend(obras)
            cot_sql = f"""
                SELECT
                    c.Empresa_cot AS empresa, c.Num_cot AS cotacao, c.Data_cot AS data,
                    c.Cotador_cot AS comprador, c.Status_cot AS status_code, c.Descricao_cot AS descricao,
                    i.CodIns_item AS insumo, i.CodForn_item AS fornecedor_cod, i.Preco_item AS preco,
                    i.Status_item AS item_status, i.QtdeSolicit_item AS qtd_solicitada,
                    i.QtdeCot_item AS qtd_cotada, i.Unid_item AS unidade,
                    p.nome_pes AS credor, g.Descr_ins AS insumo_nome, oc.Obra_Ocp AS obra
                FROM Cotacao c
                LEFT JOIN ItensCotacao i ON i.Empresa_item = c.Empresa_cot AND i.NumCot_item = c.Num_cot
                LEFT JOIN Pessoas p ON p.cod_pes = i.CodForn_item
                LEFT JOIN InsumosGeral g ON g.Cod_ins = i.CodIns_item
                OUTER APPLY (
                    SELECT TOP 1 oc1.Obra_Ocp
                    FROM OrdemCompra oc1
                    WHERE oc1.Empresa_Ocp = c.Empresa_cot AND oc1.NumCot_Ocp = c.Num_cot
                    ORDER BY oc1.DataGer_Ocp DESC
                ) oc
                WHERE {" AND ".join(where_cot)}
                ORDER BY c.Data_cot DESC, c.Num_cot DESC
            """
            cur.execute("SET DATEFORMAT mdy; SET NOCOUNT ON;\n" + cot_sql, *cot_params)
            cot_rows = _uau_rows_to_dicts(cur)
        except Exception as exc:
            sup_errors.append(f"cotacoes: {exc}")

        oc_rows = []
        try:
            oc_params = []
            where_oc = _uau_period_where("o", "DataGer_Ocp", oc_params)
            if empresa:
                where_oc.append("o.Empresa_Ocp = ?")
                oc_params.append(empresa)
            if obras:
                placeholders = ",".join(["?"] * len(obras))
                where_oc.append(f"o.Obra_Ocp IN ({placeholders})")
                oc_params.extend(obras)
            oc_sql = f"""
                SELECT
                    o.Empresa_Ocp AS empresa, o.Obra_Ocp AS obra, o.NumeroOC_Ocp AS pedido,
                    o.CodForn_Ocp AS fornecedor_cod, o.DataGer_Ocp AS data, o.Usuario_Ocp AS comprador,
                    o.NumCot_Ocp AS cotacao, o.StatusConfForn_Ocp AS status_code,
                    o.DataConfReceb_Ocp AS confirmado_em,
                    i.Item_Ioc AS item, i.CodInsumo_Ioc AS insumo, i.Unidade_Ioc AS unidade,
                    i.Qtde_Ioc AS qtd, i.Preco_Ioc AS preco, i.ValorIpi_Ioc AS valor_ipi,
                    i.Cotacao_Ioc AS cotacao_item, i.ObsInsumo_Ioc AS detalhe,
                    p.nome_pes AS credor, g.Descr_ins AS insumo_nome
                FROM OrdemCompra o
                LEFT JOIN ItensOrdemCompra i
                    ON i.Empresa_Ioc = o.Empresa_Ocp AND i.Obra_Ioc = o.Obra_Ocp AND i.NumeroOC_Ioc = o.NumeroOC_Ocp
                LEFT JOIN Pessoas p ON p.cod_pes = o.CodForn_Ocp
                LEFT JOIN InsumosGeral g ON g.Cod_ins = i.CodInsumo_Ioc
                INNER JOIN ObrUsr ON o.Empresa_Ocp = Emp_uo AND o.Obra_Ocp = Obr_uo AND Usr_uo = 'FCRUZ'
                WHERE {" AND ".join(where_oc)}
                ORDER BY o.DataGer_Ocp DESC, o.NumeroOC_Ocp DESC
            """
            cur.execute("SET DATEFORMAT mdy; SET NOCOUNT ON;\n" + oc_sql, *oc_params)
            oc_rows = _uau_rows_to_dicts(cur)
        except Exception as exc:
            sup_errors.append(f"pedidos: {exc}")

        ct_rows = []
        ct_item_rows = []
        ct_pl_rows = []
        ct_med_rows = []
        ct_prev_rows = []
        try:
            ct_params = []
            where_ct = ["1=1"]
            if empresa:
                where_ct.append("c.Empresa_cont = ?")
                ct_params.append(empresa)
            if obras:
                placeholders = ",".join(["?"] * len(obras))
                where_ct.append(f"c.Obra_cont IN ({placeholders})")
                ct_params.extend(obras)
            ct_sql = f"""
                SELECT
                    c.Empresa_cont AS empresa, c.Obra_cont AS obra, c.Cod_cont AS contrato,
                    c.CodPes_cont AS fornecedor_cod, c.Objeto_cont AS objeto,
                    c.DtCriacao_cont AS data_contrato, c.DtInicio_cont AS data_inicio, c.DtFim_cont AS data_fim,
                    c.Situacao_cont AS situacao_code, c.Status_cont AS status_code,
                    p.nome_pes AS credor, p.NomeFant_Pes AS nome_fantasia,
                    o.Descr_Obr AS obra_nome
                FROM Contratos c
                LEFT JOIN Pessoas p ON p.cod_pes = c.CodPes_cont
                LEFT JOIN Obras o ON o.Empresa_obr = c.Empresa_cont AND o.Cod_Obr = c.Obra_cont
                INNER JOIN ObrUsr ON c.Empresa_cont = Emp_uo AND c.Obra_cont = Obr_uo AND Usr_uo = 'FCRUZ'
                WHERE {" AND ".join(where_ct)}
                ORDER BY c.Empresa_cont, c.Obra_cont, c.Cod_cont
            """
            cur.execute("SET DATEFORMAT mdy; SET NOCOUNT ON;\n" + ct_sql, *ct_params)
            ct_rows = _uau_rows_to_dicts(cur)

            cti_params = []
            where_cti = ["1=1"]
            if empresa:
                where_cti.append("c.Empresa_cont = ?")
                cti_params.append(empresa)
            if obras:
                placeholders = ",".join(["?"] * len(obras))
                where_cti.append(f"c.Obra_cont IN ({placeholders})")
                cti_params.extend(obras)
            ct_item_sql = f"""
                SELECT
                    i.Empresa_itens AS empresa, c.Obra_cont AS obra, i.Contrato_itens AS contrato,
                    i.Item_itens AS item, i.Serv_itens AS servico, i.Descr_itens AS descricao,
                    i.Unid_itens AS unidade, i.Qtde_itens AS qtd, i.Preco_itens AS preco,
                    i.SaldoMedicao_Itens AS saldo_medicao
                FROM ItensContrato i
                INNER JOIN Contratos c ON c.Empresa_cont = i.Empresa_itens AND c.Cod_cont = i.Contrato_itens
                INNER JOIN ObrUsr ON c.Empresa_cont = Emp_uo AND c.Obra_cont = Obr_uo AND Usr_uo = 'FCRUZ'
                WHERE {" AND ".join(where_cti)}
                ORDER BY i.Empresa_itens, c.Obra_cont, i.Contrato_itens, i.Item_itens
            """
            cur.execute("SET DATEFORMAT mdy; SET NOCOUNT ON;\n" + ct_item_sql, *cti_params)
            ct_item_rows = _uau_rows_to_dicts(cur)

            ctpl_params = []
            where_ctpl = ["1=1"]
            if empresa:
                where_ctpl.append("cp.Empresa_ContPl = ?")
                ctpl_params.append(empresa)
            if obras:
                placeholders = ",".join(["?"] * len(obras))
                where_ctpl.append(f"cp.Obra_ContPl IN ({placeholders})")
                ctpl_params.extend(obras)
            ct_pl_sql = f"""
                SELECT
                    cp.Empresa_ContPl AS empresa, cp.Obra_ContPl AS obra, cp.CodCont_ContPl AS contrato,
                    cp.Ins_ContPl AS insumo, cp.Serv_ContPl AS servico, cp.Item_ContPl AS eap,
                    cp.Produto_ContPl AS produto, cp.Contrato_ContPl AS contrato_pl,
                    cp.ValorVinc_ContPl AS valor_vinc, cp.ValorAprov_ContPl AS valor_aprov,
                    p.Descr_plt AS descricao_eap
                FROM ContratoPl cp
                LEFT JOIN PlanTotal p
                    ON p.Empresa_plt = cp.Empresa_ContPl
                   AND p.Obra_plt = cp.Obra_ContPl
                   AND p.Item_plt = cp.Item_ContPl
                   AND p.Serv_plt = cp.Serv_ContPl
                INNER JOIN ObrUsr ON cp.Empresa_ContPl = Emp_uo AND cp.Obra_ContPl = Obr_uo AND Usr_uo = 'FCRUZ'
                WHERE {" AND ".join(where_ctpl)}
                ORDER BY cp.Empresa_ContPl, cp.Obra_ContPl, cp.CodCont_ContPl, cp.Item_ContPl
            """
            cur.execute("SET DATEFORMAT mdy; SET NOCOUNT ON;\n" + ct_pl_sql, *ctpl_params)
            ct_pl_rows = _uau_rows_to_dicts(cur)

            ctm_params = []
            where_ctm = ["1=1"]
            if empresa:
                where_ctm.append("m.Empresa_med = ?")
                ctm_params.append(empresa)
            if obras:
                placeholders = ",".join(["?"] * len(obras))
                where_ctm.append(f"c.Obra_cont IN ({placeholders})")
                ctm_params.extend(obras)
            ct_med_sql = f"""
                SELECT
                    m.Empresa_med AS empresa, c.Obra_cont AS obra, m.Contrato_med AS contrato,
                    m.Cod_med AS medicao, m.DtCriacao_med AS data_medicao, m.Valor_med AS valor_medicao,
                    m.Status_med AS status_code, m.Quem_med AS usuario, m.DataAprov_med AS data_aprovacao,
                    i.ItensCont_Item AS item_contrato, i.Ins_Item AS insumo, i.Qtde_Item AS qtd,
                    i.PrecoUnit_Item AS preco_unit, i.Unid_Item AS unidade, i.Total_Item AS valor_item
                FROM Medicoes m
                INNER JOIN Contratos c ON c.Empresa_cont = m.Empresa_med AND c.Cod_cont = m.Contrato_med
                LEFT JOIN ItensMedicao i
                    ON i.Empresa_Item = m.Empresa_med
                   AND i.Contrato_Item = m.Contrato_med
                   AND i.CodMed_Item = m.Cod_med
                INNER JOIN ObrUsr ON c.Empresa_cont = Emp_uo AND c.Obra_cont = Obr_uo AND Usr_uo = 'FCRUZ'
                WHERE {" AND ".join(where_ctm)}
                ORDER BY m.Empresa_med, c.Obra_cont, m.Contrato_med, m.Cod_med, i.ItensCont_Item
            """
            cur.execute("SET DATEFORMAT mdy; SET NOCOUNT ON;\n" + ct_med_sql, *ctm_params)
            ct_med_rows = _uau_rows_to_dicts(cur)

            ctp_params = []
            where_ctp = ["1=1"]
            if empresa:
                where_ctp.append("pp.Empresa_Proc = ?")
                ctp_params.append(empresa)
            if obras:
                placeholders = ",".join(["?"] * len(obras))
                where_ctp.append(f"c.Obra_cont IN ({placeholders})")
                ctp_params.extend(obras)
            ct_prev_sql = f"""
                SELECT
                    pp.Empresa_Proc AS empresa, c.Obra_cont AS obra, pp.Contrato_Proc AS contrato,
                    pp.CodMed_Proc AS medicao, pp.Num_Proc AS titulo, pp.NumParc_Proc AS parcela,
                    pp.DtVencParc_Proc AS vencimento, pp.DtPagParc_Proc AS pagamento,
                    pp.ValorParc_Proc AS valor, pp.Status_proc AS status_code
                FROM ParcProcMedicao pp
                INNER JOIN Contratos c ON c.Empresa_cont = pp.Empresa_Proc AND c.Cod_cont = pp.Contrato_Proc
                INNER JOIN ObrUsr ON c.Empresa_cont = Emp_uo AND c.Obra_cont = Obr_uo AND Usr_uo = 'FCRUZ'
                WHERE {" AND ".join(where_ctp)}
                ORDER BY pp.Empresa_Proc, c.Obra_cont, pp.Contrato_Proc, pp.CodMed_Proc, pp.Num_Proc, pp.NumParc_Proc
            """
            cur.execute("SET DATEFORMAT mdy; SET NOCOUNT ON;\n" + ct_prev_sql, *ctp_params)
            ct_prev_rows = _uau_rows_to_dicts(cur)
        except Exception as exc:
            sup_errors.append(f"contratos: {exc}")

    def _uau_clean_ref(v):
        if v is None:
            return ""
        s = str(v).strip()
        if not s:
            return ""
        try:
            n = float(s.replace(",", "."))
            if n.is_integer():
                return str(int(n))
        except Exception:
            pass
        return s

    titles_by_key = {}
    titles_by_pedido = {}
    titles_by_pedido_item = {}
    titles_by_contract = {}
    titles_by_contract_med = {}
    totals_by_key = {}
    financeiro = []
    for r in tit_rows:
        status_code = int(f2(r.get("status_parcela")))
        valor = f2(r.get("valor"))
        pago = valor if status_code == 2 else 0.0
        apagar = 0.0 if status_code == 2 else valor
        eap = str(r.get("eap") or "").strip()
        key = (str(r.get("empresa") or "").strip(), str(r.get("obra") or "").strip(), eap)
        pedido_oc = _uau_clean_ref(r.get("pedido_oc"))
        item_pedido = str(r.get("item_pedido") or "").strip()
        contrato_serv = _uau_clean_ref(r.get("contrato_serv"))
        cod_medicao = _uau_clean_ref(r.get("cod_medicao"))
        totals = totals_by_key.setdefault(key, {"realizado": 0.0, "a_pagar": 0.0})
        totals["realizado"] = f2(totals["realizado"] + pago)
        totals["a_pagar"] = f2(totals["a_pagar"] + apagar)
        title = {
            "credor": str(r.get("credor") or "").strip(),
            "numtit": str(r.get("numtit") or "").strip(),
            "doc": str(r.get("doc_fiscal") or "").strip(),
            "documento": str(r.get("doc_fiscal") or "").strip(),
            "dt_baixa": _uau_json_safe(r.get("data_movimento")) or "—",
            "vlr_ap": valor,
            "vlr_pg": pago,
            "vlr_a_pagar": apagar,
            "status": "PAGO" if pago else "A PAGAR",
            "observ": str(r.get("item_processo") or r.get("insumo_nome") or r.get("servico_nome") or "").strip(),
            "pedido_oc": pedido_oc,
            "item_pedido": item_pedido,
            "contrato_serv": contrato_serv,
            "cod_medicao": cod_medicao,
            "unidade": str(r.get("unidade_item") or "").strip(),
            "qtd": f2(r.get("qtd_item")),
        }
        titles_by_key.setdefault(key, []).append(title)
        if pedido_oc:
            pk = (key[0], key[1], pedido_oc)
            titles_by_pedido.setdefault(pk, []).append(title)
            if item_pedido:
                titles_by_pedido_item.setdefault((key[0], key[1], pedido_oc, item_pedido), []).append(title)
        if contrato_serv:
            ck = (key[0], key[1], contrato_serv)
            titles_by_contract.setdefault(ck, []).append(title)
            if cod_medicao:
                titles_by_contract_med.setdefault((key[0], key[1], contrato_serv, cod_medicao), []).append(title)
        financeiro.append({
            "empresa": key[0],
            "obra": key[1],
            "eap": eap,
            "pedido": pedido_oc,
            "item_pedido": item_pedido,
            "titulo": title["numtit"],
            "documento": title["documento"],
            "credor": title["credor"],
            "data": title["dt_baixa"],
            "valor": valor,
            "pago": pago,
            "a_pagar": apagar,
            "status": title["status"],
            "apropriacao": "Apropriado" if eap else "Sem apropriação",
        })

    roots_by_obra = {}
    node_by_key = {}
    eap_leaf_keys = set()
    cj0305_map = _uau_orcado_cj0305_por_cc()
    for p in plan_rows:
        emp = str(p.get("empresa") or "").strip()
        obr = str(p.get("obra") or "").strip()
        eap = str(p.get("eap") or "").strip()
        if not eap:
            continue
        root = roots_by_obra.setdefault((emp, obr), {
            "cod": f"{emp}-{obr}",
            "nm": str(p.get("obra_nome") or f"Obra {obr}").strip(),
            "orcado": 0.0,
            "realizado": 0.0,
            "a_pagar": 0.0,
            "saldo": 0.0,
            "children": [],
            "titulos": [],
        })
        parent = root
        parts = _uau_split_eap(eap)
        for i in range(1, len(parts) + 1):
            code = ".".join(parts[:i])
            parent = _uau_child(parent, code, str(p.get("descricao") or "").strip() if code == eap else code)
            node_by_key[(emp, obr, code)] = parent
        eap_leaf_keys.add((emp, obr, eap))
        totals = totals_by_key.get((emp, obr, eap), {})
        parent["un"] = str(p.get("unidade") or "").strip()
        parent["qt"] = f2(p.get("quantidade"))
        parent["realizado"] = f2(totals.get("realizado", parent.get("realizado", 0)))
        parent["a_pagar"] = f2(totals.get("a_pagar", parent.get("a_pagar", 0)))
        # PlanTotal nao tem valor orcado - pros centros de custo do
        # CJ03/CJ05, usa o valor real da planilha conferida a mao.
        cc_valores = cj0305_map.get(f"{emp}-{obr}")
        if cc_valores and eap in cc_valores:
            parent["orcado"] = f2(cc_valores[eap])
        parent["titulos"] = titles_by_key.get((emp, obr, eap), [])

    for key, totals in totals_by_key.items():
        emp, obr, eap = key
        if key in eap_leaf_keys:
            continue  # já contabilizado acima, como folha exata do PlanTotal
        if key in node_by_key:
            # título lançado direto num nível da EAP que também é "pai" de
            # itens mais detalhados (ex: alguém apropriou em "02" em vez de
            # "02.01.05"). Sem isso esse valor sumia: o roll() reescreve o
            # nó com a soma dos filhos, então precisa entrar como um filho
            # "sem detalhamento" pra não ser descartado.
            existing = node_by_key[key]
            alvo = _uau_child(existing, f"{eap}.SEM_DETALHE", "Lançado direto neste nível (sem detalhamento)")
            alvo["realizado"] = f2(alvo.get("realizado", 0) + totals.get("realizado", 0))
            alvo["a_pagar"] = f2(alvo.get("a_pagar", 0) + totals.get("a_pagar", 0))
            alvo["titulos"] = (alvo.get("titulos") or []) + titles_by_key.get(key, [])
            continue
        root = roots_by_obra.setdefault((emp, obr), {
            "cod": f"{emp}-{obr}",
            "nm": f"Obra {obr}",
            "orcado": 0.0,
            "realizado": 0.0,
            "a_pagar": 0.0,
            "saldo": 0.0,
            "children": [],
            "titulos": [],
        })
        parent = root
        parts = _uau_split_eap(eap) or ["SEM_EAP"]
        for i in range(1, len(parts) + 1):
            code = ".".join(parts[:i])
            parent = _uau_child(parent, code, "Sem apropriação" if code == "SEM_EAP" else code)
            node_by_key[(emp, obr, code)] = parent
        parent["realizado"] = f2(totals.get("realizado", 0))
        parent["a_pagar"] = f2(totals.get("a_pagar", 0))
        cc_valores = cj0305_map.get(f"{emp}-{obr}")
        if cc_valores and eap in cc_valores:
            parent["orcado"] = f2(cc_valores[eap])
        parent["titulos"] = titles_by_key.get(key, [])

    def roll(node):
        for child in node["children"]:
            roll(child)
        if node["children"]:
            node["orcado"] = f2(sum(c.get("orcado", 0) for c in node["children"]))
            node["realizado"] = f2(sum(c.get("realizado", 0) for c in node["children"]))
            node["a_pagar"] = f2(sum(c.get("a_pagar", 0) for c in node["children"]))
        node["saldo"] = f2(node.get("orcado", 0) - node.get("realizado", 0) - node.get("a_pagar", 0))
        node["children"].sort(key=lambda x: _uau_nat(x.get("cod", "")))

    tree = list(roots_by_obra.values())
    for root in tree:
        roll(root)
    tree.sort(key=lambda x: _uau_nat(x.get("cod", "")))
    summary = {
        "orcado": f2(sum(n.get("orcado", 0) for n in tree)),
        "realizado": f2(sum(n.get("realizado", 0) for n in tree)),
        "a_pagar": f2(sum(n.get("a_pagar", 0) for n in tree)),
    }
    summary["saldo"] = f2(summary["orcado"] - summary["realizado"] - summary["a_pagar"])

    def _uau_status_aprovacao(v):
        code = str(v if v is not None else "").strip()
        return {"0": "Pendente", "1": "Aprovada", "2": "Reprovada", "3": "Cancelada"}.get(code, f"Status {code}" if code else "Pendente")

    def _uau_status_compra(v, confirmed=None):
        code = str(v if v is not None else "").strip()
        if confirmed:
            return "Confirmado"
        return {"0": "Pendente", "1": "Confirmado", "2": "Recusado", "3": "Cancelado"}.get(code, f"Status {code}" if code else "Pendente")

    solicitacoes_map = {}
    for r in sol_rows:
        emp = str(r.get("empresa") or "").strip()
        obr = str(r.get("obra") or "").strip()
        num = str(r.get("solicitacao") or "").strip()
        key = (emp, obr, num)
        row = solicitacoes_map.setdefault(key, {
            "cc": f"{emp}-{obr}",
            "solicitacao": num,
            "data": _uau_json_safe(r.get("data")) or "",
            "previsto": _uau_json_safe(r.get("aprovado_em")) or "",
            "status": _uau_status_aprovacao(r.get("status_code")),
            "solicitante": str(r.get("solicitante") or "").strip(),
            "itens": 0,
            "qtd": 0.0,
            "valor": 0.0,
            "saldo": 0.0,
            "saldo_valor": 0.0,
            "cotacao": str(r.get("cotacao") or "").strip(),
        })
        row["itens"] += 1
        row["qtd"] = f2(row["qtd"] + f2(r.get("qtd_solicitada")))
        row["valor"] = f2(row["valor"] + f2(r.get("valor_solicitado")))
        row["saldo"] = f2(row["saldo"] + f2(r.get("saldo_planejado")))
        row["saldo_valor"] = row["saldo"]
    solicitacoes = sorted(solicitacoes_map.values(), key=lambda x: (x.get("data", ""), x.get("solicitacao", "")), reverse=True)

    cotacoes = []
    for r in cot_rows:
        emp = str(r.get("empresa") or "").strip()
        obr = str(r.get("obra") or "").strip()
        qtd = f2(r.get("qtd_cotada") or r.get("qtd_solicitada"))
        preco = f2(r.get("preco"))
        valor = f2(qtd * preco)
        status_txt = _uau_status_compra(r.get("status_code"))
        cotacoes.append({
            "cc": f"{emp}-{obr}" if obr else emp,
            "cotacao": str(r.get("cotacao") or "").strip(),
            "solicitacao": "",
            "data": _uau_json_safe(r.get("data")) or "",
            "aprovado": status_txt.lower() not in ("pendente", "cancelado", "recusado"),
            "status": status_txt,
            "status_pedido": status_txt,
            "insumo": str(r.get("insumo") or "").strip(),
            "detalhe": str(r.get("insumo_nome") or r.get("descricao") or "").strip(),
            "eap": "",
            "item_orc": "",
            "comprador": str(r.get("comprador") or "").strip(),
            "credor": str(r.get("credor") or "").strip(),
            "orcado_eap": 0.0,
            "realizado_eap": 0.0,
            "comprometido_eap": valor,
            "saldo_eap": 0.0,
            "apropriado": valor,
            "realizado": 0.0,
            "comprometido": valor,
            "impacto": valor,
        })

    pedidos_map = {}
    for r in oc_rows:
        emp = str(r.get("empresa") or "").strip()
        obr = str(r.get("obra") or "").strip()
        pedido = str(r.get("pedido") or "").strip()
        key = (emp, obr, pedido)
        status_txt = _uau_status_compra(r.get("status_code"), r.get("confirmado_em"))
        row = pedidos_map.setdefault(key, {
            "cc": f"{emp}-{obr}",
            "pedido": pedido,
            "data": _uau_json_safe(r.get("data")) or "",
            "aprovacao": _uau_json_safe(r.get("confirmado_em")) or "",
            "aprovado": status_txt.lower() == "confirmado",
            "status": status_txt,
            "comprador": str(r.get("comprador") or "").strip(),
            "credores": [],
            "itens": 0,
            "valor": 0.0,
            "realizado": 0.0,
            "a_pagar": 0.0,
            "saldo": 0.0,
            "status_pagamento": "Sem título vinculado",
            "titulos": [],
            "nfs": [],
            "vinculos": [],
            "cotacao": str(r.get("cotacao") or "").strip(),
        })
        cred = str(r.get("credor") or "").strip()
        if cred and cred not in row["credores"]:
            row["credores"].append(cred)
        qtd = f2(r.get("qtd"))
        preco = f2(r.get("preco"))
        valor = f2((qtd * preco) + f2(r.get("valor_ipi")))
        insumo = str(r.get("insumo") or "").strip()
        pay_rows = titles_by_pedido_item.get((emp, obr, pedido, insumo), [])
        pago_item = f2(sum(f2(t.get("vlr_pg")) for t in pay_rows))
        apagar_item = f2(sum(f2(t.get("vlr_a_pagar")) for t in pay_rows))
        saldo_item = f2(valor - pago_item - apagar_item)
        titulos_item = []
        nfs_item = []
        pagamentos_item = []
        for t in pay_rows:
            numtit = str(t.get("numtit") or "").strip()
            doc = str(t.get("documento") or t.get("doc") or "").strip()
            if numtit and numtit not in titulos_item:
                titulos_item.append(numtit)
            if doc and doc not in nfs_item:
                nfs_item.append(doc)
            pagamentos_item.append({
                "titulo": numtit,
                "documento": doc,
                "credor": t.get("credor", ""),
                "data": t.get("dt_baixa", ""),
                "status": t.get("status", ""),
                "valor": f2(t.get("vlr_ap")),
                "pago": f2(t.get("vlr_pg")),
                "a_pagar": f2(t.get("vlr_a_pagar")),
            })
        row["itens"] += 1
        row["valor"] = f2(row["valor"] + valor)
        row["realizado"] = f2(row["realizado"] + pago_item)
        row["a_pagar"] = f2(row["a_pagar"] + apagar_item)
        row["saldo"] = f2(row["saldo"] + saldo_item)
        for numtit in titulos_item:
            if numtit not in row["titulos"]:
                row["titulos"].append(numtit)
        for nf in nfs_item:
            if nf not in row["nfs"]:
                row["nfs"].append(nf)
        row["vinculos"].append({
            "insumo": insumo,
            "detalhe": str(r.get("insumo_nome") or r.get("detalhe") or "").strip(),
            "eap": "",
            "item_orc": "",
            "credores": [cred] if cred else [],
            "titulos": titulos_item,
            "nfs": nfs_item,
            "valor": valor,
            "realizado": pago_item,
            "a_pagar": apagar_item,
            "saldo": saldo_item,
            "pagamentos": pagamentos_item,
        })
    for row in pedidos_map.values():
        title_count = len(row.get("titulos") or [])
        pago = f2(row.get("realizado"))
        apagar = f2(row.get("a_pagar"))
        saldo = f2(row.get("saldo"))
        valor = f2(row.get("valor"))
        if not title_count:
            row["status_pagamento"] = "Sem título vinculado"
        elif pago >= valor - 0.01 and apagar <= 0.01:
            row["status_pagamento"] = "Pago"
        elif pago > 0.01 and apagar > 0.01:
            row["status_pagamento"] = "Pago parcial / há a pagar"
        elif apagar > 0.01:
            row["status_pagamento"] = "A pagar"
        elif saldo > 0.01:
            row["status_pagamento"] = "Sem financeiro completo"
        else:
            row["status_pagamento"] = "Vinculado"
    pedidos = sorted(pedidos_map.values(), key=lambda x: (x.get("data", ""), x.get("pedido", "")), reverse=True)

    def _uau_contract_status(situacao, status):
        sit = str(situacao if situacao is not None else "").strip()
        sta = str(status if status is not None else "").strip()
        if sit in ("2", "3", "4") or sta in ("2", "3", "4"):
            return "Rescindindo"
        if sit in ("1", "5") or sta in ("5", "9"):
            return "Concluido"
        return "Ativo"

    contract_items_by_key = {}
    for r in ct_item_rows:
        emp = str(r.get("empresa") or "").strip()
        obr = str(r.get("obra") or "").strip()
        contrato = _uau_clean_ref(r.get("contrato"))
        key = (emp, obr, contrato)
        qtd = f2(r.get("qtd"))
        preco = f2(r.get("preco"))
        contract_items_by_key.setdefault(key, []).append({
            "item": _uau_clean_ref(r.get("item")),
            "servico_cod": _uau_clean_ref(r.get("servico")),
            "servico": str(r.get("descricao") or "").strip(),
            "un": str(r.get("unidade") or "").strip(),
            "qtd_contrato": qtd,
            "preco_unit": preco,
            "valor_contrato": f2(qtd * preco),
        })

    pl_by_contract = {}
    pl_by_contract_ins = {}
    for r in ct_pl_rows:
        emp = str(r.get("empresa") or "").strip()
        obr = str(r.get("obra") or "").strip()
        contrato = _uau_clean_ref(r.get("contrato"))
        key = (emp, obr, contrato)
        pl = {
            "eap": str(r.get("eap") or "").strip(),
            "insumo": _uau_clean_ref(r.get("insumo")),
            "servico": _uau_clean_ref(r.get("servico")),
            "descricao": str(r.get("descricao_eap") or "").strip(),
            "valor": f2(r.get("valor_vinc") or r.get("valor_aprov")),
        }
        if not pl["eap"] and not pl["descricao"] and pl["valor"] == 0:
            continue
        pl_by_contract.setdefault(key, []).append(pl)
        if pl["insumo"]:
            pl_by_contract_ins.setdefault((emp, obr, contrato, pl["insumo"]), []).append(pl)
        if pl["servico"]:
            pl_by_contract_ins.setdefault((emp, obr, contrato, pl["servico"]), []).append(pl)

    med_totals_by_contract_item = {}
    med_rows_by_contract = {}
    med_codes_by_contract = {}
    for r in ct_med_rows:
        emp = str(r.get("empresa") or "").strip()
        obr = str(r.get("obra") or "").strip()
        contrato = _uau_clean_ref(r.get("contrato"))
        medicao = _uau_clean_ref(r.get("medicao"))
        item_code = _uau_clean_ref(r.get("item_contrato"))
        key = (emp, obr, contrato)
        if medicao:
            med_codes_by_contract.setdefault(key, set()).add(medicao)
        qtd = f2(r.get("qtd"))
        valor_item = f2(r.get("valor_item"))
        if item_code:
            tk = (emp, obr, contrato, item_code)
            totals = med_totals_by_contract_item.setdefault(tk, {"qtd": 0.0, "valor": 0.0})
            totals["qtd"] = f2(totals["qtd"] + qtd)
            totals["valor"] = f2(totals["valor"] + valor_item)
        med_titles = titles_by_contract_med.get((emp, obr, contrato, medicao), [])
        title = med_titles[0] if med_titles else {}
        med_rows_by_contract.setdefault(key, []).append({
            "medicao": medicao,
            "data_medicao": _uau_json_safe(r.get("data_medicao")) or "",
            "vencimento": title.get("dt_baixa", ""),
            "titulo": title.get("numtit", ""),
            "status_titulo": title.get("status", ""),
            "item": item_code,
            "servico": "",
            "qtd": qtd,
            "valor": valor_item,
            "usuario": str(r.get("usuario") or "").strip(),
            "dt_baixa": title.get("dt_baixa", ""),
        })

    prev_by_contract = {}
    for r in ct_prev_rows:
        emp = str(r.get("empresa") or "").strip()
        obr = str(r.get("obra") or "").strip()
        contrato = _uau_clean_ref(r.get("contrato"))
        titulo = _uau_clean_ref(r.get("titulo"))
        parcela = _uau_clean_ref(r.get("parcela"))
        pago = bool(r.get("pagamento")) or str(r.get("status_code") or "").strip() == "2"
        prev_by_contract.setdefault((emp, obr, contrato), []).append({
            "titulo": f"{titulo}/{parcela}" if parcela else titulo,
            "data": _uau_json_safe(r.get("vencimento")) or "",
            "valor": f2(r.get("valor")),
            "situacao": "Paga" if pago else "Em aberto",
        })

    contratos = []
    itens_ct = {}
    contract_analysis = {}
    for r in ct_rows:
        emp = str(r.get("empresa") or "").strip()
        obr = str(r.get("obra") or "").strip()
        contrato = _uau_clean_ref(r.get("contrato"))
        if not contrato:
            continue
        key = (emp, obr, contrato)
        doc = f"CTS/{contrato}"
        items = contract_items_by_key.get(key, [])
        pls = pl_by_contract.get(key, [])
        titles = titles_by_contract.get(key, [])
        valor_total = f2(sum(f2(i.get("valor_contrato")) for i in items))
        if valor_total <= 0.01:
            valor_total = f2(sum(f2(p.get("valor")) for p in pls))
        valor_medido = f2(sum(
            f2(t.get("valor")) for k, t in med_totals_by_contract_item.items()
            if k[0] == emp and k[1] == obr and k[2] == contrato
        ))
        if valor_medido <= 0.01:
            valor_medido = f2(sum(f2(t.get("vlr_pg")) for t in titles))
        a_pagar = f2(sum(f2(t.get("vlr_a_pagar")) for t in titles))
        saldo = f2(valor_total - valor_medido)
        pct = f2((valor_medido / valor_total * 100) if valor_total else 0)
        credor = str(r.get("credor") or "").strip()
        fantasia = str(r.get("nome_fantasia") or "").strip()
        status = _uau_contract_status(r.get("situacao_code"), r.get("status_code"))
        contratos.append({
            "docnuct": doc,
            "contrato": contrato,
            "codccusto": f"{emp}-{obr}" if obr else emp,
            "nmcc": str(r.get("obra_nome") or f"Obra {obr}").strip(),
            "nmcredor": credor or fantasia or "Sem credor",
            "nome_fantasia": fantasia,
            "cnpj": "",
            "vlrtotct": valor_total,
            "vlr_total": valor_total,
            "vlrmedidoct": valor_medido,
            "vlr_medido": valor_medido,
            "vlrpendct": a_pagar,
            "a_pagar": a_pagar,
            "saldo_contrato": saldo,
            "pct": pct,
            "status": status,
            "itens": len(items),
            "data_contrato": _uau_json_safe(r.get("data_contrato")) or "",
            "objeto": str(r.get("objeto") or "").strip(),
            "fonte": "UAU",
        })

        detail_rows = []
        seen_budget = set()
        for item in items:
            item_code = item.get("item", "")
            med = med_totals_by_contract_item.get((emp, obr, contrato, item_code), {"qtd": 0.0, "valor": 0.0})
            item_pls = pl_by_contract_ins.get((emp, obr, contrato, item.get("servico_cod", "")), []) or pls[:1]
            aprop = " | ".join(sorted({p.get("eap", "") for p in item_pls if p.get("eap")})) or "Apropriação pendente"
            aprop_nm = " | ".join(sorted({p.get("descricao", "") for p in item_pls if p.get("descricao")})) or ""
            orcado_aprop = 0.0
            for p in item_pls:
                budget_key = (p.get("eap", ""), p.get("insumo", ""), p.get("servico", ""))
                if budget_key in seen_budget:
                    continue
                seen_budget.add(budget_key)
                orcado_aprop = f2(orcado_aprop + f2(p.get("valor")))
            consumido_aprop = f2(med.get("valor"))
            saldo_aprop = f2(orcado_aprop - consumido_aprop)
            qtd_contrato = f2(item.get("qtd_contrato"))
            qtd_medida = f2(med.get("qtd"))
            detail_rows.append({
                "item": item_code,
                "servico": item.get("servico") or item.get("servico_cod") or "",
                "detalhe": item.get("servico") or "",
                "un": item.get("un") or "",
                "preco_unit": f2(item.get("preco_unit")),
                "qtd_contrato": qtd_contrato,
                "valor_contrato": f2(item.get("valor_contrato")),
                "qtd_medida": qtd_medida,
                "valor_medido": consumido_aprop,
                "qtd_saldo": f2(qtd_contrato - qtd_medida),
                "valor_saldo": f2(item.get("valor_contrato") - consumido_aprop),
                "apropriacao": aprop,
                "apropriacao_nm": aprop_nm,
                "n2_nm": "",
                "is_empty_item": f2(item.get("valor_contrato")) <= 0 and qtd_contrato <= 0,
                "orcado_apropriacao": orcado_aprop,
                "consumido_apropriacao": consumido_aprop,
                "a_pagar_apropriacao": 0.0,
                "saldo_apropriacao": saldo_aprop,
            })
        med_rows = med_rows_by_contract.get(key, [])
        item_name_by_code = {str(i.get("item") or ""): i.get("servico", "") for i in items}
        for m in med_rows:
            m["servico"] = item_name_by_code.get(str(m.get("item") or ""), m.get("servico") or "")
        previsoes = prev_by_contract.get(key) or [{
            "titulo": t.get("numtit", ""),
            "data": t.get("dt_baixa", ""),
            "valor": f2(t.get("vlr_ap")),
            "situacao": t.get("status", ""),
        } for t in titles]
        itens_ct[doc] = [{
            "cod": i.get("item", ""),
            "item": i.get("servico", ""),
            "total": f2(i.get("valor_contrato")),
            "medido": f2(med_totals_by_contract_item.get((emp, obr, contrato, i.get("item", "")), {}).get("valor", 0)),
            "pendente": f2(i.get("valor_contrato") - med_totals_by_contract_item.get((emp, obr, contrato, i.get("item", "")), {}).get("valor", 0)),
        } for i in items]
        contract_analysis[doc] = {
            "contrato": doc,
            "fornecedor": credor or fantasia or "Sem credor",
            "centro_custo": str(r.get("obra_nome") or f"{emp}-{obr}").strip(),
            "codccusto": f"{emp}-{obr}" if obr else emp,
            "orcamento_nominal_obra": summary.get("orcado", 0),
            "valor_total": valor_total,
            "valor_medido": valor_medido,
            "saldo_contrato": saldo,
            "a_pagar_lancado": a_pagar,
            "pct_medido": pct,
            "medicoes_vinculadas": len(med_codes_by_contract.get(key, set())),
            "itens": detail_rows,
            "medicoes": med_rows,
            "previsoes": previsoes,
            "total_orcado_apropriacao": f2(sum(f2(x.get("orcado_apropriacao")) for x in detail_rows)),
            "total_consumido_apropriacao": f2(sum(f2(x.get("consumido_apropriacao")) for x in detail_rows)),
            "total_a_pagar_apropriacao": f2(sum(f2(x.get("a_pagar_apropriacao")) for x in detail_rows)),
            "total_saldo_apropriacao": f2(sum(f2(x.get("saldo_apropriacao")) for x in detail_rows)),
        }
    contratos.sort(key=lambda x: (x.get("codccusto", ""), _uau_nat(x.get("contrato", ""))))

    return {
        "ok": True,
        "tree": tree,
        "financeiro": financeiro,
        "contratos": contratos,
        "itens_ct": itens_ct,
        "contract_analysis": contract_analysis,
        "solicitacoes": solicitacoes,
        "cotacoes": cotacoes,
        "pedidos": pedidos,
        "suprimentos": {"solicitacoes": solicitacoes, "cotacoes": cotacoes, "pedidos": pedidos},
        "summary": summary,
        "meta": {
            "fonte": "SQL UAU",
            "empresa": empresa,
            "obra": obra,
            "planejamento_itens": len(plan_rows),
            "titulos": len(tit_rows),
            "solicitacoes": len(solicitacoes),
            "cotacoes": len(cotacoes),
            "pedidos": len(pedidos),
            "contratos": len(contratos),
            "suprimentos_errors": sup_errors,
            "updated_at": time.strftime("%d/%m/%Y %H:%M:%S"),
            "message": f"{len(plan_rows)} itens de EAP, {len(tit_rows)} títulos, {len(solicitacoes)} solicitações, {len(cotacoes)} cotações e {len(pedidos)} pedidos carregados do UAU.",
        },
    }


@app.route("/")
def index():
    template_path = Path(BASE_DIR) / "templates" / "index.html"
    html = template_path.read_text(encoding="utf-8")
    html = html.replace("{{ asset_v|default(0) }}", str(int(time.time())))

    # esconde as abas não liberadas já no HTML enviado — sem isso, a tela
    # mostrava TODAS as abas primeiro e só escondia as bloqueadas depois que
    # o /api/whoami respondia (dava pra ver um flash com tudo liberado).
    user = _current_auth_user()
    if user and user.get("role") != "admin":
        allowed = set(user.get("allowed_tabs") or [])
        if allowed:
            escondidas = [k for k, _ in DASHBOARD_TABS if k not in allowed]
            if escondidas:
                regras = "".join(f".tab[onclick*=\"showTab('{k}'\"]{{display:none!important}}" for k in escondidas)
                html = html.replace("</head>", f"<style id=\"srv-tab-perms\">{regras}</style></head>")

    resp = make_response(html)
    resp.headers["Content-Type"] = "text/html; charset=utf-8"
    resp.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
    resp.headers['Pragma'] = 'no-cache'
    resp.headers['Expires'] = '0'
    return resp

@app.route("/ifc_viewer.js")
def ifc_viewer_js():
    return send_from_directory("templates", "ifc_viewer.js")

@app.route("/ifcviz.js")
def ifcviz_js():
    return send_from_directory("templates", "ifcviz.js")

@app.route("/planejamento")
def planejamento():
    resp = send_from_directory("templates", "planejamento.html")
    resp.headers["Cache-Control"] = "no-cache, must-revalidate"
    return resp


def _planning_tree_paid_months(nodes):
    """Consolida somente valores efetivamente pagos dos títulos da árvore."""
    monthly = {}

    def walk(items):
        for node in items or []:
            for title in node.get("titulos", []) or []:
                paid = f2(title.get("vlr_pg", 0))
                paid_date = str(title.get("dt_baixa", "") or "")
                if paid <= 0 or len(paid_date) < 7 or paid_date in ("—", "nan", "None"):
                    continue
                month = paid_date[:7]
                monthly[month] = monthly.get(month, 0) + paid
            walk(node.get("children", []))

    walk(nodes)
    return [{"mes": month, "v": f2(value)} for month, value in sorted(monthly.items())]


def _planning_eap_count(nodes):
    count = 0
    for node in nodes or []:
        children = node.get("children", []) or []
        count += _planning_eap_count(children) if children else 1
    return count


@app.route("/api/planejamento_data")
def api_planejamento_data():
    """Entrega orçamento/EAP e pagamentos reais para o módulo de planejamento."""
    data, err = build_data()
    if err:
        return jsonify({"ok": False, "error": err}), 500

    budget_source = request.args.get("budget_source", "api")
    cost_source = request.args.get("cost_source", "api")
    api_tree = data.get("orc_tree", []) or []
    report_tree = data.get("orc_rel_tree", []) or []
    api_summary = (data.get("summary_per_cc", {}) or {}).get("1016") or data.get("summary", {}) or {}
    report_meta = data.get("orc_rel_meta", {}) or {}

    api_monthly = ((data.get("fluxo_per_cc", {}) or {}).get("1016") or
                   _planning_tree_paid_months(api_tree))
    report_monthly = _planning_tree_paid_months(report_tree)

    budget_sources = [
        {
            "id": "api",
            "label": "Sienge / API",
            "bac": f2(api_summary.get("vlr_orcado", 0)),
            "eap_items": _planning_eap_count(api_tree),
            "file": os.path.basename(data.get("excel_source", "") or "bd_sienge_tmz.xlsx"),
        },
        {
            "id": "relatorio",
            "label": "Relatório local",
            "bac": f2(report_meta.get("total_orcado", 0)),
            "eap_items": _planning_eap_count(report_tree),
            "file": report_meta.get("fonte_orcamento", "ORC-EXE-SYN-EXE_R00.xlsx"),
        },
    ]
    cost_sources = [
        {
            "id": "api",
            "label": "Pagos no Sienge / API",
            "paid_total": f2(sum(f2(row.get("v", 0)) for row in api_monthly)),
            "monthly": api_monthly,
            "file": os.path.basename(data.get("excel_source", "") or "bd_sienge_tmz.xlsx"),
        },
        {
            "id": "relatorio",
            "label": "Pagos nos relatórios locais",
            "paid_total": f2(sum(f2(row.get("v", 0)) for row in report_monthly)),
            "monthly": report_monthly,
            "file": report_meta.get("fonte_realizado", "Relatório Analítico de Apropriações por Obra.xlsx"),
        },
    ]

    selected_budget = next((src for src in budget_sources if src["id"] == budget_source), budget_sources[0])
    selected_cost = next((src for src in cost_sources if src["id"] == cost_source), cost_sources[0])
    return jsonify({
        "ok": True,
        "budget_sources": budget_sources,
        "cost_sources": cost_sources,
        "budget": selected_budget,
        "cost": selected_cost,
        "updated_at": data.get("excel_updated_at", ""),
    })

@app.route("/dash_data.js")
def dash_data_js():
    if os.path.exists(DASH_JS_CACHE_PATH):
        resp = send_from_directory(DATA_DIR, os.path.basename(DASH_JS_CACHE_PATH), mimetype="application/javascript")
        resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        resp.headers["Pragma"] = "no-cache"
        resp.headers["Expires"] = "0"
        return resp
    data, err = build_data()
    headers = {
        'Content-Type':'application/javascript; charset=utf-8',
        'Cache-Control':'no-store, no-cache, must-revalidate, max-age=0',
        'Pragma':'no-cache',
        'Expires':'0',
    }
    if err:
        js = f"console.error({json.dumps(err)});const DASH={{}};"
        return js, 500, headers
    js = "const DASH = " + json.dumps(data, ensure_ascii=False, default=str) + ";"
    return js, 200, headers


@app.route("/api/debug")
def api_debug():
    from flask import jsonify
    data, err = build_data()
    if err:
        return jsonify({"ok": False, "error": err}), 500
    keys = list(data.keys()) if data else []
    sizes = {k: len(str(data[k])) for k in keys} if data else {}
    return jsonify({"ok": True, "keys": keys, "sizes": sizes})

@app.route("/api/mobile-link")
def api_mobile_link():
    user = _current_auth_user()
    if not user or user.get("role") != "admin":
        return jsonify({"ok": False, "error": "Restrito ao administrador."}), 403
    path = os.path.join(BASE_DIR, "LINK_MOBILE_ATUAL.txt")
    if not os.path.exists(path):
        return jsonify({"ok": False, "error": "Link ainda não gerado nesta máquina."})
    try:
        with open(path, "r", encoding="utf-8-sig") as fh:
            text = fh.read()
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)})
    match = re.search(r"https://[a-z0-9-]+\.trycloudflare\.com", text)
    gerado = re.search(r"Gerado em:\s*(.+)", text)
    if not match:
        return jsonify({"ok": False, "error": "Nenhum link do Cloudflare encontrado no arquivo ainda.", "raw": text})
    return jsonify({
        "ok": True,
        "url": match.group(0),
        "gerado_em": gerado.group(1).strip() if gerado else "",
    })


@app.route("/api/status")
def api_status():
    excel_path = resolve_excel_path()
    # rota pública (health check): não expõe caminho de pasta do servidor
    payload = {"ok": os.path.exists(excel_path), "build": BUILD_ID}
    if request.remote_addr in ("127.0.0.1", "::1"):  # detalhes só para o health check interno
        payload["excel"] = os.path.basename(excel_path)
        payload["source"] = "sharepoint_sync" if os.path.abspath(excel_path).lower() != os.path.abspath(LOCAL_EXCEL_PATH).lower() else "local_backup"
    # Teste "profundo" usado só pelo script de atualização (de dentro do container):
    # confirma que a versão nova consegue carregar os dados do painel antes de entrar no ar.
    if request.args.get("deep") == "1" and request.remote_addr in ("127.0.0.1", "::1"):
        t0 = time.time()
        try:
            data, err = build_data()
            payload["deep"] = {"ok": bool(data) and not err, "chaves": len(data or {}),
                               "segundos": round(time.time() - t0, 1), "erro": str(err or "")}
        except Exception as exc:
            payload["deep"] = {"ok": False, "erro": f"{type(exc).__name__}: {exc}"}
    return jsonify(payload)


_tunnel_lock = threading.Lock()

def _atualizar_link_tunel():
    """Mata o cloudflared atual (se houver) e sobe um novo tunel rapido,
    igual ao iniciar_dashboard_online.ps1 faz — usado tanto pelo botao
    'Atualizar link de acesso' quanto por um teste manual futuro. O link
    do tunel rapido (*.trycloudflare.com) muda sozinho de tempos em tempos
    (o processo local pode continuar "vivo" enquanto o registro dele na
    borda da Cloudflare ja caiu, silenciosamente) — esse botao evita
    precisar reiniciar tudo manualmente quando isso acontece."""
    if os.name != "nt":
        return {"ok": False, "error": "No servidor o endereço é fixo (domínio próprio) — não há link temporário pra atualizar."}
    cloudflared_exe = os.path.join(BASE_DIR, "tools", "cloudflared.exe")
    log_dir = os.path.join(BASE_DIR, "logs")
    out_log = os.path.join(log_dir, "cloudflared_out.log")
    err_log = os.path.join(log_dir, "cloudflared_err.log")
    link_file = os.path.join(BASE_DIR, "LINK_MOBILE_ATUAL.txt")
    os.makedirs(log_dir, exist_ok=True)

    if not os.path.exists(cloudflared_exe):
        return {"ok": False, "error": f"cloudflared.exe não encontrado em {cloudflared_exe}."}

    local_port = request.host.split(":")[-1] if ":" in request.host else "5000"
    local_url = f"http://127.0.0.1:{local_port}"

    with _tunnel_lock:
        try:
            subprocess.run(
                ["taskkill", "/IM", "cloudflared.exe", "/F"],
                capture_output=True, timeout=10,
            )
        except Exception:
            pass
        time.sleep(1)

        for path in (out_log, err_log):
            try:
                if os.path.exists(path):
                    os.remove(path)
            except Exception:
                pass

        try:
            out_fh = open(out_log, "w", encoding="utf-8")
            err_fh = open(err_log, "w", encoding="utf-8")
            subprocess.Popen(
                [cloudflared_exe, "tunnel", "--url", local_url, "--no-autoupdate"],
                cwd=BASE_DIR, stdout=out_fh, stderr=err_fh,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
        except Exception as exc:
            return {"ok": False, "error": f"Falha ao iniciar cloudflared: {exc}"}

        link = None
        pattern = re.compile(r"https://[a-z0-9-]+\.trycloudflare\.com")
        for _ in range(60):
            text = ""
            for path in (out_log, err_log):
                try:
                    with open(path, "r", encoding="utf-8", errors="ignore") as fh:
                        text += fh.read()
                except Exception:
                    pass
            m = pattern.search(text)
            if m:
                link = m.group(0)
                break
            time.sleep(1)

    if link:
        conteudo = (
            f"Link mobile atual do Dashboard TEMAZEC:\n{link}\n\n"
            f"Dashboard local publicado:\n{local_url}\n\n"
            f"Gerado em: {time.strftime('%d/%m/%Y %H:%M:%S')}\n\n"
            "Observacao: este link rapido do Cloudflare pode mudar quando o computador reiniciar.\n"
            "Para ter um link fixo permanente, e necessario configurar um tunel nomeado com dominio proprio na Cloudflare.\n"
        )
        try:
            with open(link_file, "w", encoding="utf-8") as fh:
                fh.write(conteudo)
        except Exception:
            pass
        return {"ok": True, "link": link, "local_url": local_url}

    return {
        "ok": False,
        "error": "O dashboard local está ativo, mas o link público do Cloudflare não foi capturado a tempo. Tente de novo em alguns segundos.",
    }


@app.route("/api/tunnel/atualizar", methods=["POST"])
def api_tunnel_atualizar():
    user = _current_auth_user()
    if not user or user.get("role") != "admin":
        return jsonify({"ok": False, "error": "Restrito ao administrador."}), 403
    resultado = _atualizar_link_tunel()
    return jsonify(resultado), (200 if resultado.get("ok") else 502)


_MELIUS_INTERROMPIDA = "Atualização interrompida (o servidor foi reiniciado enquanto rodava). Clique em atualizar para rodar de novo."


def _status_sem_thread_vira_erro(status, holder):
    """O status dos robôs fica gravado em arquivo e sobrevive ao reinício do
    servidor: se ele diz 'rodando' mas não há thread viva neste processo, a
    execução morreu. Sem isso a tela ficava 'em andamento' pra sempre e o botão
    ficava travado."""
    if isinstance(status, dict) and status.get("state") in ("starting", "running"):
        thread = holder.get("thread")
        # thread criada e ainda não iniciada (ident None) conta como viva
        if not (thread and (thread.is_alive() or thread.ident is None)):
            status = dict(status, state="error", message=_MELIUS_INTERROMPIDA)
    return status


_melius_rompimentos_sync = {"lock": threading.Lock(), "thread": None}


def _read_melius_rompimentos_status():
    try:
        return _status_sem_thread_vira_erro(
            json.loads(Path(MELIUS_ROMPIMENTOS_STATUS_PATH).read_text(encoding="utf-8")), _melius_rompimentos_sync)
    except Exception:
        return {"state": "idle", "message": "", "updated_at": ""}


def _write_melius_rompimentos_status(state, message):
    payload = {"state": state, "message": message, "updated_at": time.strftime("%Y-%m-%d %H:%M:%S")}
    os.makedirs(DATA_DIR, exist_ok=True)
    Path(MELIUS_ROMPIMENTOS_STATUS_PATH).write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    return payload


def _load_melius_rompimentos_cache():
    try:
        data = json.loads(Path(MELIUS_ROMPIMENTOS_CACHE_PATH).read_text(encoding="utf-8"))
        if isinstance(data, dict):
            return data
    except Exception:
        pass
    return {"ok": False, "updated_at": "", "concretagens": []}


def _load_sobreconsumo_concreto():
    """Sobreconsumo/subconsumo de concreto por local, direto da Melius —
    cada concretagem já vem com volume_projetado (calculado no projeto) e
    volume_realizado (efetivamente entregue pelos caminhões), então não
    precisa de nenhum cruzamento com o modelo IFC. Diferença positiva =
    sobreconsumo (gastou mais concreto que o projetado); negativa =
    entregou menos que o projetado."""
    out = []
    try:
        data = _load_melius_rompimentos_cache()
        for c in data.get("concretagens", []) or []:
            proj, real = c.get("volume_projetado"), c.get("volume_realizado")
            if proj is None or real is None:
                continue
            try:
                proj, real = float(proj), float(real)
            except (TypeError, ValueError):
                continue
            diff = f2(real - proj)
            out.append({
                "id": c.get("id", ""),
                "local": c.get("local", "") or "—",
                "data_aplicacao": c.get("data_aplicacao", ""),
                "fornecedor": c.get("fornecedor", ""),
                "fck_projeto": c.get("fck_projeto"),
                "volume_projetado": f2(proj),
                "volume_realizado": f2(real),
                "diferenca": diff,
                "pct_diferenca": f2((diff / proj * 100)) if proj else 0,
            })
        out.sort(key=lambda x: x["diferenca"], reverse=True)
    except Exception:
        pass
    return out


# Itens de orçamento que representam o MATERIAL concreto usinado em si (não
# mão de obra de lançamento, controle tecnológico ou taxa de bombeamento) —
# são os que correspondem 1:1 ao volume de concreto que a Melius registra
# como entregue pelos caminhões em cada concretagem. Conferido manualmente
# contra a árvore real de orçamento (ver conversa) pra não misturar com
# itens de serviço que também têm "concreto" no nome mas são outra coisa.
_CONCRETO_USINADO_EAP_CODES = {
    "02.003.001.007", "02.003.002.007", "02.003.003.004", "02.004.003.001",
}
# Cada item acima já pertence a um ramo do orçamento que representa uma etapa
# física específica — não precisa adivinhar, o próprio Sienge já separa assim.
_CONCRETO_EAP_CATEGORIA = {
    "02.003.001.007": "Estacas de Fundação",   # 02.003.001-FUNDAÇÃO - ESTACAS HÉLICE
    "02.003.002.007": "Blocos",                 # 02.003.002-FUNDAÇÃO - BLOCOS
    "02.003.003.004": "Contenção",              # 02.003.003-CONTENÇÕES ESTACAS ESCAVADA
    "02.004.003.001": "Estrutura (Laje/Viga/Pilar)",  # 02.004.003-CONCRETO USINADO (ESTRUTURA)
}
_CONCRETO_CATEGORIAS_ORDEM = ["Estacas de Fundação", "Contenção", "Blocos", "Estrutura (Laje/Viga/Pilar)", "Outros"]
# O código do item de concreto usinado dentro da planilha ORIGINAL de
# orçamento (fonte) não bate dígito a dígito com o código do mesmo item na
# árvore do Sienge (numeração foi re-sequenciada ao longo da obra) — o que
# não muda é o RAMO (3 primeiros níveis): cada um desses ramos só tem UM
# item de "concreto usinado" na planilha original, então o match é seguro
# por ramo, não por código completo.
_CONCRETO_EAP_RAMO = {
    "02.003.001.007": "02.03.01",
    "02.003.002.007": "02.03.02",
    "02.003.003.004": "02.03.03",
    "02.004.003.001": "02.04.03",
}


def _concreto_orcamento_referencia():
    """Lê a planilha ORIGINAL de orçamento (fonte, não o que foi lançado no
    Sienge) pra pegar a quantidade e o preço unitário orçados de concreto
    usinado de cada etapa — usado pra calcular quanto DEVERIA ter sido
    apropriado (volume real da Melius × preço unitário orçado) e comparar
    com o que foi realmente apropriado nos títulos. Esse preço unitário é o
    da composição sintética do orçamento (pode incluir mais que só o
    material) — tratar como referência, não como preço de mercado do
    concreto puro."""
    out = {}
    if not os.path.exists(ORC_REL_PATH):
        return out
    try:
        orc = pd.read_excel(ORC_REL_PATH, sheet_name=2, header=10)
        orc.columns = [str(c).strip() for c in orc.columns]
    except Exception:
        return out
    por_ramo = {}
    current = None
    for _, r in orc.iterrows():
        raw_eap = _clean_txt(r.get("ITEMIZAÇÃO"))
        if raw_eap:
            current = _norm_rel_code(raw_eap)
            if len(current.split(".")) == 4:
                descricao = _clean_txt(r.get("DESCRIÇÃO"))
                if "concreto" in descricao.lower() and "usinad" in descricao.lower():
                    ramo = ".".join(current.split(".")[:3])
                    por_ramo[ramo] = {
                        "codigo_orcamento": current, "descricao_orcamento": descricao,
                        "quantidade_orcada": f2(r.get("QUANT./COEF.")),
                        "preco_unitario_orcado": f2(r.get("PREÇO FINAL (UNIT.)(R$)")),
                        "unidade": _clean_txt(r.get("UNID.")),
                    }
    for cod, ramo in _CONCRETO_EAP_RAMO.items():
        if ramo in por_ramo:
            out[cod] = por_ramo[ramo]
    return out


def _classificar_local_concreto(local):
    """Classifica o campo de texto livre 'local' de uma concretagem da Melius
    numa das mesmas 4 etapas em que o Sienge já separa o orçamento. Texto
    livre nunca é perfeito — quando o local junta mais de uma etapa na mesma
    descrição (ex: "Estacas, blocos e vigas"), fica na primeira que bater,
    na ordem de execução da obra (estaca → contenção → bloco → estrutura).
    'EPT' é o prefixo usado nesta obra pra estaca de contenção (hélice
    contínua), por isso conta como Contenção antes da regra genérica de
    'estaca'."""
    s = _ascii_upper(local or "")
    if not s:
        return "Outros"
    if re.search(r"\bEPT\d*\b", s) or "CONTEN" in s or "TIRANTE" in s:
        return "Contenção"
    if "ESTACA" in s:
        return "Estacas de Fundação"
    if "BLOCO" in s or "RADIER" in s:
        return "Blocos"
    if "LAJE" in s or "VIGA" in s or "PILAR" in s:
        return "Estrutura (Laje/Viga/Pilar)"
    return "Outros"


def _concreto_titulos_apropriados():
    """Títulos lançados nos itens de orçamento de concreto usinado (ver
    _CONCRETO_USINADO_EAP_CODES) — usado pra comparar o valor apropriado
    financeiramente com o volume executado registrado na Melius."""
    try:
        data, err = build_data(force=False)
    except Exception:
        data, err = None, "falha ao carregar build_data"
    if err or not data:
        return [], {"orcado": 0.0, "realizado": 0.0, "a_pagar": 0.0}, []

    def walk(nodes):
        out = []
        for n in nodes or []:
            if n.get("cod") in _CONCRETO_USINADO_EAP_CODES:
                out.append(n)
            out.extend(walk(n.get("children")))
        return out

    nodes = walk(data.get("orc_rel_tree"))
    titulos = []
    totais = {"orcado": 0.0, "realizado": 0.0, "a_pagar": 0.0}
    itens = []
    for n in nodes:
        totais["orcado"] = f2(totais["orcado"] + f2(n.get("orcado")))
        totais["realizado"] = f2(totais["realizado"] + f2(n.get("realizado")))
        totais["a_pagar"] = f2(totais["a_pagar"] + f2(n.get("a_pagar")))
        categoria = _CONCRETO_EAP_CATEGORIA.get(n.get("cod", ""), "Outros")
        itens.append({"cod": n.get("cod", ""), "nm": n.get("nm", ""), "categoria": categoria,
                       "orcado": f2(n.get("orcado")),
                       "realizado": f2(n.get("realizado")), "a_pagar": f2(n.get("a_pagar"))})
        for t in n.get("titulos") or []:
            valor_pago = f2(t.get("vlr_pg"))
            valor_a_pagar = f2(t.get("vlr_a_pagar"))
            titulos.append({
                "eap": n.get("cod", ""), "eap_nome": n.get("nm", ""), "categoria": categoria,
                "numtit": t.get("numtit", ""), "credor": t.get("credor", ""),
                "valor_pago": valor_pago, "valor_a_pagar": valor_a_pagar,
                "valor": f2(valor_pago + valor_a_pagar),
                "status": t.get("status", ""), "data": t.get("dt_baixa") or t.get("dt_venc") or "",
            })
    return titulos, totais, itens


def _load_analise_concreto_execucao_aprop():
    """Confronta o executado (volume de concreto entregue, por local e data,
    direto da Melius) com o apropriado (valor dos títulos lançados nos itens
    de orçamento de concreto usinado, agrupado por mês) — mostra se o que foi
    fisicamente concretado bate com o que foi reconhecido financeiramente no
    mesmo período, ou se há atraso/diferença de lançamento."""
    melius = _load_melius_rompimentos_cache()
    concretagens = [dict(c, categoria=_classificar_local_concreto(c.get("local")))
                    for c in (melius.get("concretagens") or []) if c.get("volume_realizado") is not None]
    titulos, totais_orcamento, itens_orcamento = _concreto_titulos_apropriados()
    ref_orcamento = _concreto_orcamento_referencia()
    for item in itens_orcamento:
        ref = ref_orcamento.get(item.get("cod", "")) or {}
        item["preco_unitario_orcado"] = ref.get("preco_unitario_orcado", 0.0)
        item["quantidade_orcada"] = ref.get("quantidade_orcada", 0.0)
        item["codigo_orcamento_origem"] = ref.get("codigo_orcamento", "")
    preco_ref_por_categoria = {}
    for cod, categoria in _CONCRETO_EAP_CATEGORIA.items():
        ref = ref_orcamento.get(cod) or {}
        if ref.get("preco_unitario_orcado"):
            preco_ref_por_categoria[categoria] = ref["preco_unitario_orcado"]

    def mes_dmy(s):
        try:
            d, m, y = str(s or "").strip().split("/")
            return f"{y}-{m.zfill(2)}"
        except Exception:
            return ""

    def mes_iso(s):
        s = str(s or "").strip()
        return s[:7] if len(s) >= 7 else ""

    por_mes = {}

    def linha_mes(mes):
        return por_mes.setdefault(mes, {
            "mes": mes, "volume_projetado": 0.0, "volume_realizado": 0.0,
            "valor_pago": 0.0, "valor_a_pagar": 0.0, "n_concretagens": 0,
        })

    for c in concretagens:
        mes = mes_dmy(c.get("data_aplicacao", ""))
        if not mes:
            continue
        row = linha_mes(mes)
        row["volume_projetado"] = f2(row["volume_projetado"] + f2(c.get("volume_projetado")))
        row["volume_realizado"] = f2(row["volume_realizado"] + f2(c.get("volume_realizado")))
        row["n_concretagens"] += 1

    for t in titulos:
        mes = mes_iso(t.get("data", ""))
        if not mes:
            continue
        row = linha_mes(mes)
        row["valor_pago"] = f2(row["valor_pago"] + f2(t.get("valor_pago")))
        row["valor_a_pagar"] = f2(row["valor_a_pagar"] + f2(t.get("valor_a_pagar")))

    for row in por_mes.values():
        row["valor_apropriado"] = f2(row["valor_pago"] + row["valor_a_pagar"])
        row["custo_medio_m3"] = f2(row["valor_apropriado"] / row["volume_realizado"]) if row["volume_realizado"] else 0.0

    meses = sorted(por_mes.values(), key=lambda r: r["mes"])
    total_volume_proj = f2(sum(r["volume_projetado"] for r in meses))
    total_volume_real = f2(sum(r["volume_realizado"] for r in meses))
    total_apropriado = f2(sum(r["valor_apropriado"] for r in meses))

    # por etapa/categoria (Estacas, Contenção, Blocos, Estrutura) — o lado
    # financeiro vem direto do EAP de cada título (exato, o Sienge já separa
    # assim); o lado físico vem de classificar o texto livre do "local" da
    # Melius na mesma categoria (aproximado — ver _classificar_local_concreto).
    por_categoria = {cat: {
        "categoria": cat, "n_concretagens": 0, "volume_projetado": 0.0, "volume_realizado": 0.0,
        "valor_pago": 0.0, "valor_a_pagar": 0.0,
    } for cat in _CONCRETO_CATEGORIAS_ORDEM}
    for c in concretagens:
        row = por_categoria[c["categoria"]]
        row["n_concretagens"] += 1
        row["volume_projetado"] = f2(row["volume_projetado"] + f2(c.get("volume_projetado")))
        row["volume_realizado"] = f2(row["volume_realizado"] + f2(c.get("volume_realizado")))
    for t in titulos:
        row = por_categoria[t.get("categoria", "Outros")]
        row["valor_pago"] = f2(row["valor_pago"] + f2(t.get("valor_pago")))
        row["valor_a_pagar"] = f2(row["valor_a_pagar"] + f2(t.get("valor_a_pagar")))
    for row in por_categoria.values():
        row["valor_apropriado"] = f2(row["valor_pago"] + row["valor_a_pagar"])
        row["custo_medio_m3"] = f2(row["valor_apropriado"] / row["volume_realizado"]) if row["volume_realizado"] else 0.0
        row["preco_unitario_orcado"] = preco_ref_por_categoria.get(row["categoria"], 0.0)

    # por etapa x mês — cruzamento fino pra achar ONDE ajustar. O preço de
    # referência não vem da planilha sintética de orçamento (confirmado com o
    # usuário que é um preço composto, não serve) nem de um valor fixo —
    # vem da MEDIANA histórica real de custo/m³ já apropriada pelo Sienge
    # NAQUELA MESMA etapa, nos meses em que apropriação e volume batem bem.
    # A quantidade usada é sempre a da Melius (a correta, conforme pedido).
    # Mês/etapa cujo custo/m³ foge muito dessa mediana é onde a apropriação
    # provavelmente está errada (lançada no período errado, incompleta, ou
    # título ainda não baixado).
    import statistics
    por_cat_mes = {}
    for c in concretagens:
        mes = mes_dmy(c.get("data_aplicacao", ""))
        if not mes:
            continue
        key = (c["categoria"], mes)
        row = por_cat_mes.setdefault(key, {
            "categoria": c["categoria"], "mes": mes, "n_concretagens": 0,
            "volume_projetado": 0.0, "volume_realizado": 0.0, "valor_pago": 0.0, "valor_a_pagar": 0.0,
        })
        row["n_concretagens"] += 1
        row["volume_projetado"] = f2(row["volume_projetado"] + f2(c.get("volume_projetado")))
        row["volume_realizado"] = f2(row["volume_realizado"] + f2(c.get("volume_realizado")))
    for t in titulos:
        mes = mes_iso(t.get("data", ""))
        if not mes:
            continue
        key = (t.get("categoria", "Outros"), mes)
        row = por_cat_mes.setdefault(key, {
            "categoria": t.get("categoria", "Outros"), "mes": mes, "n_concretagens": 0,
            "volume_projetado": 0.0, "volume_realizado": 0.0, "valor_pago": 0.0, "valor_a_pagar": 0.0,
        })
        row["valor_pago"] = f2(row["valor_pago"] + f2(t.get("valor_pago")))
        row["valor_a_pagar"] = f2(row["valor_a_pagar"] + f2(t.get("valor_a_pagar")))
        row.setdefault("titulos", []).append({
            "numtit": t.get("numtit", ""), "credor": t.get("credor", ""),
            "valor": f2(t.get("valor")), "status": t.get("status", ""), "data": t.get("data", ""),
        })
    for row in por_cat_mes.values():
        row.setdefault("titulos", [])
        row["valor_apropriado"] = f2(row["valor_pago"] + row["valor_a_pagar"])
        row["custo_medio_m3"] = f2(row["valor_apropriado"] / row["volume_realizado"]) if row["volume_realizado"] else 0.0

    amostras_por_categoria = {}
    for (categoria, mes), row in por_cat_mes.items():
        if row["volume_realizado"] > 0 and row["valor_apropriado"] > 0:
            amostras_por_categoria.setdefault(categoria, []).append(row["custo_medio_m3"])
    preco_referencia_por_categoria = {}
    for categoria, amostras in amostras_por_categoria.items():
        preco_referencia_por_categoria[categoria] = {
            "preco_m3": f2(statistics.median(amostras)), "n_amostras": len(amostras),
        }

    ajustes = []
    for (categoria, mes), row in por_cat_mes.items():
        ref = preco_referencia_por_categoria.get(categoria) or {}
        preco_ref = ref.get("preco_m3", 0.0)
        n_amostras = ref.get("n_amostras", 0)
        valor_esperado = f2(row["volume_realizado"] * preco_ref) if preco_ref else 0.0
        ajuste = f2(valor_esperado - row["valor_apropriado"]) if preco_ref else f2(-row["valor_apropriado"])
        desvio_pct = f2((row["custo_medio_m3"] - preco_ref) / preco_ref * 100) if preco_ref and row["volume_realizado"] else 0.0
        # com só 1 amostra a "mediana" é o próprio mês — não tem base de
        # comparação real (desvio sempre dá 0), então nunca marca pra revisar
        # por desvio de custo (mas o caso de apropriar sem volume continua valendo).
        confiavel = n_amostras >= 2
        sem_volume_com_valor = row["volume_realizado"] <= 0 and row["valor_apropriado"] > 0
        motivo = ""
        if sem_volume_com_valor:
            motivo = "Tem título apropriado nesse mês/etapa, mas nenhuma concretagem com esse volume registrada na Melius no mesmo período — conferir se é de um mês anterior ou se falta registrar na Melius."
        elif confiavel and row["volume_realizado"] > 0 and abs(desvio_pct) > 35:
            motivo = (f"Custo/m³ apropriado (R$ {row['custo_medio_m3']:.2f}) {'bem acima' if desvio_pct>0 else 'bem abaixo'} da "
                      f"referência histórica dessa etapa (R$ {preco_ref:.2f}/m³, baseada em {n_amostras} outros meses) — "
                      f"{'pode faltar título lançar' if desvio_pct<0 else 'conferir se não duplicou/lançou valor de outro período'}.")
        ajustes.append({
            "categoria": categoria, "mes": mes, "n_concretagens": row["n_concretagens"],
            "volume_realizado": row["volume_realizado"], "volume_projetado": row["volume_projetado"],
            "valor_apropriado": row["valor_apropriado"], "custo_medio_m3": row["custo_medio_m3"],
            "preco_referencia_m3": preco_ref, "n_amostras_referencia": n_amostras,
            "valor_esperado": valor_esperado, "ajuste_necessario": ajuste, "desvio_pct": desvio_pct,
            "revisar": bool(sem_volume_com_valor or (confiavel and row["volume_realizado"] > 0 and abs(desvio_pct) > 35)),
            "motivo": motivo,
            "titulos": sorted(row.get("titulos") or [], key=lambda x: x.get("data") or "", reverse=True),
            "n_titulos": len(row.get("titulos") or []),
        })
    ajustes.sort(key=lambda r: abs(r["ajuste_necessario"]) if r["revisar"] else -1, reverse=True)

    for row in por_categoria.values():
        ref = preco_referencia_por_categoria.get(row["categoria"]) or {}
        row["preco_referencia_m3"] = ref.get("preco_m3", 0.0)
        row["n_amostras_referencia"] = ref.get("n_amostras", 0)

    categorias = [por_categoria[c] for c in _CONCRETO_CATEGORIAS_ORDEM
                  if por_categoria[c]["n_concretagens"] or por_categoria[c]["valor_apropriado"]]

    return {
        "ok": True,
        "meses": meses,
        "categorias": categorias,
        "ajustes": ajustes,
        "concretagens": sorted(concretagens, key=lambda c: c.get("data_aplicacao", ""), reverse=True),
        "titulos": sorted(titulos, key=lambda t: t.get("data") or "", reverse=True),
        "itens_orcamento": itens_orcamento,
        "totais_orcamento": totais_orcamento,
        "totais": {
            "volume_projetado": total_volume_proj,
            "volume_realizado": total_volume_real,
            "valor_apropriado": total_apropriado,
            "custo_medio_m3": f2(total_apropriado / total_volume_real) if total_volume_real else 0.0,
        },
        "updated_at_melius": melius.get("updated_at", ""),
    }


@app.route("/api/analise-concreto")
def api_analise_concreto():
    try:
        return jsonify(_load_analise_concreto_execucao_aprop())
    except Exception as exc:
        traceback.print_exc()
        return jsonify({"ok": False, "error": str(exc)}), 500


def _run_melius_rompimentos_sync(email, password):
    def status(state, message):
        _write_melius_rompimentos_status(state, message)

    try:
        from melius_sync import run_sync
        cache_anterior = _load_melius_rompimentos_cache()
        result = run_sync(email, password, Path(DATA_DIR), status, cache_anterior=cache_anterior)
        os.makedirs(DATA_DIR, exist_ok=True)
        tmp = MELIUS_ROMPIMENTOS_CACHE_PATH + ".tmp"
        Path(tmp).write_text(json.dumps(result, ensure_ascii=False, default=str), encoding="utf-8")
        os.replace(tmp, MELIUS_ROMPIMENTOS_CACHE_PATH)
        _write_melius_rompimentos_status("done", f"{len(result.get('concretagens', []))} concretagens sincronizadas com sucesso.")
    except Exception as exc:
        _write_melius_rompimentos_status("error", str(exc))
    finally:
        with _melius_rompimentos_sync["lock"]:
            _melius_rompimentos_sync["thread"] = None


@app.route("/api/melius_rompimentos/status")
def api_melius_rompimentos_status():
    status = _read_melius_rompimentos_status()
    cache = _load_melius_rompimentos_cache()
    status["cache_updated_at"] = cache.get("updated_at", "")
    status["total"] = len(cache.get("concretagens", []))
    return jsonify(status)


@app.route("/api/melius_rompimentos/data")
def api_melius_rompimentos_data():
    return jsonify(_load_melius_rompimentos_cache())


@app.route("/api/melius_rompimentos/start", methods=["POST"])
def api_melius_rompimentos_start():
    payload = request.get_json(silent=True) or {}
    password = str(payload.get("password", "")).strip() or _load_melius_site_senha()
    if not password:
        return jsonify({"ok": False, "error": "Informe a senha de login do site da Melius."}), 400
    creds = _load_melius_credentials()
    email = str(payload.get("email", "") or creds.get("user", "")).strip()
    if not email:
        return jsonify({"ok": False, "error": "Não há um e-mail configurado para o login da Melius."}), 400
    with _melius_rompimentos_sync["lock"]:
        thread = _melius_rompimentos_sync.get("thread")
        if thread and thread.is_alive():
            return jsonify({"ok": False, "error": "Já existe uma sincronização em andamento."}), 409
        _write_melius_rompimentos_status("starting", "Preparando a coleta de rompimentos...")
        thread = threading.Thread(target=_run_melius_rompimentos_sync, args=(email, password), daemon=True)
        _melius_rompimentos_sync["thread"] = thread
        thread.start()
    return jsonify({"ok": True, "status": _read_melius_rompimentos_status()})


_melius_diario_sync = {"lock": threading.Lock(), "thread": None}


def _read_melius_diario_status():
    try:
        return _status_sem_thread_vira_erro(
            json.loads(Path(MELIUS_DIARIO_STATUS_PATH).read_text(encoding="utf-8")), _melius_diario_sync)
    except Exception:
        return {"state": "idle", "message": "", "updated_at": ""}


def _write_melius_diario_status(state, message):
    payload = {"state": state, "message": message, "updated_at": time.strftime("%Y-%m-%d %H:%M:%S")}
    os.makedirs(DATA_DIR, exist_ok=True)
    Path(MELIUS_DIARIO_STATUS_PATH).write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    return payload


def _load_melius_diario_cache():
    try:
        data = json.loads(Path(MELIUS_DIARIO_CACHE_PATH).read_text(encoding="utf-8"))
        if isinstance(data, dict):
            return data
    except Exception:
        pass
    return {"ok": False, "updated_at": "", "registros": []}


def _run_melius_diario_sync(email, password):
    def status(state, message):
        _write_melius_diario_status(state, message)

    try:
        from melius_sync import run_diario_sync
        cache_anterior = _load_melius_diario_cache()
        result = run_diario_sync(email, password, Path(DATA_DIR), status, cache_anterior=cache_anterior)
        os.makedirs(DATA_DIR, exist_ok=True)
        tmp = MELIUS_DIARIO_CACHE_PATH + ".tmp"
        Path(tmp).write_text(json.dumps(result, ensure_ascii=False, default=str), encoding="utf-8")
        os.replace(tmp, MELIUS_DIARIO_CACHE_PATH)
        _write_melius_diario_status("done", f"{len(result.get('registros', []))} registros de diário sincronizados com sucesso.")
    except Exception as exc:
        _write_melius_diario_status("error", str(exc))
    finally:
        with _melius_diario_sync["lock"]:
            _melius_diario_sync["thread"] = None


@app.route("/api/melius_diario/status")
def api_melius_diario_status():
    status = _read_melius_diario_status()
    cache = _load_melius_diario_cache()
    status["cache_updated_at"] = cache.get("updated_at", "")
    status["total"] = len(cache.get("registros", []))
    return jsonify(status)


@app.route("/api/melius_diario/data")
def api_melius_diario_data():
    return jsonify(_load_melius_diario_cache())


@app.route("/api/melius_diario/start", methods=["POST"])
def api_melius_diario_start():
    payload = request.get_json(silent=True) or {}
    password = str(payload.get("password", "")).strip() or _load_melius_site_senha()
    if not password:
        return jsonify({"ok": False, "error": "Informe a senha de login do site da Melius."}), 400
    creds = _load_melius_credentials()
    email = str(payload.get("email", "") or creds.get("user", "")).strip()
    if not email:
        return jsonify({"ok": False, "error": "Não há um e-mail configurado para o login da Melius."}), 400
    with _melius_diario_sync["lock"]:
        thread = _melius_diario_sync.get("thread")
        if thread and thread.is_alive():
            return jsonify({"ok": False, "error": "Já existe uma sincronização em andamento."}), 409
        _write_melius_diario_status("starting", "Preparando a coleta do diário de obra...")
        thread = threading.Thread(target=_run_melius_diario_sync, args=(email, password), daemon=True)
        _melius_diario_sync["thread"] = thread
        thread.start()
    return jsonify({"ok": True, "status": _read_melius_diario_status()})


MELIUS_FULL_SYNC_STATUS_PATH = os.path.join(DATA_DIR, "melius_full_sync_status.json")
_melius_full_sync = {"lock": threading.Lock(), "thread": None}


def _read_melius_full_sync_status():
    try:
        return _status_sem_thread_vira_erro(
            json.loads(Path(MELIUS_FULL_SYNC_STATUS_PATH).read_text(encoding="utf-8")), _melius_full_sync)
    except Exception:
        return {"state": "idle", "message": "", "etapa": 0, "total_etapas": 3, "updated_at": ""}


def _write_melius_full_sync_status(state, message, etapa=0):
    payload = {
        "state": state, "message": message, "etapa": etapa, "total_etapas": 3,
        "updated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
    }
    os.makedirs(DATA_DIR, exist_ok=True)
    Path(MELIUS_FULL_SYNC_STATUS_PATH).write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    return payload


def _run_melius_full_sync(email, password, chave_obra):
    """Roda as 3 fontes da Melius em sequência (mesmo Chrome seria pesado
    demais em paralelo, e evita bater na Melius com duas sessões ao mesmo
    tempo): API pública (sem senha) → robô de rompimentos → robô de diário
    de obra. Uma etapa falhando não cancela as seguintes — cada uma grava
    o próprio status/cache, igual já fazem os botões separados."""
    try:
        _write_melius_full_sync_status("running", "Etapa 1/3: atualizando dados públicos da API (FVS, serviços, rastreabilidade de concreto)...", etapa=1)
        try:
            sync_melius_data(chave_obra)
        except Exception as exc:
            print(f"  [MELIUS] etapa 1/3 (API pública) falhou: {exc}")

        from melius_sync import run_sync, run_diario_sync

        _write_melius_full_sync_status("running", "Etapa 2/3: robô entrando no site pra buscar rompimentos de concreto...", etapa=2)
        try:
            cache_anterior = _load_melius_rompimentos_cache()
            result = run_sync(email, password, Path(DATA_DIR), lambda s, m: _write_melius_full_sync_status("running", f"Etapa 2/3: {m}", etapa=2), cache_anterior=cache_anterior)
            tmp = MELIUS_ROMPIMENTOS_CACHE_PATH + ".tmp"
            Path(tmp).write_text(json.dumps(result, ensure_ascii=False, default=str), encoding="utf-8")
            os.replace(tmp, MELIUS_ROMPIMENTOS_CACHE_PATH)
            _write_melius_rompimentos_status("done", f"{len(result.get('concretagens', []))} concretagens sincronizadas com sucesso.")
        except Exception as exc:
            _write_melius_rompimentos_status("error", str(exc))
            _write_melius_full_sync_status("running", f"Etapa 2/3 falhou ({exc}) — seguindo pra etapa 3...", etapa=2)

        _write_melius_full_sync_status("running", "Etapa 3/3: robô entrando no site pra buscar o diário de obra...", etapa=3)
        try:
            cache_anterior_diario = _load_melius_diario_cache()
            result2 = run_diario_sync(email, password, Path(DATA_DIR), lambda s, m: _write_melius_full_sync_status("running", f"Etapa 3/3: {m}", etapa=3), cache_anterior=cache_anterior_diario)
            tmp2 = MELIUS_DIARIO_CACHE_PATH + ".tmp"
            Path(tmp2).write_text(json.dumps(result2, ensure_ascii=False, default=str), encoding="utf-8")
            os.replace(tmp2, MELIUS_DIARIO_CACHE_PATH)
            _write_melius_diario_status("done", f"{len(result2.get('registros', []))} registros de diário sincronizados com sucesso.")
        except Exception as exc:
            _write_melius_diario_status("error", str(exc))
            _write_melius_full_sync_status("error", f"Etapa 3/3 (diário de obra) falhou: {exc}", etapa=3)
            return

        _write_melius_full_sync_status("done", "Melius atualizada por completo: API pública + rompimentos + diário de obra.", etapa=3)
    except Exception as exc:
        _write_melius_full_sync_status("error", str(exc))
    finally:
        with _melius_full_sync["lock"]:
            _melius_full_sync["thread"] = None


@app.route("/api/melius_full_sync/status")
def api_melius_full_sync_status():
    return jsonify(_read_melius_full_sync_status())


@app.route("/api/melius_full_sync/start", methods=["POST"])
def api_melius_full_sync_start():
    payload = request.get_json(silent=True) or {}
    creds = _load_melius_credentials()
    # a senha digitada na hora tem prioridade; sem ela, cai pra MELIUS_SITE_SENHA
    # (iniciar_dashboard.bat) — senha de login do SITE, diferente da senha da
    # API (MELIUS_PASSWORD/creds["password"]) usada só pro Basic Auth.
    password = str(payload.get("password", "")).strip() or _load_melius_site_senha()
    if not password:
        return jsonify({"ok": False, "error": "Nenhuma senha configurada (MELIUS_SITE_SENHA no iniciar_dashboard.bat) e nenhuma foi informada.", "needs_password": True}), 400
    email = str(payload.get("email", "") or creds.get("user", "")).strip()
    if not email:
        return jsonify({"ok": False, "error": "Não há um e-mail configurado para o login da Melius."}), 400
    chave_obra = str(payload.get("chave_obra") or "").strip()
    with _melius_full_sync["lock"]:
        thread = _melius_full_sync.get("thread")
        if thread and thread.is_alive():
            return jsonify({"ok": False, "error": "Já existe uma atualização completa em andamento."}), 409
        _write_melius_full_sync_status("starting", "Preparando atualização completa da Melius...", etapa=0)
        thread = threading.Thread(target=_run_melius_full_sync, args=(email, password, chave_obra), daemon=True)
        _melius_full_sync["thread"] = thread
        thread.start()
    return jsonify({"ok": True, "status": _read_melius_full_sync_status()})


@app.route("/api/melius/status")
def api_melius_status():
    data = _load_melius_cache()
    creds = _load_melius_credentials()
    return jsonify({
        "ok": bool(data.get("ok")),
        "status": data.get("status", "nao_atualizado"),
        "updated_at": data.get("updated_at", ""),
        "last_attempt_at": data.get("last_attempt_at", ""),
        "error": data.get("error", ""),
        "configured": all(bool(creds.get(k)) for k in ("user", "password", "token")),
        "selected_work": data.get("selected_work", ""),
    })


@app.route("/api/melius/data")
def api_melius_data():
    return jsonify(_load_melius_cache())


@app.route("/api/melius/sync", methods=["POST"])
def api_melius_sync():
    payload = request.get_json(silent=True) or {}
    data = sync_melius_data(payload.get("chave_obra", ""))
    return jsonify(data), (200 if data.get("ok") else 502)

@app.route("/api/melius/sync_local", methods=["POST"])
def api_melius_sync_local():
    payload = request.get_json(silent=True) or {}
    servico_cod = str(payload.get("servico_cod", "")).strip()
    if not servico_cod:
        return jsonify({"ok": False, "error": "servico_cod é obrigatório."}), 400
    data = sync_melius_location(servico_cod)
    return jsonify(data), (200 if data.get("ok") else 502)

@app.route("/api/uau/obras")
def api_uau_obras():
    return jsonify(load_uau_obras())

@app.route("/api/uau/dashboard")
def api_uau_dashboard():
    empresa = request.args.get("empresa", "")
    obra = request.args.get("obra", "")
    dt_ini = request.args.get("dt_ini", "")
    dt_fim = request.args.get("dt_fim", "")
    try:
        return jsonify(load_uau_dashboard_data(empresa, obra, dt_ini, dt_fim))
    except Exception as exc:
        return jsonify({
            "ok": False,
            "error": str(exc),
            "tree": [],
            "financeiro": [],
            "summary": {},
        }), 500


@app.route("/api/export/orcamento-sql.xlsx")
def api_export_orcamento_sql():
    """Exporta em Excel a árvore de orçamento SQL do UAU (mesma fonte da aba
    'Orçamento SQL' — PlanTotal/VwDesembolso ao vivo), com os mesmos filtros
    de empresa/obra/período que estiverem selecionados na tela."""
    empresa = request.args.get("empresa", "")
    obra = request.args.get("obra", "")
    dt_ini = request.args.get("dt_ini", "")
    dt_fim = request.args.get("dt_fim", "")
    try:
        dados = load_uau_dashboard_data(empresa, obra, dt_ini, dt_fim)
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 500

    import openpyxl
    from openpyxl.styles import Font
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Orcamento SQL"
    ws.append(["Nível", "Código", "Descrição", "Orçado (R$)", "Realizado (R$)", "A Pagar (R$)", "Saldo (R$)"])
    for cell in ws[1]:
        cell.font = Font(bold=True)

    def escrever(nos, nivel):
        for n in (nos or []):
            ws.append([nivel, n.get("cod", ""), n.get("nm", ""), n.get("orcado", 0) or 0,
                       n.get("realizado", 0) or 0, n.get("a_pagar", 0) or 0, n.get("saldo", 0) or 0])
            escrever(n.get("children"), nivel + 1)

    escrever(dados.get("tree"), 1)
    for i, w in enumerate((8, 16, 55, 16, 16, 16, 16), start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws.freeze_panes = "A2"

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    nome = f"orcamento_sql_uau_{datetime.date.today():%Y%m%d}.xlsx"
    return send_file(bio, as_attachment=True, download_name=nome,
                      mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def _uau_conn(database_override=None):
    """Conexão pyodbc direta com o SQL do UAU, lendo credenciais de
    dados/uau/.env — usada pelas rotas de exploração de schema e pelo
    financeiro (portadas do antigo dashboard_uau standalone)."""
    import pyodbc
    cfg = _uau_read_env_file()
    if database_override:
        cfg = dict(cfg)
        cfg["UAU_DB_NAME"] = database_override
    return pyodbc.connect(_uau_conn_str(cfg), timeout=12)


@app.route("/api/uau/comparacao-orcamentos")
@_require_admin
def api_uau_comparacao_orcamentos():
    """Compara a planilha original de orçamento do FAR CJ03/CJ05 com a
    planilha de análise de custo mais recente — não mexe em nada das abas
    existentes, é só uma visão de comparação lado a lado."""
    return jsonify(_uau_comparar_orcamentos())


@app.route("/api/uau/cj0305")
@_require_admin
def api_uau_cj0305():
    """Árvore standalone da distribuição CJ03/CJ05 (mesma planilha usada
    pra injetar o orçado na árvore principal do UAU) — pros 3 centros de
    custo que ela alimenta."""
    mapa = _uau_orcado_cj0305_por_cc()
    if not mapa or not any(mapa.values()):
        return jsonify({"tree": [], "summary": {}, "meta": {"ok": False, "message": "Planilha CJ03/CJ05 não encontrada ou sem itens."}})
    nomes_cc = {
        "116-11602": "FAR - CONSTRUCAO CONJUNTO 3 - REGINA CELIA",
        "116-11603": "FAR - CONSTRUCAO CONJUNTO 5 - JOAQUIM RORIZ",
        "116-11611": "FAR - CUSTOS INDIRETOS 50% CONJ. 3 E CONJ. 5",
    }
    tree = []
    for cod, itens in mapa.items():
        children = [{"cod": item, "nm": item, "orcado": valor, "realizado": 0.0, "a_pagar": 0.0, "saldo": valor, "children": []}
                    for item, valor in itens.items()]
        total = f2(sum(c["orcado"] for c in children))
        tree.append({"cod": cod, "nm": nomes_cc.get(cod, cod), "orcado": total, "realizado": 0.0, "a_pagar": 0.0, "saldo": total, "children": children})
    total_geral = f2(sum(n["orcado"] for n in tree))
    return jsonify({
        "tree": tree,
        "summary": {"orcado": total_geral, "realizado": 0.0, "a_pagar": 0.0, "saldo": total_geral},
        "meta": {"ok": True, "message": "Fonte: planilha Orçamento CJ 03 CJ 05.xlsx (dados/uau).", "database": "Orçamento CJ 03 CJ 05.xlsx", "host": "planilha local"},
    })


@app.route("/api/uau/valor-agregado")
def api_uau_valor_agregado():
    """Análise de valor agregado do FAR CJ03/CJ05 (planilha 'VALOR AGREGADO...xlsx' em
    dados/uau): controle físico/financeiro, VP/CR/VA e índices, orçamento planejado x medido
    por mês (CJ03 e CJ05), fluxo de saídas do UAU e ficha CEF."""
    try:
        import uau_valor_agregado
        dados = uau_valor_agregado.carregar(UAU_DATA_DIR)
    except Exception as exc:
        return jsonify({"ok": False, "error": f"Falha ao ler a planilha de valor agregado: {exc}"}), 500
    return jsonify(dados), (200 if dados.get("ok") else 404)


# ───────────── Medição PLS (Caixa) — planilhas PLS_HIS_MED*.xlsx em dados/uau ─────────────
def _pls_conj():
    cj = str(request.values.get("conj") or (request.get_json(silent=True) or {}).get("conj") or "").lower()
    import pls_medicao
    return cj if cj in pls_medicao.CONJUNTOS else ""


@app.route("/api/pls/dados")
def api_pls_dados():
    import pls_medicao
    cj = _pls_conj()
    if not cj:
        return jsonify({"ok": False, "error": "Conjunto inválido."}), 400
    d = pls_medicao.carregar(UAU_DATA_DIR, cj)
    if d.get("ok"):
        d = dict(d)
        d["marks"] = pls_medicao.estado_do_conjunto(UAU_DATA_DIR, cj, d["medicao_atual"])["marks"]
    return jsonify(d)


@app.route("/api/pls/marcas")
def api_pls_marcas():
    import pls_medicao
    cj = _pls_conj()
    d = pls_medicao.carregar(UAU_DATA_DIR, cj) if cj else {}
    if not d.get("ok"):
        return jsonify({"ok": False, "marks": {}})
    return jsonify({"ok": True, "marks": pls_medicao.estado_do_conjunto(UAU_DATA_DIR, cj, d["medicao_atual"])["marks"]})


@app.route("/api/pls/marcar", methods=["POST"])
def api_pls_marcar():
    import pls_medicao
    payload = request.get_json(silent=True) or {}
    cj = _pls_conj()
    if not cj:
        return jsonify({"ok": False, "error": "Conjunto inválido."}), 400
    user = _current_auth_user() or {}
    try:
        linha = int(payload.get("linha"))
    except (TypeError, ValueError):
        return jsonify({"ok": False, "error": "Serviço inválido."}), 400
    cols = payload.get("colunas") or []
    if not isinstance(cols, list):
        return jsonify({"ok": False, "error": "Colunas inválidas."}), 400
    r = pls_medicao.marcar(UAU_DATA_DIR, cj, linha, [str(c) for c in cols], bool(payload.get("ligar")),
                           user.get("name") or user.get("email") or "")
    return jsonify(r), (200 if r.get("ok") else 400)


@app.route("/api/pls/limpar", methods=["POST"])
def api_pls_limpar():
    import pls_medicao
    payload = request.get_json(silent=True) or {}
    cj = _pls_conj()
    if not cj:
        return jsonify({"ok": False, "error": "Conjunto inválido."}), 400
    linha = payload.get("linha")
    return jsonify(pls_medicao.limpar(UAU_DATA_DIR, cj, int(linha) if linha not in (None, "") else None))


@app.route("/api/pls/resumo")
def api_pls_resumo():
    import pls_medicao
    return jsonify(pls_medicao.resumo(UAU_DATA_DIR))


@app.route("/api/pls/auditoria", methods=["POST"])
def api_pls_auditoria():
    """Excel de auditoria (títulos, contratos, OCs, adiantamento) montado a partir do que o painel calculou."""
    import pls_medicao
    payload = request.get_json(silent=True) or {}
    if not isinstance(payload.get("conjuntos"), list) or not payload["conjuntos"]:
        return jsonify({"ok": False, "error": "Nada para exportar."}), 400
    user = _current_auth_user() or {}
    try:
        bio, nome = pls_medicao.auditoria_xlsx(payload, user.get("name") or user.get("email") or "")
    except Exception as exc:
        return jsonify({"ok": False, "error": "Falha ao montar a planilha: %s" % exc}), 500
    return send_file(bio, as_attachment=True, download_name=nome,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/api/pls/auditoria-titulos", methods=["POST"])
def api_pls_auditoria_titulos():
    """Excel com todos os títulos de pagamento (credor, valor, datas) + conferência contratos/OCs x lançado."""
    import pls_medicao
    payload = request.get_json(silent=True) or {}
    if not isinstance(payload.get("centros"), list) or not payload["centros"]:
        return jsonify({"ok": False, "error": "Nada para exportar."}), 400
    user = _current_auth_user() or {}
    try:
        bio, nome = pls_medicao.auditoria_titulos_xlsx(payload, user.get("name") or user.get("email") or "")
    except Exception as exc:
        return jsonify({"ok": False, "error": "Falha ao montar a planilha: %s" % exc}), 500
    return send_file(bio, as_attachment=True, download_name=nome,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/api/pls/contratar")
def api_pls_contratar():
    import pls_medicao
    return jsonify({"ok": True, "itens": pls_medicao.contratar_ler(UAU_DATA_DIR)})


@app.route("/api/pls/contratar/salvar", methods=["POST"])
def api_pls_contratar_salvar():
    import pls_medicao
    p = request.get_json(silent=True) or {}
    user = _current_auth_user() or {}
    r = pls_medicao.contratar_salvar(UAU_DATA_DIR, p.get("chave"), p.get("valor"), p.get("obs"),
                                     user.get("name") or user.get("email") or "")
    return jsonify(r), (200 if r.get("ok") else 400)


@app.route("/api/pls/depara")
def api_pls_depara():
    import pls_medicao
    return jsonify(pls_medicao.depara_tabela(UAU_DATA_DIR))


@app.route("/api/pls/depara/salvar", methods=["POST"])
def api_pls_depara_salvar():
    import pls_medicao
    p = request.get_json(silent=True) or {}
    item = str(p.get("item") or "").strip()
    if not item:
        return jsonify({"ok": False, "error": "Item inválido."}), 400
    cod = p.get("cod", None)
    r = pls_medicao.depara_salvar(UAU_DATA_DIR, item, None if cod is None else str(cod))
    return jsonify(r), (200 if r.get("ok") else 400)


@app.route("/api/pls/depara/restaurar", methods=["POST"])
def api_pls_depara_restaurar():
    import pls_medicao
    return jsonify(pls_medicao.depara_restaurar(UAU_DATA_DIR))


@app.route("/api/pls/fisico-financeiro")
def api_pls_fisico_financeiro():
    import pls_medicao
    return jsonify(pls_medicao.fisico_financeiro(UAU_DATA_DIR))


@app.route("/api/pls/relatorio")
def api_pls_relatorio():
    """Relatório da medição PLS em PDF ou Excel — um conjunto ou os dois."""
    import pls_medicao
    cj = str(request.args.get("conj") or "").lower()
    conjuntos = list(pls_medicao.CONJUNTOS) if cj in ("", "todos") else ([cj] if cj in pls_medicao.CONJUNTOS else [])
    if not conjuntos:
        return jsonify({"ok": False, "error": "Conjunto inválido."}), 400
    formato = str(request.args.get("formato") or "pdf").lower()
    user = _current_auth_user() or {}
    quem = user.get("name") or user.get("email") or ""
    if formato == "xlsx":
        bio, nome = pls_medicao.relatorio_xlsx(UAU_DATA_DIR, conjuntos, quem)
        mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    else:
        bio, nome = pls_medicao.relatorio_pdf(UAU_DATA_DIR, conjuntos, quem)
        mime = "application/pdf"
    if bio is None:
        return jsonify({"ok": False, "error": nome}), 400
    return send_file(bio, as_attachment=True, download_name=nome, mimetype=mime)


@app.route("/api/pls/export")
def api_pls_export():
    import pls_medicao
    cj = _pls_conj()
    if not cj:
        return jsonify({"ok": False, "error": "Conjunto inválido."}), 400
    bio, nome = pls_medicao.exportar_xlsx(UAU_DATA_DIR, cj)
    if bio is None:
        return jsonify({"ok": False, "error": nome}), 400
    return send_file(bio, as_attachment=True, download_name=nome,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/api/uau/financeiro")
@_require_admin
def api_uau_financeiro():
    """Tabela plana de títulos do UAU (VwDesembolso) — porta de
    load_financeiro_dataset do antigo dashboard_uau standalone."""
    empresa = request.args.get("empresa", "")
    obra_raw = request.args.get("obra", "")
    dt_ini = re.sub(r"\D", "", request.args.get("dt_ini", "")) or "20000101"
    dt_fim = re.sub(r"\D", "", request.args.get("dt_fim", "")) or "21001231"
    # obra pode vir com varios centros de custo separados por virgula, cada um
    # no formato "empresa|obra" (mesmo parse usado em load_uau_dashboard_data)
    obras = []
    for item in str(obra_raw or "").split(","):
        item = item.strip()
        if not item:
            continue
        obras.append(item.split("|", 1)[1].strip() if "|" in item else item)
    where = ["d.DtPgto_Des >= CONVERT(datetime, ?, 112)", "d.DtPgto_Des < DATEADD(day, 1, CONVERT(datetime, ?, 112))"]
    params = [dt_ini, dt_fim]
    if empresa:
        where.append("d.Empresa_Des = ?")
        params.append(empresa)
    if obras:
        placeholders = ",".join(["?"] * len(obras))
        where.append(f"d.Obra_Des IN ({placeholders})")
        params.extend(obras)
    sql = f"""
        SELECT TOP (2000)
            d.Empresa_Des AS Empresa, d.Obra_Des AS Obra,
            CAST(d.NumProc_Des AS varchar(30)) + '/' + CAST(d.NumParc_Des AS varchar(20)) AS Titulo,
            CAST(d.DocFiscal_Des AS varchar(80)) AS Documento,
            d.DtPgto_Des AS DataMovimento,
            CASE WHEN ISNULL(d.StatusParc_Des, 0) = 2 THEN 'Pago' ELSE 'A pagar' END AS StatusTitulo,
            d.NominalProc_Des AS Credor, d.ItemPl_Des AS EAP,
            d.DescCompPl_Des AS Servico, d.TotalLiq_Des AS ValorTitulo,
            CASE WHEN ISNULL(d.StatusParc_Des, 0) = 2 THEN d.TotalLiq_Des ELSE 0 END AS ValorPago,
            CASE WHEN ISNULL(d.StatusParc_Des, 0) <> 2 THEN d.TotalLiq_Des ELSE 0 END AS ValorAPagar
        FROM VwDesembolso d
        INNER JOIN ObrUsr ON d.Empresa_Des = Emp_uo AND d.Obra_Des = Obr_uo AND Usr_uo = 'FCRUZ'
        WHERE {" AND ".join(where)}
        ORDER BY d.DtPgto_Des DESC
    """
    try:
        with _uau_conn() as conn:
            cur = conn.cursor()
            cur.execute("SET NOCOUNT ON;\n" + sql, *params)
            rows = _uau_rows_to_dicts(cur)
    except Exception as exc:
        return jsonify({"rows": [], "columns": [], "summary": {}, "meta": {"ok": False, "message": str(exc)}}), 500
    rows = [{k: _uau_json_safe(v) for k, v in r.items()} for r in rows]
    columns = ["Empresa", "Obra", "Titulo", "Documento", "DataMovimento", "StatusTitulo", "Credor", "EAP", "Servico", "ValorTitulo", "ValorPago", "ValorAPagar"]
    return jsonify({
        "rows": rows, "columns": columns,
        "summary": {
            "linhas": len(rows),
            "valor_total": f2(sum(f2(r.get("ValorTitulo")) for r in rows)),
            "valor_pago": f2(sum(f2(r.get("ValorPago")) for r in rows)),
            "valor_a_pagar": f2(sum(f2(r.get("ValorAPagar")) for r in rows)),
        },
        "meta": {"ok": True, "message": f"{len(rows)} títulos carregados."},
    })


@app.route("/api/uau/databases")
@_require_admin
def api_uau_databases():
    try:
        with _uau_conn(database_override="master") as conn:
            cur = conn.cursor()
            cur.execute("SELECT name FROM sys.databases ORDER BY name")
            return jsonify([r["name"] for r in _uau_rows_to_dicts(cur)])
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 500


@app.route("/api/uau/tables")
@_require_admin
def api_uau_tables():
    q = request.args.get("q", "").strip()
    sql = """
        SELECT TABLE_SCHEMA, TABLE_NAME, TABLE_TYPE
        FROM INFORMATION_SCHEMA.TABLES
        WHERE (? = '' OR TABLE_NAME LIKE ? OR TABLE_SCHEMA LIKE ?)
        ORDER BY TABLE_SCHEMA, TABLE_NAME
    """
    like = f"%{q}%"
    try:
        with _uau_conn() as conn:
            cur = conn.cursor()
            cur.execute(sql, q, like, like)
            return jsonify(_uau_rows_to_dicts(cur))
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 500


@app.route("/api/uau/columns")
@_require_admin
def api_uau_columns():
    table = request.args.get("table", "")
    schema = request.args.get("schema", "dbo")
    sql = """
        SELECT COLUMN_NAME, DATA_TYPE
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_SCHEMA = ? AND TABLE_NAME = ?
        ORDER BY ORDINAL_POSITION
    """
    try:
        with _uau_conn() as conn:
            cur = conn.cursor()
            cur.execute(sql, schema, table)
            return jsonify(_uau_rows_to_dicts(cur))
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 500


@app.route("/api/protestos", methods=["GET", "POST"])
def api_protestos():
    if request.method == "GET":
        return jsonify({"ok": True, **load_protestos_state()})
    payload = request.get_json(silent=True) or {}
    saved = save_protestos_state(payload)
    data = _cache.get("data") or load_dashboard_cache()
    if isinstance(data, dict):
        data["protestos"] = saved
        save_dashboard_cache(data)
    return jsonify({"ok": True, **saved})

@app.route("/api/protestos/consulta-cnpj", methods=["POST"])
def api_protestos_consulta_cnpj():
    payload = request.get_json(silent=True) or {}
    state = load_protestos_state()
    cnpj = re.sub(r"\D", "", str(payload.get("cnpj") or state.get("cnpj") or PROTESTOS_DEFAULT_CNPJ))
    consulta = consulta_protesto_cnpj(cnpj)
    cartorios = consulta.get("cartorios") if isinstance(consulta.get("cartorios"), list) else []
    _tem_titulos_cadastrados = any(
        isinstance(it, dict) and it.get("origem") != "cenprot"
        for it in (state.get("items") or [])
    )
    if consulta.get("status") == "autenticacao_necessaria" and not _tem_titulos_cadastrados and not cartorios:
        # a mensagem padrão pressupõe que já existe um título cadastrado
        # aguardando confirmação do cartório — sem nenhum título no registro
        # ainda, isso é enganoso (parece que algo foi encontrado quando não
        # há nada). Sem token da Direct Data, a CENPROT pública exige
        # captcha e não pode ser contornada por automação.
        consulta["message"] = (
            "A CENPROT/Resolve exige autenticação ou hCaptcha — a consulta automática não retorna "
            "nada sem um título já cadastrado aqui pra cruzar. Adicione manualmente os títulos que "
            "você já sabe que foram protestados (botão \"+ Adicionar título\"), ou configure "
            "DIRECTD_API_TOKEN no iniciar_dashboard.bat pra consulta nacional automática sem captcha."
        )

    def _titulo_em_cartorio(titulo, cartorio):
        base = re.sub(r"\D", "", str(titulo or ""))
        if not base:
            return False
        for raw in cartorio.get("titulos") or []:
            if isinstance(raw, dict):
                valores = " ".join(str(v) for v in raw.values())
            else:
                valores = str(raw)
            if base and base in re.sub(r"\D", "", valores):
                return True
        return False

    def _apply_cartorio(item, cartorio, motivo):
        nome = str(cartorio.get("nome") or "").strip()
        if nome:
            item["cartorio"] = nome
        cidade = str(cartorio.get("cidade") or "").strip()
        if cidade:
            item["cidade"] = cidade
        uf = str(cartorio.get("uf") or "").strip().upper()[:2]
        if uf:
            item["uf"] = uf
        item["consulta_em"] = executed_today
        detalhe = "Cartório retornado pela CENPROT"
        if motivo:
            detalhe += f" ({motivo})"
        if cartorio.get("numero_cartorio"):
            detalhe += f" - Nº {cartorio.get('numero_cartorio')}"
        obs = str(item.get("observacao") or "")
        if detalhe not in obs:
            item["observacao"] = (obs + " " + detalhe + ".").strip()

    executed_today = time.strftime("%Y-%m-%d")
    state["cnpj"] = cnpj
    state["consulta"] = consulta

    # títulos cadastrados manualmente (origem != 'cenprot') seguem no registro,
    # só recebem o cartório se a consulta ao vivo localizar o título deles.
    manual_items = [it for it in (state.get("items", []) or []) if isinstance(it, dict) and it.get("origem") != "cenprot"]
    for item in manual_items:
        item["cnpj"] = cnpj
        item["consulta_em"] = executed_today
        item.setdefault("observacao", "")
        matched = next((c for c in cartorios if _titulo_em_cartorio(item.get("titulo"), c)), None)
        if matched:
            _apply_cartorio(item, matched, "título localizado no retorno")
        elif len(cartorios) == 1:
            _apply_cartorio(item, cartorios[0], "único cartório retornado para o CNPJ")
        elif consulta["status"] in {"bloqueado", "erro", "autenticacao_necessaria", "sem_retorno"} and "CENPROT" not in item.get("observacao", ""):
            motivo = "consulta CENPROT exige autenticação/captcha" if consulta["status"] == "autenticacao_necessaria" else "CENPROT não retornou cartório para este CNPJ"
            item["observacao"] = (item.get("observacao", "") + f" Cartório não retornado automaticamente: {motivo}.").strip()

    # descoberta ao vivo: cada consulta substitui os títulos descobertos anteriormente
    # pela CENPROT — nunca reaproveita um resultado antigo salvo em disco.
    discovered = []
    for cartorio in cartorios:
        for t in cartorio.get("titulos") or []:
            numero = str(t.get("numero", "")).strip()
            if not numero:
                continue
            obs = "Descoberto automaticamente na consulta CENPROT"
            if t.get("data"):
                obs += f" · protesto em {t.get('data')}"
            discovered.append({
                "titulo": numero,
                "cnpj": cnpj,
                "status": "protestado",
                "cartorio": cartorio.get("nome", ""),
                "cidade": cartorio.get("cidade", ""),
                "uf": cartorio.get("uf", ""),
                "consulta_em": executed_today,
                "observacao": obs + ".",
                "valor_cenprot": t.get("valor", 0),
                "origem": "cenprot",
            })

    state["items"] = manual_items + discovered
    saved = save_protestos_state(state)
    data = _cache.get("data") or load_dashboard_cache()
    if isinstance(data, dict):
        data["protestos"] = saved
        save_dashboard_cache(data)
    return jsonify({"ok": True, **saved})

@app.route("/api/refresh", methods=["POST"])
def api_refresh():
    global _cache
    started = time.time()
    full = str(request.args.get("full", "")).lower() in {"1", "true", "yes"}
    force = str(request.args.get("force", "")).lower() in {"1", "true", "yes"}
    changed_files = dashboard_changed_sources()
    if not force and not full and not changed_files:
        data = _cache.get("data") or load_dashboard_cache()
        if data:
            _cache = {"data": data, "ts": time.time()}
            return jsonify({
                "ok": True,
                "changed": False,
                "changed_files": [],
                "excel_updated_at": data.get("excel_updated_at", ""),
                "cnpj_updated": 0,
                "elapsed_seconds": round(time.time() - started, 3),
                "message": "Os arquivos já estão atualizados.",
            })

    _cache = {"data": None, "ts": 0}
    data, err = build_data(force=True)
    if err:
        return jsonify({"ok": False, "error": err}), 500
    cnpj_updated = 0
    if full:
        cnpjs = [r.get("cnpj", "") for r in data.get("contratos", [])]
        _, cnpj_updated = refresh_cnpj_cache(cnpjs)
        if cnpj_updated:
            _cache = {"data": None, "ts": 0}
            data, err = build_data(force=True)
    return jsonify({
        "ok": not bool(err),
        "error": err or "",
        "changed": True,
        "changed_files": changed_files,
        "excel_updated_at": data.get("excel_updated_at", "") if data else "",
        "cnpj_updated": cnpj_updated,
        "elapsed_seconds": round(time.time() - started, 3),
        "message": "Arquivos locais atualizados.",
    }), 200 if not err else 500


def _export_report_view(tree, mode):
    """Replica no Excel as visoes gerenciais do Orçamento Relatório."""
    original = json.loads(json.dumps(tree or []))
    moved = []

    def money(value):
        return f2(value or 0)

    def strip(nodes):
        for node in nodes or []:
            children = node.get("children") or []
            if children:
                strip(children)
                node["realizado"] = money(sum(ch.get("realizado", 0) for ch in children))
                node["a_pagar"] = money(sum(ch.get("a_pagar", 0) for ch in children))
                node["saldo"] = money(node.get("orcado", 0) - node["realizado"] - node["a_pagar"])
                continue
            titles = node.get("titulos") or []
            transferring = [
                title for title in titles
                if title.get("incorporacao_reap") and str(title.get("cc_real", "")) != "1011"
            ]
            for title in transferring:
                moved.append({
                    **title,
                    "reap_origem_eap": node.get("cod", ""),
                    "reap_origem_item": node.get("nm", ""),
                })
            node["titulos"] = [title for title in titles if title not in transferring]
            node["realizado"] = money(node.get("realizado", 0) - sum(t.get("vlr_pg", 0) for t in transferring))
            node["a_pagar"] = money(node.get("a_pagar", 0) - sum(t.get("vlr_a_pagar", 0) for t in transferring))
            node["saldo"] = money(node.get("orcado", 0) - node["realizado"] - node["a_pagar"])

    strip(original)
    groups = {}
    for title in moved:
        creditor = str(title.get("credor") or "Credor não informado")
        groups.setdefault(creditor, []).append(title)
    leaves = []
    for index, (creditor, titles) in enumerate(sorted(groups.items()), start=1):
        paid = money(sum(t.get("vlr_pg", 0) for t in titles))
        payable = money(sum(t.get("vlr_a_pagar", 0) for t in titles))
        leaves.append({
            "cod": f"INC.REAP.001.{index:03d}",
            "nm": creditor,
            "orcado": 0,
            "realizado": paid,
            "a_pagar": payable,
            "saldo": money(-paid - payable),
            "titulos": titles,
        })
    paid = money(sum(node["realizado"] for node in leaves))
    payable = money(sum(node["a_pagar"] for node in leaves))
    virtual = {
        "cod": "INC",
        "nm": "INCORPORAÇÃO - REAPROPRIAÇÃO",
        "orcado": 0,
        "realizado": paid,
        "a_pagar": payable,
        "saldo": money(-paid - payable),
        "children": [{
            "cod": "INC.REAP",
            "nm": "Credores classificados como incorporação",
            "orcado": 0, "realizado": paid, "a_pagar": payable, "saldo": money(-paid - payable),
            "children": [{
                "cod": "INC.REAP.001",
                "nm": "Títulos transferidos do custo da obra",
                "orcado": 0, "realizado": paid, "a_pagar": payable, "saldo": money(-paid - payable),
                "children": leaves,
            }],
        }],
    }
    if mode == "real":
        return json.loads(json.dumps(tree or []))
    if mode == "inc_reap":
        return [virtual] if moved else []
    if mode == "obra_reap":
        return original
    return original + ([virtual] if moved else [])


def _export_period_tree(tree, start_date="", end_date=""):
    if not start_date and not end_date:
        return tree

    def date_in_period(value):
        text = str(value or "")
        if not text or text in ("—", "nan", "None"):
            return False
        try:
            key = pd.to_datetime(text, dayfirst=True, errors="coerce")
            if pd.isna(key):
                return False
            iso = key.strftime("%Y-%m-%d")
            return (not start_date or iso >= start_date) and (not end_date or iso <= end_date)
        except Exception:
            return False

    cloned = json.loads(json.dumps(tree or []))

    def roll(node):
        children = node.get("children") or []
        if children:
            for child in children:
                roll(child)
            node["realizado"] = f2(sum(ch.get("realizado", 0) for ch in children))
            node["a_pagar"] = f2(sum(ch.get("a_pagar", 0) for ch in children))
        else:
            paid = [
                title for title in node.get("titulos", [])
                if str(title.get("status", "")).upper() != "A PAGAR" and date_in_period(title.get("dt_baixa"))
            ]
            payable = [
                title for title in node.get("titulos", [])
                if str(title.get("status", "")).upper() == "A PAGAR"
                and date_in_period(title.get("dt_venc") or title.get("dt_baixa"))
            ]
            node["titulos"] = paid + payable
            node["realizado"] = f2(sum(title.get("vlr_pg", 0) for title in paid))
            node["a_pagar"] = f2(sum(title.get("vlr_a_pagar", title.get("vlr_ap", 0)) for title in payable))
        node["saldo"] = f2(node.get("orcado", 0) - node.get("realizado", 0) - node.get("a_pagar", 0))

    for root in cloned:
        roll(root)
    return cloned


@app.route("/api/export/orcamento-relatorio.xlsx")
def api_export_orcamento_relatorio():
    data, err = build_data()
    if err:
        return jsonify({"ok": False, "error": err}), 500
    mode = request.args.get("mode", "reap_all")
    if mode not in {"reap_all", "obra_reap", "inc_reap", "real"}:
        mode = "reap_all"
    start_date = request.args.get("start", "")
    end_date = request.args.get("end", "")
    tree = _export_period_tree(_export_report_view(data.get("orc_rel_tree", []), mode), start_date, end_date)

    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Orçamento Relatório"
    labels = {
        "reap_all": "Gerencial - Obra + Incorporação",
        "obra_reap": "Obra - sem custos da incorporação",
        "inc_reap": "Somente Incorporação - reapropriação",
        "real": "Custo real - como lançado",
    }
    root_totals = {
        "orcado": f2(sum(float(n.get("orcado", 0) or 0) for n in tree)),
        "realizado": f2(sum(float(n.get("realizado", 0) or 0) for n in tree)),
        "a_pagar": f2(sum(float(n.get("a_pagar", 0) or 0) for n in tree)),
        "saldo": f2(sum(float(n.get("saldo", 0) or 0) for n in tree)),
    }

    ws.append(["ORÇAMENTO RELATÓRIO - SPE T2 BUENO"])
    ws.append(["Visão selecionada", labels[mode]])
    ws.append(["Período", f"{start_date or 'início'} a {end_date or 'fim'}"])
    ws.append([
        "Total igual ao dashboard", "", "", "", "", "", "", "", "", "",
        root_totals["orcado"], "", root_totals["realizado"], root_totals["a_pagar"],
        root_totals["saldo"], "", "", "", "",
    ])
    ws.append([
        "Conferência", "Some apenas linhas com Tipo = EAP e Nível = 1. "
        "Os demais níveis são subtotais hierárquicos e não devem ser somados junto com o nível 1.",
    ])
    ws.append([])
    headers = [
        "Tipo", "Nível", "Código EAP", "Descrição / Credor", "Título", "NF(s)", "Pedido(s)",
        "Contrato(s)", "Status", "CC real", "Orçado", "Apropriado", "Pago", "A pagar",
        "Saldo", "Data baixa", "Vencimento", "Observação", "Origem reapropriação",
    ]
    ws.append(headers)
    header_row = ws.max_row
    header_fill = PatternFill("solid", fgColor="A50E24")
    thin = Side(style="thin", color="D7DEE8")
    for cell in ws[header_row]:
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(bottom=thin)
    for cell in ws[4]:
        cell.fill = PatternFill("solid", fgColor="F4E8EA")
        cell.font = Font(bold=True, color="23344D")
        cell.border = Border(bottom=thin)
    ws["A5"].font = Font(bold=True, color="A50E24")
    ws["B5"].font = Font(color="4B5D75")
    ws.freeze_panes = f"A{header_row + 1}"

    level_fills = ["F4E8EA", "EAF0F6", "F4F7FA", "FFFFFF"]
    currency_columns = range(11, 16)

    def append_node(node, level):
        row = ws.max_row + 1
        ws.append([
            "EAP", level, node.get("cod", ""), node.get("nm", ""), "", "", "", "", "", "",
            node.get("orcado", 0), "", node.get("realizado", 0), node.get("a_pagar", 0), node.get("saldo", 0),
            "", "", "", "",
        ])
        for cell in ws[row]:
            cell.fill = PatternFill("solid", fgColor=level_fills[min(level - 1, 3)])
            cell.font = Font(bold=level <= 3)
            cell.border = Border(bottom=thin)
        ws.row_dimensions[row].outlineLevel = max(0, level - 1)
        for title in node.get("titulos", []) or []:
            title_row = ws.max_row + 1
            origin = ""
            if title.get("reap_origem_eap"):
                origin = f"{title.get('reap_origem_eap')} - {title.get('reap_origem_item', '')}"
            ws.append([
                "TÍTULO", level + 1, node.get("cod", ""), title.get("credor", ""), title.get("numtit", ""),
                ", ".join(map(str, title.get("nfs", []) or [])),
                ", ".join(map(str, title.get("pedidos", []) or [])),
                ", ".join(map(str, title.get("contratos", []) or [])),
                title.get("status", ""), title.get("cc_real", ""),
                "", title.get("vlr_ap", 0), title.get("vlr_pg", 0), title.get("vlr_a_pagar", 0), "",
                title.get("dt_baixa", ""), title.get("dt_venc", ""), title.get("observ", ""), origin,
            ])
            ws.row_dimensions[title_row].outlineLevel = min(7, level)
            for cell in ws[title_row]:
                cell.border = Border(bottom=Side(style="hair", color="E7EBF0"))
        for child in node.get("children", []) or []:
            append_node(child, level + 1)

    for root in tree:
        append_node(root, 1)
    ws.auto_filter.ref = f"A{header_row}:S{ws.max_row}"
    for col in currency_columns:
        for row in range(header_row + 1, ws.max_row + 1):
            ws.cell(row, col).number_format = 'R$ #,##0.00;[Red]-R$ #,##0.00'
    widths = [12, 8, 22, 48, 16, 18, 18, 18, 14, 12, 16, 16, 16, 16, 16, 14, 14, 60, 48]
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    ws.sheet_properties.outlinePr.summaryBelow = False
    ws.sheet_view.showGridLines = False
    ws["A1"].font = Font(size=16, bold=True, color="A50E24")
    ws.merge_cells("A1:S1")

    resumo = wb.create_sheet("Resumo", 0)
    resumo.append(["ORÇAMENTO RELATÓRIO - RESUMO DE CONFERÊNCIA"])
    resumo.append(["Visão selecionada", labels[mode]])
    resumo.append(["Período", f"{start_date or 'início'} a {end_date or 'fim'}"])
    resumo.append([])
    resumo.append(["Total exibido no dashboard", "Valor"])
    resumo.append(["Orçado", root_totals["orcado"]])
    resumo.append(["Realizado / Pago", root_totals["realizado"]])
    resumo.append(["A pagar", root_totals["a_pagar"]])
    resumo.append(["Saldo", root_totals["saldo"]])
    resumo.append([])
    resumo.append(["Como conferir", "Na aba Orçamento Relatório, filtre Tipo = EAP e Nível = 1. Não some níveis 2, 3, 4 ou linhas TÍTULO junto com o nível 1, porque eles detalham a mesma hierarquia."])
    resumo.append(["Desconsiderado nos cálculos", "Documentos PPC, PCT e documentos/observações com PREVIS."])
    resumo.append(["Realizado", "Só entra quando o título aparece no Relatório de Contas Pagas local."])
    resumo.append(["A pagar", "Entra pelo Relatório de Contas a Pagar local quando o título não está pago."])
    resumo.append(["Reapropriação", "Na visão gerencial, credores de incorporação e seus impostos/retenções são movidos para INC - INCORPORAÇÃO - REAPROPRIAÇÃO."])
    resumo.sheet_view.showGridLines = False
    resumo["A1"].font = Font(size=15, bold=True, color="A50E24")
    resumo.merge_cells("A1:D1")
    for cell in resumo[5]:
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center")
    for row in range(6, 10):
        resumo.cell(row, 2).number_format = 'R$ #,##0.00;[Red]-R$ #,##0.00'
    resumo.column_dimensions["A"].width = 28
    resumo.column_dimensions["B"].width = 120

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    filename = f"Orcamento_Relatorio_{mode}_{time.strftime('%Y%m%d_%H%M')}.xlsx"
    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

@app.route("/api/sienge_sync/status")
def api_sienge_sync_status():
    status = _read_sienge_sync_status()
    if status.get("state") == "awaiting_code":
        try:
            started = datetime.strptime(str(status.get("started_at") or ""), "%Y-%m-%dT%H:%M:%S")
            if (datetime.now() - started).total_seconds() > 300:
                status = _write_sienge_sync_status(
                    "idle",
                    "Código expirado. Informe e-mail e senha novamente para solicitar um novo código.",
                    {"error": "", "details": "", "finished_at": time.strftime("%Y-%m-%dT%H:%M:%S")},
                )
        except Exception:
            pass
    return jsonify(status)

@app.route("/api/sienge_sync/reset", methods=["POST"])
def api_sienge_sync_reset():
    _sienge_sync["code"] = ""
    _sienge_sync["code_event"].set()
    _sienge_sync["confirmation_event"].set()
    _write_sienge_sync_status(
        "idle",
        "Login reiniciado. Informe e-mail e senha para solicitar um novo código.",
        {"error": "", "details": "", "files": [], "finished_at": time.strftime("%Y-%m-%dT%H:%M:%S")},
    )
    return jsonify({"ok": True})

@app.route("/api/sienge_sync/start", methods=["POST"])
def api_sienge_sync_start():
    payload = request.get_json(silent=True) or {}
    email = str(payload.get("email", "") or "").strip()
    password = str(payload.get("password", "") or "")
    if _read_sienge_sync_status().get("needs_password") and not password:
        return jsonify({
            "ok": False,
            "error": "Digite a senha correta do Sienge ID para recomeçar o processo.",
        }), 400
    if not email:
        return jsonify({"ok": False, "error": "Digite o email do Sienge ID para atualizar os relatórios."}), 400
    with _sienge_sync["lock"]:
        thread = _sienge_sync.get("thread")
        if thread and thread.is_alive():
            return jsonify({"ok": False, "error": "Uma atualização já está em andamento."}), 409
        thread = threading.Thread(
            target=_run_sienge_report_sync,
            args=(email, password),
            name="sienge-report-sync",
            daemon=True,
        )
        _sienge_sync["thread"] = thread
        thread.start()
    return jsonify({"ok": True, "message": "Atualização iniciada."})

@app.route("/api/sienge_sync/code", methods=["POST"])
def api_sienge_sync_code():
    code = str((request.get_json(silent=True) or {}).get("code", "") or "").strip()
    if not code:
        return jsonify({"ok": False, "error": "Informe o código de verificação."}), 400
    _sienge_sync["code"] = code
    _sienge_sync["code_event"].set()
    _write_sienge_sync_status("login", "Código recebido. Concluindo o login...")
    return jsonify({"ok": True})

@app.route("/api/sienge_sync/confirm", methods=["POST"])
def api_sienge_sync_confirm():
    confirmed = bool((request.get_json(silent=True) or {}).get("confirmed"))
    if not confirmed:
        return jsonify({"ok": False, "error": "A confirmação para prosseguir não foi informada."}), 400
    _sienge_sync["confirmed"] = True
    _sienge_sync["confirmation_event"].set()
    _write_sienge_sync_status("login", "Confirmação recebida. Abrindo a nova sessão do Sienge...")
    return jsonify({"ok": True})

@app.route("/api/rso_state", methods=["GET", "POST"])
def api_rso_state():
    if request.method == "POST":
        return jsonify(save_rso_state(request.get_json(silent=True) or {}))
    return jsonify(load_rso_state())

@app.route("/api/pa_medicao", methods=["POST"])
def api_pa_medicao():
    global _cache
    payload = request.get_json(silent=True) or {}
    task_id = str(payload.get("id", "")).strip()
    if not task_id.startswith("PA-"):
        return jsonify({"ok": False, "error": "Atividade inválida."}), 400
    pct = max(0, min(100, f2(payload.get("pct", 0))))
    medicoes = load_pa_medicoes()
    medicoes[task_id] = pct
    save_pa_medicoes(medicoes)
    data = _cache.get("data") or load_dashboard_cache()
    if data:
        for task in data.get("plan_cronograma", []):
            if task.get("id") == task_id:
                task["pct_real"] = pct
                break
        _cache = {"data": data, "ts": time.time()}
        save_dashboard_cache(data)
    return jsonify({"ok": True, "id": task_id, "pct": pct})


@app.route("/api/ppc_week")
def api_ppc_week():
    start = request.args.get("start", "")
    end = request.args.get("end", "")
    return jsonify(load_ppc_week(start, end))


def _orc_fonte(valor=None):
    """Os tiques de 'item concluído' são guardados por fonte: Synergia (padrão) ou UAU (Orçamento SQL)."""
    v = request.args.get("fonte") if valor is None else valor
    return "uau" if str(v or "").lower() == "uau" else "synergia"


def _orc_saves_da_fonte(fonte):
    return [s for s in load_orc_saves() if (s.get("fonte") or "synergia") == fonte]


@app.route("/api/orc_item_status")
def api_orc_item_status():
    saves = _orc_saves_da_fonte(_orc_fonte())
    latest = saves[-1] if saves else None
    return jsonify({"items": (latest or {}).get("items", {}), "save": _orc_save_summary(latest) if latest else None})


def _orc_save_summary(save):
    if not save:
        return None
    return {
        "id": save.get("id"),
        "label": save.get("label"),
        "created_at": save.get("created_at"),
        "created_by": save.get("created_by"),
        "total_itens": len(save.get("items", {})),
    }


@app.route("/api/orc_item_status/saves")
def api_orc_item_status_saves():
    saves = _orc_saves_da_fonte(_orc_fonte())
    out = [_orc_save_summary(s) for s in saves]
    out.sort(key=lambda r: r.get("created_at") or "", reverse=True)
    return jsonify(out)


@app.route("/api/orc_item_status/saves/<save_id>")
def api_orc_item_status_get_save(save_id):
    saves = load_orc_saves()
    match = next((s for s in saves if str(s.get("id")) == str(save_id)), None)
    if not match:
        return jsonify({"ok": False, "error": "Salvamento não encontrado."}), 404
    return jsonify({"ok": True, "save": match})


@app.route("/api/orc_item_status/save", methods=["POST"])
def api_orc_item_status_save():
    payload = request.get_json(silent=True) or {}
    items = payload.get("items")
    if not isinstance(items, dict):
        return jsonify({"ok": False, "error": "Lista de itens inválida."}), 400
    fonte = _orc_fonte(payload.get("fonte"))
    codes = {}
    for k, v in items.items():
        if not v:
            continue
        if fonte == "uau" and isinstance(v, (int, float)) and not isinstance(v, bool):
            codes[str(k)] = max(0.0, min(100.0, round(float(v), 2)))  # % medido do item
        else:
            codes[str(k)] = True
    user = _current_auth_user()
    entry = {
        "id": time.strftime("%Y%m%d%H%M%S"),
        "label": str(payload.get("label") or "").strip() or time.strftime("%d/%m/%Y %H:%M"),
        "created_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "created_by": (user or {}).get("name") or (user or {}).get("email") or "",
        "fonte": _orc_fonte(payload.get("fonte")),
        "items": codes,
    }
    with _orc_item_status_lock:
        saves = load_orc_saves()
        saves.append(entry)
        save_orc_saves(saves)
    return jsonify({"ok": True, "save": _orc_save_summary(entry)})


@app.route("/api/orc_item_status/saves/<save_id>", methods=["DELETE"])
def api_orc_item_status_delete_save(save_id):
    with _orc_item_status_lock:
        saves = load_orc_saves()
        remaining = [s for s in saves if str(s.get("id")) != str(save_id)]
        if len(remaining) == len(saves):
            return jsonify({"ok": False, "error": "Salvamento não encontrado."}), 404
        save_orc_saves(remaining)
    return jsonify({"ok": True})


def _ifc_medicao_summary(save):
    if not save:
        return None
    return {
        "id": save.get("id"),
        "data": save.get("data"),
        "created_at": save.get("created_at"),
        "created_by": save.get("created_by"),
        "model_name": save.get("model_name"),
        "total_elementos": save.get("total_elementos"),
        "medidos": save.get("medidos"),
        "parciais": save.get("parciais"),
        "pct_avanco": save.get("pct_avanco"),
        "volume_total": save.get("volume_total"),
        "volume_medido": save.get("volume_medido"),
    }


@app.route("/api/ifc_medicao/saves")
def api_ifc_medicao_saves():
    saves = load_ifc_medicao_saves()
    out = [_ifc_medicao_summary(s) for s in saves]
    out.sort(key=lambda r: r.get("data") or r.get("created_at") or "", reverse=True)
    return jsonify(out)


@app.route("/api/ifc_medicao/saves/<save_id>")
def api_ifc_medicao_get_save(save_id):
    saves = load_ifc_medicao_saves()
    match = next((s for s in saves if str(s.get("id")) == str(save_id)), None)
    if not match:
        return jsonify({"ok": False, "error": "Medição salva não encontrada."}), 404
    return jsonify({"ok": True, "save": match})


@app.route("/api/ifc_medicao/save", methods=["POST"])
def api_ifc_medicao_save():
    payload = request.get_json(silent=True) or {}
    items = payload.get("items")
    if not isinstance(items, dict):
        return jsonify({"ok": False, "error": "Lista de itens inválida."}), 400
    data_ref = str(payload.get("data") or "").strip() or time.strftime("%Y-%m-%d")
    user = _current_auth_user()
    entry = {
        "id": time.strftime("%Y%m%d%H%M%S"),
        "data": data_ref,
        "created_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "created_by": (user or {}).get("name") or (user or {}).get("email") or "",
        "model_name": str(payload.get("model_name") or ""),
        "total_elementos": int(payload.get("total_elementos") or 0),
        "medidos": int(payload.get("medidos") or 0),
        "parciais": int(payload.get("parciais") or 0),
        "pct_avanco": float(payload.get("pct_avanco") or 0),
        "volume_total": float(payload.get("volume_total") or 0),
        "volume_medido": float(payload.get("volume_medido") or 0),
        "items": items,
    }
    with _ifc_medicao_lock:
        saves = load_ifc_medicao_saves()
        saves.append(entry)
        save_ifc_medicao_saves(saves)
    return jsonify({"ok": True, "save": _ifc_medicao_summary(entry)})


@app.route("/api/ifc_medicao/saves/<save_id>", methods=["DELETE"])
def api_ifc_medicao_delete_save(save_id):
    with _ifc_medicao_lock:
        saves = load_ifc_medicao_saves()
        remaining = [s for s in saves if str(s.get("id")) != str(save_id)]
        if len(remaining) == len(saves):
            return jsonify({"ok": False, "error": "Medição salva não encontrada."}), 404
        save_ifc_medicao_saves(remaining)
    return jsonify({"ok": True})


_ATA_FILENAME_RE = re.compile(r"AR-?0*(\d+).*?(\d{4})\.(\d{2})\.(\d{2})", re.IGNORECASE)


def _parse_ata_filename(fname):
    m = _ATA_FILENAME_RE.search(fname)
    if not m:
        return {"numero": "", "data": ""}
    numero = "AR-" + m.group(1).zfill(2)
    data = f"{m.group(2)}-{m.group(3)}-{m.group(4)}"
    return {"numero": numero, "data": data}


@app.route("/api/atas/list")
def api_atas_list():
    try:
        if not os.path.isdir(ATAS_DIR):
            return jsonify({"ok": True, "atas": []})
        rows = []
        for fname in os.listdir(ATAS_DIR):
            if not fname.lower().endswith(".pdf"):
                continue
            full = os.path.join(ATAS_DIR, fname)
            meta = _parse_ata_filename(fname)
            rows.append({
                "arquivo": fname,
                "numero": meta["numero"],
                "data": meta["data"],
                "tamanho_kb": round(os.path.getsize(full) / 1024, 1),
            })
        rows.sort(key=lambda r: (r["data"], r["numero"]), reverse=True)
        return jsonify({"ok": True, "atas": rows})
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 500


@app.route("/api/atas/file/<path:filename>")
def api_atas_file(filename):
    # So permite servir arquivos que realmente existem na listagem da pasta —
    # evita path traversal a partir do nome de arquivo vindo da URL.
    safe_name = os.path.basename(filename)
    full = os.path.join(ATAS_DIR, safe_name)
    if not os.path.isfile(full) or os.path.dirname(os.path.abspath(full)) != os.path.abspath(ATAS_DIR):
        return jsonify({"ok": False, "error": "Arquivo não encontrado."}), 404
    return send_file(full, mimetype="application/pdf")


# Pendências recorrentes identificadas manualmente a partir da leitura de
# todas as atas de reunião (AR-01 a AR-23). Não é um parser automático — é
# uma curadoria de assuntos que reaparecem sem confirmação de encerramento
# nas atas seguintes, atendendo ao pedido da própria Construtora (AR-15,
# item 02: "compilado entre as Atas para resolver as pendências").
#
# Cada pendência carrega "encerrada" (True/False) e, quando fechada,
# "encerramento" com a ata/data/trecho que confirmou o fechamento — é o
# único jeito de uma pendência sair do quadro de "aberta" sem reler tudo
# de novo a cada ata nova. QUANDO UMA ATA NOVA CHEGAR: reler o PDF inteiro
# (não só a tabela de ações — os "Registros da Reunião" é onde costuma
# aparecer a confirmação de fechamento de pendência de ata anterior, ex.:
# "na visita foi verificado que a pendência foi solucionada") e:
#   1) atualizar "ultima_mencao"/"ocorrencias" de quem foi citado de novo;
#   2) marcar "encerrada": True só com confirmação EXPLÍCITA de fechamento
#      (não deduzir por ausência de menção — isso é "sem confirmação", não
#      "resolvido", e continua contando como pendência aberta);
#   3) se abriu um novo desdobramento (ex.: "acesso liberado, mas
#      treinamento pendente"), manter aberta com o texto atualizado.
_ATAS_PENDENCIAS_SEED = [
    {
        "titulo": "Compra e revenda de patrimônios sem emissão de Nota Fiscal",
        "palavras_chave": ["nota fiscal", "revenda de patrimôni", "aquisição de patrimôni", "metodologia de aquisição"],
        "descricao": "Metodologia para aquisição/revenda de patrimônios da obra sem emissão de NF, cobrada repetidamente sem confirmação de fechamento nas atas seguintes. Reconfirmada aberta na AR-23 (item R01: processo de aquisição de grua/cremalheira/escoramento \"foi aprovado parcialmente, mas não foi documentado\") e virou ação formal com prazo (A07: oficializar a metodologia até 24/08/2026).",
        "responsavel": "GAMA DIAS",
        "primeira_mencao": {"ata": "AR-06", "data": "2026-04-27"},
        "ultima_mencao": {"ata": "AR-23", "data": "2026-08-17"},
        "ocorrencias": 5,
        "status": "aberta — ação formal A07 com prazo 24/08/2026",
        "encerrada": False,
        "encerramento": None,
    },
    {
        "titulo": "Contrato de construção da obra sem assinatura dos sócios",
        "palavras_chave": ["contrato de construção", "assinatura dos sócios", "assinado pelos sócios"],
        "descricao": "\"Novamente foi comentado que o contrato de construção da obra ainda não foi assinado pelos sócios\" — cobrança à direção da Temazec, sem menção de resolução nas atas posteriores lidas (inclusive a AR-23, que não voltou a citar o assunto).",
        "responsavel": "TEMAZEC",
        "primeira_mencao": {"ata": "AR-09", "data": "2026-05-18"},
        "ultima_mencao": {"ata": "AR-09", "data": "2026-05-18"},
        "ocorrencias": 1,
        "status": "sem acompanhamento formal nas atas seguintes",
        "encerrada": False,
        "encerramento": None,
    },
    {
        "titulo": "Divergência entre dados extraídos via API do Sienge (gerenciadora x construtora)",
        "palavras_chave": ["divergência", "api do sienge", "evolução financeira", "dados extraídos"],
        "descricao": "Identificada divergência nos arquivos de API do Sienge extraídos pela gerenciadora e pela construtora; cogitada videoconferência para resolução.",
        "responsavel": "ENGIN",
        "primeira_mencao": {"ata": "AR-09", "data": "2026-05-18"},
        "ultima_mencao": {"ata": "AR-23", "data": "2026-08-17"},
        "ocorrencias": 2,
        "status": "encerrada",
        "encerrada": True,
        "encerramento": {
            "ata": "AR-23", "data": "2026-08-17",
            "observacao": "AR-23, item R03: \"Em atas anteriores havia ficado em aberto o assunto a respeito da divergência entre os valores da evolução financeira levantada pela Gerenciadora e Construtora. Na visita foi verificado que pendencia foi solucionada.\"",
        },
    },
    {
        "titulo": "Acesso do engenheiro da obra a módulos do Sienge",
        "palavras_chave": ["acesso", "apropriação de custos", "controle de estoque", "estoque e patrimônio", "módulos do sienge", "treinamento"],
        "descricao": "Elvis (engenheiro da obra) sem acesso à apropriação de custos (Financeiro > Títulos > Apropriação de Obra) e ao controle de estoque (Suprimentos > Apoio > Estoque). PARCIALMENTE resolvida na AR-23: acesso ao módulo de estoque/patrimônio foi liberado, mas abriu um novo desdobramento — falta o treinamento da equipe da construtora pela PSA antes do uso efetivo.",
        "responsavel": "GAMA DIAS",
        "primeira_mencao": {"ata": "AR-09", "data": "2026-05-18"},
        "ultima_mencao": {"ata": "AR-23", "data": "2026-08-17"},
        "ocorrencias": 2,
        "status": "parcial — acesso liberado, treinamento da equipe pela PSA ainda pendente",
        "encerrada": False,
        "encerramento": None,
    },
    {
        "titulo": "Relatório personalizado no Sienge (orçamento x cotação)",
        "palavras_chave": ["relatório personalizado", "saldo de orçamento", "aprovação de cotaç", "relatório no sienge"],
        "descricao": "Solicitação de relatório no Sienge para comparar saldo de orçamento com o custo a ser aprovado, auxiliando a aprovação de cotações — pedido de Rodrigo (Gama Dias) e da Temazec.",
        "responsavel": "GAMA DIAS",
        "primeira_mencao": {"ata": "AR-09", "data": "2026-05-18"},
        "ultima_mencao": {"ata": "AR-09", "data": "2026-05-18"},
        "ocorrencias": 1,
        "status": "sem confirmação de encerramento",
        "encerrada": False,
        "encerramento": None,
    },
    {
        "titulo": "Contratos no Sienge não consideram adiantamentos corretamente",
        "palavras_chave": ["adiantamentos", "contratos no sienge", "adiantamentos de contratos"],
        "descricao": "Identificado que contratos lançados no Sienge não refletem adiantamentos reais — exemplo citado: contrato de aquisição do elevador aparece como \"concluído\" indevidamente. Reconfirmada aberta na AR-23 (item R05: \"ainda não foi possível evidenciar o valor real de adiantamentos dos contratos da obra\") e virou ação formal (A09, prazo 24/08/2026).",
        "responsavel": "TEMAZEC / ENGIN",
        "primeira_mencao": {"ata": "AR-12", "data": "2026-06-08"},
        "ultima_mencao": {"ata": "AR-23", "data": "2026-08-17"},
        "ocorrencias": 3,
        "status": "aberta — ação formal A09 com prazo 24/08/2026",
        "encerrada": False,
        "encerramento": None,
    },
    {
        "titulo": "Oficialização do contrato de aquisição da grua e cremalheira",
        "palavras_chave": ["grua", "cremalheira", "contrato de aquisição", "oficializado", "escoramento"],
        "descricao": "Grua já instalada e elevador cremalheira já em reforma, mas contrato de aquisição ainda não oficializado com os sócios. Risco explícito de sobrecusto e impacto no cronograma se houver reabertura de concorrência. Mesma raiz do item \"Compra e revenda de patrimônios sem NF\" — AR-23 confirma que segue sem documentação (item R01), agora com ação formal A07 (prazo 24/08/2026).",
        "responsavel": "GAMA DIAS / TEMAZEC / CONTEC",
        "primeira_mencao": {"ata": "AR-19", "data": "2026-07-20"},
        "ultima_mencao": {"ata": "AR-23", "data": "2026-08-17"},
        "ocorrencias": 2,
        "status": "urgente — equipamento em uso sem contrato formal",
        "encerrada": False,
        "encerramento": None,
    },
    {
        "titulo": "Apuração e pagamento da taxa de construção",
        "palavras_chave": ["taxa de construção", "pagamento da taxa", "desembolso acumulado"],
        "descricao": "Construtora solicita auxílio para apurar as informações necessárias e iniciar o pagamento da taxa de construção. AR-23 (item R02) levantou o desembolso acumulado de R$ 4.869.106,55 até julho/2026 (taxa correspondente: R$ 96.382,05) mas a Gerenciadora identificou um possível ERRO na consideração do pagamento da taxa corrigido — abriu ação formal A08 (validar com o Financeiro da Gama Dias, prazo 24/08/2026). Ainda sem confirmação de pagamento.",
        "responsavel": "GAMA DIAS / ENGIN",
        "primeira_mencao": {"ata": "AR-17", "data": "2026-07-06"},
        "ultima_mencao": {"ata": "AR-23", "data": "2026-08-17"},
        "ocorrencias": 3,
        "status": "aberta — possível erro de cálculo identificado, ação formal A08 com prazo 24/08/2026",
        "encerrada": False,
        "encerramento": None,
    },
    {
        "titulo": "Apropriação indevida de R$ 40.500,00 — Projeto de Arquitetura de Interiores",
        "palavras_chave": ["apropriação indevida", "40.500", "arquitetura de interiores", "reapropriar"],
        "descricao": "Identificada NOVAMENTE apropriação indevida no custo da obra referente à elaboração de projeto de arquitetura de interiores; equipe de Desenvolvimento de Projetos deve reapropriar para o centro de custo de Incorporação. Não voltou a ser citada na AR-23.",
        "responsavel": "GAMA DIAS",
        "primeira_mencao": {"ata": "AR-22", "data": "2026-08-10"},
        "ultima_mencao": {"ata": "AR-22", "data": "2026-08-10"},
        "ocorrencias": 1,
        "status": "recorrente — já ocorreu antes ('novamente'), sem confirmação de encerramento",
        "encerrada": False,
        "encerramento": None,
    },
    {
        "titulo": "Desvio negativo de evolução física",
        "palavras_chave": ["evolução física", "desvio negativo", "planilha de medição"],
        "descricao": "No mês de referência a evolução física foi de 0,43% contra 2,28% planejado; acumulado de 7,47% contra 8,60% planejado (desvio de -1,13%). AR-23 (item R07) reavaliou com a planilha de medição validada: projeção aponta desvio negativo persistindo até dezembro/26, mas com tendência mês a mês de DIMINUIÇÃO do desvio, caminho crítico preservado, e alívio esperado com o início da medição do banco. Segue em acompanhamento — não é uma pendência que \"fecha\", é indicador contínuo.",
        "responsavel": "TEMAZEC / ENGIN",
        "primeira_mencao": {"ata": "AR-17", "data": "2026-07-06"},
        "ultima_mencao": {"ata": "AR-23", "data": "2026-08-17"},
        "ocorrencias": 2,
        "status": "em acompanhamento — tendência de melhora, caminho crítico preservado",
        "encerrada": False,
        "encerramento": None,
    },
]

ATAS_PENDENCIAS_PATH = os.path.join(DATA_DIR, "atas_pendencias_state.json")
ATAS_TEXTO_CACHE_PATH = os.path.join(DATA_DIR, "atas_texto_cache.json")

# Frases que indicam que um assunto foi FECHADO na ata em que aparecem —
# usadas pra decidir se uma pendência pode ser marcada como encerrada.
# Case/acento-insensível (comparado sobre texto já normalizado).
_FRASES_ENCERRAMENTO = [
    "foi solucionad", "foi resolvid", "foi encerrad", "esta encerrad",
    "foi concluid", "foi finalizad", "foi regularizad", "ja foi resolvid",
    "pendencia foi solucionada", "problema foi resolvido", "questao foi resolvida",
]
# Frases que indicam que o assunto CONTINUA aberto (mesmo remencionado) —
# usadas pra, se aparecerem numa pendência já marcada como encerrada,
# reabri-la em vez de deixar encerrada silenciosamente errada.
_FRASES_REABERTURA = [
    "ainda nao", "segue pendente", "continua em aberto", "nao foi possivel",
    "ainda em aberto", "segue sem", "permanece em aberto", "novamente",
]


def _norm_txt(s):
    return unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode("ascii").lower()


def _load_atas_pendencias():
    if not os.path.exists(ATAS_PENDENCIAS_PATH):
        estado = {"pendencias": copy.deepcopy(_ATAS_PENDENCIAS_SEED), "ultima_releitura": None}
        _save_atas_pendencias(estado)
        return estado
    try:
        with open(ATAS_PENDENCIAS_PATH, "r", encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return {"pendencias": copy.deepcopy(_ATAS_PENDENCIAS_SEED), "ultima_releitura": None}


def _save_atas_pendencias(estado):
    os.makedirs(DATA_DIR, exist_ok=True)
    tmp = ATAS_PENDENCIAS_PATH + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(estado, fh, ensure_ascii=False, indent=2)
    os.replace(tmp, ATAS_PENDENCIAS_PATH)


def _extrair_texto_atas():
    """Extrai o texto de todas as atas em ATAS_DIR, cacheado por mtime do
    arquivo — só reprocessa PDF novo ou alterado. Retorna lista ordenada
    por data, cada item com {arquivo, numero, data, texto}."""
    import pdfplumber

    resultado = []
    if not os.path.isdir(ATAS_DIR):
        return resultado

    cache = {}
    if os.path.exists(ATAS_TEXTO_CACHE_PATH):
        try:
            with open(ATAS_TEXTO_CACHE_PATH, "r", encoding="utf-8") as fh:
                cache = json.load(fh)
        except Exception:
            cache = {}

    mudou = False
    for fname in os.listdir(ATAS_DIR):
        if not fname.lower().endswith(".pdf"):
            continue
        full = os.path.join(ATAS_DIR, fname)
        mtime = os.path.getmtime(full)
        cached = cache.get(fname)
        if cached and cached.get("mtime") == mtime:
            texto = cached["texto"]
        else:
            try:
                with pdfplumber.open(full) as pdf:
                    texto = "\n".join(p.extract_text() or "" for p in pdf.pages)
            except Exception as exc:
                texto = ""
                print(f"  [ATAS] falha ao ler {fname}: {exc}")
            cache[fname] = {"mtime": mtime, "texto": texto}
            mudou = True
        meta = _parse_ata_filename(fname)
        if not meta["data"]:
            continue
        resultado.append({"arquivo": fname, "numero": meta["numero"], "data": meta["data"], "texto": texto})

    if mudou:
        tmp = ATAS_TEXTO_CACHE_PATH + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(cache, fh, ensure_ascii=False)
        os.replace(tmp, ATAS_TEXTO_CACHE_PATH)

    resultado.sort(key=lambda t: t["data"])
    return resultado


def reanalisar_pendencias():
    """Regra de releitura automática: pra cada pendência já tracked, varre
    as atas mais novas que a última menção registrada procurando as
    palavras-chave do assunto. Se achar, junto com uma frase de
    encerramento explícita ("foi solucionado", "foi encerrado" etc.),
    marca a pendência como encerrada citando o trecho encontrado — nunca
    por ausência de menção, só por confirmação textual. Se uma pendência
    JÁ encerrada for citada de novo com sinal de reabertura ("ainda não",
    "novamente", "segue pendente"), ela volta a ficar aberta. Também
    junta, à parte, trechos de "Registros da Reunião" que soam a pendência
    (linguagem de pendência/atraso) mas não bateram com nenhuma
    palavra-chave já cadastrada — isso não vira pendência sozinho, fica
    como sugestão pra revisão humana."""
    estado = _load_atas_pendencias()
    pendencias = estado["pendencias"]
    atas = _extrair_texto_atas()

    atas_norm = [{**a, "texto_norm": _norm_txt(a["texto"]), "linhas_norm": [_norm_txt(l) for l in a["texto"].split("\n")], "linhas": a["texto"].split("\n")} for a in atas]

    linhas_casadas = set()  # (arquivo, indice_linha) já atribuídas a alguma pendência — sobra pra sugestões

    for p in pendencias:
        kws = [_norm_txt(k) for k in (p.get("palavras_chave") or [])]
        if not kws:
            continue
        ultima_data = (p.get("ultima_mencao") or {}).get("data") or ""
        for ata in atas_norm:
            if ata["data"] <= ultima_data:
                continue
            for i, linha_norm in enumerate(ata["linhas_norm"]):
                if not linha_norm.strip():
                    continue
                if not any(kw in linha_norm for kw in kws):
                    continue
                linhas_casadas.add((ata["arquivo"], i))
                # janela de contexto (linha +- 1) pra pegar frase de encerramento
                # que esteja no período/frase seguinte à citação do assunto
                janela = " ".join(ata["linhas_norm"][max(0, i - 1):i + 2])
                trecho_original = " ".join(ata["linhas"][max(0, i - 1):i + 2]).strip()
                fechou = any(f in janela for f in _FRASES_ENCERRAMENTO)
                reabriu = any(f in janela for f in _FRASES_REABERTURA)
                p["ocorrencias"] = (p.get("ocorrencias") or 0) + 1
                p["ultima_mencao"] = {"ata": ata["numero"], "data": ata["data"]}
                if fechou and not reabriu:
                    p["encerrada"] = True
                    p["encerramento"] = {"ata": ata["numero"], "data": ata["data"], "observacao": trecho_original[:400]}
                    p["status"] = "encerrada (releitura automática)"
                elif p.get("encerrada") and reabriu:
                    p["encerrada"] = False
                    p["encerramento"] = None
                    p["status"] = "reaberta (releitura automática) — " + trecho_original[:200]
                # se nem fechou nem reabriu e a pendência já está aberta, só
                # a reconfirmação de ultima_mencao/ocorrencias acima é o que muda

    # ── sugestões: trechos de "registros da reunião" com cheiro de pendência
    # que não bateram com nenhuma palavra-chave já tracked ──────────────────
    # O texto extraído do PDF quebra a frase em várias linhas curtas (o
    # layout da tabela, não a pontuação) — pegar só a linha que bateu a
    # palavra-chave dava um pedaço cortado no meio da frase. Agrupa linhas
    # marcadas que estão perto umas das outras (mesmo parágrafo/célula) numa
    # única sugestão, juntando um pouco de contexto antes/depois.
    _SINAIS_PENDENCIA = ["ainda nao", "nao foi possivel", "em aberto", "pendente", "sem confirmacao", "novamente"]
    sugestoes = []
    for ata in atas_norm:
        linhas = ata["linhas"]
        total_linhas = len(linhas)
        marcadas = [
            i for i, linha_norm in enumerate(ata["linhas_norm"])
            if (ata["arquivo"], i) not in linhas_casadas
            and len(linha_norm.strip()) >= 25
            and any(s in linha_norm for s in _SINAIS_PENDENCIA)
        ]
        grupos = []
        for i in marcadas:
            if grupos and i - grupos[-1][-1] <= 2:
                grupos[-1].append(i)
            else:
                grupos.append([i])
        for grupo in grupos:
            ini = max(0, grupo[0] - 2)
            fim = min(total_linhas, grupo[-1] + 3)
            trecho = " ".join(l.strip() for l in linhas[ini:fim] if l.strip())
            trecho = re.sub(r"\s+", " ", trecho).strip()
            if len(trecho) < 25:
                continue
            sugestoes.append({
                "ata": ata["numero"], "data": ata["data"],
                "trecho": trecho[:600],
            })
    # limita e deduplica trechos muito parecidos (mesmo prefixo)
    vistos = set()
    sugestoes_dedup = []
    for s in sugestoes:
        chave = s["trecho"][:60].lower()
        if chave in vistos:
            continue
        vistos.add(chave)
        sugestoes_dedup.append(s)

    estado["pendencias"] = pendencias
    estado["ultima_releitura"] = time.strftime("%Y-%m-%d %H:%M:%S")
    estado["sugestoes"] = sugestoes_dedup[:30]
    _save_atas_pendencias(estado)
    return estado


@app.route("/api/atas/pendencias")
def api_atas_pendencias():
    estado = _load_atas_pendencias()
    return jsonify({"ok": True, "pendencias": estado["pendencias"], "ultima_releitura": estado.get("ultima_releitura"), "sugestoes": estado.get("sugestoes") or []})


@app.route("/api/atas/pendencias/reanalisar", methods=["POST"])
def api_atas_pendencias_reanalisar():
    try:
        estado = reanalisar_pendencias()
        return jsonify({"ok": True, "pendencias": estado["pendencias"], "ultima_releitura": estado.get("ultima_releitura"), "sugestoes": estado.get("sugestoes") or []})
    except Exception as exc:
        return jsonify({"ok": False, "error": f"Falha ao reler atas: {exc}"}), 500


@app.route("/api/atas/pendencias/<int:idx>/toggle", methods=["POST"])
def api_atas_pendencias_toggle(idx):
    """Override manual — pra quando a releitura automática (regra por
    palavra-chave) errar o veredito, um humano pode corrigir na hora."""
    estado = _load_atas_pendencias()
    pendencias = estado["pendencias"]
    if idx < 0 or idx >= len(pendencias):
        return jsonify({"ok": False, "error": "Pendência não encontrada."}), 404
    p = pendencias[idx]
    p["encerrada"] = not p.get("encerrada")
    if p["encerrada"]:
        user = _current_auth_user()
        quem = (user or {}).get("name") or (user or {}).get("email") or "usuário"
        p["encerramento"] = {"ata": "manual", "data": time.strftime("%Y-%m-%d"), "observacao": f"Marcada como encerrada manualmente por {quem}."}
        p["status"] = "encerrada (marcação manual)"
    else:
        p["encerramento"] = None
        p["status"] = "reaberta manualmente"
    _save_atas_pendencias(estado)
    return jsonify({"ok": True, "item": p})


@app.route("/api/atas/pendencias/nova", methods=["POST"])
def api_atas_pendencias_nova():
    payload = request.get_json(silent=True) or {}
    titulo = str(payload.get("titulo") or "").strip()
    if not titulo:
        return jsonify({"ok": False, "error": "Título é obrigatório."}), 400
    estado = _load_atas_pendencias()
    item = {
        "titulo": titulo,
        "descricao": str(payload.get("descricao") or "").strip(),
        "responsavel": str(payload.get("responsavel") or "").strip(),
        "palavras_chave": [w.strip() for w in str(payload.get("palavras_chave") or "").split(",") if w.strip()],
        "primeira_mencao": {"ata": payload.get("ata") or "", "data": payload.get("data") or ""},
        "ultima_mencao": {"ata": payload.get("ata") or "", "data": payload.get("data") or ""},
        "ocorrencias": 1,
        "status": "aberta (adicionada manualmente)",
        "encerrada": False,
        "encerramento": None,
    }
    estado["pendencias"].append(item)
    _save_atas_pendencias(estado)
    return jsonify({"ok": True, "item": item})


# ═══ MINUTAS DE CONTRATO ═════════════════════════════════════════════════
# Gera o .docx da minuta a partir do template base (dados/templates/
# minuta_contrato_base.docx), preenchendo automaticamente os dados do
# CONTRATADO (via credor/contrato do Sienge + consulta de CNPJ na Receita
# Federal) e os itens do contrato, mantendo intactas todas as clausulas
# juridicas do modelo original.

def _load_minuta_counter():
    try:
        with open(MINUTA_COUNTER_PATH, "r", encoding="utf-8-sig") as fh:
            d = json.load(fh)
        return int(d.get("ultimo", 1000))
    except Exception:
        return 1000

def _save_minuta_counter(n):
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(MINUTA_COUNTER_PATH, "w", encoding="utf-8") as fh:
        json.dump({"ultimo": int(n)}, fh)

def _minuta_fmt_brl(v):
    try:
        v = float(v)
    except Exception:
        v = 0.0
    return f"{v:,.2f}".replace(",", "\x00").replace(".", ",").replace("\x00", ".")

def _minuta_distinct_cells(row):
    seen = {}
    for c in row.cells:
        k = id(c._tc)
        if k not in seen:
            seen[k] = c
    return list(seen.values())

def _minuta_set_cell(cell, text):
    para = cell.paragraphs[0]
    runs = para.runs
    if runs:
        runs[0].text = text
        for r in runs[1:]:
            r.text = ""
    else:
        para.add_run(text)
    for p in cell.paragraphs[1:]:
        for r in p.runs:
            r.text = ""

def _minuta_set_para(doc, idx, text):
    para = doc.paragraphs[idx]
    runs = para.runs
    if runs:
        runs[0].text = text
        for r in runs[1:]:
            r.text = ""
    else:
        para.add_run(text)

def gerar_minuta_docx(dados):
    if _docxlib is None:
        raise RuntimeError("python-docx não está instalado no servidor.")
    doc = _docxlib.Document(MINUTA_TEMPLATE_PATH)

    _minuta_set_para(doc, 0, f"CONTRATO DE PRESTAÇÃO DE SERVIÇO N° {dados.get('numero','')}")
    _minuta_set_para(doc, 10, dados.get("objeto", ""))
    valor_total = sum(float(it.get("total") or 0) for it in dados.get("itens", []))
    _minuta_set_para(doc, 13, f"VALOR ESTIMADO TOTAL DO CONTRATO: R$ {_minuta_fmt_brl(valor_total)}")
    if dados.get("data_extenso"):
        _minuta_set_para(doc, 140, dados["data_extenso"])

    c = dados.get("contratado", {})
    t = doc.tables[1]
    _minuta_set_cell(_minuta_distinct_cells(t.rows[1])[1], c.get("empresa", ""))
    _minuta_set_cell(_minuta_distinct_cells(t.rows[2])[1], c.get("endereco", ""))
    r3 = _minuta_distinct_cells(t.rows[3])
    _minuta_set_cell(r3[1], c.get("bairro", ""))
    _minuta_set_cell(r3[3], c.get("cep", ""))
    _minuta_set_cell(_minuta_distinct_cells(t.rows[4])[1], c.get("telefone", ""))
    r5 = _minuta_distinct_cells(t.rows[5])
    _minuta_set_cell(r5[1], c.get("cnpj", ""))
    _minuta_set_cell(r5[3], c.get("ie", ""))
    _minuta_set_cell(_minuta_distinct_cells(t.rows[6])[1], c.get("representante", ""))
    _minuta_set_cell(_minuta_distinct_cells(t.rows[7])[1], c.get("cpf", ""))
    _minuta_set_cell(_minuta_distinct_cells(t.rows[8])[1], c.get("email", ""))

    ti = doc.tables[2]
    itens = dados.get("itens", [])
    if itens:
        template_tr = ti.rows[1]._tr
        for row in list(ti.rows[1:]):
            ti._tbl.remove(row._tr)
        for it in itens:
            new_tr = copy.deepcopy(template_tr)
            ti._tbl.append(new_tr)
            new_row = ti.rows[-1]
            cells = _minuta_distinct_cells(new_row)
            total = float(it.get("total") or 0)
            vals = [it.get("descricao", ""), it.get("unidade") or "UND",
                    _minuta_fmt_brl(it.get("qtd", 0)), _minuta_fmt_brl(it.get("preco_unit", 0)), _minuta_fmt_brl(total)]
            for cell, v in zip(cells, vals):
                _minuta_set_cell(cell, str(v))

    _minuta_set_cell(_minuta_distinct_cells(doc.tables[5].rows[0])[1], c.get("empresa", ""))

    if dados.get("forma_pagamento"):
        _minuta_set_para(doc, 25, dados["forma_pagamento"])
    if dados.get("prazo"):
        _minuta_set_cell(_minuta_distinct_cells(doc.tables[3].rows[0])[0], dados["prazo"])

    testemunhas = dados.get("testemunhas")
    if testemunhas and len(testemunhas) >= 2:
        _minuta_set_cell(_minuta_distinct_cells(doc.tables[6].rows[0])[1], testemunhas[0].get("nome", ""))
        _minuta_set_cell(_minuta_distinct_cells(doc.tables[6].rows[1])[1], testemunhas[0].get("cpf", ""))
        _minuta_set_cell(_minuta_distinct_cells(doc.tables[7].rows[0])[1], testemunhas[1].get("nome", ""))
        _minuta_set_cell(_minuta_distinct_cells(doc.tables[7].rows[1])[1], testemunhas[1].get("cpf", ""))

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

@app.route("/api/minutas/proximo_numero")
def api_minutas_proximo_numero():
    return jsonify({"ok": True, "proximo": _load_minuta_counter() + 1})

def _minuta_pendentes_controle():
    """Itens de 'Controle de contratações.xlsx' — contratações em andamento
    (aprovadas internamente ou ainda pendentes) que ainda não têm itens
    lançados no Sienge, por isso nunca aparecem em fsup_ctitensct. Sem CNPJ
    nem itens detalhados: ao selecionar, o usuário preenche o CONTRATADO
    manualmente na minuta."""
    out = []
    for r in load_controle_contratacoes():
        item = str(r.get("item") or "").strip()
        desc = str(r.get("descricao") or "").strip()
        obs = str(r.get("observacoes") or "").strip()
        if not desc and not obs:
            continue
        label = desc or (obs[:40] + "…" if len(obs) > 40 else obs)
        valor = r.get("custo") or r.get("orc") or 0
        aprov = str(r.get("status") or "").strip().upper()
        out.append({
            "docnuct": "", "nuct": "", "nmcredor": label, "nmcc": r.get("etapa", ""),
            "cnpj": "", "vlrtotct": valor, "itens": 0,
            "pct": 0, "status": "sem_sienge",
            "item_controle": item, "observacoes": obs,
            "aprovacao_interna": aprov or "—",
        })
    return out


@app.route("/api/minutas/credores")
def api_minutas_credores():
    data = load_dashboard_cache() or {}
    contratos = data.get("contratos") or []
    itens_ct = data.get("itens_ct") or {}
    out = []
    for c in contratos:
        docnuct = c.get("docnuct")
        nmcredor = c.get("nmcredor", "")
        pct = float(c.get("pct") or 0)
        sem_credor = str(nmcredor).strip().lower().startswith("credor n")
        status = "sem_credor" if sem_credor else ("concluido" if pct >= 100 else "ativo")
        out.append({
            "docnuct": docnuct, "nuct": c.get("nuct", ""), "nmcredor": nmcredor, "nmcc": c.get("nmcc", ""),
            "cnpj": c.get("cnpj", ""), "vlrtotct": c.get("vlrtotct", 0),
            "itens": len(itens_ct.get(docnuct, [])),
            "pct": pct, "status": status,
            "status_sienge": c.get("status_sienge", ""), "aprovacao_sienge": c.get("aprovacao_sienge", ""),
        })
    out.extend(_minuta_pendentes_controle())
    out.sort(key=lambda x: x["vlrtotct"] or 0)
    return jsonify({"ok": True, "credores": out})

@app.route("/api/minutas/itens/<path:docnuct>")
def api_minutas_itens(docnuct):
    data = load_dashboard_cache() or {}
    itens_ct = data.get("itens_ct") or {}
    return jsonify({"ok": True, "itens": itens_ct.get(docnuct, [])})

@app.route("/api/minutas/cnpj/<cnpj>")
def api_minutas_cnpj(cnpj):
    info = lookup_cnpj_full(cnpj)
    if not info:
        return jsonify({"ok": False, "error": "CNPJ não encontrado, inválido, ou serviço de consulta indisponível no momento."}), 404
    return jsonify({"ok": True, "dados": info})

@app.route("/api/minutas/gerar", methods=["POST"])
def api_minutas_gerar():
    if _docxlib is None:
        return jsonify({"ok": False, "error": "python-docx não está instalado no servidor."}), 500
    body = request.get_json(force=True, silent=True) or {}
    numero = str(body.get("numero") or "").strip()
    if not numero:
        return jsonify({"ok": False, "error": "Informe o número do contrato."}), 400
    contratado = body.get("contratado") or {}
    if not (contratado.get("empresa") or "").strip():
        return jsonify({"ok": False, "error": "Informe ao menos a empresa do CONTRATADO."}), 400
    itens = body.get("itens") or []
    for it in itens:
        try:
            qtd = float(it.get("qtd") or 0)
            pu = float(it.get("preco_unit") or 0)
            it["total"] = round(qtd * pu, 2)
        except Exception:
            it["total"] = float(it.get("total") or 0)
    dados = {
        "numero": numero,
        "objeto": body.get("objeto", ""),
        "data_extenso": body.get("data_extenso") or "",
        "contratado": contratado,
        "itens": itens,
        "forma_pagamento": body.get("forma_pagamento", ""),
        "prazo": body.get("prazo", ""),
        "testemunhas": body.get("testemunhas"),
    }
    try:
        buf = gerar_minuta_docx(dados)
    except Exception as ex:
        traceback.print_exc()
        return jsonify({"ok": False, "error": f"Erro ao gerar minuta: {ex}"}), 500
    try:
        num_int = int(re.sub(r"\D", "", numero) or 0)
        if num_int and num_int > _load_minuta_counter():
            _save_minuta_counter(num_int)
    except Exception:
        pass
    empresa_safe = re.sub(r'[\\/:*?"<>|]', "", (contratado.get("empresa") or "MINUTA")).strip()
    fname = f"CONTRATO {numero} - {empresa_safe}.docx"
    return send_file(buf, mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                      as_attachment=True, download_name=fname)


# ═══ CARTAS CONVITE ═══════════════════════════════════════════════════════
# Mesma ideia da minuta: textos fixos no JSON de configuração
# (dados/synergia/templates/carta_convite_<pacote>.json) e o que muda vem
# sozinho — pranchas atuais do AutoDoc (última revisão APROVADA) e prazos do
# Cronograma Base do PCP. A lógica toda fica em carta_convite.py, que também
# roda sozinho pelo agendador do Windows (ROBO_CARTA_CONVITE.bat).
_cartas_sync_lock = threading.Lock()
_cartas_envio_lock = threading.Lock()


def _carta_pacote():
    import carta_convite as cc
    pacote = (request.args.get("pacote") or (request.get_json(silent=True) or {}).get("pacote") or "mdo_bruta").strip()
    if not re.fullmatch(r"[a-z0-9_]+", pacote) or not os.path.exists(cc.config_path(pacote)):
        return None
    return pacote


@app.route("/api/cartas/pacotes")
def api_cartas_pacotes():
    import carta_convite as cc
    return jsonify({"ok": True, "pacotes": cc.listar_pacotes()})


@app.route("/api/cartas/estado")
def api_cartas_estado():
    import carta_convite as cc
    pacote = _carta_pacote()
    if not pacote:
        return jsonify({"ok": False, "error": "Pacote de carta convite não encontrado."}), 404
    st = cc.estado_pacote(pacote)
    cfg = cc.carregar_config(pacote)
    prazos = st.get("prazos") or {}
    email = cc.config_email(pacote, cfg, st)
    return jsonify({
        "ok": True, "pacote": pacote, "nome": cfg.get("nome", pacote),
        "grupos": st.get("grupos") or [], "sync_at": st.get("sync_at", ""),
        "disciplinas_autodoc": st.get("disciplinas_autodoc") or [],
        "disciplinas_selecionadas": cc.disciplinas_selecionadas(pacote, cfg, st),
        "pranchas_selecionadas": sorted(cc.chaves_selecionadas(pacote, cfg, st)),
        "sync_progresso": st.get("sync_progresso", ""),
        "frentes": cc.frentes_brutas(pacote, cfg, st),
        "matriz": cc.matriz_efetiva(pacote, cfg, st, so_escopo=False),
        "anexo": cc.anexo_efetivo(pacote, cfg, st),
        "email": email, "envio": st.get("envio") or {}, "enviando": _cartas_envio_lock.locked(),
        "historico_envios": (st.get("historico_envios") or [])[:10],
        "sync_status": st.get("sync_status", ""), "sync_erro": st.get("sync_erro"),
        "ultimas_mudancas": st.get("ultimas_mudancas") or [], "historico_mudancas": st.get("historico_mudancas") or [],
        "prazos_fonte": prazos.get("fonte", ""), "prazos_arquivo": prazos.get("arquivo", ""),
        "prazos_pavimentos": len(prazos.get("linhas") or []), "prazos_lido_em": prazos.get("lido_em", ""),
        "ultima_geracao": st.get("ultima_geracao", ""), "ultimos_avisos": st.get("ultimos_avisos") or [],
        "sincronizando": _cartas_sync_lock.locked(),
    })


@app.route("/api/cartas/sincronizar", methods=["POST"])
def api_cartas_sincronizar():
    """Relê as pranchas no AutoDoc (~1 min, login real) em thread de fundo —
    o front consulta /api/cartas/estado até sync_status sair de
    'sincronizando'."""
    import carta_convite as cc
    pacote = _carta_pacote()
    if not pacote:
        return jsonify({"ok": False, "error": "Pacote de carta convite não encontrado."}), 404
    if not _cartas_sync_lock.acquire(blocking=False):
        return jsonify({"ok": True, "background": True, "ja_rodando": True})
    cc._atualizar_estado(pacote, sync_status="sincronizando", sync_erro=None)

    def _bg():
        try:
            cc.sincronizar_pranchas(pacote)
        except Exception as exc:
            cc._atualizar_estado(pacote, sync_status="erro", sync_erro=str(exc))
        finally:
            _cartas_sync_lock.release()

    threading.Thread(target=_bg, daemon=True).start()
    return jsonify({"ok": True, "background": True})


@app.route("/api/cartas/selecao", methods=["POST"])
def api_cartas_selecao():
    """Salva os tipos de projeto (disciplinas) e as pranchas marcadas."""
    import carta_convite as cc
    pacote = _carta_pacote()
    if not pacote:
        return jsonify({"ok": False, "error": "Pacote de carta convite não encontrado."}), 404
    body = request.get_json(force=True, silent=True) or {}
    disc = body.get("disciplinas")
    pr = body.get("pranchas")
    cc.salvar_selecao(pacote, disciplinas=disc if isinstance(disc, list) else None,
                      pranchas=pr if isinstance(pr, list) else None)
    return jsonify({"ok": True})


@app.route("/api/cartas/escopo", methods=["POST"])
def api_cartas_escopo():
    """Salva frentes de serviço marcadas, quantidades, responsáveis da
    matriz e as marcações do quadro resumo (Anexo I)."""
    import carta_convite as cc
    pacote = _carta_pacote()
    if not pacote:
        return jsonify({"ok": False, "error": "Pacote de carta convite não encontrado."}), 404
    body = request.get_json(force=True, silent=True) or {}
    frentes = body.get("frentes")
    if body.get("recarregar_disciplinas"):
        # volta ao automático: frentes sugeridas pelos tipos de projeto marcados
        cc._atualizar_estado(pacote, frentes=None)
        frentes = None
    st = cc.salvar_escopo(pacote,
                          frentes=frentes if isinstance(frentes, list) else None,
                          matriz=body["matriz"] if isinstance(body.get("matriz"), dict) else None,
                          matriz_frente=body["matriz_frente"] if isinstance(body.get("matriz_frente"), dict) else None,
                          anexo=body["anexo"] if isinstance(body.get("anexo"), dict) else None)
    # devolve as frentes com os ids definitivos (frente nova ganha id no servidor)
    return jsonify({"ok": True, "frentes": cc.frentes_brutas(pacote, None, st)})


@app.route("/api/cartas/planilha", methods=["POST"])
def api_cartas_planilha():
    import carta_convite as cc
    pacote = _carta_pacote()
    if not pacote:
        return jsonify({"ok": False, "error": "Pacote de carta convite não encontrado."}), 404
    body = request.get_json(force=True, silent=True) or {}
    buf = cc.gerar_planilha_proposta(pacote, _carta_opcoes(body))
    return send_file(buf, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True, download_name=cc.nome_planilha(pacote))


@app.route("/api/cartas/disciplinas/atualizar", methods=["POST"])
def api_cartas_disciplinas_atualizar():
    """Relê só a lista de tipos de projeto (raiz PROJETOS do AutoDoc)."""
    import carta_convite as cc
    pacote = _carta_pacote()
    if not pacote:
        return jsonify({"ok": False, "error": "Pacote de carta convite não encontrado."}), 404
    if not _cartas_sync_lock.acquire(blocking=False):
        return jsonify({"ok": False, "error": "Já existe uma leitura do AutoDoc em andamento."}), 409
    try:
        r = cc.atualizar_disciplinas(pacote)
    finally:
        _cartas_sync_lock.release()
    return jsonify(r), (200 if r.get("ok") else 502)


def _carta_opcoes(body):
    def _data(s):
        try:
            return datetime.date.fromisoformat(str(s or "").strip())
        except ValueError:
            return None
    opcoes = {k: str(body.get(k) or "").strip() for k in ("empresa", "ac", "eng_responsavel", "telefone", "revisao")}
    opcoes["data"] = _data(body.get("data"))
    prazo = _data(body.get("prazo_proposta"))
    opcoes["prazo_proposta"] = f"**{prazo.strftime('%d/%m/%Y')}**" if prazo else ""
    return opcoes


@app.route("/api/cartas/enviar", methods=["POST"])
def api_cartas_enviar():
    """Gera a carta, baixa as pranchas marcadas do AutoDoc e envia por
    e-mail a partir da caixa informada — em thread de fundo; o front
    acompanha por /api/cartas/estado (campo envio)."""
    import carta_convite as cc
    pacote = _carta_pacote()
    if not pacote:
        return jsonify({"ok": False, "error": "Pacote de carta convite não encontrado."}), 404
    body = request.get_json(force=True, silent=True) or {}
    em = body.get("email") or {}
    remetente = str(em.get("remetente") or "").strip()
    senha = str(em.get("senha") or "")
    if not cc.destinatarios_de(em.get("destinatarios")):
        return jsonify({"ok": False, "error": "Informe ao menos um e-mail de destinatário."}), 400
    if not remetente or not (senha or cc.senha_email_salva(remetente)):
        return jsonify({"ok": False, "error": "Informe o e-mail e a senha do remetente."}), 400
    salvar = {k: em.get(k) for k in ("remetente", "nome_remetente", "smtp_host", "smtp_port", "destinatarios",
                                      "assunto", "mensagem")}
    salvar["anexar_pranchas"] = bool(body.get("anexar_pranchas", True))
    # assunto/mensagem iguais ao padrão não são gravados — assim continuam
    # acompanhando o tipo de projeto marcado (título/frentes mudam)
    padrao = cc.config_email(pacote, st={**cc.estado_pacote(pacote), "email": {}})
    for k in ("assunto", "mensagem"):
        if (salvar.get(k) or "").strip() == (padrao.get(k) or "").strip():
            salvar[k] = ""
    cc._atualizar_estado(pacote, email=salvar)
    if em.get("lembrar") and senha:
        cc.salvar_senha_email(remetente, senha)
    elif em.get("lembrar") is False:
        cc.esquecer_senha_email(remetente)
    if not _cartas_envio_lock.acquire(blocking=False):
        return jsonify({"ok": False, "error": "Já existe um envio em andamento."}), 409
    opcoes = _carta_opcoes(body)

    def _bg():
        try:
            cc.enviar_carta(pacote, opcoes, salvar, senha or None, anexar_pranchas=salvar["anexar_pranchas"])
        except Exception as exc:
            traceback.print_exc()
            cc._atualizar_estado(pacote, envio={"status": "erro", "etapa": "", "erros": [str(exc)],
                                                "fim": time.strftime("%Y-%m-%d %H:%M:%S")})
        finally:
            _cartas_envio_lock.release()

    cc._atualizar_estado(pacote, envio={"status": "enviando", "etapa": "Iniciando...", "enviados": [], "erros": []})
    threading.Thread(target=_bg, daemon=True).start()
    return jsonify({"ok": True, "background": True})


@app.route("/api/cartas/gerar", methods=["POST"])
def api_cartas_gerar():
    import carta_convite as cc
    pacote = _carta_pacote()
    if not pacote:
        return jsonify({"ok": False, "error": "Pacote de carta convite não encontrado."}), 404
    body = request.get_json(force=True, silent=True) or {}
    opcoes = _carta_opcoes(body)
    try:
        buf, avisos = cc.gerar_docx(pacote, opcoes=opcoes)
    except Exception as ex:
        traceback.print_exc()
        return jsonify({"ok": False, "error": f"Erro ao gerar carta convite: {ex}"}), 500
    cc._atualizar_estado(pacote, ultima_geracao=time.strftime("%Y-%m-%d %H:%M:%S"), ultimos_avisos=avisos)
    resp = send_file(buf, mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                     as_attachment=True, download_name=cc.nome_arquivo(pacote, opcoes))
    resp.headers["X-Carta-Avisos"] = str(len(avisos))
    return resp


# ═══ DOTAÇÃO (planejado x medido x comprometido, por mês, na EAP) ══════════
def _dotacao_month_of(date_str):
    s = str(date_str or "")
    return s[:7] if len(s) >= 7 and s[4] == "-" else ""

def _dotacao_walk(node, fisico_rows):
    norm_cod = _norm_rel_code(node.get("cod"))
    f_row = fisico_rows.get(norm_cod)
    orcado = f2(node.get("orcado", 0))
    monthly_planejado = {}
    if f_row:
        for mk, pct in (f_row.get("monthly_pct") or {}).items():
            monthly_planejado[mk] = f2(orcado * pct)
    monthly_medido = {}
    monthly_comprometido = {}
    for t in (node.get("titulos") or []):
        vp = f2(t.get("vlr_pg", 0))
        if vp:
            mk = _dotacao_month_of(t.get("dt_baixa"))
            if mk:
                monthly_medido[mk] = f2(monthly_medido.get(mk, 0) + vp)
        va = f2(t.get("vlr_a_pagar", 0))
        if va:
            mk = _dotacao_month_of(t.get("dt_venc"))
            if mk:
                monthly_comprometido[mk] = f2(monthly_comprometido.get(mk, 0) + va)
    children_out = []
    for ch in (node.get("children") or []):
        child_out = _dotacao_walk(ch, fisico_rows)
        children_out.append(child_out)
        for mk, v in child_out["monthly_medido"].items():
            monthly_medido[mk] = f2(monthly_medido.get(mk, 0) + v)
        for mk, v in child_out["monthly_comprometido"].items():
            monthly_comprometido[mk] = f2(monthly_comprometido.get(mk, 0) + v)
    realizado = f2(node.get("realizado", 0))
    return {
        "cod": node.get("cod"), "nm": node.get("nm"),
        "orcado": orcado, "medido": realizado, "comprometido": f2(node.get("a_pagar", 0)),
        "saldo_a_medir": f2(orcado - realizado),
        "monthly_planejado": monthly_planejado,
        "monthly_medido": monthly_medido,
        "monthly_comprometido": monthly_comprometido,
        "children": children_out,
    }

def _dotacao_apply_indireto_pct_direto(tree_out):
    """O avanço físico do Custo Indireto (01) não é medido de forma
    independente — a obra usa o MESMO percentual medido no Custo Direto (02)
    aplicado sobre o orçado do indireto, mês a mês (overhead reconhecido
    proporcional ao avanço da obra em si)."""
    direto = next((n for n in tree_out if str(n.get("cod", "")).strip() == "02"), None)
    indireto = next((n for n in tree_out if str(n.get("cod", "")).strip() == "01"), None)
    if not direto or not indireto:
        return
    orcado_direto = f2(direto.get("orcado", 0))
    if orcado_direto <= 0:
        return
    pct_by_month = {m: (v / orcado_direto) for m, v in (direto.get("monthly_planejado") or {}).items()}
    def apply(node):
        node["monthly_planejado"] = {m: f2(f2(node.get("orcado", 0)) * p) for m, p in pct_by_month.items() if p}
        for ch in (node.get("children") or []):
            apply(ch)
    apply(indireto)

def _load_planejamento_financeiro_curva():
    """Curva de desembolso planejado do cronograma (aba feng_plan, importada
    do MS Project/Sienge), agregada em % do total do projeto por mês.
    Usada como base da Previsão de desembolso na Dotação — o detalhamento
    mês a mês POR ITEM da EAP nessa fonte é esparso demais pra ser usável
    (só um punhado de ~770 itens tem distribuição mensal preenchida), então
    a curva é do projeto inteiro, aplicada proporcionalmente ao orçado de
    cada item."""
    out = {"ok": False, "months": {}, "source": ""}
    try:
        excel_path = resolve_excel_path()
        if not excel_path or not os.path.exists(excel_path):
            return out
        xl = pd.ExcelFile(excel_path, engine="openpyxl")
        if "feng_plan" not in xl.sheet_names:
            out["error"] = "Aba feng_plan não encontrada na base do Sienge."
            return out
        plan = xl.parse("feng_plan")
        plan["mesplan"] = pd.to_datetime(plan.get("mesplan"), errors="coerce")
        plan["vlrplan"] = pd.to_numeric(plan.get("vlrplan"), errors="coerce").fillna(0)
        src = plan[plan["mesplan"].notna()].copy()
        if src.empty:
            out["error"] = "Cronograma sem datas de planejamento mensal preenchidas."
            return out
        grp = src.groupby(src["mesplan"].dt.to_period("M"))["vlrplan"].sum()
        total = float(grp.sum())
        if total <= 0:
            out["error"] = "Cronograma sem valores planejados."
            return out
        months = {str(k): f2(v / total) for k, v in grp.items()}
        out.update({"ok": True, "months": months, "source": os.path.basename(excel_path)})
    except Exception as exc:
        out["error"] = str(exc)
    return out

def _dotacao_apply_previsao_desembolso(tree_out, pct_by_month):
    """Previsão de desembolso = curva de planejamento do cronograma (% do
    total do projeto por mês) aplicada sobre o orçado de cada item —
    substitui qualquer cálculo derivado de medido/comprometido."""
    def apply(node):
        node["monthly_previsao"] = {m: f2(f2(node.get("orcado", 0)) * p) for m, p in pct_by_month.items() if p}
        for ch in (node.get("children") or []):
            apply(ch)
    for n in tree_out:
        apply(n)

@app.route("/api/dotacao")
def api_dotacao():
    data = load_dashboard_cache() or {}
    orc_rel_tree = data.get("orc_rel_tree") or []
    if not orc_rel_tree:
        return jsonify({"ok": False, "error": "Orçamento ainda não carregado — abra a aba Orçamento primeiro."}), 400
    fisico = _load_dotacao_fisico_preliminar()
    tree_out = [_dotacao_walk(n, fisico.get("rows") or {}) for n in orc_rel_tree]
    _dotacao_apply_indireto_pct_direto(tree_out)
    planejamento = _load_planejamento_financeiro_curva()
    _dotacao_apply_previsao_desembolso(tree_out, planejamento.get("months") or {})
    return jsonify({
        "ok": True,
        "months": fisico.get("months", []),
        "tree": tree_out,
        "previsao_ok": planejamento.get("ok", False),
        "previsao_source": planejamento.get("source", ""),
        "previsao_error": planejamento.get("error", ""),
        "fisico_ok": fisico.get("ok", False),
        "fisico_source": fisico.get("source", ""),
        "fisico_error": fisico.get("error", ""),
    })


@app.route("/api/mapa_controle")
def api_mapa_controle():
    return jsonify(_load_mapa_controle())

@app.route("/api/export/mapa-controle.xlsx")
def api_export_mapa_controle():
    data = _load_mapa_controle()
    if not data.get("ok"):
        return jsonify({"ok": False, "error": data.get("error") or "Mapa de Controle indisponível."}), 500

    status = request.args.get("status", "").strip()
    q = request.args.get("q", "").strip().lower()

    def matches(row):
        if status == "PC_SEM_NF":
            if row["status"] not in ("Aguardando entrega", "Atrasado"):
                return False
        elif status == "Atraso de entrada":
            if row.get("atraso_entrada_status") != "Atraso de entrada":
                return False
        elif status and row["status"] != status:
            return False
        if q:
            aprop_txt = " ".join(
                str(a.get(k, "")) for a in (row.get("apropriacoes") or []) for k in ("refitemorc", "apropriacao")
            )
            hay = " ".join(
                str(row.get(k, "")) for k in ("item", "sc", "pc", "insumo_cod", "insumo", "nf", "status", "titulo_gerado")
            ).lower() + " " + aprop_txt.lower()
            if q not in hay:
                return False
        return True

    rows = [r for r in data.get("rows", []) if matches(r)]

    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Mapa de Controle"
    ws.append(["MAPA DE CONTROLE - SC → PC → ENTREGA/NF - SPE T2 BUENO"])
    ws.append(["Gerado em", time.strftime("%d/%m/%Y %H:%M")])
    ws.append(["Filtro de status", status or "Todos"])
    ws.append(["Busca", q or "—"])
    ws.append([])
    headers = [
        "Item", "SC", "Data SC", "Cód. insumo", "Insumo", "Qt. solicitada",
        "PC", "Data PC", "Previsão entrega", "Qt. entregue", "Saldo",
        "NF", "Título gerado", "Apropriação", "Valor apropriado", "Data NF", "Data entrada obra",
        "Prazo de entrada", "Data vencimento", "Status",
    ]
    ws.append(headers)
    header_row = ws.max_row
    header_fill = PatternFill("solid", fgColor="A50E24")
    thin = Side(style="thin", color="D7DEE8")
    for cell in ws[header_row]:
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(bottom=thin)
    ws.freeze_panes = f"A{header_row + 1}"

    status_fill = {
        "Entregue": "E7F5EC",
        "Aguardando entrega": "EAF0F6",
        "Aguardando PC": "FBF0E4",
        "Atrasado": "FBEBEA",
    }
    for row in rows:
        r = ws.max_row + 1
        apropriacoes = row.get("apropriacoes") or []
        aprop_nomes = ", ".join(dict.fromkeys(a.get("apropriacao", "") for a in apropriacoes if a.get("apropriacao")))
        aprop_valor = sum(a.get("vlr_ap", 0) or 0 for a in apropriacoes)
        atraso_status = row.get("atraso_entrada_status", "")
        atraso_dias = row.get("atraso_entrada_dias")
        prazo_entrada = f"{atraso_status} ({atraso_dias}d)" if atraso_status and atraso_dias is not None else ""
        ws.append([
            row.get("item", ""), row.get("sc", ""), row.get("data_sc", ""),
            row.get("insumo_cod", ""), row.get("insumo", ""), row.get("qtd_solicitada", 0),
            row.get("pc", ""), row.get("data_pc", ""), row.get("previsao_entrega", ""),
            row.get("qtd_entregue", 0), row.get("saldo", 0),
            row.get("nf", ""), row.get("titulo_gerado", ""), aprop_nomes, aprop_valor, row.get("data_nf", ""),
            row.get("data_entrada_obra", ""), prazo_entrada, row.get("data_vencimento", ""), row.get("status", ""),
        ])
        fill = status_fill.get(row.get("status", ""))
        if fill:
            for cell in ws[r]:
                cell.fill = PatternFill("solid", fgColor=fill)
        for cell in ws[r]:
            cell.border = Border(bottom=Side(style="hair", color="E7EBF0"))

    ws.auto_filter.ref = f"A{header_row}:T{ws.max_row}"
    widths = [8, 10, 12, 12, 40, 14, 10, 12, 16, 14, 12, 14, 14, 34, 16, 12, 16, 20, 16, 18]
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    ws.sheet_view.showGridLines = False
    ws["A1"].font = Font(size=15, bold=True, color="A50E24")
    ws.merge_cells("A1:T1")

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    filename = f"Mapa_Controle_{time.strftime('%Y%m%d_%H%M')}.xlsx"
    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/api/locacoes/itens-disponiveis")
def api_locacoes_itens_disponiveis():
    """Lista os nomes de item realmente usados no Sienge (feng_itensobra),
    pra popular sugestões no campo "adicionar manualmente" — sem isso o
    usuário digita um nome que não bate com nada e a inclusão não aparece
    em lugar nenhum, sem nenhum aviso do motivo."""
    try:
        excel_path = resolve_excel_path()
        if not os.path.exists(excel_path):
            return jsonify({"ok": False, "itens": []})
        xl = pd.ExcelFile(excel_path, engine="openpyxl")
        itens_obra = xl.parse("feng_itensobra")
        nomes = sorted({str(n).strip() for n in itens_obra.get("descitem", pd.Series(dtype=str)).dropna() if str(n).strip()})
        return jsonify({"ok": True, "itens": nomes})
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc), "itens": []})


@app.route("/api/locacoes/item", methods=["POST"])
def api_locacoes_item():
    payload = request.get_json(silent=True) or {}
    nome = str(payload.get("item") or "").strip()
    acao = str(payload.get("acao") or "").strip()
    if not nome or acao not in {"incluir", "excluir", "remover_excecao", "remover_inclusao"}:
        return jsonify({"ok": False, "error": "Parâmetros inválidos."}), 400
    state = _load_locacoes_state()
    incluidos = set(state.get("incluidos_manual", []))
    excluidos = set(state.get("excluidos_manual", []))
    if acao == "incluir":
        incluidos.add(nome); excluidos.discard(nome)
    elif acao == "excluir":
        excluidos.add(nome); incluidos.discard(nome)
    elif acao == "remover_excecao":
        excluidos.discard(nome)
    elif acao == "remover_inclusao":
        incluidos.discard(nome)
    _save_locacoes_state({
        "incluidos_manual": sorted(incluidos),
        "excluidos_manual": sorted(excluidos),
        "categorias_manual": state.get("categorias_manual", {}),
    })

    # o recalculo completo (build_data force=True) le o Excel inteiro de novo
    # e hoje leva ~1 min — rodar isso na resposta travava a tela só pra
    # marcar "não é locação". A troca fica salva na hora (acima); o
    # recalculo roda em segundo plano e os números batem certinho assim que
    # terminar (a tela já reflete a mudança na hora, de forma otimista).
    def _rebuild_bg():
        global _cache
        _cache = {"data": None, "ts": 0}
        try:
            build_data(force=True)
        except Exception as exc:
            print(f"  [LOCACOES] falha ao recalcular em segundo plano: {exc}")
    threading.Thread(target=_rebuild_bg, daemon=True).start()

    return jsonify({"ok": True, "background": True, "item": nome, "acao": acao})


@app.route("/api/locacoes/categoria", methods=["POST"])
def api_locacoes_categoria():
    payload = request.get_json(silent=True) or {}
    nome = str(payload.get("item") or "").strip()
    categoria = str(payload.get("categoria") or "").strip()
    if not nome or categoria not in {"escoramento", "ferramentas"}:
        return jsonify({"ok": False, "error": "Parâmetros inválidos."}), 400
    state = _load_locacoes_state()
    categorias = dict(state.get("categorias_manual", {}))
    categorias[nome] = categoria
    _save_locacoes_state({
        "incluidos_manual": state.get("incluidos_manual", []),
        "excluidos_manual": state.get("excluidos_manual", []),
        "categorias_manual": categorias,
    })

    def _rebuild_bg():
        global _cache
        _cache = {"data": None, "ts": 0}
        try:
            build_data(force=True)
        except Exception as exc:
            print(f"  [LOCACOES] falha ao recalcular categoria em segundo plano: {exc}")
    threading.Thread(target=_rebuild_bg, daemon=True).start()

    return jsonify({"ok": True, "background": True, "item": nome, "categoria": categoria})


# ── Almoxarifado: ficha manual de entrada/saída de equipamentos, ferramentas
# e materiais. O almoxarife lança a entrada (com custo, opcionalmente puxando
# os itens direto da chave da NF) e a saída (local de aplicação + pra quem foi
# entregue). Tudo fica em almoxarifado_state.json — sem depender da planilha
# do Sienge pra existir. A planilha .xls gerada aqui serve só pra Inicialização
# de Estoque do Sienge, pra quem já tem código de barras cadastrado.
ALMOX_STATE_PATH = os.path.join(DATA_DIR, "almoxarifado_state.json")
ALMOX_CATEGORIAS = {"equipamento", "ferramenta", "material"}
_almox_lock = threading.Lock()


def _load_almox_state():
    default = {"movimentos": [], "codigos_barras": {}, "patrimonio": [], "estoque_minimo": {}}
    if not os.path.exists(ALMOX_STATE_PATH):
        return default
    try:
        with open(ALMOX_STATE_PATH, "r", encoding="utf-8-sig") as fh:
            data = json.load(fh)
        if isinstance(data, dict):
            if isinstance(data.get("movimentos"), list):
                default["movimentos"] = [m for m in data["movimentos"] if isinstance(m, dict)]
            if isinstance(data.get("codigos_barras"), dict):
                default["codigos_barras"] = {str(k): str(v) for k, v in data["codigos_barras"].items()}
            if isinstance(data.get("patrimonio"), list):
                default["patrimonio"] = [p for p in data["patrimonio"] if isinstance(p, dict)]
            if isinstance(data.get("estoque_minimo"), dict):
                default["estoque_minimo"] = {str(k): _almox_num(v) for k, v in data["estoque_minimo"].items()}
    except Exception as exc:
        print(f"  [ALMOXARIFADO] falha ao ler estado: {exc}")
    return default


def _save_almox_state(state):
    tmp = ALMOX_STATE_PATH + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(state, fh, ensure_ascii=False, indent=2)
    os.replace(tmp, ALMOX_STATE_PATH)


def _almox_num(v, default=0.0):
    try:
        f = float(v)
        return default if f != f else f
    except (TypeError, ValueError):
        return default


def _almox_item_key(categoria, nome):
    return f"{categoria}|{_ascii_upper(nome).strip()}"


def _almox_cc_param(src):
    return re.sub(r"\D", "", str(src.get("cc") or "")) or "1016"


def _almox_build(cc):
    state = _load_almox_state()
    barras = state["codigos_barras"]
    minimos = state["estoque_minimo"]
    movs_cc = [m for m in state["movimentos"] if str(m.get("cc") or "1016") == cc]
    itens = {}
    for m in movs_cc:
        categoria = m.get("categoria") or "material"
        nome = str(m.get("item_nome") or "").strip()
        if not nome:
            continue
        key = _almox_item_key(categoria, nome)
        it = itens.setdefault(key, {
            "key": key, "nome": nome, "categoria": categoria, "grupo": str(m.get("grupo") or "").strip(),
            "unidade": str(m.get("unidade") or "").strip(),
            "qt_entrada": 0.0, "valor_entrada": 0.0, "qt_saida": 0.0,
        })
        qt = _almox_num(m.get("qtd"))
        if m.get("tipo") == "entrada":
            it["qt_entrada"] += qt
            it["valor_entrada"] += qt * _almox_num(m.get("custo_unit"))
            if not it["unidade"] and m.get("unidade"):
                it["unidade"] = str(m.get("unidade") or "").strip()
            if not it["grupo"] and m.get("grupo"):
                it["grupo"] = str(m.get("grupo") or "").strip()
        else:
            it["qt_saida"] += qt
    out = []
    for it in itens.values():
        it["custo_medio"] = it["valor_entrada"] / it["qt_entrada"] if it["qt_entrada"] else 0.0
        it["saldo"] = it["qt_entrada"] - it["qt_saida"]
        it["valor_saldo"] = it["saldo"] * it["custo_medio"]
        it["codigo_barras"] = barras.get(it["key"], "")
        it["estoque_minimo"] = minimos.get(it["key"], 0.0)
        it["abaixo_minimo"] = bool(it["estoque_minimo"]) and it["saldo"] < it["estoque_minimo"]
        out.append(it)
    out.sort(key=lambda i: (i["grupo"] or "zzz", i["nome"].lower()))
    ccs = sorted({str(m.get("cc") or "1016") for m in state["movimentos"]}) or ["1016"]
    return {
        "cc": cc, "ccs": ccs, "itens": out,
        "movimentos": sorted(movs_cc, key=lambda m: (m.get("data") or "", m.get("criado_em") or ""), reverse=True),
    }


@app.route("/api/almoxarifado")
def api_almoxarifado():
    try:
        return jsonify({"ok": True, **_almox_build(_almox_cc_param(request.args))})
    except Exception as exc:
        traceback.print_exc()
        return jsonify({"ok": False, "error": str(exc)}), 500


@app.route("/api/almoxarifado/movimento", methods=["POST"])
def api_almoxarifado_movimento():
    p = request.get_json(silent=True) or {}
    tipo = str(p.get("tipo") or "")
    nome = str(p.get("item_nome") or "").strip()
    categoria = str(p.get("categoria") or "").strip().lower()
    qtd = _almox_num(p.get("qtd"))
    data = str(p.get("data") or "").strip()[:10]
    if tipo not in {"entrada", "saida"} or not nome or categoria not in ALMOX_CATEGORIAS or qtd <= 0 or not re.match(r"^\d{4}-\d{2}-\d{2}$", data):
        return jsonify({"ok": False, "error": "Preencha item, categoria, data e uma quantidade maior que zero."}), 400
    local = str(p.get("local_aplicacao") or "").strip()
    entregue_para = str(p.get("entregue_para") or "").strip()
    if tipo == "saida" and (not local or not entregue_para):
        return jsonify({"ok": False, "error": "Informe o local de aplicação e pra quem foi entregue."}), 400
    user = _current_auth_user() or {}
    mov = {
        "id": f"{int(time.time() * 1000)}-{os.urandom(3).hex()}",
        "tipo": tipo, "cc": _almox_cc_param(p),
        "item_nome": nome, "categoria": categoria, "grupo": str(p.get("grupo") or "").strip(),
        "unidade": str(p.get("unidade") or "").strip(),
        "qtd": qtd, "data": data,
        "custo_unit": _almox_num(p.get("custo_unit")) if tipo == "entrada" else None,
        "fornecedor": str(p.get("fornecedor") or "").strip() if tipo == "entrada" else "",
        "chave_nf": re.sub(r"\D", "", str(p.get("chave_nf") or "")) if tipo == "entrada" else "",
        "local_aplicacao": local, "entregue_para": entregue_para,
        "obs": str(p.get("obs") or "").strip(),
        "usuario": user.get("name") or user.get("email") or "",
        "criado_em": datetime.datetime.now().isoformat(timespec="seconds"),
    }
    with _almox_lock:
        state = _load_almox_state()
        state["movimentos"].append(mov)
        _save_almox_state(state)
    return jsonify({"ok": True, "movimento": mov})


@app.route("/api/almoxarifado/movimento/<mov_id>", methods=["DELETE"])
def api_almoxarifado_movimento_excluir(mov_id):
    with _almox_lock:
        state = _load_almox_state()
        antes = len(state["movimentos"])
        state["movimentos"] = [m for m in state["movimentos"] if m.get("id") != mov_id]
        if len(state["movimentos"]) == antes:
            return jsonify({"ok": False, "error": "Movimento não encontrado."}), 404
        _save_almox_state(state)
    return jsonify({"ok": True})


def _almox_pedidos_preview(cc):
    """Lê fsup_sol_ped (pedidos de compra já lançados no Sienge) e agrupa por
    insumo+detalhe — fonte local, sem depender da consulta por chave de NF na
    SEFAZ (que ainda está com falha de schema). A categoria é um palpite pelo
    grupo de insumo do próprio Sienge; o usuário confere e ajusta antes de
    importar."""
    path = resolve_excel_path()
    if not path or not os.path.exists(path):
        raise FileNotFoundError("Base do Sienge (bd_sienge_tmz.xlsx) não encontrada.")
    xl = pd.ExcelFile(path, engine="openpyxl")
    ped = xl.parse("fsup_sol_ped")
    ped = ped.copy()
    ped["_cc"] = ped["codccusto"].map(_almox_ped_int_str)
    ped = ped[(ped["_cc"] == cc) & ped["codpedi"].notna()]
    ped = ped[ped["statpedi"].astype(str).str.strip().str.lower() != "cancelado"]

    itens = {}
    for _, r in ped.iterrows():
        nome = _almox_ped_txt(r.get("nminsum"))
        if not nome:
            continue
        det = _almox_ped_txt(r.get("nmdetinsum"))
        nome_completo = f"{nome} - {det}" if det else nome
        grupo = _almox_ped_txt(r.get("nmgrins")).lower()
        if "ferramenta" in grupo:
            categoria = "ferramenta"
        elif "equipamento" in grupo or "locação" in grupo or "locacao" in grupo:
            categoria = "equipamento"
        else:
            categoria = "material"
        key = f"{categoria}|{nome_completo}"
        it = itens.setdefault(key, {
            "nome": nome_completo, "categoria": categoria,
            "unidade": _almox_ped_txt(r.get("codunidmed")),
            "qtd": 0.0, "valor": 0.0,
        })
        qt = _almox_num(r.get("qtpedi"))
        pr = _almox_num(r.get("vlrinsumpedi"))
        it["qtd"] += qt
        it["valor"] += qt * pr
    out = []
    for it in itens.values():
        if it["qtd"] <= 0:
            continue
        it["valor_unit"] = it["valor"] / it["qtd"] if it["qtd"] else 0.0
        del it["valor"]
        out.append(it)
    out.sort(key=lambda i: i["nome"].lower())
    return out


def _almox_ped_txt(v):
    return "" if v is None or (isinstance(v, float) and v != v) else str(v).strip()


def _almox_ped_int_str(v):
    f = _almox_num(v, None)
    return str(int(f)) if f is not None else ""


@app.route("/api/almoxarifado/importar-pedidos/preview")
def api_almox_importar_pedidos_preview():
    cc = _almox_cc_param(request.args)
    try:
        return jsonify({"ok": True, "itens": _almox_pedidos_preview(cc)})
    except Exception as exc:
        traceback.print_exc()
        return jsonify({"ok": False, "error": str(exc)}), 500


@app.route("/api/almoxarifado/importar-pedidos", methods=["POST"])
def api_almox_importar_pedidos():
    """Importa em lote os itens marcados (vindos do preview acima) como
    entrada no almoxarifado. Reimportar substitui as entradas anteriores
    dessa mesma origem (não duplica ao clicar de novo)."""
    p = request.get_json(silent=True) or {}
    cc = _almox_cc_param(p)
    data = str(p.get("data") or "").strip()[:10] or datetime.date.today().isoformat()
    user = _current_auth_user() or {}
    novos = []
    for l in (p.get("itens") or []):
        nome = str(l.get("nome") or "").strip()
        categoria = str(l.get("categoria") or "material").strip().lower()
        qtd = _almox_num(l.get("qtd"))
        if not nome or categoria not in ALMOX_CATEGORIAS or qtd <= 0:
            continue
        novos.append({
            "id": f"{int(time.time() * 1000)}-{os.urandom(3).hex()}",
            "tipo": "entrada", "cc": cc,
            "item_nome": nome, "categoria": categoria, "unidade": str(l.get("unidade") or "").strip(),
            "qtd": qtd, "data": data,
            "custo_unit": _almox_num(l.get("valor_unit")),
            "fornecedor": "", "chave_nf": "",
            "local_aplicacao": "", "entregue_para": "",
            "obs": "Importado dos pedidos de compra do Sienge",
            "usuario": user.get("name") or user.get("email") or "",
            "criado_em": datetime.datetime.now().isoformat(timespec="seconds"),
            "origem_importacao": "pedidos_sienge",
        })
    if not novos:
        return jsonify({"ok": False, "error": "Nenhum item válido pra importar."}), 400
    with _almox_lock:
        state = _load_almox_state()
        state["movimentos"] = [m for m in state["movimentos"]
                                if not (m.get("origem_importacao") == "pedidos_sienge" and str(m.get("cc") or "1016") == cc)]
        state["movimentos"].extend(novos)
        _save_almox_state(state)
    return jsonify({"ok": True, "importados": len(novos)})


@app.route("/api/almoxarifado/item/<path:key>", methods=["DELETE"])
@_require_admin
def api_almoxarifado_item_excluir(key):
    """Apaga um item do estoque de vez — some da lista, junto com TODO o
    histórico de entrada e saída dele (não é uma baixa, é remoção mesmo).
    Só admin, pra não sumir com histórico por engano."""
    cc = _almox_cc_param(request.args)
    with _almox_lock:
        state = _load_almox_state()
        antes = len(state["movimentos"])
        state["movimentos"] = [
            m for m in state["movimentos"]
            if not (_almox_item_key(m.get("categoria") or "material", m.get("item_nome") or "") == key
                    and str(m.get("cc") or "1016") == cc)
        ]
        removidos = antes - len(state["movimentos"])
        state["codigos_barras"].pop(key, None)
        _save_almox_state(state)
    if removidos == 0:
        return jsonify({"ok": False, "error": "Item não encontrado."}), 404
    return jsonify({"ok": True, "movimentos_removidos": removidos})


@app.route("/api/almoxarifado/codigo-barras", methods=["POST"])
def api_almoxarifado_codigo_barras():
    p = request.get_json(silent=True) or {}
    key = str(p.get("key") or "").strip()
    codigo = re.sub(r"\s", "", str(p.get("codigo") or ""))
    if not key:
        return jsonify({"ok": False, "error": "Item não informado."}), 400
    with _almox_lock:
        state = _load_almox_state()
        if codigo:
            state["codigos_barras"][key] = codigo
        else:
            state["codigos_barras"].pop(key, None)
        _save_almox_state(state)
    return jsonify({"ok": True, "key": key, "codigo": codigo})


@app.route("/api/almoxarifado/estoque-minimo", methods=["POST"])
def api_almoxarifado_estoque_minimo():
    p = request.get_json(silent=True) or {}
    key = str(p.get("key") or "").strip()
    if not key:
        return jsonify({"ok": False, "error": "Item não informado."}), 400
    valor = _almox_num(p.get("estoque_minimo"))
    with _almox_lock:
        state = _load_almox_state()
        if valor > 0:
            state["estoque_minimo"][key] = valor
        else:
            state["estoque_minimo"].pop(key, None)
        _save_almox_state(state)
    return jsonify({"ok": True, "key": key, "estoque_minimo": valor})


@app.route("/api/almoxarifado/custo-item", methods=["POST"])
def api_almoxarifado_custo_item():
    """Ajusta manualmente o custo unitário das entradas de um item que ainda
    estão sem custo (custo_unit 0/None) — usado pros itens que não têm
    histórico de compra no Sienge pra casar automaticamente."""
    p = request.get_json(silent=True) or {}
    nome = str(p.get("item_nome") or "").strip()
    categoria = str(p.get("categoria") or "").strip().lower()
    if not nome or not categoria:
        return jsonify({"ok": False, "error": "Item não informado."}), 400
    valor = _almox_num(p.get("custo_unit"))
    with _almox_lock:
        state = _load_almox_state()
        atualizados = 0
        for mov in state["movimentos"]:
            if (mov.get("tipo") == "entrada" and mov.get("item_nome") == nome
                    and (mov.get("categoria") or "").lower() == categoria
                    and not _almox_num(mov.get("custo_unit"))):
                mov["custo_unit"] = valor
                atualizados += 1
        _save_almox_state(state)
    if not atualizados:
        return jsonify({"ok": False, "error": "Nenhuma entrada sem custo encontrada pra esse item."}), 404
    return jsonify({"ok": True, "atualizados": atualizados})


# ── Patrimônio: cadastro de bens duráveis (equipamentos/ferramentas próprios),
# separado do estoque consumível — aqui é registro individual (com nº de
# patrimônio, valor de aquisição, localização, responsável), não saldo por
# quantidade.
ALMOX_PATRIMONIO_SITUACOES = {"ativo", "manutencao", "baixado"}


@app.route("/api/almoxarifado/patrimonio")
def api_almoxarifado_patrimonio_listar():
    cc = _almox_cc_param(request.args)
    state = _load_almox_state()
    itens = [p for p in state["patrimonio"] if str(p.get("cc") or "1016") == cc]
    itens.sort(key=lambda p: (p.get("nome") or "").lower())
    return jsonify({"ok": True, "itens": itens})


@app.route("/api/almoxarifado/patrimonio", methods=["POST"])
def api_almoxarifado_patrimonio_criar():
    p = request.get_json(silent=True) or {}
    nome = str(p.get("nome") or "").strip()
    categoria = str(p.get("categoria") or "equipamento").strip().lower()
    if not nome or categoria not in ALMOX_CATEGORIAS:
        return jsonify({"ok": False, "error": "Informe o nome e a categoria do item."}), 400
    user = _current_auth_user() or {}
    item = {
        "id": f"{int(time.time() * 1000)}-{os.urandom(3).hex()}",
        "cc": _almox_cc_param(p),
        "nome": nome, "categoria": categoria,
        "numero_patrimonio": str(p.get("numero_patrimonio") or "").strip(),
        "valor_aquisicao": _almox_num(p.get("valor_aquisicao")),
        "data_aquisicao": str(p.get("data_aquisicao") or "").strip()[:10],
        "fornecedor": str(p.get("fornecedor") or "").strip(),
        "nota_fiscal": str(p.get("nota_fiscal") or "").strip(),
        "localizacao": str(p.get("localizacao") or "").strip(),
        "responsavel": str(p.get("responsavel") or "").strip(),
        "situacao": "ativo",
        "obs": str(p.get("obs") or "").strip(),
        "usuario": user.get("name") or user.get("email") or "",
        "criado_em": datetime.datetime.now().isoformat(timespec="seconds"),
    }
    with _almox_lock:
        state = _load_almox_state()
        state["patrimonio"].append(item)
        _save_almox_state(state)
    return jsonify({"ok": True, "item": item})


@app.route("/api/almoxarifado/patrimonio/<item_id>", methods=["POST"])
def api_almoxarifado_patrimonio_atualizar(item_id):
    p = request.get_json(silent=True) or {}
    with _almox_lock:
        state = _load_almox_state()
        item = next((x for x in state["patrimonio"] if x.get("id") == item_id), None)
        if not item:
            return jsonify({"ok": False, "error": "Item não encontrado."}), 404
        if "situacao" in p and str(p["situacao"]) in ALMOX_PATRIMONIO_SITUACOES:
            item["situacao"] = str(p["situacao"])
        for campo in ("localizacao", "responsavel", "obs", "numero_patrimonio"):
            if campo in p:
                item[campo] = str(p[campo] or "").strip()
        _save_almox_state(state)
    return jsonify({"ok": True, "item": item})


@app.route("/api/almoxarifado/patrimonio/<item_id>", methods=["DELETE"])
def api_almoxarifado_patrimonio_excluir(item_id):
    with _almox_lock:
        state = _load_almox_state()
        antes = len(state["patrimonio"])
        state["patrimonio"] = [x for x in state["patrimonio"] if x.get("id") != item_id]
        if len(state["patrimonio"]) == antes:
            return jsonify({"ok": False, "error": "Item não encontrado."}), 404
        _save_almox_state(state)
    return jsonify({"ok": True})


@app.route("/api/almoxarifado/nf-lookup")
def api_almoxarifado_nf_lookup():
    """Busca uma NFe pela chave de acesso, direto na SEFAZ, pra pré-preencher
    a entrada sem digitar item por item. Usa notas_fiscais.consultar_nfe_xml_completo
    (consChNFe) — quando o CNPJ que consulta não é o emitente/destinatário de
    registro dessa nota, a SEFAZ só libera o resumo (fornecedor e valor
    total, sem os itens); nesse caso volta itens=[] e resumo_apenas=True, pra
    a tela avisar e deixar cadastrar os itens manualmente."""
    chave = re.sub(r"\D", "", str(request.args.get("chave") or ""))
    if len(chave) != 44:
        return jsonify({"ok": False, "error": "Chave de acesso precisa ter 44 dígitos."}), 400
    try:
        import notas_fiscais as nf
        resultado = nf.consultar_nfe_xml_completo(chave)
        if not resultado.get("xml"):
            return jsonify({"ok": False, "error": resultado.get("xMotivo") or "NFe não encontrada na SEFAZ pra essa chave."}), 404
        itens, fornecedor, valor_total, resumo_apenas = nf.extrair_itens_nfe(resultado["xml"])
        return jsonify({"ok": True, "fornecedor": fornecedor, "valor_total": valor_total, "itens": itens, "resumo_apenas": resumo_apenas})
    except Exception as exc:
        traceback.print_exc()
        return jsonify({"ok": False, "error": str(exc)}), 500


@app.route("/api/almoxarifado/nf-xml-parse", methods=["POST"])
def api_almoxarifado_nf_xml_parse():
    """Lê os itens direto de um arquivo XML de NFe que o usuário já tem em
    mãos (ex: recebido por e-mail do fornecedor) — não depende da SEFAZ nem
    do certificado, então sempre traz os produtos completos, mesmo quando a
    consulta por chave só libera o resumo."""
    p = request.get_json(silent=True) or {}
    xml_text = str(p.get("xml") or "").strip()
    if not xml_text:
        return jsonify({"ok": False, "error": "Nenhum conteúdo de XML recebido."}), 400
    try:
        import notas_fiscais as nf
        from xml.etree import ElementTree as ET
        itens, fornecedor, valor_total, resumo_apenas = nf.extrair_itens_nfe(xml_text)
        if not itens:
            return jsonify({"ok": False, "error": "Não encontrei produtos (<det>) nesse XML — confira se é o arquivo certo (a NFe completa, não um recibo ou DANFE em PDF)."}), 400
        return jsonify({"ok": True, "fornecedor": fornecedor, "valor_total": valor_total, "itens": itens})
    except ET.ParseError:
        return jsonify({"ok": False, "error": "Esse arquivo não é um XML válido."}), 400
    except Exception as exc:
        traceback.print_exc()
        return jsonify({"ok": False, "error": str(exc)}), 500


@app.route("/api/export/almoxarifado-inicializacao.xls")
def api_export_almoxarifado_inicializacao():
    """Planilha no layout da Inicialização de Estoques de Insumos do Sienge:
    A = código de barras, B = quantidade, C = preço unitário (.xls 97-2003).
    Só entram itens com código de barras cadastrado; os demais vão pra 2ª aba,
    pra completar antes de importar."""
    try:
        import xlwt
    except ImportError:
        return jsonify({"ok": False, "error": "Biblioteca xlwt não instalada (pip install xlwt)."}), 500
    cc = _almox_cc_param(request.args)
    usar_saldo = request.args.get("base") == "saldo"
    cabecalho = request.args.get("cabecalho") == "1"
    dados = _almox_build(cc)

    wb = xlwt.Workbook(encoding="utf-8")
    h = xlwt.easyxf("font: bold on")
    txt = xlwt.easyxf("", num_format_str="@")
    qtd_fmt = xlwt.easyxf("", num_format_str="0.0000")
    prc_fmt = xlwt.easyxf("", num_format_str="0.00")
    ws = wb.add_sheet("Importacao")
    pend = wb.add_sheet("Sem codigo de barras")
    for i, w in enumerate((22, 14, 16)):
        ws.col(i).width = 256 * w
    for i, (t, w) in enumerate([("Categoria", 14), ("Item", 40), ("Unid.", 7), ("Quantidade", 14), ("Preço unitário", 16)]):
        pend.write(0, i, t, h)
        pend.col(i).width = 256 * w
    r = 0
    if cabecalho:
        for i, t in enumerate(("Código de Barras", "Quantidade", "Preço Unitário")):
            ws.write(0, i, t, h)
        r = 1
    rp = 1
    for it in dados["itens"]:
        qt = it["saldo"] if usar_saldo else it["qt_entrada"]
        if qt <= 0:
            continue
        preco = round(it["custo_medio"], 2)
        if it["codigo_barras"]:
            ws.write(r, 0, it["codigo_barras"], txt)
            ws.write(r, 1, round(qt, 4), qtd_fmt)
            ws.write(r, 2, preco, prc_fmt)
            r += 1
        else:
            for i, v in enumerate((it["categoria"], it["nome"], it["unidade"])):
                pend.write(rp, i, v)
            pend.write(rp, 3, round(qt, 4), qtd_fmt)
            pend.write(rp, 4, preco, prc_fmt)
            rp += 1
    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    nome = f"inicializacao_almoxarifado_{cc}_{datetime.date.today():%Y%m%d}.xls"
    return send_file(bio, as_attachment=True, download_name=nome, mimetype="application/vnd.ms-excel")


@app.route("/api/notas-fiscais/status")
def api_notas_fiscais_status():
    import notas_fiscais as nf
    return jsonify(nf.certificado_status())


@app.route("/api/notas-fiscais")
def api_notas_fiscais():
    import notas_fiscais as nf
    state = nf._load_state()
    notas = [{k: v for k, v in n.items() if k != "xml"} for n in (state.get("notas") or {}).values()]
    notas.sort(key=lambda n: n.get("data_emissao", ""), reverse=True)
    return jsonify({
        "ok": True,
        "certificado": nf.certificado_status(),
        "ultima_sync_nfe": state.get("ultima_sync_nfe", ""),
        "ultima_sync_nfse": state.get("ultima_sync_nfse", ""),
        "nfe_bloqueado_ate": state.get("nfe_bloqueado_ate", ""),
        "ultima_tentativa_nfe": state.get("ultima_tentativa_nfe", ""),
        "notas": notas,
    })


@app.route("/api/notas-fiscais/marcar-custo-obra", methods=["POST"])
def api_notas_fiscais_marcar_custo_obra():
    """Marca (ou desmarca) manualmente que uma nota sinalizada como
    'Possível Custo Obra' NÃO é, de fato, custo da obra 1016 — descarte
    fica salvo e sobrevive a novas sincronizações (ver o carry-over em
    api_notas_fiscais_sincronizar). Aplica em TODAS as notas do mesmo CNPJ
    de uma vez (o mesmo fornecedor não é/é da obra em todas as notas dele,
    não faz sentido confirmar nota por nota)."""
    import notas_fiscais as nf
    payload = request.get_json(silent=True) or {}
    tipo = str(payload.get("tipo") or "").strip()
    chave = str(payload.get("chave") or "").strip()
    numero = str(payload.get("numero") or "").strip()
    cnpj_fornecedor = re.sub(r"\D", "", str(payload.get("cnpj_fornecedor") or ""))
    descartar = bool(payload.get("descartar", True))
    if not tipo or (not chave and not numero):
        return jsonify({"ok": False, "error": "Informe tipo e chave (ou número) da nota."}), 400

    key = f"{tipo}:{chave}" if chave else f"{tipo}:{cnpj_fornecedor}:{numero}"
    state = nf._load_state()
    notas_dict = state.get("notas") or {}
    nota = notas_dict.get(key)
    if not nota:
        return jsonify({"ok": False, "error": "Nota não encontrada."}), 404

    credores_obra = None
    cc1016_nm = "OBRA - SPE T2 BUENO"
    if not descartar:
        try:
            credores_obra = _credores_obra_1016()
            cc1016_nm = next(
                (str(r.get("nm", "")) for r in (load_dashboard_cache() or {}).get("tit_cc") or []
                 if str(r.get("cod")) == "1016"), cc1016_nm
            )
        except Exception:
            credores_obra = None

    alvo_cnpj = re.sub(r"\D", "", str(nota.get("cnpj_fornecedor") or cnpj_fornecedor or ""))
    afetadas = []
    for outra in notas_dict.values():
        if alvo_cnpj and re.sub(r"\D", "", str(outra.get("cnpj_fornecedor") or "")) != alvo_cnpj:
            continue
        if not alvo_cnpj and outra is not nota:
            continue
        outra["custo_obra_descartado"] = descartar
        if descartar:
            outra["possivel_custo_obra"] = False
        elif credores_obra is not None:
            # desfazer o descarte: recalcula do zero (pode ter deixado de
            # bater com os credores da obra desde que foi descartada, ou o
            # contrario)
            outra["possivel_custo_obra"] = bool(
                _eh_possivel_custo_obra(outra.get("cnpj_fornecedor"), outra.get("fornecedor"), credores_obra)
                and outra.get("centro_custo") != cc1016_nm
            )
        afetadas.append({k: v for k, v in outra.items() if k != "xml"})

    state["notas"] = notas_dict
    nf._save_state(state)
    return jsonify({"ok": True, "nota": {k: v for k, v in nota.items() if k != "xml"}, "notas_afetadas": afetadas})


@app.route("/api/export/notas-fiscais.xlsx")
def api_export_notas_fiscais():
    import notas_fiscais as nf
    state = nf._load_state()
    notas = [{k: v for k, v in n.items() if k != "xml"} for n in (state.get("notas") or {}).values()]
    notas.sort(key=lambda n: n.get("data_emissao", ""), reverse=True)

    status_f = request.args.get("status", "").strip()
    lancado_f = request.args.get("lancado", "").strip()
    custo_obra_f = request.args.get("custo_obra", "").strip()
    tipo_f = request.args.get("tipo", "").strip()
    q = request.args.get("q", "").strip().lower()

    def matches(n):
        if status_f == "normal" and n.get("cancelada"):
            return False
        if status_f == "cancelada" and not n.get("cancelada"):
            return False
        if lancado_f == "lancada" and not n.get("lancado_sienge"):
            return False
        if lancado_f == "pendente" and n.get("lancado_sienge"):
            return False
        if custo_obra_f == "custo_obra" and not n.get("possivel_custo_obra"):
            return False
        if tipo_f in ("NFe", "NFSe") and n.get("tipo") != tipo_f:
            return False
        if q:
            hay = " ".join(str(n.get(k, "")) for k in ("numero", "fornecedor", "titulo_sienge")).lower()
            if q not in hay:
                return False
        return True

    rows = [n for n in notas if matches(n)]

    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Notas Fiscais"
    ws.append(["CONTROLE DE NOTAS FISCAIS (NFe + NFSe) x SIENGE - SPE T2 BUENO"])
    ws.append(["Gerado em", time.strftime("%d/%m/%Y %H:%M")])
    filtros_ativos = ", ".join(f"{lbl}: {val}" for lbl, val in [
        ("Status", status_f), ("Lançado", lancado_f), ("Possível custo obra", custo_obra_f), ("Tipo", tipo_f),
    ] if val) or "Todas"
    ws.append(["Filtro", filtros_ativos])
    ws.append(["Busca", q or "—"])
    ws.append([])
    headers = ["Tipo", "Número", "Chave/RPS", "Fornecedor", "CNPJ fornecedor", "Valor", "Emissão", "Status", "Lançado no Sienge", "Título Sienge", "Centro de Custo", "Apropriação", "Possível Custo Obra"]
    ws.append(headers)
    header_row = ws.max_row
    header_fill = PatternFill("solid", fgColor="A50E24")
    thin = Side(style="thin", color="D7DEE8")
    for cell in ws[header_row]:
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(bottom=thin)
    ws.freeze_panes = f"A{header_row + 1}"

    for r in rows:
        row = ws.max_row + 1
        ws.append([
            r.get("tipo", ""), r.get("numero", ""), r.get("chave", ""),
            r.get("fornecedor", ""), r.get("cnpj_fornecedor", ""), r.get("valor", 0),
            r.get("data_emissao", ""), "Cancelada" if r.get("cancelada") else "Normal",
            "Sim" if r.get("lancado_sienge") else "Não", r.get("titulo_sienge", ""),
            r.get("centro_custo", "") if r.get("lancado_sienge") else "",
            r.get("apropriacao", "") if r.get("lancado_sienge") else "",
            "Sim" if r.get("possivel_custo_obra") else "",
        ])
        fill = "FBEBEA" if r.get("cancelada") else ("E7F5EC" if r.get("lancado_sienge") else "FBF0E4")
        for cell in ws[row]:
            cell.fill = PatternFill("solid", fgColor=fill)
            cell.border = Border(bottom=Side(style="hair", color="E7EBF0"))
        ws.cell(row, 6).number_format = 'R$ #,##0.00;[Red]-R$ #,##0.00'

    ws.auto_filter.ref = f"A{header_row}:M{ws.max_row}"
    widths = [8, 10, 52, 40, 18, 14, 22, 12, 16, 14, 26, 26, 18]
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width
    ws.sheet_view.showGridLines = False
    ws["A1"].font = Font(size=15, bold=True, color="A50E24")
    ws.merge_cells("A1:M1")

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    filename = f"Notas_Fiscais_{time.strftime('%Y%m%d_%H%M')}.xlsx"
    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/api/notas-fiscais/xml")
def api_notas_fiscais_xml():
    import notas_fiscais as nf
    tipo = request.args.get("tipo", "")
    numero = request.args.get("numero", "")
    chave = request.args.get("chave", "")
    if tipo not in {"NFe", "NFSe"} or not numero:
        return jsonify({"ok": False, "error": "Parâmetros inválidos."}), 400

    notas = (nf._load_state().get("notas") or {})
    item = notas.get(f"{tipo}:{chave}") if chave else None
    if not item:
        # numero sozinho não é único entre fornecedores diferentes — usa a
        # chave quando disponível; sem ela, pega a 1ª nota desse tipo/número.
        item = next(
            (v for v in notas.values() if v.get("tipo") == tipo and str(v.get("numero")) == numero
             and (not chave or v.get("chave") == chave)),
            None,
        )
    if not item:
        return jsonify({"ok": False, "error": "Nota não encontrada — sincronize novamente."}), 404

    if tipo == "NFSe":
        xml_text = item.get("xml")
        if not xml_text:
            return jsonify({"ok": False, "error": "XML não guardado para esta nota (sincronize de novo)."}), 404
    else:
        cert = nf.certificado_status()
        if not cert.get("ok"):
            return jsonify({"ok": False, "error": cert.get("error") or "Certificado não configurado."}), 400
        chave = item.get("chave", "")
        if not chave:
            return jsonify({"ok": False, "error": "Esta NFe não tem chave de acesso registrada."}), 400
        try:
            resultado = nf.consultar_nfe_xml_completo(chave)
        except Exception as exc:
            return jsonify({"ok": False, "error": f"Falha ao buscar o XML na SEFAZ: {exc}"}), 502
        if not resultado or not resultado.get("xml"):
            motivo = (resultado or {}).get("xMotivo") or "SEFAZ não retornou o documento."
            return jsonify({"ok": False, "error": f"SEFAZ: {motivo}"}), 502
        xml_text = resultado["xml"]

    resp = make_response(xml_text)
    resp.headers["Content-Type"] = "application/xml; charset=utf-8"
    resp.headers["Content-Disposition"] = f'attachment; filename="{tipo}_{numero}.xml"'
    return resp


def _fmt_cnpj_cpf(digits):
    d = re.sub(r"\D", "", str(digits or ""))
    if len(d) == 14:
        return f"{d[0:2]}.{d[2:5]}.{d[5:8]}/{d[8:12]}-{d[12:14]}"
    if len(d) == 11:
        return f"{d[0:3]}.{d[3:6]}.{d[6:9]}-{d[9:11]}"
    return digits or "—"


def _extrair_extras_nfse(xml_text):
    """Complementa o comprovante em PDF com campos que só existem no XML
    completo da NFSe (já salvo junto com a nota) — descrição do serviço e
    ISSQN — sem precisar consultar a prefeitura/ADN de novo."""
    extras = {}
    try:
        from xml.etree import ElementTree as ET
        root = ET.fromstring(xml_text)
        ns = {"n": "http://www.sped.fazenda.gov.br/nfse"}
        desc = root.findtext(".//n:xDescServ", default="", namespaces=ns)
        if desc:
            extras["descricao"] = re.sub(r"\s+", " ", desc).strip()
        v_iss = root.findtext(".//n:vISSQN", default="", namespaces=ns)
        if v_iss:
            try:
                extras["iss"] = f"R$ {float(v_iss):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            except Exception:
                pass
    except Exception:
        pass
    return extras


def _gerar_pdf_nota_fiscal(item):
    """Comprovante em PDF de uma nota já reconciliada — montado só com o que
    já está salvo localmente (sem nova consulta à SEFAZ/prefeitura), então
    funciona mesmo com o certificado indisponível ou a SEFAZ bloqueada."""
    from fpdf import FPDF

    def safe(s):
        # as fontes core do fpdf2 (Helvetica etc.) só suportam latin-1 — um
        # "—" ou outro caractere fora disso derruba a geração inteira.
        return str(s if s not in (None, "") else "-").encode("latin-1", "replace").decode("latin-1")

    pdf = FPDF(unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    pdf.set_font("Helvetica", "B", 15)
    pdf.set_text_color(165, 14, 36)
    pdf.cell(0, 9, safe(f"{item.get('tipo', '')} - Comprovante"), new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", "", 8)
    pdf.set_text_color(110, 110, 110)
    pdf.cell(0, 5, safe("Gerado pelo Dashboard de Obra (SPE T2 BUENO) a partir dos dados oficiais recebidos - nao substitui o DANFE/DANFSE."),
              new_x="LMARGIN", new_y="NEXT")
    pdf.ln(4)
    pdf.set_draw_color(165, 14, 36)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(4)

    def campo(label, valor):
        pdf.set_font("Helvetica", "B", 10)
        pdf.set_text_color(30, 30, 30)
        pdf.cell(48, 7, safe(label))
        pdf.set_font("Helvetica", "", 10)
        pdf.multi_cell(0, 7, safe(valor), new_x="LMARGIN", new_y="NEXT")

    valor_fmt = f"R$ {float(item.get('valor') or 0):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    campo("Numero:", item.get("numero"))
    campo("Chave de acesso:", item.get("chave"))
    campo("Fornecedor:", item.get("fornecedor"))
    campo("CNPJ fornecedor:", _fmt_cnpj_cpf(item.get("cnpj_fornecedor")))
    campo("Valor:", valor_fmt)
    campo("Emissao:", item.get("data_emissao"))
    campo("Status:", "CANCELADA" if item.get("cancelada") else "Normal")
    if item.get("lancado_sienge"):
        campo("Lancado no Sienge:", f"Titulo {item.get('titulo_sienge')}")
        campo("Centro de custo:", item.get("centro_custo"))
        campo("Apropriacao:", item.get("apropriacao"))
    else:
        campo("Lancado no Sienge:", "Nao lancada")

    if item.get("tipo") == "NFSe" and item.get("xml"):
        extras = _extrair_extras_nfse(item["xml"])
        if extras.get("iss"):
            campo("ISSQN:", extras["iss"])
        if extras.get("descricao"):
            pdf.ln(2)
            pdf.set_font("Helvetica", "B", 10)
            pdf.set_text_color(30, 30, 30)
            pdf.cell(0, 7, safe("Descricao do servico:"), new_x="LMARGIN", new_y="NEXT")
            pdf.set_font("Helvetica", "", 9)
            pdf.multi_cell(0, 5.5, safe(extras["descricao"]), new_x="LMARGIN", new_y="NEXT")

    return bytes(pdf.output())


def _gerar_pdf_nfse_completo(item):
    """Documento auxiliar da NFS-e no layout de grade/tabela usado pelo
    DANFSE oficial (o que a prefeitura emite) — mesmas seções, mesmos
    campos, montado com TUDO que já está no XML oficial já salvo, sem
    precisar consultar a prefeitura/ADN de novo. Não é o PDF que o
    município gera de fato (não tem QR Code nem assinatura digital deles),
    mas reproduz o layout e o conteúdo dele fielmente."""
    from fpdf import FPDF
    from xml.etree import ElementTree as ET

    ns = {"n": "http://www.sped.fazenda.gov.br/nfse"}
    root = ET.fromstring(item["xml"])
    inf = root.find("n:infNFSe", ns)
    if inf is None:
        raise ValueError("XML sem infNFSe")
    emit = inf.find("n:emit", ns)
    dps = inf.find("n:DPS/n:infDPS", ns)
    prest = dps.find("n:prest", ns) if dps is not None else None
    toma = dps.find("n:toma", ns) if dps is not None else None
    serv = dps.find("n:serv", ns) if dps is not None else None
    valores_top = inf.find("n:valores", ns)
    valores_dps = dps.find("n:valores", ns) if dps is not None else None
    ibscbs_top = inf.find("n:IBSCBS", ns)
    trib_dps = valores_dps.find("n:trib", ns) if valores_dps is not None else None

    def tx(node, path):
        return (node.findtext(path, default="", namespaces=ns) or "").strip() if node is not None else ""

    def money(v):
        try:
            return f"R$ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except Exception:
            return "-"

    def data_fmt(iso):
        m = re.match(r"(\d{4})-(\d{2})-(\d{2})[T ](\d{2}:\d{2}:\d{2})", str(iso or ""))
        return f"{m.group(3)}/{m.group(2)}/{m.group(1)} {m.group(4)}" if m else (iso or "-")

    def data_curta(iso):
        m = re.match(r"(\d{4})-(\d{2})-(\d{2})", str(iso or ""))
        return f"{m.group(3)}/{m.group(2)}/{m.group(1)}" if m else (iso or "-")

    def cep_fmt(c):
        c = re.sub(r"\D", "", c or "")
        return f"{c[:5]}-{c[5:]}" if len(c) == 8 else c

    def endereco(node, cep_no_proprio):
        if node is None:
            return "", ""
        partes = [p for p in [tx(node, "n:xLgr"), tx(node, "n:nro"), tx(node, "n:xCpl"), tx(node, "n:xBairro")] if p]
        linha = " - ".join(partes)
        cep_node = node if cep_no_proprio else node.find("n:endNac", ns)
        cep = cep_fmt(tx(cep_node, "n:CEP")) if cep_node is not None else ""
        return linha, cep

    endereco_prest, cep_prest = endereco(emit.find("n:enderNac", ns) if emit is not None else None, True)
    endereco_toma, cep_toma = endereco(toma.find("n:end", ns) if toma is not None else None, False)

    simples_map = {"1": "Nao optante", "2": "Optante - Microempreendedor Individual (MEI)",
                    "3": "Optante - Microempresa ou Empresa de Pequeno Porte (ME/EPP)"}
    regime_sn = simples_map.get(tx(prest, "n:regTrib/n:opSimpNac"), "-") if prest is not None else "-"
    chave = (inf.get("Id") or "").replace("NFS", "")
    municipio = tx(inf, "n:xLocEmi") or "-"

    def safe(s):
        return str(s if s not in (None, "") else "-").encode("latin-1", "replace").decode("latin-1")

    pdf = FPDF(unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=14)
    pdf.set_left_margin(8)
    pdf.set_right_margin(8)
    pdf.add_page()
    LARG = pdf.w - pdf.l_margin - pdf.r_margin

    def secao(titulo):
        pdf.set_fill_color(230, 230, 230)
        pdf.set_draw_color(0, 0, 0)
        pdf.set_line_width(0.2)
        pdf.set_text_color(0, 0, 0)
        pdf.set_font("Helvetica", "B", 8.5)
        pdf.cell(LARG, 5.5, safe(titulo), border=1, fill=True, align="C", new_x="LMARGIN", new_y="NEXT")

    def barra(texto):
        pdf.set_fill_color(255, 255, 255)
        pdf.set_font("Helvetica", "B", 7.5)
        pdf.set_text_color(0, 0, 0)
        pdf.cell(LARG, 5, safe(texto), border=1, align="C", new_x="LMARGIN", new_y="NEXT")

    def grade(campos, altura=10.5):
        """campos: lista de (label, valor, peso). peso = largura relativa
        (pesos iguais -> colunas iguais). Desenha uma linha de caixas com
        borda, legenda pequena em cima e valor embaixo — mesmo estilo do
        DANFSE oficial."""
        pdf.set_draw_color(0, 0, 0)
        pdf.set_line_width(0.15)
        total_peso = sum(c[2] if len(c) > 2 else 1 for c in campos)
        larguras = [LARG * ((c[2] if len(c) > 2 else 1) / total_peso) for c in campos]
        larguras[-1] += LARG - sum(larguras)  # fecha exatamente na borda direita
        x0, y0 = pdf.get_x(), pdf.get_y()
        x = x0
        for (label, valor, *_resto), w in zip(campos, larguras):
            pdf.rect(x, y0, w, altura)
            pdf.set_xy(x + 1.3, y0 + 0.8)
            pdf.set_font("Helvetica", "B", 6.3)
            pdf.set_text_color(60, 60, 60)
            pdf.cell(w - 2.6, 3, safe(label))
            pdf.set_xy(x + 1.3, y0 + 4.2)
            pdf.set_font("Helvetica", "", 8)
            pdf.set_text_color(0, 0, 0)
            pdf.multi_cell(w - 2.6, 3.4, safe(valor))
            x += w
        pdf.set_xy(x0, y0 + altura)

    # ── cabeçalho ────────────────────────────────────────────────────────
    x0, y0 = pdf.get_x(), pdf.get_y()
    pdf.set_font("Helvetica", "B", 12)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(0, 5.5, safe(f"Prefeitura Municipal de {municipio}"), new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", "", 9)
    pdf.cell(0, 5, safe("Secretaria Municipal da Fazenda"), new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", "B", 9)
    pdf.cell(0, 6, safe("NFS-e - Nota Fiscal de Servico Eletronica"), new_x="LMARGIN", new_y="NEXT")
    y1 = pdf.get_y()
    box_w = 46
    pdf.rect(x0 + LARG - box_w, y0, box_w, y1 - y0)
    pdf.set_xy(x0 + LARG - box_w, y0 + 1)
    pdf.set_font("Helvetica", "", 7)
    pdf.cell(box_w, 4, safe("Numero da Nota Fiscal"), align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.set_x(x0 + LARG - box_w)
    pdf.set_font("Helvetica", "B", 16)
    pdf.cell(box_w, 9, safe(tx(inf, "n:nNFSe")), align="C")
    pdf.set_xy(x0, y1 + 1)

    grade([
        ("Data de Geracao da NFS-e", data_fmt(tx(inf, "n:dhProc")), 1),
        ("Data de Competencia", data_curta(tx(dps, "n:dCompet")), 1),
        ("Codigo de Autenticidade", chave, 2),
    ], altura=11)

    barra(f"Consulte a autenticidade desta nota em: https://www.nfse.gov.br/consultapublica  -  Chave: {chave}")
    pdf.ln(1)

    # ── prestador ────────────────────────────────────────────────────────
    secao("IDENTIFICACAO DO PRESTADOR")
    grade([
        ("CNPJ/CPF/NIF", _fmt_cnpj_cpf(tx(emit, "n:CNPJ")), 1.3),
        ("Inscricao Municipal", tx(emit, "n:IM") or "-", 1),
        ("Telefone", tx(emit, "n:fone") or "-", 1),
    ])
    grade([("Nome/Razao Social", tx(emit, "n:xNome"), 1)], altura=9)
    if tx(emit, "n:xFant"):
        grade([("Nome Fantasia", tx(emit, "n:xFant"), 1)], altura=9)
    grade([
        ("Endereco", endereco_prest, 3),
        ("CEP", cep_prest or "-", 1),
    ], altura=9)
    grade([
        ("Cidade", municipio, 1),
        ("Estado", "GO", 1),
        ("Pais", "Brasil", 1),
    ], altura=9)
    grade([("E-mail", tx(emit, "n:email") or "-", 1)], altura=9)
    grade([("Situacao Simples Nacional", regime_sn, 1)], altura=9)

    # ── tomador ──────────────────────────────────────────────────────────
    secao("IDENTIFICACAO DO TOMADOR")
    grade([
        ("CNPJ/CPF/NIF", _fmt_cnpj_cpf(tx(toma, "n:CNPJ")), 1.3),
        ("Inscricao Municipal", tx(toma, "n:IM") or "-", 1),
        ("Telefone", tx(toma, "n:fone") or "-", 1),
    ])
    grade([("Nome/Razao Social", tx(toma, "n:xNome"), 1)], altura=9)
    grade([
        ("Endereco", endereco_toma, 3),
        ("CEP", cep_toma or "-", 1),
    ], altura=9)
    grade([("E-mail", tx(toma, "n:email") or "-", 1)], altura=9)

    barra("INTERMEDIARIO DO SERVICO NAO IDENTIFICADO NA NFS-E")
    barra("DESTINATARIO E O PROPRIO TOMADOR IDENTIFICADO NA NFS-E")

    # ── serviço ──────────────────────────────────────────────────────────
    secao("DADOS DO SERVICO PRESTADO")
    grade([
        ("Cod. Trib. Nacional", tx(serv, "n:cServ/n:cTribNac") or "-", 1),
        ("NBS", tx(serv, "n:cServ/n:cNBS") or "-", 1),
        ("Local da Prestacao", municipio, 1),
    ])
    if tx(inf, "n:xTribNac"):
        grade([("Atividade", tx(inf, "n:xTribNac"), 1)], altura=17)
    grade([
        ("Vl. do Servico", money(tx(valores_dps, "n:vServPrest/n:vServ")), 1),
        ("Vl. do Desc. Incondicionado", "-", 1),
        ("Vl. do Desc. Condicionado", "-", 1),
    ])
    grade([("Descricao do Servico", tx(serv, "n:cServ/n:xDescServ"), 1)], altura=17)
    info_comp = tx(serv, "n:infoCompl/n:xInfComp")

    # ── issqn ────────────────────────────────────────────────────────────
    secao("IMPOSTO SOBRE SERVICO DE QUALQUER NATUREZA - ISSQN")
    tp_ret_map = {"1": "Nao Retido", "2": "Retido pelo Tomador", "3": "Retido pelo Intermediario"}
    tp_ret = tp_ret_map.get(tx(trib_dps, "n:tribMun/n:tpRetISSQN"), "-") if trib_dps is not None else "-"
    grade([
        ("Municipio de Incidencia", municipio, 1.5),
        ("Tipo de Retencao", tp_ret, 1),
        ("Valor Deducao", "R$ 0,00", 1),
    ])
    grade([
        ("Base de Calculo", money(tx(valores_top, "n:vBC")), 1),
        ("Aliquota", f"{tx(valores_top, 'n:pAliqAplic') or '0'}%", 1),
        ("Vl. ISSQN", money(tx(valores_top, "n:vISSQN")), 1),
    ])

    # ── tributação nacional (Reforma Tributária) ────────────────────────
    v_cp_ret = tx(trib_dps, "n:tribFed/n:vRetCP") if trib_dps is not None else ""
    if v_cp_ret or tx(valores_top, "n:vTotalRet"):
        secao("TRIBUTACAO NACIONAL")
        grade([
            ("Tipo de Retencao", "PIS/COFINS/CSLL Nao Retidos", 2),
            ("Vl. CP Retido", money(v_cp_ret) if v_cp_ret else "-", 1),
        ], altura=9)

    # ── IBS/CBS (Reforma Tributária) ─────────────────────────────────────
    if ibscbs_top is not None:
        secao("IMPOSTO E CONTRIBUICAO SOBRE BENS E SERVICOS - IBS/CBS")
        tot = ibscbs_top.find("n:totCIBS", ns)
        v_ret_total = tx(valores_top, "n:vTotalRet")
        v_cbs = tx(tot, "n:gCBS/n:vCBS") if tot is not None else ""
        v_ibs = tx(tot, "n:gIBS/n:vIBSTot") if tot is not None else ""
        grade([
            ("Total de Retencao", money(v_ret_total) if v_ret_total else "-", 1),
            ("Valor Total do CBS", money(v_cbs) if v_cbs else "-", 1),
            ("Valor Total do IBS", money(v_ibs) if v_ibs else "-", 1),
            ("Valor Total Liquido", money(tx(valores_top, "n:vLiq")), 1),
        ])
    else:
        grade([("Valor Liquido da Nota", money(tx(valores_top, "n:vLiq")), 1)], altura=9)

    if info_comp:
        secao("INFORMACOES COMPLEMENTARES")
        pdf.set_font("Helvetica", "", 8)
        pdf.set_text_color(0, 0, 0)
        pdf.multi_cell(LARG, 4.2, safe(info_comp), border=1, new_x="LMARGIN", new_y="NEXT")

    complementar = tx(inf, "n:xOutInf")
    if complementar:
        pdf.set_font("Helvetica", "", 7.5)
        pdf.multi_cell(LARG, 4, safe(complementar), border=1, new_x="LMARGIN", new_y="NEXT")

    pdf.ln(2)
    pdf.set_font("Helvetica", "I", 7)
    pdf.set_text_color(140, 140, 140)
    pdf.multi_cell(0, 3.6, safe(
        "Documento gerado pelo Dashboard de Obra (SPE T2 BUENO) reproduzindo o layout e o conteudo do DANFSE oficial "
        "a partir do XML recebido da NFS-e Nacional (Ambiente de Dados Nacional). Nao tem QR Code nem assinatura "
        "digital do municipio - para a via oficial, use a chave de acesso acima em www.nfse.gov.br/consultapublica."
    ))

    return bytes(pdf.output())


def _localizar_nota_fiscal(tipo, numero, chave):
    import notas_fiscais as nf
    notas = (nf._load_state().get("notas") or {})
    item = notas.get(f"{tipo}:{chave}") if chave else None
    if not item:
        item = next(
            (v for v in notas.values() if v.get("tipo") == tipo and str(v.get("numero")) == numero
             and (not chave or v.get("chave") == chave)),
            None,
        )
    return item


@app.route("/api/notas-fiscais/pdf")
def api_notas_fiscais_pdf():
    tipo = request.args.get("tipo", "")
    numero = request.args.get("numero", "")
    chave = request.args.get("chave", "")
    if tipo not in {"NFe", "NFSe"} or not numero:
        return jsonify({"ok": False, "error": "Parâmetros inválidos."}), 400

    item = _localizar_nota_fiscal(tipo, numero, chave)
    if not item:
        return jsonify({"ok": False, "error": "Nota não encontrada — sincronize novamente."}), 404

    try:
        if tipo == "NFSe" and item.get("xml"):
            try:
                pdf_bytes = _gerar_pdf_nfse_completo(item)
            except Exception:
                pdf_bytes = _gerar_pdf_nota_fiscal(item)  # XML com formato inesperado — cai pro comprovante simples
        else:
            pdf_bytes = _gerar_pdf_nota_fiscal(item)
    except Exception as exc:
        return jsonify({"ok": False, "error": f"Falha ao gerar o PDF: {exc}"}), 500

    resp = make_response(pdf_bytes)
    resp.headers["Content-Type"] = "application/pdf"
    resp.headers["Content-Disposition"] = f'attachment; filename="{tipo}_{numero}.pdf"'
    return resp


@app.route("/api/notas-fiscais/portal-oficial")
def api_notas_fiscais_portal_oficial():
    """Devolve o link + a chave de acesso pra achar o PDF/DANFE OFICIAL — o
    que o órgão emissor realmente emitiu, com layout, QR Code e assinatura.
    NFSe usa a Consulta Pública NACIONAL (nfse.gov.br), cobre qualquer
    município aderido. NFe usa a Consulta NACIONAL da SEFAZ
    (nfe.fazenda.gov.br), cobre qualquer UF. Não dá pra baixar isso
    automaticamente: a consulta exige resolver um captcha manualmente, é
    uma proteção do próprio governo."""
    tipo = request.args.get("tipo", "")
    numero = request.args.get("numero", "")
    chave = request.args.get("chave", "")
    if tipo not in ("NFSe", "NFe") or not numero:
        return jsonify({"ok": False, "error": "Só disponível para NFe/NFSe."}), 400

    item = _localizar_nota_fiscal(tipo, numero, chave)
    if not item:
        return jsonify({"ok": False, "error": "Nota não encontrada."}), 404

    if tipo == "NFe":
        chave_acesso = re.sub(r"\D", "", str(item.get("chave") or ""))
        if len(chave_acesso) != 44:
            return jsonify({"ok": False, "error": "Nota sem chave de acesso de 44 dígitos registrada."}), 404
        return jsonify({
            "ok": True,
            "url": "https://www.nfe.fazenda.gov.br/portal/consultaRecaptcha.aspx?tipoConsulta=resumo&tipoConteudo=XbSeqxE8pl8=",
            "aviso": "Cole a chave abaixo no campo \"Chave de Acesso\", resolva o desafio de segurança (captcha) e clique em "
                     "Continuar. Funciona pra nota de qualquer estado do Brasil (portal nacional da SEFAZ).",
            "campos": {
                "Chave de Acesso da NF-e": chave_acesso,
            },
        })

    if not item.get("xml"):
        return jsonify({"ok": False, "error": "Nota não encontrada."}), 404
    try:
        from xml.etree import ElementTree as ET
        ns = {"n": "http://www.sped.fazenda.gov.br/nfse"}
        root = ET.fromstring(item["xml"])
        inf = root.find("n:infNFSe", ns)
        cod_autenticidade = (inf.get("Id") or "").replace("NFS", "")
    except Exception as exc:
        return jsonify({"ok": False, "error": f"Falha ao ler o XML: {exc}"}), 500

    if not cod_autenticidade:
        return jsonify({"ok": False, "error": "Nota sem chave de acesso registrada."}), 404

    return jsonify({
        "ok": True,
        "url": "https://www.nfse.gov.br/consultapublica",
        "aviso": "Cole a chave abaixo no campo \"Chave de acesso da NFS-e\" (modo \"Por Chave de Acesso\", já vem selecionado), "
                 "resolva o desafio de segurança (hCaptcha) e clique em Consultar. Funciona pra nota de qualquer município do Brasil.",
        "campos": {
            "Chave de acesso da NFS-e": cod_autenticidade,
        },
    })


@app.route("/api/notas-fiscais/sincronizar", methods=["POST"])
def api_notas_fiscais_sincronizar():
    import notas_fiscais as nf
    cert = nf.certificado_status()
    if not cert.get("ok"):
        return jsonify({"ok": False, "error": cert.get("error") or "Certificado não configurado."}), 400

    state = nf._load_state()
    now = time.strftime("%Y-%m-%d %H:%M:%S")
    erros = []

    MAX_PAGINAS = 40  # até ~2000 documentos por sincronização — bem acima de qualquer backlog normal

    notas_nfe = []
    eventos_nfe = []
    NFE_INTERVALO_MIN_MINUTOS = 65  # SEFAZ não recomenda consultar mais de 1x/hora — damos margem
    nfe_bloqueado_ate = state.get("nfe_bloqueado_ate", "")
    ultima_tentativa_nfe = state.get("ultima_tentativa_nfe", "")
    proxima_liberacao = ""
    if ultima_tentativa_nfe:
        try:
            ts = time.mktime(time.strptime(ultima_tentativa_nfe, "%Y-%m-%d %H:%M:%S"))
            proxima_liberacao = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(ts + NFE_INTERVALO_MIN_MINUTOS * 60))
        except Exception:
            proxima_liberacao = ""

    if nfe_bloqueado_ate and now < nfe_bloqueado_ate:
        # a SEFAZ já rejeitou por "Consumo Indevido" recentemente — nem tenta
        # de novo. Insistir durante o bloqueio reinicia a janela de espera,
        # que é exatamente o que estava fazendo o bloqueio nunca acabar.
        erros.append(f"SEFAZ (NFe): bloqueada por consumo indevido até {nfe_bloqueado_ate[11:16]} — aguardando, sem tentar de novo antes disso.")
    elif proxima_liberacao and now < proxima_liberacao:
        # trava própria (proativa): nem deixa consultar antes do intervalo
        # mínimo entre chamadas, pra nunca chegar a disparar o bloqueio da
        # SEFAZ em primeiro lugar — independente de já ter sido bloqueada.
        erros.append(f"SEFAZ (NFe): última consulta às {ultima_tentativa_nfe[11:16]} — aguardando intervalo mínimo de {NFE_INTERVALO_MIN_MINUTOS} min entre consultas (libera às {proxima_liberacao[11:16]}) pra não arriscar bloqueio.")
    else:
        state["ultima_tentativa_nfe"] = now
        nf._save_state(state)  # grava a tentativa já aqui, antes de chamar a SEFAZ — protege mesmo se o processo cair no meio da consulta
        try:
            ult_nsu = state.get("ult_nsu", "0")
            paginas = 0
            while paginas < MAX_PAGINAS:
                dist = nf.consultar_nfe_distribuicao(ult_nsu)
                if dist.get("cStat") == "656":
                    bloqueado_ate = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(time.time() + 70 * 60))
                    state["nfe_bloqueado_ate"] = bloqueado_ate
                    erros.append(f"SEFAZ (NFe): cStat 656 — Consumo Indevido. Bloqueada até {bloqueado_ate[11:16]}, não tente antes disso (cada tentativa durante o bloqueio reinicia a espera).")
                    break
                if dist.get("cStat") and dist["cStat"] not in ("138", "137"):
                    erros.append(f"SEFAZ (NFe): cStat {dist['cStat']} — {dist.get('xMotivo', '')}")
                    break
                notas_nfe.extend(dist.get("notas", []))
                eventos_nfe.extend(dist.get("eventos", []))
                novo_nsu = dist.get("max_nsu") or ult_nsu
                state["ult_nsu"] = novo_nsu
                state["ultima_sync_nfe"] = now
                state["nfe_bloqueado_ate"] = ""
                paginas += 1
                if str(novo_nsu) == str(ult_nsu) or not dist.get("notas"):
                    break  # sem documentos novos nesta página — chegamos ao fim do backlog
                ult_nsu = novo_nsu
                time.sleep(1)  # dá um respiro entre páginas, evita parecer "consumo indevido"
        except Exception as exc:
            erros.append(f"Falha ao consultar SEFAZ (NFe): {exc}")

    notas_nfse, eventos_nfse = [], []
    try:
        ult_nsu_nfse = state.get("ult_nsu_nfse", 0)
        paginas = 0
        while paginas < MAX_PAGINAS:
            dist_nfse = nf.consultar_nfse_nacional(ult_nsu_nfse)
            if dist_nfse.get("erros"):
                erros.append(f"NFSe (ADN nacional): {'; '.join(dist_nfse['erros'])}")
                break
            notas_nfse.extend(dist_nfse.get("notas", []))
            eventos_nfse.extend(dist_nfse.get("eventos", []))
            novo_nsu = dist_nfse.get("max_nsu", ult_nsu_nfse)
            state["ult_nsu_nfse"] = novo_nsu
            state["ultima_sync_nfse"] = now
            paginas += 1
            if novo_nsu == ult_nsu_nfse or not (dist_nfse.get("notas") or dist_nfse.get("eventos")):
                break
            ult_nsu_nfse = novo_nsu
            time.sleep(1)
    except Exception as exc:
        erros.append(f"Falha ao consultar NFSe (ADN nacional): {exc}")

    try:
        nf_titulo_candidatos = _nf_titulo_candidatos()
    except Exception:
        nf_titulo_candidatos = {}
    linhas = nf.reconciliar_com_sienge(notas_nfe, notas_nfse, nf_titulo_candidatos, eventos_nfse, eventos_nfe)

    try:
        cc_aprop_map = _nf_titulo_ccusto_apropriacao()
    except Exception:
        cc_aprop_map = {}
    try:
        credores_obra = _credores_obra_1016()
    except Exception:
        credores_obra = {"cnpjs": set(), "cnpj_roots": set(), "nomes": set()}
    try:
        _cc1016_nm = next(
            (str(r.get("nm", "")) for r in (load_dashboard_cache() or {}).get("tit_cc") or []
             if str(r.get("cod")) == "1016"), "OBRA - SPE T2 BUENO"
        )
    except Exception:
        _cc1016_nm = "OBRA - SPE T2 BUENO"
    for linha in linhas:
        info = cc_aprop_map.get(linha.get("titulo_sienge") or "", {})
        linha["centro_custo"] = info.get("centro_custo", "")
        linha["apropriacao"] = info.get("apropriacao", "")
        linha["possivel_custo_obra"] = bool(
            _eh_possivel_custo_obra(linha.get("cnpj_fornecedor"), linha.get("fornecedor"), credores_obra)
            and linha.get("centro_custo") != _cc1016_nm
        )

    notas_dict = state.get("notas") or {}
    for linha in linhas:
        # a chave de acesso (chNFe/ChaveAcesso) é única por documento; o
        # número sozinho NÃO é — dois fornecedores diferentes podem emitir,
        # cada um, sua própria NF/NFSe nº 50, e usar "tipo:numero" como
        # chave do dicionário fazia uma sobrescrever a outra (nota some).
        chave = linha.get("chave") or ""
        if chave:
            key = f"{linha['tipo']}:{chave}"
        else:
            key = f"{linha['tipo']}:{linha.get('cnpj_fornecedor', '')}:{linha['numero']}"
        # "Possível Custo Obra" pode ter sido descartado manualmente (usuário
        # confirmou que não é da obra) — esse descarte precisa sobreviver a
        # uma nova sincronização, senão a nota recalculada apaga a marcação.
        if (notas_dict.get(key) or {}).get("custo_obra_descartado"):
            linha["custo_obra_descartado"] = True
            linha["possivel_custo_obra"] = False
        notas_dict[key] = linha

    # Varredura de cancelamento retroativo: um evento de cancelamento chega
    # na mesma consulta em que é liberado pela SEFAZ, não necessariamente
    # junto com a nota original (que pode já estar guardada de uma sync
    # anterior) — sem isso, uma nota cancelada depois de já lançada aqui
    # ficava presa mostrando "normal" pra sempre, mesmo com o evento
    # chegando certinho.
    chaves_canceladas_agora = {e.get("chave") for e in eventos_nfse if e.get("chave")}
    chaves_canceladas_agora |= {
        e.get("chave") for e in eventos_nfe
        if e.get("chave") and str(e.get("tipo_evento", "")).strip() == "110111"
    }
    for nota in notas_dict.values():
        if nota.get("chave") in chaves_canceladas_agora and not nota.get("cancelada"):
            nota["cancelada"] = True

    state["notas"] = notas_dict
    nf._save_state(state)

    return jsonify({
        "ok": not erros or bool(linhas),
        "novas_notas": len(linhas),
        "erros": erros,
        "notas": list(notas_dict.values()),
    })


@app.route("/api/projetos/status")
def api_projetos_status():
    import autodoc_sync as ad
    return jsonify(ad.credenciais_status())


@app.route("/api/projetos")
def api_projetos():
    import autodoc_sync as ad
    state = ad._load_state()
    atrasados = ad.calcular_atrasados(state)
    analise = ad.analisar_documentos(state)
    return jsonify({
        "ok": True,
        "cronograma_atual": state.get("cronograma_atual", []),
        "cronograma_historico": state.get("cronograma_historico", []),
        "disciplinas": state.get("disciplinas", []),
        "disciplinas_detalhe": state.get("disciplinas_detalhe", {}),
        "sync_at": state.get("sync_at", ""),
        "sync_completo_at": state.get("sync_completo_at", ""),
        "atrasados": atrasados["atrasados"],
        "em_dia_qtd": atrasados["em_dia_qtd"],
        "total_avaliado": atrasados["total_avaliado"],
        "registro": analise["registro"],
        "status_por_disciplina": analise["status_por_disciplina"],
        "retrabalho_documentos": analise["retrabalho_documentos"],
        "retrabalho_ranking": analise["retrabalho_ranking"],
        "curva_entregas": analise["curva_entregas"],
        "total_documentos": analise["total_documentos"],
        "pct_aprovados": analise["pct_aprovados"],
    })


@app.route("/api/projetos/sincronizar-completo", methods=["POST"])
def api_projetos_sincronizar_completo():
    """Entra em cada uma das 30 pastas de disciplina (não só lista os
    nomes) e lê os documentos + histórico de cada uma. Demora ~5-10 min —
    é o que alimenta o painel de atrasados e a análise por disciplina."""
    import autodoc_sync as ad
    resultado = ad.sincronizar_projetos_completo()
    if not resultado.get("ok"):
        return jsonify(resultado), 502
    return jsonify(resultado)


@app.route("/api/projetos/sincronizar", methods=["POST"])
def api_projetos_sincronizar():
    """Loga no AutoDoc (Selenium) e relê a pasta de gestão de projetos —
    cronograma de projetos (revisão atual + histórico) e disciplinas.
    Demora ~30-60s (login real + navegação de página); a tela deve avisar
    isso enquanto espera."""
    import autodoc_sync as ad
    resultado = ad.sincronizar_projetos()
    if not resultado.get("ok"):
        return jsonify(resultado), 502
    return jsonify(resultado)


@app.route("/api/ifc/autodoc/listar")
def api_ifc_autodoc_listar():
    """Lista (sem baixar) os .ifc disponíveis na pasta dedicada de modelos
    do AutoDoc — usada pra UI mostrar o que já pode ser trazido pro
    visualizador de compatibilização."""
    import autodoc_sync as ad
    resultado = ad.listar_modelos_ifc_autodoc()
    if not resultado.get("ok"):
        return jsonify(resultado), 502
    return jsonify(resultado)


@app.route("/api/ifc/autodoc/baixar", methods=["POST"])
def api_ifc_autodoc_baixar():
    """Baixa um .ifc específico (login + clique real no AutoDoc, ~15-30s) —
    roda em thread de fundo, o front consulta /api/ifc/modelos pra saber
    quando terminou (status muda de 'baixando' pra 'pronto' ou 'erro')."""
    import autodoc_sync as ad
    payload = request.get_json(silent=True) or {}
    nome = str(payload.get("arquivo") or "").strip()
    if not nome:
        return jsonify({"ok": False, "error": "Arquivo não informado."}), 400
    entry = ad._marcar_baixando(nome)

    def _bg():
        try:
            ad.baixar_modelo_ifc(nome)
        except Exception as exc:
            ad._marcar_erro(nome, str(exc))

    threading.Thread(target=_bg, daemon=True).start()
    return jsonify({"ok": True, "background": True, "item": entry})


@app.route("/api/ifc/modelos")
def api_ifc_modelos():
    import autodoc_sync as ad
    return jsonify({"ok": True, "modelos": ad._load_ifc_index()})


@app.route("/api/ifc/modelos/<model_id>/arquivo")
def api_ifc_modelo_arquivo(model_id):
    import autodoc_sync as ad
    idx = ad._load_ifc_index()
    entry = next((e for e in idx if e["id"] == model_id), None)
    if not entry or entry.get("status") != "pronto" or not entry.get("arquivo"):
        return jsonify({"ok": False, "error": "Modelo não encontrado ou ainda não baixado."}), 404
    return send_from_directory(ad.IFC_DIR, entry["arquivo"], as_attachment=False, download_name=entry["nome"])


@app.route("/api/ifc/modelos/<model_id>", methods=["DELETE"])
def api_ifc_modelo_excluir(model_id):
    import autodoc_sync as ad
    idx = ad._load_ifc_index()
    entry = next((e for e in idx if e["id"] == model_id), None)
    if not entry:
        return jsonify({"ok": False, "error": "Modelo não encontrado."}), 404
    if entry.get("arquivo"):
        caminho = os.path.join(ad.IFC_DIR, entry["arquivo"])
        if os.path.exists(caminho):
            os.remove(caminho)
    restante = [e for e in idx if e["id"] != model_id]
    ad._save_ifc_index(restante)
    return jsonify({"ok": True})


# ── NWD/NWC (Navisworks) — não dá pra renderizar o conteúdo 3D no navegador
# (formato binário fechado da Autodesk, sem parser aberto como o do IFC
# acima), então isso é só um lugar pra guardar/organizar esses arquivos —
# quem quiser abrir de verdade usa o Navisworks. Serve de referência pra
# aba de medição física por elemento (que já lia dados/synergia pra achar
# esses arquivos, mas não tinha upload).
@app.route("/api/ifc/nwd/listar")
def api_ifc_nwd_listar():
    return jsonify({"ok": True, "arquivos": list_nwd_files()})


@app.route("/api/ifc/nwd/upload", methods=["POST"])
def api_ifc_nwd_upload():
    f = request.files.get("file")
    if not f or not f.filename:
        return jsonify({"ok": False, "error": "Nenhum arquivo enviado."}), 400
    nome = secure_filename(f.filename)
    ext = os.path.splitext(nome)[1].lower()
    if ext not in (".nwd", ".nwc"):
        return jsonify({"ok": False, "error": "Envie um arquivo .nwd ou .nwc."}), 400
    os.makedirs(NWD_DIR, exist_ok=True)
    f.save(os.path.join(NWD_DIR, nome))
    return jsonify({"ok": True, "arquivos": list_nwd_files()})


@app.route("/api/ifc/nwd/arquivo/<path:nome>")
def api_ifc_nwd_download(nome):
    nome = secure_filename(nome)
    # Só modelos .nwd/.nwc e só da pasta própria deles. Antes havia um fallback para
    # DATA_DIR que deixava qualquer usuário logado baixar access_users.json (hashes de
    # senha), caches e planilhas financeiras.
    if os.path.splitext(nome)[1].lower() not in (".nwd", ".nwc"):
        return jsonify({"ok": False, "error": "Arquivo não encontrado."}), 404
    if os.path.isfile(os.path.join(NWD_DIR, nome)):
        return send_from_directory(NWD_DIR, nome, as_attachment=True)
    return jsonify({"ok": False, "error": "Arquivo não encontrado."}), 404


@app.route("/api/ifc/nwd/<path:nome>", methods=["DELETE"])
def api_ifc_nwd_excluir(nome):
    nome = secure_filename(nome)
    if os.path.splitext(nome)[1].lower() not in (".nwd", ".nwc"):
        return jsonify({"ok": False, "error": "Arquivo inválido."}), 400
    removido = False
    for pasta in (NWD_DIR, DATA_DIR):
        caminho = os.path.join(pasta, nome)
        if not os.path.exists(caminho):
            continue
        # arquivo grande recém-baixado pode ainda estar com handle aberto no
        # Windows por um instante (streaming do download) — tenta algumas
        # vezes antes de desistir.
        for tentativa in range(5):
            try:
                os.remove(caminho)
                removido = True
                break
            except PermissionError:
                if tentativa == 4:
                    return jsonify({"ok": False, "error": "Arquivo em uso no momento — tente excluir de novo em alguns segundos."}), 409
                time.sleep(0.3)
    if not removido:
        return jsonify({"ok": False, "error": "Arquivo não encontrado."}), 404
    return jsonify({"ok": True, "arquivos": list_nwd_files()})


@app.route("/api/cronograma/sincronizar-mpp", methods=["POST"])
def api_cronograma_sincronizar_mpp():
    """Relê o CRONOGRAMA_SYNERGIA_BUENO.mpp direto do MS Project (instalado
    nesta máquina) e recarrega o cache do dashboard. Demora ~1-2 min — a
    tela deve mostrar isso claramente enquanto espera."""
    resultado = sync_mpp_cronograma()
    if not resultado.get("ok"):
        return jsonify(resultado), 502
    build_data(force=True)
    return jsonify(resultado)


if __name__ == "__main__":
    # Desenvolvimento/Windows. Em produção (Linux) quem sobe é o gunicorn (gunicorn.conf.py).
    port = int(os.environ.get("DASHBOARD_PORT", "5000") or 5000)
    host = os.environ.get("DASHBOARD_HOST", "127.0.0.1")
    print("\n"+"="*55)
    print("  Dashboard Obra v4 - SPE T2 BUENO")
    print(f"  Acesse: http://{host}:{port}/")
    print("="*55+"\n")
    app.run(debug=False, host=host, port=port, threaded=True)

