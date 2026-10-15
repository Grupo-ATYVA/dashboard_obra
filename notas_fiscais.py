"""Controle de Notas Fiscais (NFe/NFSe) emitidas contra o CNPJ da SPE T2
BUENO, confrontadas com os títulos já lançados no Sienge.

Duas fontes de notas, ambas autenticadas com o mesmo certificado e-CNPJ A1
(arquivo .pfx + senha):

  - NFe (material)      -> Distribuição DFe da SEFAZ (nacional, qualquer UF)
  - NFSe (mão de obra)   -> webservice ABRASF 2.04 da Prefeitura de Goiânia

Sem o certificado configurado (CERT_PFX_PATH + CERT_PFX_SENHA), nenhuma
consulta externa acontece — as funções de sincronização retornam um erro
claro para a tela, em vez de travar o dashboard.

IMPORTANTE: os clientes de NFe/NFSe abaixo foram implementados a partir da
documentação pública dos dois webservices (Manual de Orientação do
Contribuinte da NFe e Manual ABRASF 2.04 da Prefeitura de Goiânia), mas
nunca foram testados contra o serviço real, porque isso exige um
certificado válido que ainda não existe neste ambiente. É esperado que a
primeira sincronização de verdade precise de ajustes finos depois de ver
as respostas reais dos dois serviços.
"""

from __future__ import annotations

import base64
import difflib
import gzip
import io
import json
import os
import re
import tempfile
import time
import unicodedata
from datetime import datetime, timezone
from pathlib import Path
from xml.etree import ElementTree as ET

CNPJ_SPE = "45573855000140"
UF_AUTOR_GO = "52"

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "dados" / "synergia"
CERT_DIR = DATA_DIR / "certificado"
STATE_PATH = DATA_DIR / "notas_fiscais_state.json"
CREDOR_EQUIVALENCIAS_PATH = DATA_DIR / "credores_equivalentes.json"

CERT_PFX_PATH = os.environ.get("CERT_PFX_PATH", str(CERT_DIR / "certificado.pfx"))

NFE_DIST_DFE_URL = "https://www1.nfe.fazenda.gov.br/NFeDistribuicaoDFe/NFeDistribuicaoDFe.asmx"

NS_NFE = "http://www.portalfiscal.inf.br/nfe"

_cert_cache = {"mtime": None, "cert_path": None, "key_path": None, "info": None}


def _cert_senha():
    return os.environ.get("CERT_PFX_SENHA", "")


def _load_cert_pem():
    """Converte o .pfx (uma vez) em arquivos PEM temporários (cert + chave
    privada), reaproveitados enquanto o arquivo do certificado não mudar."""
    if not os.path.exists(CERT_PFX_PATH):
        raise FileNotFoundError(f"Certificado não encontrado em {CERT_PFX_PATH}.")
    senha = _cert_senha()
    if not senha:
        raise ValueError("Variável de ambiente CERT_PFX_SENHA não definida.")

    mtime = os.path.getmtime(CERT_PFX_PATH)
    if _cert_cache["mtime"] == mtime and _cert_cache["cert_path"] and os.path.exists(_cert_cache["cert_path"]):
        return _cert_cache["cert_path"], _cert_cache["key_path"], _cert_cache["info"]

    from cryptography.hazmat.primitives.serialization import pkcs12, Encoding, PrivateFormat, NoEncryption

    with open(CERT_PFX_PATH, "rb") as fh:
        pfx_bytes = fh.read()
    key, cert, extra_certs = pkcs12.load_key_and_certificates(pfx_bytes, senha.encode("utf-8"))
    if key is None or cert is None:
        raise ValueError("O arquivo .pfx não contém certificado e chave privada válidos (senha errada?).")

    tmp_dir = tempfile.mkdtemp(prefix="temazec_cert_")  # mkdtemp ja cria com permissao 0700
    # a chave privada em PEM nao pode ficar no disco depois que o processo acabar
    import atexit
    import shutil
    atexit.register(shutil.rmtree, tmp_dir, ignore_errors=True)
    cert_path = os.path.join(tmp_dir, "cert.pem")
    key_path = os.path.join(tmp_dir, "key.pem")
    with open(cert_path, "wb") as fh:
        fh.write(cert.public_bytes(Encoding.PEM))
        for extra in extra_certs or []:
            fh.write(extra.public_bytes(Encoding.PEM))
    key_fd = os.open(key_path, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
    with os.fdopen(key_fd, "wb") as fh:
        fh.write(key.private_bytes(Encoding.PEM, PrivateFormat.TraditionalOpenSSL, NoEncryption()))

    subject = cert.subject.rfc4514_string()
    cnpj_match = re.search(r"\d{14}", subject.replace(".", "").replace("/", "").replace("-", ""))
    not_after = getattr(cert, "not_valid_after_utc", None) or cert.not_valid_after
    if not_after.tzinfo is None:
        not_after = not_after.replace(tzinfo=timezone.utc)
    info = {
        "titular": subject,
        "cnpj_certificado": cnpj_match.group(0) if cnpj_match else "",
        "cnpj_confere": bool(cnpj_match) and cnpj_match.group(0) == CNPJ_SPE,
        "valido_ate": not_after.strftime("%d/%m/%Y"),
        "expirado": not_after < datetime.now(timezone.utc),
    }
    _cert_cache.update({"mtime": mtime, "cert_path": cert_path, "key_path": key_path, "info": info})
    return cert_path, key_path, info


def certificado_status():
    """Só confere se o certificado está presente/válido — não faz nenhuma consulta externa."""
    if not os.path.exists(CERT_PFX_PATH):
        return {"configurado": False, "ok": False, "error": f"Nenhum certificado em {CERT_PFX_PATH}."}
    try:
        _, _, info = _load_cert_pem()
    except Exception as exc:
        return {"configurado": True, "ok": False, "error": str(exc)}
    if info["expirado"]:
        return {"configurado": True, "ok": False, "error": f"Certificado expirado em {info['valido_ate']}.", **info}
    if not info["cnpj_confere"]:
        return {
            "configurado": True, "ok": False,
            "error": f"O certificado é do CNPJ {info['cnpj_certificado'] or '?'}, não do {CNPJ_SPE} (SPE T2 BUENO).",
            **info,
        }
    return {"configurado": True, "ok": True, **info}


# ── estado local (cursor NSU + notas já descobertas) ────────────────────────

def _load_state():
    default = {"ult_nsu": "0", "ult_nsu_nfse": 0, "notas": {}, "ultima_sync_nfe": "", "ultima_sync_nfse": ""}
    if not STATE_PATH.exists():
        return default
    try:
        with open(STATE_PATH, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        if isinstance(data, dict):
            default.update(data)
    except Exception:
        pass
    return default


def _save_state(state):
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    tmp = str(STATE_PATH) + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(state, fh, ensure_ascii=False, indent=2)
    os.replace(tmp, str(STATE_PATH))


# ── NFe — Distribuição DFe (SEFAZ, nacional) ────────────────────────────────

def _nfe_soap_envelope(ult_nsu):
    return (
        '<?xml version="1.0" encoding="utf-8"?>'
        '<soap12:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
        'xmlns:xsd="http://www.w3.org/2001/XMLSchema" '
        'xmlns:soap12="http://www.w3.org/2003/05/soap-envelope">'
        '<soap12:Body>'
        '<nfeDistDFeInteresse xmlns="http://www.portalfiscal.inf.br/nfe/wsdl/NFeDistribuicaoDFe">'
        '<nfeDadosMsg>'
        f'<distDFeInt versao="1.35" xmlns="{NS_NFE}">'
        '<tpAmb>1</tpAmb>'
        f'<cUFAutor>{UF_AUTOR_GO}</cUFAutor>'
        f'<CNPJ>{CNPJ_SPE}</CNPJ>'
        f'<distNSU><ultNSU>{str(ult_nsu).zfill(15)}</ultNSU></distNSU>'
        '</distDFeInt>'
        '</nfeDadosMsg>'
        '</nfeDistDFeInteresse>'
        '</soap12:Body>'
        '</soap12:Envelope>'
    )


def _parse_nfe_dist_response(xml_text):
    """Extrai os docZip (resNFe/resEvento, compactados em gzip+base64) da
    resposta do webservice de Distribuição DFe."""
    root = ET.fromstring(xml_text)
    ns = {"n": NS_NFE}
    ret = root.find(".//n:retDistDFeInt", ns)
    if ret is None:
        return {"cStat": "", "xMotivo": "Resposta sem retDistDFeInt.", "max_nsu": None, "notas": [], "eventos": []}
    cstat = (ret.findtext("n:cStat", default="", namespaces=ns) or "").strip()
    xmotivo = (ret.findtext("n:xMotivo", default="", namespaces=ns) or "").strip()
    max_nsu = ret.findtext("n:maxNSU", default=None, namespaces=ns)
    notas, eventos = [], []
    for doc_zip in ret.findall(".//n:docZip", ns):
        schema = doc_zip.get("schema", "")
        nsu = doc_zip.get("NSU", "")
        try:
            raw = gzip.decompress(base64.b64decode(doc_zip.text or ""))
            inner = ET.fromstring(raw)
        except Exception:
            continue
        if schema.startswith("resNFe"):
            notas.append({
                "nsu": nsu,
                "chave": inner.findtext("n:chNFe", default="", namespaces=ns),
                "cnpj_emitente": inner.findtext("n:CNPJ", default="", namespaces=ns),
                "nome_emitente": inner.findtext("n:xNome", default="", namespaces=ns),
                "valor": inner.findtext("n:vNF", default="0", namespaces=ns),
                "data_emissao": inner.findtext("n:dhEmi", default="", namespaces=ns),
                "situacao": inner.findtext("n:cSitNFe", default="", namespaces=ns),  # 1=autorizada, 2=cancelada, 3=denegada
            })
        elif schema.startswith("resEvento"):
            eventos.append({
                "nsu": nsu,
                "chave": inner.findtext("n:chNFe", default="", namespaces=ns),
                "tipo_evento": inner.findtext("n:tpEvento", default="", namespaces=ns),  # 110111 = Cancelamento
                "data_evento": inner.findtext("n:dhEvento", default="", namespaces=ns),
            })
    return {"cStat": cstat, "xMotivo": xmotivo, "max_nsu": max_nsu, "notas": notas, "eventos": eventos}


def consultar_nfe_distribuicao(ult_nsu="0"):
    import requests
    cert_path, key_path, _ = _load_cert_pem()
    body = _nfe_soap_envelope(ult_nsu)
    headers = {"Content-Type": "application/soap+xml; charset=utf-8"}
    resp = requests.post(NFE_DIST_DFE_URL, data=body.encode("utf-8"), headers=headers,
                          cert=(cert_path, key_path), timeout=45)
    resp.raise_for_status()
    return _parse_nfe_dist_response(resp.text)


def _nfe_soap_envelope_chave(chave):
    # A consulta por chave (consChNFe) rejeita com cStat 215 "Falha no
    # esquema xml" se usar versao="1.35" (a mesma versão que a consulta por
    # NSU exige) — testado ao vivo: só funciona com versao="1.01" pra esse
    # tipo específico de consulta. Achado em 24/09/2026.
    return (
        '<?xml version="1.0" encoding="utf-8"?>'
        '<soap12:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
        'xmlns:xsd="http://www.w3.org/2001/XMLSchema" '
        'xmlns:soap12="http://www.w3.org/2003/05/soap-envelope">'
        '<soap12:Body>'
        '<nfeDistDFeInteresse xmlns="http://www.portalfiscal.inf.br/nfe/wsdl/NFeDistribuicaoDFe">'
        '<nfeDadosMsg>'
        f'<distDFeInt versao="1.01" xmlns="{NS_NFE}">'
        '<tpAmb>1</tpAmb>'
        f'<cUFAutor>{UF_AUTOR_GO}</cUFAutor>'
        f'<CNPJ>{CNPJ_SPE}</CNPJ>'
        f'<consChNFe><chNFe>{chave}</chNFe></consChNFe>'
        '</distDFeInt>'
        '</nfeDadosMsg>'
        '</nfeDistDFeInteresse>'
        '</soap12:Body>'
        '</soap12:Envelope>'
    )


def consultar_nfe_xml_completo(chave):
    """Busca o resumo (ou, se autorizado, o documento completo) de uma NFe
    específica pela chave de acesso — usa consChNFe em vez de distNSU, na
    mesma Distribuição DFe da SEFAZ. Testado ao vivo: quando o CNPJ que
    consulta não é o emitente/destinatário de registro dessa nota, a SEFAZ
    só libera o resumo (resNFe — chave, fornecedor, valor total, sem os
    itens); ver extrair_itens_nfe() pra como isso é tratado."""
    import requests
    cert_path, key_path, _ = _load_cert_pem()
    chave_digits = re.sub(r"\D", "", str(chave or ""))
    body = _nfe_soap_envelope_chave(chave_digits)
    headers = {"Content-Type": "application/soap+xml; charset=utf-8"}
    resp = requests.post(NFE_DIST_DFE_URL, data=body.encode("utf-8"), headers=headers,
                          cert=(cert_path, key_path), timeout=45)
    resp.raise_for_status()
    root = ET.fromstring(resp.text)
    ns = {"n": NS_NFE}
    ret = root.find(".//n:retDistDFeInt", ns)
    if ret is None:
        return {"cStat": "", "xMotivo": "Resposta sem retDistDFeInt.", "xml": None}
    cstat = (ret.findtext("n:cStat", default="", namespaces=ns) or "").strip()
    xmotivo = (ret.findtext("n:xMotivo", default="", namespaces=ns) or "").strip()
    for doc_zip in ret.findall(".//n:docZip", ns):
        xml_text = _decodificar_dfe(doc_zip.text)
        if xml_text:
            return {"cStat": cstat, "xMotivo": xmotivo, "xml": xml_text}
    return {"cStat": cstat, "xMotivo": xmotivo, "xml": None}


def extrair_itens_nfe(xml_text):
    """Lê os produtos (tag <det>/<prod>) do XML de uma NFe, quando o
    documento completo (nfeProc/NFe) está disponível. A consulta por chave
    (consChNFe) na Distribuição DFe geralmente só libera o RESUMO (resNFe) —
    sem os itens — quando o CNPJ que consulta não é o emitente/destinatário
    de registro dessa nota; nesse caso a lista de itens vem vazia, mas ainda
    assim devolve fornecedor e valor total do resumo, e sinaliza
    `resumo_apenas=True` pra quem chamou saber que não vai achar produtos."""
    root = ET.fromstring(xml_text)
    ns = {"n": NS_NFE}
    itens = []
    for det in root.findall(".//n:det", ns):
        prod = det.find("n:prod", ns)
        if prod is None:
            continue
        itens.append({
            "cprod": (prod.findtext("n:cProd", default="", namespaces=ns) or "").strip(),
            "nome": (prod.findtext("n:xProd", default="", namespaces=ns) or "").strip(),
            "unidade": (prod.findtext("n:uCom", default="", namespaces=ns) or "").strip(),
            "qtd": prod.findtext("n:qCom", default="0", namespaces=ns) or "0",
            "valor_unit": prod.findtext("n:vUnCom", default="0", namespaces=ns) or "0",
            "valor_total": prod.findtext("n:vProd", default="0", namespaces=ns) or "0",
            "cean": (prod.findtext("n:cEAN", default="", namespaces=ns) or "").strip(),
        })
    emit = root.find(".//n:emit", ns)
    if emit is not None:
        fornecedor = (emit.findtext("n:xNome", default="", namespaces=ns) or "").strip()
    else:
        # resNFe (resumo): não tem <emit>, o nome do emitente vem direto na raiz
        fornecedor = (root.findtext("n:xNome", default="", namespaces=ns) or "").strip()
    valor_total = (root.findtext(".//n:vNF", default="", namespaces=ns) or "").strip()
    resumo_apenas = (not itens) and root.tag.split("}")[-1] == "resNFe"
    return itens, fornecedor, valor_total, resumo_apenas


# ── NFSe — Ambiente de Dados Nacional (ADN), sistema unificado da NFS-e ────
# Goiânia (e a maioria dos municípios) migrou do webservice ABRASF municipal
# para este sistema nacional único a partir de 01/01/2026 (LC 214/2025). É
# REST autenticado por mTLS com o mesmo certificado — testado e confirmado
# funcionando em produção com dados reais do CNPJ da SPE T2 BUENO.
ADN_NFSE_URL = "https://adn.nfse.gov.br/contribuintes/DFe"
NS_NFSE = "http://www.sped.fazenda.gov.br/nfse"
_TIPOS_EVENTO_CANCELAMENTO = {
    "CANCELAMENTO", "CANCELAMENTO_POR_SUBSTITUICAO",
    "CANCELAMENTO_DEFERIDO_ANALISE_FISCAL", "CANCELAMENTO_POR_OFICIO",
}


def _decodificar_dfe(arquivo_xml_b64):
    if not arquivo_xml_b64:
        return None
    try:
        return gzip.decompress(base64.b64decode(arquivo_xml_b64)).decode("utf-8", "ignore")
    except Exception:
        return None


def _parse_nfse_xml(xml_text):
    try:
        root = ET.fromstring(xml_text)
    except Exception:
        return None
    ns = {"n": NS_NFSE}
    inf = root.find(".//n:infNFSe", ns)
    if inf is None:
        return None
    emit = inf.find("n:emit", ns)
    valores = inf.find("n:valores", ns)
    # vServ (em infDPS/.../valores/vServPrest) é o valor BRUTO da nota (valor
    # da operação/serviço, antes de qualquer retenção) — é o que aparece como
    # "VALOR TOTAL DA NFS-e / VALOR DA OPERAÇÃO/SERVIÇO" no DANFSe. vLiq (em
    # infNFSe/valores) é o valor LÍQUIDO (já descontado ISSQN/INSS retidos).
    # O painel mostra o valor bruto; vLiq fica guardado à parte como referência.
    dps_valores = root.find(".//n:infDPS/n:valores", ns)
    v_serv = dps_valores.findtext("n:vServPrest/n:vServ", default="", namespaces=ns) if dps_valores is not None else ""
    v_liq = valores.findtext("n:vLiq", default="0", namespaces=ns) if valores is not None else "0"
    return {
        "numero": inf.findtext("n:nNFSe", default="", namespaces=ns),
        "cnpj_prestador": emit.findtext("n:CNPJ", default="", namespaces=ns) if emit is not None else "",
        "prestador": emit.findtext("n:xNome", default="", namespaces=ns) if emit is not None else "",
        "valor": v_serv or v_liq,
        "valor_liquido": v_liq,
        "data_emissao": inf.findtext("n:dhProc", default="", namespaces=ns),
        "xml": xml_text,
    }


def consultar_nfse_nacional(ult_nsu=0):
    """Consulta o lote de Documentos Fiscais de Serviço (NFSe + eventos) a
    partir do NSU informado. Sistema nacional único — cobre qualquer
    município aderido, não só Goiânia."""
    import requests
    cert_path, key_path, _ = _load_cert_pem()
    url = f"{ADN_NFSE_URL}/{int(ult_nsu or 0)}"
    resp = requests.get(url, params={"cnpjConsulta": CNPJ_SPE, "lote": "true"},
                         cert=(cert_path, key_path), timeout=45)
    if resp.status_code == 404:
        # NSU maior que o disponível — fim do backlog, não é erro.
        return {"status": "FIM_DO_BACKLOG", "erros": [], "max_nsu": int(ult_nsu or 0), "notas": [], "eventos": []}
    resp.raise_for_status()
    data = resp.json()
    max_nsu = int(ult_nsu or 0)
    notas, eventos = [], []
    for item in data.get("LoteDFe") or []:
        nsu = int(item.get("NSU") or 0)
        max_nsu = max(max_nsu, nsu)
        xml_text = _decodificar_dfe(item.get("ArquivoXml"))
        if not xml_text:
            continue
        tipo_doc = item.get("TipoDocumento")
        if tipo_doc == "NFSE":
            parsed = _parse_nfse_xml(xml_text)
            if parsed:
                parsed["chave"] = item.get("ChaveAcesso", "")
                notas.append(parsed)
        elif tipo_doc == "EVENTO":
            tipo_evento = item.get("TipoEvento") or ""
            if tipo_evento in _TIPOS_EVENTO_CANCELAMENTO:
                eventos.append({"chave": item.get("ChaveAcesso", ""), "tipo_evento": tipo_evento})
    return {
        "status": data.get("StatusProcessamento", ""),
        "erros": [str(e) for e in (data.get("Erros") or [])],
        "max_nsu": max_nsu,
        "notas": notas,
        "eventos": eventos,
    }


# ── reconciliação com o Sienge ───────────────────────────────────────────────

_SUFIXOS_EMPRESA = re.compile(
    r"\b(LTDA|EPP|EIRELI|ME|MEI|S/?A|SA|CIA|COMERCIO|COM|SERVICOS|SERVICO|"
    r"INDUSTRIA|CONSTRUCOES|ENGENHARIA)\b"
)


def _norm_nome(s):
    """Normaliza nome de credor/fornecedor pra comparação: maiúsculas, sem
    acento, sem razão social genérica (LTDA/ME/S.A./...), só alfanumérico."""
    s = unicodedata.normalize("NFKD", str(s or "").upper()).encode("ascii", "ignore").decode("ascii")
    s = re.sub(r"[^A-Z0-9 ]", " ", s)
    s = _SUFIXOS_EMPRESA.sub(" ", s)
    return re.sub(r"\s+", " ", s).strip()


_credor_equiv_cache = None


def _load_credor_equivalencias():
    """Grupos de CNPJ que representam o MESMO fornecedor na prática, mas com
    cadastros diferentes no Sienge (ex.: empresa que trocou de razão social/
    CNPJ ao longo do tempo — os títulos antigos ficam sob o CNPJ anterior,
    as notas novas chegam com o CNPJ atual). Confirmado manualmente pelo
    usuário caso a caso, nunca inferido automaticamente — casar CNPJs
    diferentes sem confirmação seria arriscado (poderia vincular a nota ao
    fornecedor errado). Arquivo opcional dados/credores_equivalentes.json:
    lista de listas de CNPJ (com ou sem pontuação). Sem o arquivo, nenhuma
    equivalência é assumida."""
    global _credor_equiv_cache
    if _credor_equiv_cache is not None:
        return _credor_equiv_cache
    grupos = []
    try:
        if CREDOR_EQUIVALENCIAS_PATH.exists():
            data = json.loads(CREDOR_EQUIVALENCIAS_PATH.read_text(encoding="utf-8"))
            if isinstance(data, list):
                for grupo in data:
                    if isinstance(grupo, list):
                        limpo = {re.sub(r"\D", "", str(v)) for v in grupo if str(v).strip()}
                        if len(limpo) > 1:
                            grupos.append(limpo)
    except Exception:
        pass
    _credor_equiv_cache = grupos
    return grupos


def _cnpj_equivalente(cnpj_a, cnpj_b):
    if not cnpj_a or not cnpj_b:
        return False
    if cnpj_a == cnpj_b:
        return True
    return any(cnpj_a in grupo and cnpj_b in grupo for grupo in _load_credor_equivalencias())


def _casar_titulo(candidatos_por_chave, chave, cnpj_nota, nome_nota=""):
    """Escolhe, entre os títulos do Sienge com o mesmo número de documento,
    o que pertence de fato ao fornecedor da nota. Números de NF/NFSe não são
    exclusivos entre fornecedores — dois credores diferentes podem emitir,
    cada um, sua própria nota nº 54 — então casar só pelo número gera falso
    positivo ('lançado no Sienge' quando não é a mesma nota).

    Prioridade: 1) CNPJ idêntico (quando conhecido dos dois lados);
    2) nome do credor mais parecido com o emitente/prestador da nota
    (cobre credores sem CNPJ cadastrado na planilha de apoio, ex.: CIPLAN).
    Sem sinal nenhum pra desempatar entre candidatos de credores diferentes,
    não casa — prefere falso negativo a falso positivo."""
    candidatos = (candidatos_por_chave or {}).get(chave) or []
    if not candidatos:
        return ""
    # Dedup por número de título — várias fontes (bd_sienge_tmz +
    # relatórios manuais) podem repetir o mesmo título; ficar com a PRIMEIRA
    # ocorrência (ordem de inserção) perdia sinal quando essa primeira vinha
    # sem CNPJ/nome (ex.: credor sem CNPJ cadastrado na planilha de apoio) e
    # uma ocorrência posterior, de outra fonte, tinha o nome — mantém a mais
    # "rica" (com CNPJ e/ou nome preenchido) entre as repetidas.
    vistos = {}
    for c in candidatos:
        chave_tit = c.get("numtitcp")
        atual = vistos.get(chave_tit)
        if atual is None:
            vistos[chave_tit] = c
            continue
        score_atual = bool(atual.get("cnpj")) * 2 + bool(atual.get("nome"))
        score_novo = bool(c.get("cnpj")) * 2 + bool(c.get("nome"))
        if score_novo > score_atual:
            vistos[chave_tit] = c
    candidatos = list(vistos.values())

    if len(candidatos) == 1:
        cand = candidatos[0]
        if cnpj_nota and cand.get("cnpj") and cand["cnpj"] != cnpj_nota and not _cnpj_equivalente(cand["cnpj"], cnpj_nota):
            return ""
        return cand.get("numtitcp", "")

    if cnpj_nota:
        exatos = [c for c in candidatos if c.get("cnpj") == cnpj_nota or _cnpj_equivalente(c.get("cnpj"), cnpj_nota)]
        if len(exatos) == 1:
            return exatos[0].get("numtitcp", "")
        if len(exatos) > 1:
            candidatos = exatos

    nome_norm = _norm_nome(nome_nota)
    if nome_norm:
        melhor, melhor_score = None, 0.0
        for cand in candidatos:
            score = difflib.SequenceMatcher(None, nome_norm, _norm_nome(cand.get("nome", ""))).ratio()
            if score > melhor_score:
                melhor, melhor_score = cand, score
        if melhor is not None and melhor_score >= 0.55:
            return melhor.get("numtitcp", "")
    return ""


def reconciliar_com_sienge(notas_nfe, notas_nfse, nf_titulo_candidatos, eventos_nfse=None, eventos_nfe=None):
    """Cruza as notas descobertas (NFe + NFSe) com os títulos do Sienge que
    têm o mesmo número de documento, confirmando pelo CNPJ do
    emitente/prestador para não misturar fornecedores diferentes que usam a
    mesma numeração por coincidência.

    Cancelamento: a SEFAZ não reemite o resNFe/resNFSe original quando uma
    nota é cancelada depois de já ter sido recebida — o cancelamento chega
    como um evento à parte (resEvento tpEvento=110111 pra NFe; evento de
    cancelamento na distribuição nacional pra NFSe), com a mesma chave de
    acesso. Por isso 'cancelada' também é decidida cruzando a chave da nota
    contra os eventos de cancelamento recebidos nesta mesma consulta — tanto
    pra NFe quanto pra NFSe (antes só NFSe considerava eventos; NFe
    dependia só de 'situacao' vindo junto com a nota original, que nunca
    reflete um cancelamento posterior)."""
    chaves_canceladas = {e.get("chave") for e in (eventos_nfse or []) if e.get("chave")}
    chaves_canceladas |= {
        e.get("chave") for e in (eventos_nfe or [])
        if e.get("chave") and str(e.get("tipo_evento", "")).strip() == "110111"
    }
    linhas = []
    for n in notas_nfe or []:
        chave = re.sub(r"\D", "", str(n.get("chave", "")))
        nf_num = chave[25:34].lstrip("0") if len(chave) == 44 else ""
        cnpj_nota = re.sub(r"\D", "", str(n.get("cnpj_emitente", "")))
        titulo = _casar_titulo(nf_titulo_candidatos, nf_num, cnpj_nota, n.get("nome_emitente", "")) if nf_num else ""
        linhas.append({
            "tipo": "NFe", "numero": nf_num or n.get("chave", ""), "chave": n.get("chave", ""),
            "fornecedor": n.get("nome_emitente", ""), "cnpj_fornecedor": n.get("cnpj_emitente", ""),
            "valor": float(n.get("valor") or 0), "data_emissao": n.get("data_emissao", ""),
            "cancelada": n.get("situacao") == "2" or n.get("chave") in chaves_canceladas,
            "lancado_sienge": bool(titulo), "titulo_sienge": titulo,
            "xml_disponivel": False,  # NFe: XML completo é buscado sob demanda (não fica salvo)
        })
    for n in notas_nfse or []:
        nf_key = re.sub(r"\D", "", str(n.get("numero", ""))).lstrip("0")
        cnpj_nota = re.sub(r"\D", "", str(n.get("cnpj_prestador", "")))
        titulo = _casar_titulo(nf_titulo_candidatos, nf_key, cnpj_nota, n.get("prestador", "")) if nf_key else ""
        linhas.append({
            "tipo": "NFSe", "numero": n.get("numero", ""), "chave": n.get("chave", ""),
            "fornecedor": n.get("prestador", ""), "cnpj_fornecedor": n.get("cnpj_prestador", ""),
            "valor": float(n.get("valor") or 0), "valor_liquido": float(n.get("valor_liquido") or 0),
            "data_emissao": n.get("data_emissao", ""),
            "cancelada": n.get("chave") in chaves_canceladas,
            "lancado_sienge": bool(titulo), "titulo_sienge": titulo,
            "xml_disponivel": bool(n.get("xml")),
            "xml": n.get("xml", ""),  # documento completo, já salvo — download instantâneo
        })
    return linhas
