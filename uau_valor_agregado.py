"""Leitor da planilha 'VALOR AGREGADO_FAR_CJ03 ...xlsx' (dados/uau).

A planilha é a análise de valor agregado (EVM) da obra FAR CJ03/CJ05 — Ed. Regina Célia:
  - 'C O N T R O L E' ........ físico/financeiro planejado x realizado, mês a mês
  - 'D-AGREG' ................ VP, CR, VA, índices (IDP/IDC) e estimativas
  - '...ORÇ AGRUPADO...' ..... orçamento CJ03+CJ05 com o % e R$ PLANEJADO por mês
  - '...CJ03-MEDIÇÕES FÍS' ... orçamento CJ03 com o % e R$ MEDIDO por mês
  - '...CJ05-MEDIÇÕES FÍS' ... orçamento CJ05 com o % e R$ MEDIDO por mês
  - 'UAU-DADOS-VISUALIZADOR' .. saídas do fluxo do UAU (a pagar / paga) por mês
  - 'CEF' .................... ficha-resumo de custos enviada à Caixa

Só lê (nunca grava na planilha). Resultado cacheado pelo mtime do arquivo.
"""
import datetime
import glob
import os
import unicodedata

_cache = {"path": None, "mtime": None, "dados": None}


def _norm(txt):
    s = unicodedata.normalize("NFKD", str(txt or "")).encode("ascii", "ignore").decode("ascii")
    return s.upper().replace(" ", "")


def _num(v, casas=2):
    if isinstance(v, bool) or v is None or v == "":
        return 0.0
    try:
        return round(float(v), casas)
    except (TypeError, ValueError):
        return 0.0


def _mes(v):
    return v.strftime("%Y-%m") if isinstance(v, (datetime.datetime, datetime.date)) else ""


def _txt(v):
    return str(v).strip() if v not in (None, "") else ""


def localizar_planilha(pasta):
    cand = [p for p in glob.glob(os.path.join(pasta, "VALOR AGREGADO*.xlsx"))
            if not os.path.basename(p).startswith("~$")]
    return max(cand, key=os.path.getmtime) if cand else None


def _linhas(wb, trecho):
    alvo = _norm(trecho)
    for nome in wb.sheetnames:
        if alvo in _norm(nome):
            return nome, [list(r) for r in wb[nome].iter_rows(values_only=True)]
    return None, []


def _col(linha, i):
    return linha[i] if i < len(linha) else None


def _controle(linhas):
    meses, analise, orcamento, titulo = [], "", 0.0, ""
    if linhas:
        titulo = _txt(_col(linhas[1], 1)) if len(linhas) > 1 else ""
        analise = _mes(_col(linhas[0], 13))
        orcamento = _num(_col(linhas[2], 2)) if len(linhas) > 2 else 0.0
    for r in linhas[5:]:
        mes = _mes(_col(r, 1))
        if not mes:
            continue
        meses.append({
            "mes": mes,
            "medicao": int(_num(_col(r, 0), 0)),
            "fis_plan": _num(_col(r, 2), 6), "fis_plan_acum": _num(_col(r, 3), 6),
            "fis_real": _num(_col(r, 4), 6), "fis_real_acum": _num(_col(r, 5), 6),
            "fin_plan": _num(_col(r, 6)), "fin_plan_acum": _num(_col(r, 7)),
            "fin_real": _num(_col(r, 8)), "fin_real_acum": _num(_col(r, 9)),
            "custo_mes": _num(_col(r, 10)), "custo_acum": _num(_col(r, 11)),
            "custo_acum_sem_estoque": _num(_col(r, 12)),
        })
    return {"obra": titulo, "analise": analise, "orcamento_aprovado": orcamento, "meses": meses}


def _valor_agregado(linhas):
    out = []
    for r in linhas[4:]:
        mes = _mes(_col(r, 0))
        if not mes:
            continue
        # linhas futuras só têm o valor planejado (CR/VA vazios): marca pra tela não plotar zero
        tem_real = _col(r, 2) not in (None, "")
        out.append({
            "mes": mes, "tem_real": tem_real,
            "vp": _num(_col(r, 1)), "cr": _num(_col(r, 2)), "va": _num(_col(r, 3)),
            "pct_fis": _num(_col(r, 4), 6), "vc": _num(_col(r, 5)), "vpr": _num(_col(r, 6)),
            "idp": _num(_col(r, 7), 4), "idc": _num(_col(r, 8), 4),
            "ept": _num(_col(r, 9)), "ent": _num(_col(r, 10)),
            "pct_economia": _num(_col(r, 11), 6), "rentabilidade": _num(_col(r, 12), 6),
            "orc_meta": _num(_col(r, 14)),
        })
    return out


def _arvore(linhas):
    """Orçamento por nível: 1 = INDIRETOS/DIRETOS, 2 = grupo (item com 'Anexos' = 0),
    3 = insumo (código IPxxxx). Linhas sem Produto são só detalhe zerado e ficam de fora."""
    ini = next((i for i, r in enumerate(linhas) if _txt(_col(r, 0)) == "Produto"), None)
    if ini is None:
        return {"meses": [], "tree": [], "total": 0.0, "meta": {}}
    cab = linhas[ini]
    datas = linhas[ini - 1] if ini else []
    cols = []  # (mes, coluna do %)
    for c, v in enumerate(datas):
        if _mes(v) and _txt(_col(cab, c)).startswith("%"):
            cols.append((_mes(v), c))
    meses = [m for m, _ in cols]
    meta = {}
    for r in linhas[:ini]:
        for c, v in enumerate(r[:14]):
            t = _txt(v)
            if t.endswith(":") and _col(r, c + 1) not in (None, ""):
                prox = _col(r, c + 1)
                meta[t.rstrip(":")] = prox.strftime("%Y-%m") if isinstance(prox, datetime.datetime) else prox
    for r in linhas[:ini]:  # valores que ficam duas colunas à frente do rótulo (ex.: 'ORÇ. INCC BASE')
        for c, v in enumerate(r[:14]):
            t = _txt(v)
            if t.endswith(":") and _col(r, c + 1) in (None, "") and _col(r, c + 2) not in (None, ""):
                meta[t.rstrip(":")] = _col(r, c + 2)

    def no(r, nivel, cod):
        pct, val = [], []
        for _, c in cols:
            pct.append(_num(_col(r, c), 6))
            val.append(_num(_col(r, c + 1)))
        total = _num(_col(r, 11))
        return {"cod": cod, "nm": _txt(_col(r, 8)), "nivel": nivel, "un": _txt(_col(r, 9)),
                "total": total, "pct_serv": _num(_col(r, 13), 6),
                "pct": pct, "val": val, "exec_acum": round(sum(val), 2),
                "children": []}

    tree, n1, n2 = [], None, None
    orfaos = []
    super_g = None  # grupo-pai sem 'Produto' (ex.: 'ESTRUTURAS'), que engloba o grupo seguinte de mesmo item
    for r in linhas[ini + 1:]:
        prod, item, ins = _txt(_col(r, 0)), _txt(_col(r, 4)), _txt(_col(r, 7))
        if not prod:
            if item and not ins and _num(_col(r, 11)) and n1 is not None:
                if _txt(_col(r, 9)):
                    super_g = no(r, 2, item)
                    n1["children"].append(super_g)
                    n2 = super_g
                else:  # grupo real a que só falta o 'Produto': entra só se o nível 1 não tem grupo-pai
                    orfaos.append((n1, no(r, 2, item)))
            continue
        anexos = _col(r, 2)
        grupo = anexos is not None and _num(anexos) == 0
        if not item and not ins:
            n1 = no(r, 1, _txt(_col(r, 8)))
            tree.append(n1)
            n2 = None
        elif item and not ins and grupo and n1 is not None:
            n2 = no(r, 2, item)
            if super_g is not None and super_g["cod"] == item:
                super_g["children"].append(n2)  # sub-grupo: não soma de novo no total do nível 1
                n2["sub"] = True
                n2["nivel"] = 3
                super_g = None
            else:
                n1["children"].append(n2)
                super_g = None
        elif n2 is not None and (ins or (item and _num(_col(r, 11)))):
            # insumo (IPxxxx) — ou linha do grupo que não traz o código do insumo
            n3 = no(r, 3, ins)
            if n3["total"] or any(n3["val"]):
                n2["children"].append(n3)
    for pai, grupo_orfao in orfaos:
        if not any(g.get("children") and any(x.get("sub") for x in g["children"]) for g in pai["children"]):
            pai["children"].append(grupo_orfao)
    for n in tree:
        n["children"] = [g for g in n["children"] if g["total"] or g["children"]]
        n["children"].sort(key=lambda g: g["cod"])
        n["soma_filhos"] = round(sum(g["total"] for g in n["children"]), 2)
        for g in n["children"]:
            g["soma_filhos"] = round(sum(x["total"] for x in g["children"]), 2)
    total = round(sum(n["total"] for n in tree), 2)
    return {"meses": meses, "tree": tree, "total": total, "meta": meta}


def _fluxo(linhas):
    out = []
    for r in linhas:
        mes = _mes(_col(r, 0))
        if not mes:
            continue
        a_pagar, paga = _num(_col(r, 1)), _num(_col(r, 2))
        out.append({"mes": mes, "a_pagar": a_pagar, "paga": paga,
                    "total": _num(_col(r, 3)), "saldo_acum": _num(_col(r, 5))})
    while out and not (out[-1]["a_pagar"] or out[-1]["paga"]):
        out.pop()
    while out and not (out[0]["a_pagar"] or out[0]["paga"]):
        out.pop(0)
    return out


def _cef(linhas):
    itens, totais, infra = [], {}, []
    for i, r in enumerate(linhas):
        cod, nome, custo, obs = _txt(_col(r, 1)), _txt(_col(r, 2)), _col(r, 6), _txt(_col(r, 9))
        if cod and nome and isinstance(custo, (int, float)) and i < 48:
            itens.append({"cod": cod, "nm": nome, "custo": _num(custo), "obs": obs})
        elif cod and nome and i >= 52 and isinstance(custo, (int, float)):
            infra.append({"cod": cod, "nm": nome, "custo": _num(custo)})
    for r in linhas:
        if _txt(_col(r, 1)) == "Custo Total sem BDI":
            totais["custo_total_sem_bdi"] = _num(_col(r, 2))
            totais["bdi"] = _num(_col(r, 4), 6)
    if len(linhas) > 49 and isinstance(_col(linhas[49], 6), (int, float)):
        totais["habitacao"] = _num(_col(linhas[49], 6))
    if len(linhas) > 63 and isinstance(_col(linhas[63], 6), (int, float)):
        totais["infraestrutura"] = _num(_col(linhas[63], 6))
    titulo = _txt(_col(linhas[1], 1)) if len(linhas) > 1 else ""
    return {"titulo": titulo, "itens": itens, "infra": infra, "totais": totais}


def _conferencia(controle, cj03, cj05):
    """Confronta o físico lançado no 'C O N T R O L E' com o R$ medido nas abas CJ03/CJ05
    (soma do nível 1 de cada uma, ponderado pelo orçamento dos dois conjuntos)."""
    orc = controle.get("orcamento_aprovado") or (cj03["total"] + cj05["total"])
    por_mes = {}
    for chave, arv in (("cj03", cj03), ("cj05", cj05)):
        for i, mes in enumerate(arv["meses"]):
            por_mes.setdefault(mes, {"cj03": 0.0, "cj05": 0.0})[chave] = round(sum(n["val"][i] for n in arv["tree"]), 2)
    ctrl = {m["mes"]: m for m in controle["meses"]}
    linhas, acum = [], 0.0
    for mes in sorted(por_mes):
        c = ctrl.get(mes, {})
        total = round(por_mes[mes]["cj03"] + por_mes[mes]["cj05"], 2)
        medido = bool(total) or bool(c.get("fis_real"))
        if medido:
            acum += total
        linhas.append({
            "mes": mes, "medido_cj03": por_mes[mes]["cj03"], "medido_cj05": por_mes[mes]["cj05"],
            "medido_total": total, "medido_acum": round(acum, 2),
            "pct_ponderado_acum": round(acum / orc, 6) if orc else 0.0,
            "fis_real_controle": c.get("fis_real", 0.0), "fis_real_acum_controle": c.get("fis_real_acum", 0.0),
        })
    ultimo = next((l for l in reversed(linhas) if l["medido_total"]), None)
    return {"orcamento": orc, "meses": linhas,
            "medido_acum": ultimo["medido_acum"] if ultimo else 0.0,
            "pct_ponderado_acum": ultimo["pct_ponderado_acum"] if ultimo else 0.0,
            "pct_controle_acum": ultimo["fis_real_acum_controle"] if ultimo else 0.0,
            "ate_mes": ultimo["mes"] if ultimo else ""}


def carregar(pasta):
    """Devolve o dicionário completo (ou {'ok': False, 'error': ...})."""
    path = localizar_planilha(pasta)
    if not path:
        return {"ok": False, "error": "Planilha 'VALOR AGREGADO*.xlsx' não encontrada em dados/uau."}
    mtime = os.path.getmtime(path)
    if _cache["path"] == path and _cache["mtime"] == mtime and _cache["dados"]:
        return _cache["dados"]
    import openpyxl
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception as exc:
        return {"ok": False, "error": f"Não consegui abrir a planilha: {exc}"}
    try:
        _, l_ctrl = _linhas(wb, "C O N T R O L E")
        _, l_agreg = _linhas(wb, "D-AGREG")
        _, l_orc = _linhas(wb, "ORC AGRUPADO")
        _, l_cj03 = _linhas(wb, "CJ03-MEDI")
        _, l_cj05 = _linhas(wb, "CJ05-MEDI")
        _, l_fluxo = _linhas(wb, "VISUALIZADOR FLUXO")
        _, l_cef = _linhas(wb, "CEF")
    finally:
        wb.close()
    controle = _controle(l_ctrl)
    va = _valor_agregado(l_agreg)
    analise = controle["analise"]
    ultimo = next((r for r in reversed(va) if r["mes"] == analise and r["tem_real"]), None) \
        or next((r for r in reversed(va) if r["tem_real"]), None)
    planejado, cj03, cj05 = _arvore(l_orc), _arvore(l_cj03), _arvore(l_cj05)
    dados = {
        "ok": True,
        "arquivo": os.path.basename(path),
        "atualizado_em": datetime.datetime.fromtimestamp(mtime).strftime("%d/%m/%Y %H:%M:%S"),
        "obra": controle["obra"],
        "mes_analise": analise,
        "orcamento_aprovado": controle["orcamento_aprovado"],
        "controle": controle["meses"],
        "valor_agregado": va,
        "ultimo": ultimo or {},
        "planejado": planejado,
        "cj03": cj03,
        "cj05": cj05,
        "conferencia": _conferencia(controle, cj03, cj05),
        "fluxo_uau": _fluxo(l_fluxo),
        "cef": _cef(l_cef),
    }
    _cache.update({"path": path, "mtime": mtime, "dados": dados})
    return dados
