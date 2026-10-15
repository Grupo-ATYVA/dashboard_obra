import json
import re
import html
from pathlib import Path
from urllib.request import Request, urlopen

import pandas as pd

DATA_DIR = Path(__file__).resolve().parent / "dados" / "synergia"
URL = "https://buscadorsinapi.com.br/go"
OUT = DATA_DIR / "SINAPI_GO_2026_04.xlsx"
META = DATA_DIR / "SINAPI_GO_2026_04_FONTE.json"

request = Request(URL, headers={"User-Agent": "Mozilla/5.0"})
with urlopen(request, timeout=60) as response:
    page = response.read().decode("utf-8", errors="ignore")
rows = {}
for href, raw_text in re.findall(r'<a[^>]+href=["\']([^"\']+)["\'][^>]*>(.*?)</a>', page, flags=re.I | re.S):
    text = html.unescape(re.sub(r"<[^>]+>", " ", raw_text))
    text = " ".join(text.split())
    match = re.match(r"^(\d{1,7})\s+(.+?)\s+R\$\s*([\d.]+,\d{2})$", text)
    if not match:
        continue
    code, description, price = match.groups()
    rows[code] = {
        "CODIGO": code,
        "DESCRICAO": description,
        "UNIDADE": "",
        "PRECO_GO": float(price.replace(".", "").replace(",", ".")),
        "REFERENCIA": "04/2026",
        "UF": "GO",
        "URL_ITEM": href,
    }

df = pd.DataFrame(rows.values()).sort_values(["CODIGO"])
df.to_excel(OUT, index=False)
meta = {
    "referencia": "04/2026",
    "uf": "GO",
    "fonte": "Índice público baseado nas planilhas CAIXA / IBGE",
    "url": URL,
    "url_oficial": "https://www.caixa.gov.br/poder-publico/modernizacao-gestao/sinapi/Paginas/default.aspx",
    "arquivo": OUT.name,
    "disponivel": True,
    "itens": len(df),
    "observacao": "Referência auxiliar. Para licitações, confirme os valores no arquivo oficial da CAIXA.",
}
META.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
print(f"{OUT}: {len(df)} referências salvas")
