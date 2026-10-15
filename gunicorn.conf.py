"""Configuracao do gunicorn (producao).

Por que UM worker e varias threads: o painel guarda estado em memoria do processo (cache de
dados, travas de sincronizacao, limite de tentativas de login, presenca de usuarios) e grava
arquivos JSON em dados/. Com varios processos cada um teria uma copia diferente desse estado e
as travas deixariam de valer. Escalar = aumentar GUNICORN_THREADS, nao os workers.
"""
import os

bind = f"{os.environ.get('DASHBOARD_HOST', '0.0.0.0')}:{os.environ.get('DASHBOARD_PORT', '5070')}"
workers = 1
threads = int(os.environ.get("GUNICORN_THREADS", "16") or 16)
worker_class = "gthread"

# sincronizacoes e carga dos dados podem demorar minutos; uploads grandes (NWD/IFC) tambem
timeout = int(os.environ.get("GUNICORN_TIMEOUT", "900") or 900)
graceful_timeout = 120
keepalive = 5

# o proxy (Caddy) esta em outro container da rede dashnet; so ele fala com o gunicorn
forwarded_allow_ips = "*"

# logs para stdout/stderr (docker logs); sem IP/UA em excesso
accesslog = "-"
errorlog = "-"
loglevel = os.environ.get("GUNICORN_LOGLEVEL", "info")

# limites de requisicao (cabecalhos), protecao basica contra abuso
limit_request_line = 8190
limit_request_fields = 100
limit_request_field_size = 8190
