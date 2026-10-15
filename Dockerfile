# Imagem do Dashboard de Obra (producao, Linux).
# Montada no servidor por deploy/release.sh:  docker build --build-arg BUILD_ID=<id> -t dash:<id> .
FROM python:3.12-slim-bookworm

ARG BUILD_ID=dev
ENV DASH_BUILD_ID=${BUILD_ID} \
    PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1 \
    DASHBOARD_HOST=0.0.0.0 \
    DASHBOARD_PORT=5070 \
    DASHBOARD_BEHIND_PROXY=1 \
    DASHBOARD_HTTPS=1 \
    CHROME_BINARY_PATH=/usr/bin/chromium \
    CHROMEDRIVER_PATH=/usr/bin/chromedriver

# Sistema: curl (health check do release.sh), Chromium (robos Selenium), LibreOffice (PDF do RMO e
# Cartas), fontes e o driver ODBC do SQL Server (UAU).
RUN apt-get update \
 && DEBIAN_FRONTEND=noninteractive apt-get install -y --no-install-recommends \
      ca-certificates curl gnupg unixodbc \
      chromium chromium-driver \
      libreoffice-writer libreoffice-impress \
      fonts-liberation fonts-dejavu-core fonts-crosextra-carlito fonts-crosextra-caladea \
 && curl -fsSL https://packages.microsoft.com/keys/microsoft.asc | gpg --dearmor -o /usr/share/keyrings/microsoft-prod.gpg \
 && echo "deb [arch=amd64 signed-by=/usr/share/keyrings/microsoft-prod.gpg] https://packages.microsoft.com/debian/12/prod bookworm main" > /etc/apt/sources.list.d/mssql-release.list \
 && apt-get update \
 && ACCEPT_EULA=Y DEBIAN_FRONTEND=noninteractive apt-get install -y --no-install-recommends msodbcsql18 msodbcsql17 \
 && apt-get purge -y gnupg && apt-get autoremove -y \
 && rm -rf /var/lib/apt/lists/*

# Usuario sem privilegios (uid 1000 = dono de /opt/dashboard/dados no servidor, ver provision.sh)
RUN useradd --uid 1000 --create-home --shell /usr/sbin/nologin app \
 && mkdir -p /app/dados /app/logs \
 && chown -R app:app /app

WORKDIR /app

# Dependencias antes do codigo: a camada fica em cache enquanto requirements.txt nao mudar
COPY requirements.txt .
RUN pip install -r requirements.txt

COPY --chown=app:app . .

USER app
EXPOSE 5070

HEALTHCHECK --interval=30s --timeout=10s --start-period=60s --retries=3 \
  CMD curl -fsS http://127.0.0.1:5070/api/status || exit 1

CMD ["gunicorn", "-c", "gunicorn.conf.py", "app:app"]
