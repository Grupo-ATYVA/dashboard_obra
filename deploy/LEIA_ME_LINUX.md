# Dashboard TEMAZEC no Linux — hospedagem 24h

Este guia substitui `ONLINE_FIXO_CLOUDFLARE.md` e os arquivos antigos desta pasta
(`dashboard-obra.service`, `cloudflared-config.example.yml`), que eram uma tentativa anterior
(sem Docker) e não valem mais.

## 1. Como funciona

- O painel roda dentro de **Docker** (imagem com Python, Chromium para os robôs, LibreOffice para
  PDF e o driver do SQL Server do UAU). O servidor só precisa ter Docker.
- Na frente fica o **Caddy**, que cuida do HTTPS (certificado grátis e renovação automática).
- `dados/` (planilhas, usuários, certificado A1, caches) **não entra na imagem**: fica numa pasta
  do servidor (`/opt/dashboard/dados`) e é montada dentro do container.
- **Atualização sem queda (blue/green):** a versão nova sobe ao lado da atual, é testada
  (precisa carregar os dados do painel de verdade), e só então o Caddy troca para ela. Se qualquer
  teste falhar, a versão antiga nem chega a sair do ar. Quem usa o painel não percebe a troca.

## 2. O que você precisa providenciar

| Item | Detalhe |
|---|---|
| Servidor | Ubuntu 24.04 (ou 22.04), **2 vCPU, 8 GB de RAM** (mínimo 4 GB), 60 GB de disco. **Região Brasil** — alguns sites do governo (SEFAZ/Receita) recusam IP de fora. |
| Domínio | Um endereço (ex.: `painel.suaempresa.com.br`) com registro DNS tipo **A** apontando para o IP do servidor. |
| UAU (SQL Server) | O IP público do servidor novo precisa ser **liberado no firewall do SQL Server do UAU** (porta 1433). |
| Acesso SSH | Colocar no servidor a chave pública que está em `C:\Users\elvis_henryque\.ssh\dash_deploy.pub` (arquivo `~/.ssh/authorized_keys` do usuário `root` ou de um usuário com Docker). |

## 3. Primeira publicação (a ordem importa)

1. **Preparar o servidor** (uma vez, como root): no Windows,
   `scp deploy\provision.sh root@IP_DO_SERVIDOR:/root/` e depois
   `ssh root@IP_DO_SERVIDOR "bash /root/provision.sh"`. Instala Docker, firewall (só 22/80/443),
   fail2ban, fuso de São Paulo e cria `/opt/dashboard`.
2. **Configurar o acesso:** copie `deploy/server.conf.example` para `deploy/server.conf` e preencha.
3. **Segredos:** no Windows rode
   `powershell -ExecutionPolicy Bypass -File deploy\gerar_env.ps1 -Dominio SEU.DOMINIO -Email SEU@EMAIL`
   (monta `deploy\servidor.env` a partir do `iniciar_dashboard.bat`, sem mostrar senhas), envie por
   `scp deploy\servidor.env root@SERVIDOR:/opt/dashboard/.env`, rode
   `ssh root@SERVIDOR "chmod 600 /opt/dashboard/.env"` e **apague o `servidor.env` local**.
4. **Dados — ANTES de abrir ao público:** `powershell -ExecutionPolicy Bypass -File deploy\enviar_dados.ps1`
   (use `-SemModelosGrandes` para pular os `.nwd`, ~500 MB, e mandá-los depois). Sem o `dados/`
   o primeiro visitante poderia criar o administrador.
5. **Subir o proxy** (uma vez), no servidor: copie `deploy/docker-compose.proxy.yml` e `deploy/Caddyfile`
   para `/opt/dashboard` e rode `cd /opt/dashboard && docker compose --env-file .env -f docker-compose.proxy.yml up -d`.
6. **Primeira versão do painel:** `powershell -ExecutionPolicy Bypass -File deploy\deploy.ps1`.
7. **Agendamentos** (robô da Carta Convite e backup): no servidor, `crontab -e` e cole `deploy/crontab.example`
   (copie `run-in-active.sh` e `backup.sh` já vão junto no passo 6).
8. Abra o endereço, entre com seu login e confira as abas (principalmente Notas Fiscais, UAU e Melius).

## 4. Atualizar (toda alteração no painel)

```
powershell -ExecutionPolicy Bypass -File deploy\deploy.ps1
```

Confere sintaxe → empacota só o código (≈1 MB) → envia → o servidor monta a imagem, sobe a versão
nova, testa, troca e aposenta a antiga. Leva ~1–3 minutos e o painel **continua no ar o tempo todo**.
Para só validar sem enviar: `-SoVerificar`.

**Voltar atrás** (no servidor): `/opt/dashboard/rollback.sh` — volta para a release anterior pelo mesmo
caminho sem queda. O servidor guarda as 5 últimas.

## 5. Backup

`backup.sh` (cron das 02:30) guarda 14 dias em `/opt/dashboard/backups`. Isso protege de erro,
**não** de perda do servidor: configure uma cópia fora dele (há uma linha de `rclone` comentada no fim do script).
Restaurar: `tar -xzf dados_AAAAMMDD_HHMM.tgz -C /opt/dashboard` e `chown -R 1000:1000 /opt/dashboard/dados`.

## 6. O que muda no Linux

- **Cronograma (.mpp):** ler direto do MS Project só existe no Windows. No servidor o botão
  "Atualizar do Project" avisa disso e o painel mostra o último cronograma sincronizado. Para atualizar:
  rode a sincronização no seu computador e envie o `mpp_cronograma_cache.json` para `dados/synergia/`.
- **PDF do RMO e da Carta Convite:** no Windows saem pelo PowerPoint/Word; no servidor, pelo LibreOffice.
  O resultado é parecido, mas pode haver pequena diferença de fonte/espaçamento — confira o primeiro PDF.
- **Senha de e-mail salva na Carta Convite:** no Windows ficava protegida pelo seu usuário (DPAPI).
  No servidor usa outra proteção; a senha salva antes **não migra** — é só salvar de novo uma vez.
- **Link temporário do Cloudflare:** não existe mais; o endereço agora é fixo. O botão
  "Atualizar link de acesso" informa isso.
- **Upload de arquivos grandes** (NWD/IFC): funciona direto pelo Caddy (limite 2 GB). Se um dia
  colocar o Cloudflare na frente, o plano gratuito corta uploads acima de 100 MB.

## 7. Segurança já aplicada

- Login com **trava contra tentativa em massa** (5 erros por e-mail / 20 por IP em 15 min).
- Criar o primeiro administrador exige o `DASHBOARD_SETUP_TOKEN`.
- Cookie de sessão `HttpOnly`, `SameSite=Lax` e `Secure` (só trafega por HTTPS).
- A rota pública `/api/status` não mostra mais caminhos de pasta do servidor.
- Firewall só com 22/80/443, `fail2ban` no SSH, atualizações de segurança automáticas.
- Segredos só em `/opt/dashboard/.env` (permissão 600), nunca dentro da imagem nem do código.

## 8. Problemas comuns

| Sintoma | O que ver |
|---|---|
| `deploy.ps1` diz "A versão antiga continua no ar" | Leia o trecho do log que ele mostra; é o que falhou antes de trocar. |
| Painel abre mas Notas Fiscais dá erro de certificado | `CERT_PFX_SENHA` no `.env` e o arquivo em `dados/synergia/certificado/`. |
| UAU não carrega | IP do servidor liberado no firewall do SQL; teste `docker exec dash-<cor> python -c "import pyodbc;print(pyodbc.drivers())"`. |
| Robô do Sienge/Melius falha | `docker logs dash-<cor> --tail 100`; sem monitor o Chrome roda sempre em modo oculto. |
| Ver qual versão está no ar | `curl https://SEU.DOMINIO/api/status` (campo `build`). |
| Ver logs | `docker logs -f dash-blue` ou `dash-green` (a que estiver ativa: `cat /opt/dashboard/state/active_color`). |

## 7. Segurança (resumo para quem mantém o painel)

- O painel só fala com a internet pelo Caddy (HTTPS). A porta 5070 do container não é publicada.
- Senhas: mínimo de 12 caracteres; sessão expira em 12 h sem uso e cai ao trocar a senha.
- `DASHBOARD_FRAME_ANCESTORS` define quem pode embutir o painel em iframe (padrão do `env.example`: o SharePoint
  `https://temazeccombr.sharepoint.com`). Vazio = ninguém.
- Permissão por aba é conferida no servidor (tabela `_TAB_API_PREFIXES` em `app.py`). Ao criar uma aba/rota nova,
  inclua o prefixo ali.
- Nunca versione `dados/`, `.env`, certificados (`.pfx`), `.claude/` ou planilhas: o `.gitignore` já bloqueia.
- Se uma senha aparecer em texto no repositório ou em conversa, troque-a imediatamente.

