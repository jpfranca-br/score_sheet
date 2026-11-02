# Publicar o Gerador de Súmulas em um servidor Linux

Este guia prepara o aplicativo e um serviço que escuta apenas em
`127.0.0.1:8000`. A escolha de proxy, rede privada, domínio, HTTPS ou túnel
fica a cargo de quem administra o servidor.

O aplicativo não grava a planilha enviada no disco: o arquivo é processado em
memória e o PDF é devolvido diretamente ao navegador.

## Arquivos que vão para o servidor

```
/opt/sumulas/
├── app_web.py
├── gerar_sumulas.py
├── provas.xlsx
├── requirements.txt
├── requirements-web.txt
└── deploy/
    └── sumulas.service
```

## 1. Pacotes e usuário

```bash
sudo apt update
sudo apt install -y python3-venv
sudo useradd --system --no-create-home --shell /usr/sbin/nologin sumulas
sudo mkdir -p /opt/sumulas
sudo chown -R $USER:$USER /opt/sumulas   # seu usuário passa a poder enviar arquivos para lá
```

Agora copie os arquivos para `/opt/sumulas` ou use `git clone`.
Do Windows, no PowerShell, de dentro da pasta do projeto:

```powershell
scp -r app_web.py gerar_sumulas.py provas.xlsx requirements.txt requirements-web.txt deploy usuario@IP_DA_VPS:/opt/sumulas/
```

O usuário `sumulas` (que roda o serviço) só precisa ler os arquivos, e as permissões
padrão já permitem isso.

## 2. Ambiente Python

```bash
cd /opt/sumulas
python3 -m venv .venv
.venv/bin/pip install -r requirements-web.txt
```

Teste rápido: `.venv/bin/gunicorn --bind 127.0.0.1:8000 app_web:app`,
e em outro terminal `curl 127.0.0.1:8000/saude` deve responder `{"status":"ok"}`.

## 3. Serviço (sobe sozinho e reinicia se cair)

```bash
sudo cp deploy/sumulas.service /etc/systemd/system/sumulas.service
sudo systemctl daemon-reload
sudo systemctl enable --now sumulas
sudo systemctl status sumulas
```

## 4. Exposição do serviço

O serviço permanece acessível somente em `127.0.0.1:8000`. A próxima etapa é
escolher o método de exposição da rede: Cloudflare Tunnel, outro túnel,
proxy reverso autorizado pelo administrador ou acesso restrito à rede interna.

Qualquer método escolhido deve encaminhar as requisições para
`http://127.0.0.1:8000` e aceitar uploads de até 5 MB.

## Manutenção

- Logs: `sudo journalctl -u sumulas -f`
- Atualizar o código: atualize os arquivos em `/opt/sumulas` e execute `sudo systemctl restart sumulas`
- Trocar a planilha modelo: substitua `/opt/sumulas/provas.xlsx` (não precisa reiniciar)
- Limites (em `app_web.py`): `MAX_MB = 5` por arquivo, `MAX_PAGINAS = 300` por planilha.
