# Gerador de Súmulas de Tiro

Este projeto gera súmulas em PDF a partir de uma planilha Excel `.xlsx`.

## Baixar o projeto

No Windows, macOS ou Linux, instale o Git e execute:

```bash
git clone https://github.com/jpfranca-br/score_sheet.git
cd score_sheet
```

Para atualizar uma cópia já instalada:

```bash
git pull
```

## Arquivos principais

- `gerar_sumulas.py`: gera o PDF localmente a partir da linha de comando.
- `app_web.py`: disponibiliza um gerador de PDF em uma página web.
- `provas.xlsx`: planilha modelo com os campos da aba `Provas`.
- `requirements.txt`: dependências do gerador local.
- `requirements-web.txt`: dependências da versão web.
- `deploy/`: arquivos auxiliares de implantação em servidor.

## Usar localmente no Windows

Crie um ambiente Python, instale as dependências e gere o PDF:

```powershell
py -m venv .venv
.\.venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt
python gerar_sumulas.py
```

O comando padrão lê `provas.xlsx` na própria pasta do projeto e cria `sumulas.pdf` ao lado da planilha. A biblioteca PDF será aberta automaticamente no Windows.

Também é possível informar arquivos diferentes:

```powershell
python gerar_sumulas.py caminho\da\planilha.xlsx
python gerar_sumulas.py caminho\da\planilha.xlsx caminho\do\resultado.pdf
```

## Usar o gerador pela rede local

Instale também as dependências da versão web:

```powershell
pip install -r requirements-web.txt
python app_web.py
```

Abra `http://127.0.0.1:8000` no navegador deste mesmo computador. Para que outros dispositivos da rede acessem a aplicação, execute-a com um servidor adequado ou configure separadamente o modo de exposição desejado.

## Local ou web: qual diferença?

| Uso | Programa | Entrada | Saída |
|---|---|---|---|
| Desktop/linha de comando | `gerar_sumulas.py` | Planilha local | Arquivo PDF local |
| Navegador | `app_web.py` | Upload de uma planilha pelo navegador | PDF enviado para download |

`gerar_sumulas.py` é o modo mais simples para produzir PDFs em uma máquina. Ele grava o PDF no disco e não precisa de navegador.

`app_web.py` oferece uma interface web. Neste modo, o arquivo enviado é processado em memória; a aplicação não guarda permanentemente a planilha do usuário. Em desenvolvimento, ele usa o servidor integrado do Flask e aceita conexões apenas em `127.0.0.1`.

## Executar em um servidor Linux

O conteúdo completo está em [DEPLOY.md](DEPLOY.md).

Resumo:

1. Instale os pacotes Python da versão web.
2. Inicie o Gunicorn na porta local `127.0.0.1:8000`.
3. Configure o serviço do sistema com `deploy/sumulas.service`.
4. Escolha como expor o serviço à Internet; esta documentação não impõe uma solução de proxy ou túnel.
