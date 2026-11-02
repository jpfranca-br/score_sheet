# Shooting Score Sheet Generator

This project generates PDF score sheets from an Excel `.xlsx` spreadsheet.

## Download the project

On Windows, macOS, or Linux, install Git and run:

```bash
git clone https://github.com/jpfranca-br/score_sheet.git
cd score_sheet
```

To update an existing copy:

```bash
git pull
```

## Main files

- `gerar_sumulas.py`: generates a PDF locally from the command line.
- `app_web.py`: provides a PDF generator through a web page.
- `provas.xlsx`: sample spreadsheet containing the `Provas` worksheet fields.
- `requirements.txt`: dependencies for the local generator.
- `requirements-web.txt`: dependencies for the web version.
- `deploy/`: files for server deployment.

## Use locally on Windows

Create a Python environment, install the dependencies, and generate the PDF:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
python gerar_sumulas.py
```

By default, the command reads `provas.xlsx` from the project folder and creates `sumulas.pdf` next to the spreadsheet. On Windows, the PDF opens automatically in the default PDF application.

You can also specify different input and output files:

```powershell
python gerar_sumulas.py caminho\da\planilha.xlsx
python gerar_sumulas.py caminho\da\planilha.xlsx caminho\do\resultado.pdf
```

## Use locally on Ubuntu

Install Python and Git if they are not already available:

```bash
sudo apt update
sudo apt install -y git python3 python3-venv
```

Clone the project and create a virtual environment:

```bash
git clone https://github.com/jpfranca-br/score_sheet.git
cd score_sheet
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
```

### Generate a PDF from the command line

Install the local dependencies and run the generator:

```bash
pip install -r requirements.txt
python gerar_sumulas.py
```

The default command reads `provas.xlsx` and creates `sumulas.pdf` in the project folder. To use another spreadsheet and output location:

```bash
python gerar_sumulas.py /path/to/spreadsheet.xlsx
python gerar_sumulas.py /path/to/spreadsheet.xlsx /path/to/output.pdf
```

### Run the web generator on Ubuntu

Install the web dependencies and start the application:

```bash
pip install -r requirements-web.txt
python app_web.py
```

Open `http://127.0.0.1:8000` in a browser on the same Ubuntu computer. Stop the server with `Ctrl+C`.

## Use the web generator on the local network

Install the web dependencies as described for Windows or Ubuntu:

```bash
pip install -r requirements-web.txt
python app_web.py
```

Open `http://127.0.0.1:8000` in the browser on the same computer. The development server only listens on the loopback address `127.0.0.1`; use a production server or configure an authorized exposure method before allowing other devices to connect.

## Local or web: which one should I use?

| Use | Program | Input | Output |
|---|---|---|---|
| Desktop/command line | `gerar_sumulas.py` | Local spreadsheet | Local PDF file |
| Browser | `app_web.py` | Spreadsheet uploaded through the browser | PDF download |

`gerar_sumulas.py` is the simplest way to produce PDFs on one machine. It writes the PDF to disk and does not require a browser.

`app_web.py` provides a web interface. Uploaded files are processed in memory and are not stored permanently by the application. In development, the built-in Flask server accepts connections only on `127.0.0.1`.

## Run on a Linux server

See the complete deployment guide in [DEPLOY.md](DEPLOY.md).

Summary:

1. Install the Python packages for the web version.
2. Start Gunicorn on local port `127.0.0.1:8000`.
3. Configure the system service with `deploy/sumulas.service`.
4. Choose how the service should be exposed to the network; the documentation does not prescribe a proxy or tunnel solution.
