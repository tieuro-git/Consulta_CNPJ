from flask import Flask, request, send_file
import requests
from openpyxl import Workbook
import io

app = Flask(__name__)

BASE_URL = "https://publica.cnpj.ws/cnpj"

def limpar_cnpj(s: str) -> str:
    return s.replace(".", "").replace("/", "").replace("-", "").strip()

def consulta_publica(cnpj: str) -> dict:
    url = f"{BASE_URL}/{cnpj}"
    r = requests.get(url, timeout=30)
    return r.json() if r.status_code == 200 else {"erro": f"HTTP {r.status_code}"}

def extrair_campos(dados: dict):
    if "erro" in dados:
        return "", "", ""
    razao = dados.get("razao_social", "")
    est   = dados.get("estabelecimento") or {}
    municipio = (est.get("cidade") or {}).get("nome", "")
    uf       = (est.get("estado") or {}).get("sigla", "")
    return razao, municipio, uf

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        cnpjs_raw = request.form["cnpjs"].splitlines()
        cnpjs = [limpar_cnpj(c) for c in cnpjs_raw if c.strip()]

        wb = Workbook()
        ws = wb.active
        ws.append(["CNPJ", "RAZÃO SOCIAL", "MUNICÍPIO", "ESTADO"])

        for cnpj in cnpjs:
            dados = consulta_publica(cnpj)
            razao, municipio, uf = extrair_campos(dados)
            ws.append([cnpj, razao, municipio, uf])

        file_stream = io.BytesIO()
        wb.save(file_stream)
        file_stream.seek(0)

        return send_file(file_stream, as_attachment=True, download_name="resultado.xlsx")

    return """
    <form method="post">
        <textarea name="cnpjs" rows="10" cols="40"></textarea><br>
        <button type="submit">Consultar e baixar Excel</button>
    </form>
    """

# 🚨 IMPORTANTE: para rodar no Vercel
def handler(event, context):
    return app(event, context)
