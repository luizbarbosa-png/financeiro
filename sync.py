import os
import json
from datetime import date, timedelta

import requests
import gspread
from google.oauth2.service_account import Credentials

API = "https://api.pluggy.ai"
CLIENT_ID = os.environ["PLUGGY_CLIENT_ID"]
CLIENT_SECRET = os.environ["PLUGGY_CLIENT_SECRET"]
ITEM_IDS = [i.strip() for i in os.environ["PLUGGY_ITEM_IDS"].split(",") if i.strip()]
SHEET_ID = os.environ["SHEET_ID"]
GOOGLE_CREDS = json.loads(os.environ["GOOGLE_CREDENTIALS"])
DIAS = int(os.environ.get("DIAS", "60"))  # quantos dias para trás buscar

HEADER = ["id", "data", "banco", "conta", "descricao", "categoria",
          "tipo", "valor", "mes", "entrada", "saida", "pagador"]


def pluggy_headers():
    r = requests.post(f"{API}/auth",
                      json={"clientId": CLIENT_ID, "clientSecret": CLIENT_SECRET},
                      timeout=30)
    r.raise_for_status()
    return {"X-API-KEY": r.json()["apiKey"]}


def get(path, headers, **params):
    r = requests.get(f"{API}{path}", headers=headers, params=params, timeout=60)
    r.raise_for_status()
    return r.json()


def safe(texto):
    """Evita que o Sheets interprete texto como fórmula."""
    texto = "" if texto is None else str(texto)
    return "'" + texto if texto[:1] in ("=", "+", "-", "@") else texto


def buscar_transacoes(headers):
    inicio = (date.today() - timedelta(days=DIAS)).isoformat()
    fim = date.today().isoformat()
    linhas = []
    for item_id in ITEM_IDS:
        item = get(f"/items/{item_id}", headers)
        banco = (item.get("connector") or {}).get("name", "Banco")
        contas = get("/accounts", headers, itemId=item_id)["results"]
        for conta in contas:
            dados = get("/v2/transactions", headers, accountId=conta["id"],
                        dateFrom=inicio + "T00:00:00.000Z",
                        dateTo=fim + "T23:59:59.999Z")
            while True:
                for t in dados["results"]:
                    dia = t["date"][:10]
                    valor = abs(float(t["amount"]))
                    if t.get("type") == "DEBIT":
                        valor = -valor
                    linhas.append([
                        t["id"], dia, banco, conta.get("name", ""),
                        safe(t.get("description")), safe(t.get("category")),
                        t.get("type", ""), valor, "'" + dia[:7],
                        valor if valor > 0 else 0,
                        -valor if valor < 0 else 0,
                        safe(((t.get("paymentData") or {}).get("payer") or {})
                             .get("name")),
                    ])
                proximo = dados.get("next")
                if not proximo:
                    break
                dados = get(f"/v2/transactions{proximo}", headers)
    return linhas


def preparar_planilha(sh):
    nomes = [w.title for w in sh.worksheets()]
    if "Transacoes" not in nomes:
        ws = sh.add_worksheet("Transacoes", rows=2000, cols=len(HEADER))
        ws.append_row(HEADER)
    else:
        ws = sh.worksheet("Transacoes")
        if len(ws.row_values(1)) < len(HEADER):
            # migração única: adiciona a coluna "pagador" e recarrega o histórico
            if ws.col_count < len(HEADER):
                ws.add_cols(len(HEADER) - ws.col_count)
            ws.update(range_name="A1", values=[HEADER])
            ws.batch_clear(["A2:L"])
    if "Resumo" not in nomes:
        r = sh.add_worksheet("Resumo", rows=200, cols=12)
        T = "Transacoes!"
        r.update(range_name="A1", value_input_option="USER_ENTERED", values=[
            ["Mês", "Entradas", "Saídas", "Saldo do mês"],
            [f'=IFERROR(SORT(UNIQUE(FILTER({T}I2:I,{T}I2:I<>""))),"")',
             f'=ARRAYFORMULA(IF(A2:A="","",SUMIFS({T}J:J,{T}I:I,A2:A)))',
             f'=ARRAYFORMULA(IF(A2:A="","",SUMIFS({T}K:K,{T}I:I,A2:A)))',
             '=ARRAYFORMULA(IF(A2:A="","",B2:B-C2:C))'],
        ])
        r.update(range_name="F1", value_input_option="USER_ENTERED", values=[
            [f'=IFERROR(QUERY({T}A2:K,"select Col6, sum(Col11) where Col11>0 '
             f"group by Col6 order by sum(Col11) desc "
             f"label Col6 'Categoria', sum(Col11) 'Gasto'\",0),\"\")"]
        ])
        r.update(range_name="I1", value_input_option="USER_ENTERED", values=[
            ["Entradas totais", "Saídas totais", "Resultado"],
            [f"=SUM({T}J:J)", f"=SUM({T}K:K)", "=I2-J2"],
        ])
    return ws


def main():
    headers = pluggy_headers()
    linhas = buscar_transacoes(headers)

    creds = Credentials.from_service_account_info(
        GOOGLE_CREDS,
        scopes=["https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive"])
    sh = gspread.authorize(creds).open_by_key(SHEET_ID)
    ws = preparar_planilha(sh)

    ja_existem = set(ws.col_values(1)[1:])
    novas = sorted((l for l in linhas if l[0] not in ja_existem),
                   key=lambda l: l[1])
    if novas:
        ws.append_rows(novas, value_input_option="USER_ENTERED")
    print(f"{len(linhas)} transações lidas, {len(novas)} novas gravadas.")


if __name__ == "__main__":
    main()
