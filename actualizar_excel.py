import yfinance as yf
import pandas as pd
from datetime import date
from dateutil.relativedelta import relativedelta

tickers = {
    "ECL": "ECL.SN",
    "ITAU": "ITAUCL.SN",
    "COLBUN": "COLBUN.SN",
    "PARAUCO": "PARAUCO.SN",
    "CENCOMALLS": "CENCOMALLS.SN",
    "AGUAS-A": "AGUAS-A.SN",
    "CHILE": "CHILE.SN",
    "CAP": "CAP.SN",
    "SQM-B": "SQM-B.SN",
    "PUCOBRE": "PUCOBRE.SN",
    "BCI": "BCI.SN",
    "LATAM": "LTM.SN",
    "FALABELLA": "FALABELLA.SN",
    "NORTEGRANDE": "NORTEGRAN.SN",
    "INVERSIONES LA CONSTRUCCION": "ILC.SN",
    "SALMONES CAMANCHACA": "SALMOCAM.SN",
    "SALFACORP": "SALFACORP.SN",
    "ENAEX": "ENAEX.SN",
    "BESALCO": "BESALCO.SN",
    "AGUAS ANDINAS": "AGUAS-A.SN"
}

output_file = "input_actualizado.xlsx"

end_date = date.today()
start_date = end_date - relativedelta(years=5)

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:

    for sheet_name, yahoo_ticker in tickers.items():

        print(f"Descargando {sheet_name} ({yahoo_ticker}) desde {start_date} hasta {end_date}...")

        df = yf.download(
            yahoo_ticker,
            start=start_date,
            end=end_date,
            interval="1d",
            auto_adjust=False,
            progress=False,
            group_by="column"
        )

        if df.empty:
            print(f"Sin datos para {sheet_name}")
            continue

        if isinstance(df.columns, pd.MultiIndex):
            df.columns = [col[0] for col in df.columns]

        df = df.reset_index()

        df["Date"] = pd.to_datetime(df["Date"]).dt.date

        df["Daily Return"] = df["Adj Close"].pct_change()

        df = df[
            [
                "Date",
                "Adj Close",
                "Close",
                "High",
                "Low",
                "Open",
                "Volume",
                "Daily Return"
            ]
        ]

        # Opcional: dejar la fecha más reciente arriba
        df = df.sort_values("Date", ascending=False)

        df.to_excel(
            writer,
            sheet_name=sheet_name[:31],
            index=False
        )

print("Excel actualizado correctamente")