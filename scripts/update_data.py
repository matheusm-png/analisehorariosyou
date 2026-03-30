import pandas as pd
import json
import re
import sys
from io import StringIO
import requests

SHEET_ID = "1Een2tWifGxC0e-fgXVzVH9-QsUrzLrf97rCrq2rqdNk"
GID = "0"
CSV_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={GID}"
HTML_PATH = "index.html"


def fetch_csv():
    r = requests.get(CSV_URL, headers={"User-Agent": "Mozilla/5.0"}, timeout=30)
    r.raise_for_status()
    r.encoding = "utf-8"
    return r.text


def map_email(v):
    if pd.isna(v):
        return "Invalido"
    v = str(v).strip().lower()
    if v == "sim":
        return "Sim"
    if v in ("nao", "não"):
        return "Outro"
    return "Invalido"


def map_porte(v):
    if pd.isna(v):
        return "NI"
    v = str(v).strip().replace(" ", "")
    lut = {
        "20-30": "20-30",
        "31-60": "31-60", "30-60": "31-60", "31-100": "31-60",
        "61-120": "61-120", "61-300": "61-120", "121-300": "61-120",
        "101-500": "101-500", "201-500": "101-500",
        "301-500": "301-500",
        "500+": "500+", "501+": "500+",
    }
    return lut.get(v, "NI")


def parse_hour(v):
    if pd.isna(v):
        return 0
    try:
        return int(str(v).replace("h", "").strip())
    except Exception:
        return 0


def parse_dow(v):
    if pd.isna(v):
        return 0
    try:
        return pd.Timestamp(v).weekday()
    except Exception:
        return 0


def norm_status(v):
    if pd.isna(v):
        return "NI"
    v = str(v).strip()
    lut = {
        "vendido": "VENDIDO",
        "agendado": "Agendado",
        "descartado": "Descartado",
        "em tentativa": "Em tentativa",
        "agendamento q": "Agendamento q",
        "agendamento nq": "Agendamento nq",
    }
    return lut.get(v.lower(), v)


def build_raw(df):
    raw = []
    for _, r in df.iterrows():
        raw.append({
            "h": parse_hour(r["Hora"]),
            "d": parse_dow(r["Data"]),
            "m": str(r["Mês"]) if pd.notna(r["Mês"]) else "NI",
            "s": norm_status(r["Status"]),
            "o": str(r["Origem"]).strip() if pd.notna(r["Origem"]) else "NI",
            "e": map_email(r["Email Corporativo?"]),
            "p": map_porte(r["NR de Colab"]),
        })
    return raw


def main():
    print("Buscando dados do Google Sheets...")
    csv_text = fetch_csv()
    df = pd.read_csv(StringIO(csv_text), dtype=str)
    df["Data"] = pd.to_datetime(df["Data"], dayfirst=True, errors="coerce")

    raw = build_raw(df)
    print(f"  {len(raw)} registros processados")

    raw_json = json.dumps(raw, ensure_ascii=False, separators=(",", ":"))

    with open(HTML_PATH, "r", encoding="utf-8") as f:
        html = f.read()

    new_html = re.sub(
        r"const RAW = \[.*?\];",
        f"const RAW = {raw_json};",
        html, count=1, flags=re.DOTALL,
    )

    if new_html == html:
        print("Nenhuma alteração detectada.")
        sys.exit(0)

    with open(HTML_PATH, "w", encoding="utf-8") as f:
        f.write(new_html)

    print("index.html atualizado com sucesso.")


if __name__ == "__main__":
    main()
