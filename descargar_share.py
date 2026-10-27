import os
import json
import urllib.request
import urllib.parse
from datetime import datetime, timezone, timedelta
import pandas as pd

DEVICE_ID = "YkliaG1nckQ4RFlBZWE5YWFRWEMrZz09"
AUTHORIZE = "5MMNHG"
EST_ID = "364053"

# Zona horaria fija de Salta, Argentina (UTC-3)
TZ_ARGENTINA = timezone(timedelta(hours=-3))

def solicitar_export(fecha_str):
    url_export = "https://www.ecowitt.net/index/export_excel"
    params = urllib.parse.urlencode({
        "device_id": DEVICE_ID,
        "authorize": AUTHORIZE,
        "mode": "0",
        "sdate": f"{fecha_str} 00:00",
        "edate": f"{fecha_str} 23:59",
        "sortList": "1|3|4|5|6",
        "hideList": ""
    }).encode("utf-8")

    req = urllib.request.Request(
        url_export,
        data=params,
        headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
            "Referer": f"https://www.ecowitt.net/home/share?authorize={AUTHORIZE}&device_id={DEVICE_ID}"
        }
    )

    with urllib.request.urlopen(req, timeout=15) as response:
        return json.loads(response.read().decode("utf-8"))

def descargar_datos_sin_login():
    # 1. Calculamos la fecha según la hora local de Salta
    ahora_salta = datetime.now(TZ_ARGENTINA)
    hoy = ahora_salta.strftime("%Y-%m-%d")
    print(f"1. Solicitando exportación para la fecha local (Salta): {hoy}...")
    
    res_data = solicitar_export(hoy)
    
    # Si justo a medianoche no hay datos del nuevo día, toma el día anterior
    if res_data.get("errcode") != "0" or "url" not in res_data:
        ayer = (ahora_salta - timedelta(days=1)).strftime("%Y-%m-%d")
        print(f"   Aviso: {res_data.get('errmsg')}. Probando con el día anterior ({ayer})...")
        res_data = solicitar_export(ayer)

    if res_data.get("errcode") != "0" or "url" not in res_data:
        raise Exception(f"Error devuelto por Ecowitt: {res_data}")

    download_url = res_data["url"]
    print(f"2. URL del archivo generada: {download_url}")
    
    xlsx_filename = f"{EST_ID}_daily_station_data.xlsx"
    xls_filename = f"{EST_ID}_daily_station_data.xls"
    csv_filename = f"{EST_ID}_daily_station_data.csv"
    
    # 3. Descargar archivos
    print(f"3. Descargando {xlsx_filename}...")
    urllib.request.urlretrieve(download_url, xlsx_filename)
    urllib.request.urlretrieve(download_url, xls_filename)
    
    # 4. Convertir a CSV
    print(f"4. Convirtiendo a {csv_filename}...")
    df = pd.read_excel(xlsx_filename, header=None)
    df.to_csv(csv_filename, index=False, header=False, encoding="utf-8-sig")
    print(f"   -> CSV generado con éxito ({os.path.getsize(csv_filename):,} bytes).")
    
    print("\n¡Descarga y conversión completada exitosamente sin login!")

if __name__ == "__main__":
    descargar_datos_sin_login()
