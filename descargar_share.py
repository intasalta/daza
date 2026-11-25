import os
import json
import shutil
import urllib.request
import urllib.parse
from datetime import datetime, timezone, timedelta
import pandas as pd

# Zona horaria de las estaciones en Salta, Argentina (UTC-3)
TZ_ARGENTINA = timezone(timedelta(hours=-3))

ESTACIONES = [
    {
        "nombre": "Colegio Secundario N° 5085 Dr. Mariano Moreno",
        "url": "https://www.ecowitt.net/home/share?authorize=5MMNHG&device_id=YkliaG1nckQ4RFlBZWE5YWFRWEMrZz09",
        "default_id": "364053"
    },
    {
        "nombre": "La Charqueada",
        "url": "https://www.ecowitt.net/home/share?authorize=123DE0&device_id=bHRjQndWMmNTR2lnV1llNFczOW9TUT09",
        "default_id": "370040"
    }
]

def extraer_credenciales_url(share_url):
    """Extrae authorize y device_id de una URL de compartir de Ecowitt."""
    parsed = urllib.parse.urlparse(share_url)
    params = urllib.parse.parse_qs(parsed.query)
    auth = params.get("authorize", [""])[0]
    dev_id = params.get("device_id", [""])[0]
    return auth, dev_id

def consultar_info_estacion(auth, share_url):
    """Consulta la API de Ecowitt para obtener el share_id (ID numérico) y nombre real."""
    url = f"https://www.ecowitt.net/index/get_device_list?authorize={auth}"
    req = urllib.request.Request(
        url,
        headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
            "Referer": share_url,
            "authorize": auth
        }
    )
    try:
        with urllib.request.urlopen(req, timeout=12) as response:
            res = json.loads(response.read().decode("utf-8"))
            if res.get("errcode") == "0" and res.get("list"):
                item = res["list"][0]
                return item.get("share_id"), item.get("name")
    except Exception as e:
        print(f"   [Aviso] No se pudo obtener lista de dispositivos: {e}")
    return None, None

def solicitar_export_excel(device_id, authorize, share_url, fecha_str):
    """Intenta solicitar exportación de Excel directa si el enlace de share tiene permisos."""
    url_export = "https://www.ecowitt.net/index/export_excel"
    params = urllib.parse.urlencode({
        "device_id": device_id,
        "authorize": authorize,
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
            "Referer": share_url,
            "authorize": authorize
        }
    )

    with urllib.request.urlopen(req, timeout=15) as response:
        return json.loads(response.read().decode("utf-8"))

def descargar_via_get_data(device_id, authorize, share_url, est_id, fecha_str):
    """Descarga los registros diarios mediante la API index/get_data y genera el CSV/XLSX/XLS
    con el formato de doble encabezado estándar de Ecowitt.
    """
    url_get = "https://www.ecowitt.net/index/get_data"
    params = urllib.parse.urlencode({
        "device_id": device_id,
        "authorize": authorize,
        "is_list": "1",
        "mode": "0",
        "sdate": f"{fecha_str} 00:00",
        "edate": f"{fecha_str} 23:59",
        "page": "1"
    }).encode("utf-8")

    req = urllib.request.Request(
        url_get,
        data=params,
        headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
            "Referer": share_url,
            "authorize": authorize
        }
    )

    with urllib.request.urlopen(req, timeout=20) as response:
        res = json.loads(response.read().decode("utf-8"))

    if res.get("errcode") != "0" or not res.get("list"):
        raise Exception(f"Error devuelto por index/get_data: {res}")

    raw_items = res["list"]
    # Los registros vienen del más reciente al más antiguo; los invertimos para orden cronológico
    items = raw_items[::-1]

    header_row_0 = ["Time", "Outdoor", "", "", "", "", "Solar and UVI", "", "Rainfall", "", "", "", "", "", "", "", "Wind", "", "", "", "Pressure", ""]
    header_row_1 = [
        "Time", "Temperature(℃)", "Feels Like(℃)", "Dew Point(℃)", "Humidity(%)", "VPD(kPa)",
        "Solar(W/m²)", "UVI", "Rain Rate(mm/hr)", "Daily(mm)", "Event(mm)", "Hourly(mm)",
        "24 Hours(mm)", "Weekly(mm)", "Monthly(mm)", "Yearly(mm)", "Wind Speed(km/h)",
        "Wind Gust(km/h)", "Wind Direction(º)", "10-minute Average Wind Direction(º)",
        "Relative(hPa)", "Absolute(hPa)"
    ]

    rows = [header_row_0, header_row_1]
    for it in items:
        rows.append([
            it.get("create_time", ""),
            it.get("tempf", "-"),
            it.get("sendible_temp", "-"),
            it.get("drew_temp", "-"),
            it.get("humidity", "-"),
            it.get("vpd", "-"),
            it.get("solarradiation", "-"),
            it.get("uv", "-"),
            it.get("rainratein", "-"),
            it.get("dailyrainin", "-"),
            it.get("eventrainin", "-"),
            it.get("hourlyrainin", "-"),
            it.get("last24hrainin", "-"),
            it.get("weeklyrainin", "-"),
            it.get("monthlyrainin", "-"),
            it.get("yearlyrainin", "-"),
            it.get("windspeedmph", "-"),
            it.get("windgustmph", "-"),
            it.get("winddir", "-"),
            it.get("winddir_avg10m", "-"),
            it.get("baromrelin", "-"),
            it.get("baromabsin", "-")
        ])

    df = pd.DataFrame(rows)
    csv_filename = f"{est_id}_daily_station_data.csv"
    xlsx_filename = f"{est_id}_daily_station_data.xlsx"
    xls_filename = f"{est_id}_daily_station_data.xls"

    df.to_csv(csv_filename, index=False, header=False, encoding="utf-8-sig")
    df.to_excel(xlsx_filename, index=False, header=False)
    shutil.copyfile(xlsx_filename, xls_filename)

    print(f"   -> Generados {len(items)} registros vía API get_data:")
    print(f"      * CSV:  {csv_filename} ({os.path.getsize(csv_filename):,} bytes)")
    print(f"      * XLSX: {xlsx_filename} ({os.path.getsize(xlsx_filename):,} bytes)")
    print(f"      * XLS:  {xls_filename} ({os.path.getsize(xls_filename):,} bytes)")

def procesar_estacion(cfg):
    nombre = cfg["nombre"]
    share_url = cfg["url"]
    default_id = cfg.get("default_id", "")

    print(f"\n=======================================================")
    print(f"Procesando: {nombre}")
    print(f"URL: {share_url}")
    print(f"=======================================================")

    auth, dev_id = extraer_credenciales_url(share_url)
    if not auth or not dev_id:
        print(f"Error: URL inválida o faltan parámetros en {share_url}")
        return

    # Consultar ID real
    share_id, real_name = consultar_info_estacion(auth, share_url)
    est_id = share_id or default_id
    if real_name:
        print(f"Estación confirmada: {real_name} (ID: {est_id})")
    else:
        print(f"Estación ID: {est_id}")

    ahora_salta = datetime.now(TZ_ARGENTINA)
    hoy = ahora_salta.strftime("%Y-%m-%d")

    # Intento 1: export_excel directo
    usar_fallback_api = False
    try:
        print(f"1. Solicitando exportación directa para {hoy}...")
        res_data = solicitar_export_excel(dev_id, auth, share_url, hoy)
        if res_data.get("errcode") != "0" or "url" not in res_data:
            ayer = (ahora_salta - timedelta(days=1)).strftime("%Y-%m-%d")
            print(f"   Aviso: {res_data.get('errmsg')}. Probando día anterior ({ayer})...")
            res_data = solicitar_export_excel(dev_id, auth, share_url, ayer)

        if res_data.get("errcode") == "0" and "url" in res_data:
            download_url = res_data["url"]
            print(f"2. URL de descarga obtenida: {download_url}")

            xlsx_filename = f"{est_id}_daily_station_data.xlsx"
            xls_filename = f"{est_id}_daily_station_data.xls"
            csv_filename = f"{est_id}_daily_station_data.csv"

            print(f"3. Descargando {xlsx_filename}...")
            urllib.request.urlretrieve(download_url, xlsx_filename)
            shutil.copyfile(xlsx_filename, xls_filename)

            print(f"4. Convirtiendo a {csv_filename}...")
            df = pd.read_excel(xlsx_filename, header=None)
            df.to_csv(csv_filename, index=False, header=False, encoding="utf-8-sig")
            print(f"   -> Completado exitosamente vía export directo.")
        else:
            usar_fallback_api = True
    except Exception as e:
        print(f"   [Info] Export directo no disponible ({e}), utilizando método alternativo get_data...")
        usar_fallback_api = True

    # Intento 2: API get_data (soporta cualquier enlace compartido sin restricción de export)
    if usar_fallback_api:
        print("-> Usando sincronización por API get_data...")
        try:
            descargar_via_get_data(dev_id, auth, share_url, est_id, hoy)
        except Exception as e2:
            ayer = (ahora_salta - timedelta(days=1)).strftime("%Y-%m-%d")
            print(f"   Aviso en {hoy}: {e2}. Intentando día anterior ({ayer})...")
            descargar_via_get_data(dev_id, auth, share_url, est_id, ayer)

def descargar_todas_las_estaciones():
    print(f"Iniciando descarga y actualización climática: {datetime.now(TZ_ARGENTINA).strftime('%Y-%m-%d %H:%M:%S')} (Hora Salta)")
    for est in ESTACIONES:
        try:
            procesar_estacion(est)
        except Exception as e:
            print(f"Error procesando estación '{est.get('nombre')}': {e}")
    print("\n¡Proceso de descarga finalizado para todas las estaciones!")

if __name__ == "__main__":
    descargar_todas_las_estaciones()
