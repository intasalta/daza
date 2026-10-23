import os
import time
import subprocess
import xml.etree.ElementTree as ET
import pandas as pd
from playwright.sync_api import sync_playwright

ESTACIONES = ["364053"]

def convertir_xls_a_formatos(xls_path, est_id, export_type):
    try:
        tree = ET.parse(xls_path)
        root = tree.getroot()
        ns = {'ss': 'urn:schemas-microsoft-com:office:spreadsheet'}

        rows = []
        for row in root.findall('.//ss:Row', ns):
            row_cells = []
            for cell in row.findall('ss:Cell', ns):
                merge = int(cell.attrib.get('{urn:schemas-microsoft-com:office:spreadsheet}MergeAcross', 0))
                data = cell.find('ss:Data', ns)
                val = data.text if data is not None else ''
                row_cells.append(val)
                for _ in range(merge):
                    row_cells.append(val)
            rows.append(row_cells)

        if len(rows) < 3:
            return []

        h1 = rows[0]
        h2 = rows[1]
        headers = []
        for i in range(len(h2)):
            c1 = h1[i] if i < len(h1) and h1[i] else ''
            c2 = h2[i] if i < len(h2) and h2[i] else ''
            if c1 and c2 and c1 != c2:
                headers.append(f"{c1} - {c2}")
            elif c2:
                headers.append(c2)
            elif c1:
                headers.append(c1)
            else:
                headers.append(f"Col_{i}")

        df = pd.DataFrame(rows[2:], columns=headers)
        for col in df.columns:
            if col not in ['Fecha / Hora', 'Time']:
                try:
                    df[col] = pd.to_numeric(df[col])
                except Exception:
                    pass

        csv_filename = f"{est_id}_daily_station_data.csv"
        xlsx_filename = f"{est_id}_daily_station_data.xlsx"
        df.to_csv(csv_filename, index=False, encoding='utf-8-sig')
        df.to_excel(xlsx_filename, index=False)
        return [csv_filename, xlsx_filename]
    except Exception as conv_err:
        print(f"Error al convertir {xls_path}: {conv_err}")
        return []

def calcular_estadisticas(csv_path, est_id):
    try:
        df = pd.read_csv(csv_path)
        date_col = None
        for possible in ['Fecha / Hora', 'Date/Time', 'Timestamp', 'Time']:
            if possible in df.columns:
                date_col = possible
                break
        if not date_col:
            return
        df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
        now = pd.Timestamp.now()
        last_24h = df[(df[date_col] >= now - pd.Timedelta(hours=24))]
        if last_24h.empty:
            return
        stats = {}
        for metric in ['Temperature(°C)', 'Temperatura', 'Temp', 'Humidity(%)', 'Humedad', 'Pressure(hPa)', 'Presión']:
            if metric in last_24h.columns:
                stats[metric] = {
                    'max': float(last_24h[metric].max()),
                    'min': float(last_24h[metric].min())
                }
        import json
        stats_path = f"stats_{est_id}_daily.json"
        with open(stats_path, "w", encoding="utf-8") as f:
            json.dump(stats, f, ensure_ascii=False, indent=2)
        return stats_path
    except Exception as e:
        print(f"Error al calcular estadísticas: {e}")
        return

def run():
    print("Descargando reporte DIARIO...")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, args=["--no-sandbox", "--disable-setuid-sandbox"])
        storage_state = "auth.json" if os.path.exists("auth.json") else None
        if not storage_state:
            raise FileNotFoundError("No se encontró 'auth.json'.")

        context = browser.new_context(accept_downloads=True, storage_state=storage_state)
        page = context.new_page()

        archivos_generados = []
        for est_id in ESTACIONES:
            station_url = f"https://www.ecowitt.net/home/index?id={est_id}"
            page.goto(station_url, wait_until="networkidle", timeout=60000)
            time.sleep(5)

            # Por defecto suele estar en Daily, pero aseguramos clic si existe el botón
            try:
                page.get_by_text("Daily", exact=True).first.click()
                time.sleep(3)
            except Exception:
                pass

            btn_export = page.wait_for_selector("button:has-text('Export')", timeout=20000)
            with page.expect_download(timeout=30000) as download_info:
                btn_export.click()
            
            download = download_info.value
            xls_filename = f"{est_id}_daily_station_data.xls"
            download.save_as(xls_filename)
            archivos_generados.append(xls_filename)
            
            archivos_generados.extend(convertir_xls_a_formatos(xls_filename, est_id, "daily"))
            stats_path = calcular_estadisticas(f"{est_id}_daily_station_data.csv", est_id)
            if stats_path:
                archivos_generados.append(stats_path)

        browser.close()
        print(f"¡Diario finalizado! Archivos: {len(archivos_generados)}")

if __name__ == "__main__":
    run()
