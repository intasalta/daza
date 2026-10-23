import os
import time
import subprocess
import xml.etree.ElementTree as ET
from pathlib import Path
import pandas as pd
from datetime import datetime, timedelta
from playwright.sync_api import sync_playwright

ESTACIONES = ["364053"]
EXPORT_TYPES = {
    "daily": "Daily",
    "weekly": "Weekly"
}

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

        csv_filename = f"{est_id}_{export_type}_station_data.csv"
        xlsx_filename = f"{est_id}_{export_type}_station_data.xlsx"
        df.to_csv(csv_filename, index=False, encoding='utf-8-sig')
        df.to_excel(xlsx_filename, index=False)
        return [csv_filename, xlsx_filename]
    except Exception as conv_err:
        print(f"Error al convertir {xls_path}: {conv_err}")
        return []

def calcular_estadisticas(csv_path, est_id, export_type):
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
        stats_path = f"stats_{est_id}_{export_type}.json"
        with open(stats_path, "w", encoding="utf-8") as f:
            json.dump(stats, f, ensure_ascii=False, indent=2)
        return stats_path
    except Exception as e:
        print(f"Error al calcular estadísticas: {e}")
        return

def run():
    print("Iniciando descarga utilizando sesión guardada (auth.json)...")

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=[
                "--disable-blink-features=AutomationControlled",
                "--no-sandbox",
                "--disable-setuid-sandbox"
            ]
        )
        
        # Verificar si existe el archivo de autorización generado previamente
        storage_state = "auth.json" if os.path.exists("auth.json") else None
        if not storage_state:
            raise FileNotFoundError("No se encontró el archivo 'auth.json'. Generalo localmente primero.")

        context = browser.new_context(
            viewport={"width": 1920, "height": 1080},
            accept_downloads=True,
            storage_state=storage_state,  # <--- Inyecta las cookies de sesión automáticamente
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
        )
        page = context.new_page()

        archivos_generados = []
        for est_id in ESTACIONES:
            station_url = f"https://www.ecowitt.net/home/index?id={est_id}"
            print(f"\n--- Procesando Estación {est_id} ---")
            print(f"Navegando directamente a: {station_url}")
            
            try:
                page.goto(station_url, wait_until="networkidle", timeout=60000)
                time.sleep(5)
            except Exception as nav_err:
                print(f"Error al navegar a la estación (posible sesión expirada): {nav_err}")
                page.screenshot(path=f"error_nav_{est_id}.png")
                continue

            for export_key, view_name in EXPORT_TYPES.items():
                try:
                    print(f"Cambiando a la vista '{view_name}'...")
                    view_btn = page.locator("div, button", has_text=view_name)
                    for candidate in view_btn.all():
                        if candidate.inner_text().strip() == view_name:
                            candidate.click()
                            break
                    time.sleep(3)

                    print(f"Buscando botón 'Export' para {export_key}...")
                    btn_export = page.wait_for_selector("button:has-text('Export')", timeout=20000)
                    
                    with page.expect_download(timeout=30000) as download_info:
                        btn_export.click()
                    
                    download = download_info.value
                    xls_filename = f"{est_id}_{export_key}_station_data.xls"
                    download.save_as(xls_filename)
                    archivos_generados.append(xls_filename)
                    print(f"  -> Archivo descargado: {xls_filename}")
                    
                    formatos_extra = convertir_xls_a_formatos(xls_filename, est_id, export_key)
                    archivos_generados.extend(formatos_extra)
                    
                    csv_path = f"{est_id}_{export_key}_station_data.csv"
                    stats_path = calcular_estadisticas(csv_path, est_id, export_key)
                    if stats_path:
                        archivos_generados.append(stats_path)
                except Exception as e:
                    print(f"Error al procesar export '{export_key}' para estación {est_id}: {e}")
                    page.screenshot(path=f"error_{est_id}_{export_key}.png")
                    continue

        browser.close()

        if archivos_generados and os.path.exists('.git'):
            print("\nActualizando repositorio Git...")
            try:
                subprocess.run(["git", "config", "--global", "user.name", "GitHub Actions Bot"], check=False)
                subprocess.run(["git", "config", "--global", "user.email", "actions@github.com"], check=False)
                for f in archivos_generados:
                    subprocess.run(["git", "add", f], check=False)
                subprocess.run(["git", "commit", "-m", "Actualización automática de datos climáticos [skip ci]"], check=False)
                subprocess.run(["git", "push"], check=False)
                print("Archivos sincronizados en Git con éxito.")
            except Exception as git_err:
                print(f"Aviso Git: {git_err}")

        print(f"\n¡Proceso finalizado! Total archivos generados: {len(archivos_generados)}")

if __name__ == "__main__":
    run()
