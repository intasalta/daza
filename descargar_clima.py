import os
import time
import subprocess
import xml.etree.ElementTree as ET
from pathlib import Path
import pandas as pd
from datetime import datetime, timedelta
from playwright.sync_api import sync_playwright

# Load env variables from .env if exists (for local testing)
env_file = Path(__file__).resolve().parent / ".env"
if env_file.exists():
    with open(env_file, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip())

USER = os.environ.get("FC_USER")
PASS = os.environ.get("FC_PASS")

# Station ID for ECOWITT
ESTACIONES = ["364053"]

# Export types corresponding to the interface views
EXPORT_TYPES = {
    "daily": "Daily",
    "weekly": "Weekly"
}

def convertir_xls_a_formatos(xls_path, est_id, export_type):
    """Convierte el archivo SpreadsheetML XML (.xls) de Ecowitt a CSV y XLSX."""
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
            print(f"Aviso: El archivo {xls_path} no tiene suficientes filas para convertir.")
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
            if col != 'Fecha / Hora' and col != 'Time':
                try:
                    df[col] = pd.to_numeric(df[col])
                except Exception:
                    pass

        csv_filename = f"{est_id}_{export_type}_station_data.csv"
        xlsx_filename = f"{est_id}_{export_type}_station_data.xlsx"
        df.to_csv(csv_filename, index=False, encoding='utf-8-sig')
        df.to_excel(xlsx_filename, index=False)
        print(f"  -> Generado CSV: {csv_filename} ({os.path.getsize(csv_filename):,} bytes)")
        print(f"  -> Generado XLSX: {xlsx_filename} ({os.path.getsize(xlsx_filename):,} bytes)")
        return [csv_filename, xlsx_filename]
    except Exception as conv_err:
        print(f"Error al convertir {xls_path}: {conv_err}")
        return []

def calcular_estadisticas(csv_path, est_id, export_type):
    """Calcula max/min de temperatura, humedad y presión en las últimas 24h."""
    try:
        df = pd.read_csv(csv_path)
        date_col = None
        for possible in ['Fecha / Hora', 'Date/Time', 'Timestamp', 'Time']:
            if possible in df.columns:
                date_col = possible
                break
        if not date_col:
            print(f"No se encontró columna de fecha en {csv_path}")
            return
        df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
        now = pd.Timestamp.now()
        last_24h = df[(df[date_col] >= now - pd.Timedelta(hours=24))]
        if last_24h.empty:
            print(f"No hay datos de las últimas 24h en {csv_path}")
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
        print(f"Estadísticas guardadas en {stats_path}")
        return stats_path
    except Exception as e:
        print(f"Error al calcular estadísticas: {e}")
        return

def run():
    if not USER or not PASS:
        raise ValueError("Error: Las variables de entorno FC_USER y FC_PASS deben estar definidas.")
    print(f"Iniciando descarga para el usuario: {USER}")

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=["--disable-blink-features=AutomationControlled"]
        )
        context = browser.new_context(
            viewport={"width": 1920, "height": 1080},
            accept_downloads=True
        )
        page = context.new_page()

        print("1. Accediendo a la página de login de ECOWITT...")
        page.goto("https://www.ecowitt.net/auth/login", wait_until="networkidle", timeout=60000)
        
        # Selectores genéricos alternativos por si el input de email cambia de atributo
        page.wait_for_selector("input[type='text'], input[type='email']", timeout=60000)
        
        # Llenar usuario y contraseña usando selectores robustos
        page.locator("input[type='text'], input[type='email']").first.fill(USER)
        page.locator("input[type='password']").first.fill(PASS)
        
        # Click en botón de inicio de sesión
        page.locator("button:has-text('Sign In'), button[type='submit']").first.click()
        page.wait_for_url(lambda u: "auth/login" not in u, timeout=45000)
        print(f"Sesión iniciada correctamente. URL actual: {page.url}")
        time.sleep(3)

        archivos_generados = []
        for est_id in ESTACIONES:
            station_url = f"https://www.ecowitt.net/home/index?id={est_id}"
            print(f"\n--- Procesando Estación {est_id} ---")
            print(f"Navegando a: {station_url}")
            page.goto(station_url, wait_until="networkidle", timeout=60000)
            time.sleep(5)  # Esperar a que carguen los componentes de la estación

            for export_key, view_name in EXPORT_TYPES.items():
                try:
                    print(f"Cambiando a la vista '{view_name}'...")
                    # Intentar hacer clic en el botón de vista (Daily / Weekly) si existe
                    view_btn = page.locator(f"div, button", has_text=view_name)
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
                    print(f"Archivo XLS descargado: {xls_filename} ({os.path.getsize(xls_filename):,} bytes)")
                    
                    formatos_extra = convertir_xls_a_formatos(xls_filename, est_id, export_key)
                    archivos_generados.extend(formatos_extra)
                    
                    csv_path = f"{est_id}_{export_key}_station_data.csv"
                    stats_path = calcular_estadisticas(csv_path, est_id, export_key)
                    if stats_path:
                        archivos_generados.append(stats_path)
                except Exception as e:
                    print(f"Error al procesar export '{export_key}' para estación {est_id}: {e}")
                    # Tomar captura de pantalla para depuración si falla
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
                subprocess.run(["git", "commit", "-m", "Actualización automática de datos climáticos (.xls, .csv, .xlsx, stats) [skip ci]"], check=False)
                subprocess.run(["git", "push"], check=False)
                print("Archivos sincronizados en Git con éxito.")
            except Exception as git_err:
                print(f"Aviso Git: {git_err}")

        print(f"\n¡Proceso finalizado con éxito! Total archivos generados: {len(archivos_generados)}")

if __name__ == "__main__":
    run()
