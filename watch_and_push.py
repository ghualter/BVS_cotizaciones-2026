#!/usr/bin/env python3
"""
watch_and_push.py
==================
Monitorea el Excel en OneDrive. Cuando detecta que fue modificado,
actualiza automáticamente el index.html y hace git push a GitHub Pages.

INSTALACIÓN (una sola vez):
    pip install openpyxl pandas watchdog

CONFIGURACIÓN:
    Edita las 3 variables de la sección CONFIG más abajo.

USO:
    python watch_and_push.py

    Déjalo corriendo en segundo plano. Cada vez que guardes el Excel,
    el dashboard se actualiza solo en ~15-30 segundos.

INICIO AUTOMÁTICO CON WINDOWS:
    Crea un acceso directo a este script en:
    C:\\Users\\<TU_USUARIO>\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\Startup
    O usa el Task Scheduler (ver README).
"""

import sys, os, time, json, re, subprocess, logging
from pathlib import Path
from datetime import datetime

try:
    import pandas as pd
    from watchdog.observers import Observer
    from watchdog.events import FileSystemEventHandler
except ImportError:
    sys.exit(
        "❌  Faltan dependencias. Ejecuta:\n"
        "    pip install openpyxl pandas watchdog"
    )

# ══════════════════════════════════════════════
# ▶▶▶  CONFIGURACIÓN — EDITA ESTAS 3 LÍNEAS  ◀◀◀
# ══════════════════════════════════════════════

# Ruta completa a tu Excel en OneDrive (copia la ruta exacta)
EXCEL_PATH = r"C:\Users\TU_USUARIO\OneDrive\WcZ_Registro_Cotizaciones_2026.xlsx"

# Ruta completa a la carpeta del repo local (donde está el index.html y el .git)
REPO_PATH  = r"C:\Users\TU_USUARIO\Documents\BVS_Cotizaciones_2026"

# Nombre del archivo HTML dentro del repo
HTML_FILE  = "index.html"

# ══════════════════════════════════════════════
# FIN DE CONFIGURACIÓN
# ══════════════════════════════════════════════

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(message)s",
    datefmt="%H:%M:%S",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(Path(REPO_PATH) / "watcher.log", encoding="utf-8"),
    ],
)
log = logging.getLogger()

DEBOUNCE_SECONDS = 8   # espera N segundos tras el último cambio antes de procesar
_last_modified   = 0
_pending         = False


# ──────────────────────────────────────────────
# EXTRACCIÓN DE DATOS (igual que refresh_and_push.py)
# ──────────────────────────────────────────────

def safe_float(v):
    try:
        return round(float(v), 2)
    except (TypeError, ValueError):
        return 0


def read_excel(path: str) -> dict:
    xl = pd.ExcelFile(path)

    # T1
    df = pd.read_excel(xl, sheet_name="T1_Cotizaciones", header=5)
    cols = ["Codigo","Cliente","Referencia","Monto","Fecha_Envio",
            "Antigüedad_Dias","Estado","En_Territorio","Cuenta_Territorio",
            "Segmento","Segmento_BVS","Q_Cierre","Mes"]
    df = df[cols].dropna(subset=["Codigo"])
    df = df[~df["Codigo"].astype(str).str.upper().isin(["CODIGO","NAN"])]
    df["Codigo"]          = df["Codigo"].astype(str).str.strip()
    df["fecha_str"]       = pd.to_datetime(df["Fecha_Envio"], errors="coerce").dt.strftime("%Y-%m-%d")
    df["Monto"]           = pd.to_numeric(df["Monto"], errors="coerce").fillna(0)
    df["Antigüedad_Dias"] = pd.to_numeric(df["Antigüedad_Dias"], errors="coerce").fillna(0)

    t1 = []
    for _, r in df.iterrows():
        ref = str(r["Referencia"]) if pd.notna(r["Referencia"]) else ""
        t1.append({
            "Codigo": str(r["Codigo"]),
            "Cliente": str(r["Cliente"]) if pd.notna(r["Cliente"]) else "",
            "Referencia": ref,
            "Monto": round(float(r["Monto"]), 2),
            "fecha_str": str(r["fecha_str"]) if str(r["fecha_str"]) != "NaT" else "",
            "Antigüedad_Dias": int(r["Antigüedad_Dias"]),
            "Estado": str(r["Estado"]) if pd.notna(r["Estado"]) else "",
            "En_Territorio": str(r["En_Territorio"]) if pd.notna(r["En_Territorio"]) else "",
            "Cuenta_Territorio": str(r["Cuenta_Territorio"]) if pd.notna(r["Cuenta_Territorio"]) else "",
            "Segmento": str(r["Segmento"]) if pd.notna(r["Segmento"]) else "",
            "Segmento_BVS": str(r["Segmento_BVS"]) if pd.notna(r["Segmento_BVS"]) else "",
            "Q_Cierre": str(r["Q_Cierre"]) if pd.notna(r["Q_Cierre"]) else "",
            "Mes": str(r["Mes"]) if pd.notna(r["Mes"]) else "",
            "isCisco": "CISCO" in ref.upper(),
        })

    # T5
    df5 = pd.read_excel(xl, sheet_name="T5_Oportunidades_CRM", header=3)
    df5 = df5.iloc[:, 1:]
    df5.columns = ["op_id","tema","cuenta","tecnologia","fabricante","fase",
                   "fcst_status","ingresos","ganancia_abs","pct_ganancia","pct_exito",
                   "fecha_cierre","n_cotiz","monto_cotizado","logrado","por_conf","perdido","semaforo"]
    df5 = df5[df5["op_id"].astype(str).str.startswith("OP-")]

    t5 = []
    for _, r in df5.iterrows():
        tema = str(r["tema"]) if pd.notna(r["tema"]) else ""
        fab  = str(r["fabricante"]) if pd.notna(r["fabricante"]) else ""
        fecha = ""
        try:
            fecha = pd.to_datetime(r["fecha_cierre"]).strftime("%Y-%m-%d")
        except Exception:
            pass
        t5.append({
            "op_id": str(r["op_id"]),
            "cuenta": str(r["cuenta"]) if pd.notna(r["cuenta"]) else "",
            "tema": tema, "tecnologia": str(r["tecnologia"]) if pd.notna(r["tecnologia"]) else "",
            "fabricante": fab, "fase": str(r["fase"]) if pd.notna(r["fase"]) else "",
            "fcst_status": str(r["fcst_status"]) if pd.notna(r["fcst_status"]) else "",
            "ingresos": safe_float(r["ingresos"]),
            "ganancia_abs": safe_float(r["ganancia_abs"]),
            "pct_ganancia": round(safe_float(r["pct_ganancia"]) * 100, 1),
            "pct_exito": safe_float(r["pct_exito"]),
            "fecha_cierre_str": fecha,
            "semaforo": str(r["semaforo"]) if pd.notna(r["semaforo"]) else "",
            "isCisco": "CISCO" in tema.upper() or "CISCO" in fab.upper(),
            "logrado": safe_float(r["logrado"]),
            "por_conf": safe_float(r["por_conf"]),
            "perdido": safe_float(r["perdido"]),
        })

    # KPIs
    df_raw = pd.read_excel(xl, sheet_name="T1_Cotizaciones", header=None)
    kpis_src = {
        "cuota":         float(df_raw.iloc[1, 10]),
        "lograda":       float(df_raw.iloc[2, 3]),
        "por_confirmar": float(df_raw.iloc[2, 4]),
        "perdido":       float(df_raw.iloc[2, 5]),
        "avance_ytd":    float(df_raw.iloc[2, 10]),
    }
    df_t4 = pd.read_excel(xl, sheet_name="T4_KPIs", header=0)
    df_t4.columns = ["metrica","valor","categoria","unidad"]
    df_t4 = df_t4[df_t4["metrica"] != "Metrica"].dropna(subset=["metrica"])
    kpis_detail = {}
    for _, r in df_t4.iterrows():
        try:
            kpis_detail[str(r["metrica"])] = round(float(r["valor"]), 4)
        except (TypeError, ValueError):
            kpis_detail[str(r["metrica"])] = str(r["valor"])
    kpis_src["kpis"] = kpis_detail

    return {"T1_RAW": t1, "T5_RAW": t5, "KPIS": kpis_src}


# ──────────────────────────────────────────────
# INYECCIÓN EN HTML
# ──────────────────────────────────────────────

def update_html(data: dict, html_path: str):
    html = Path(html_path).read_text(encoding="utf-8")

    html = re.sub(r"const T1_RAW\s*=\s*\[.*?\];",
                  f"const T1_RAW = {json.dumps(data['T1_RAW'], ensure_ascii=False)};",
                  html, flags=re.DOTALL)
    html = re.sub(r"const T5_RAW\s*=\s*\[.*?\];",
                  f"const T5_RAW = {json.dumps(data['T5_RAW'], ensure_ascii=False)};",
                  html, flags=re.DOTALL)
    html = re.sub(r"const KPIS\s*=\s*\{.*?\};",
                  f"const KPIS   = {json.dumps(data['KPIS'], ensure_ascii=False)};",
                  html, flags=re.DOTALL)

    Path(html_path).write_text(html, encoding="utf-8")


# ──────────────────────────────────────────────
# GIT PUSH
# ──────────────────────────────────────────────

def git_push(repo: str, html_file: str):
    def run(cmd):
        r = subprocess.run(cmd, capture_output=True, text=True, cwd=repo)
        return r.returncode == 0, r.stderr.strip()

    ts = datetime.now().strftime("%Y-%m-%d %H:%M")
    run(["git", "add", html_file])

    status = subprocess.run(["git", "status", "--porcelain"],
                            capture_output=True, text=True, cwd=repo)
    if not status.stdout.strip():
        log.info("  ℹ️  Sin cambios en git — push omitido.")
        return

    ok, err = run(["git", "commit", "-m", f"data: auto-update {ts}"])
    if not ok:
        log.warning(f"  ⚠️  Commit falló: {err}")
        return

    ok, err = run(["git", "push", "origin", "main"])
    if not ok:
        ok, err = run(["git", "push", "origin", "master"])
    if ok:
        log.info("  🚀 Push exitoso → GitHub Pages actualizado")
    else:
        log.error(f"  ❌ Push falló: {err}")


# ──────────────────────────────────────────────
# PIPELINE COMPLETO
# ──────────────────────────────────────────────

def run_pipeline():
    html_path = str(Path(REPO_PATH) / HTML_FILE)
    log.info(f"📂 Leyendo Excel...")
    try:
        data = read_excel(EXCEL_PATH)
        log.info(f"   ✓ T1={len(data['T1_RAW'])} filas  T5={len(data['T5_RAW'])} filas")
    except Exception as e:
        log.error(f"   ❌ Error leyendo Excel: {e}")
        return

    log.info(f"✏️  Actualizando {HTML_FILE}...")
    try:
        update_html(data, html_path)
        log.info(f"   ✓ HTML actualizado")
    except Exception as e:
        log.error(f"   ❌ Error actualizando HTML: {e}")
        return

    log.info("🔀 Haciendo git push...")
    try:
        git_push(REPO_PATH, HTML_FILE)
    except Exception as e:
        log.error(f"   ❌ Error en git: {e}")

    log.info("✅ Listo.\n")


# ──────────────────────────────────────────────
# WATCHER
# ──────────────────────────────────────────────

class ExcelHandler(FileSystemEventHandler):
    def on_modified(self, event):
        global _last_modified, _pending
        if event.src_path and Path(event.src_path).name == Path(EXCEL_PATH).name:
            _last_modified = time.time()
            _pending = True

    on_created = on_modified  # por si OneDrive recrea el archivo al sincronizar


def main():
    global _pending, _last_modified

    # Validaciones
    if not Path(EXCEL_PATH).exists():
        sys.exit(
            f"❌  No se encontró el Excel en:\n   {EXCEL_PATH}\n\n"
            "   Edita la variable EXCEL_PATH en este script."
        )
    if not Path(REPO_PATH).exists():
        sys.exit(
            f"❌  No se encontró el repo en:\n   {REPO_PATH}\n\n"
            "   Edita la variable REPO_PATH en este script."
        )
    if not (Path(REPO_PATH) / HTML_FILE).exists():
        sys.exit(f"❌  No se encontró {HTML_FILE} en {REPO_PATH}")

    watch_dir = str(Path(EXCEL_PATH).parent)

    observer = Observer()
    observer.schedule(ExcelHandler(), path=watch_dir, recursive=False)
    observer.start()

    log.info("=" * 55)
    log.info("👁️  Watcher iniciado")
    log.info(f"   Excel : {EXCEL_PATH}")
    log.info(f"   Repo  : {REPO_PATH}")
    log.info(f"   HTML  : {HTML_FILE}")
    log.info("   Esperando cambios en el Excel... (Ctrl+C para detener)")
    log.info("=" * 55)

    try:
        while True:
            time.sleep(1)
            if _pending and (time.time() - _last_modified) >= DEBOUNCE_SECONDS:
                _pending = False
                log.info(f"\n🔔 Cambio detectado en el Excel → procesando...")
                run_pipeline()
    except KeyboardInterrupt:
        log.info("\n⏹  Watcher detenido.")
    finally:
        observer.stop()
        observer.join()


if __name__ == "__main__":
    main()
