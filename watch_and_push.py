# -*- coding: utf-8 -*-
#!/usr/bin/env python3
"""
watch_and_push.py — v2.0
==========================
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
        "[ERROR]  Faltan dependencias. Ejecuta:\n"
        "    pip install openpyxl pandas watchdog"
    )

# ══════════════════════════════════════════════
# ▶▶▶  CONFIGURACIÓN — EDITA ESTAS 3 LÍNEAS  ◀◀◀
# ══════════════════════════════════════════════

# Ruta completa a tu Excel en OneDrive (copia la ruta exacta)
EXCEL_PATH = r"D:\OneDrive - BVS TELEVISION SRL\@COMERCIAL\2026\WcZ_Oportunidades CRM\WcZ_Registro Cotizaciones 2026.xlsx"

# Ruta completa a la carpeta del repo local (donde está el index.html y el .git)
REPO_PATH  = r"D:\OneDrive - BVS TELEVISION SRL\@COMERCIAL\2026\WcZ_Oportunidades CRM\BVS_Cotizaciones_2026"

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

DEBOUNCE_SECONDS = 8
_last_modified   = 0
_pending         = False


# ──────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────

def safe_float(v):
    try:
        return round(float(v), 2)
    except (TypeError, ValueError):
        return 0


def safe_str(v):
    if v is None or (isinstance(v, float) and v != v):
        return ''
    return str(v).strip()


# ──────────────────────────────────────────────
# EXTRACCIÓN DE DATOS
# ──────────────────────────────────────────────

def read_excel(path: str) -> dict:
    xl = pd.ExcelFile(path)

    # ── T1_Cotizaciones ──────────────────────────────────────────
    df = pd.read_excel(xl, sheet_name="T1_Cotizaciones", header=5)
    df = df[['Codigo','Cliente','Referencia','Tipo_Venta','Monto','Fecha_Envio',
             'Antigüedad_Dias','Estado','En_Territorio','Cuenta_Territorio',
             'Segmento','Segmento_BVS','Q_Cierre','Mes']].copy()
    df = df[df['Codigo'].notna() & ~df['Codigo'].astype(str).str.upper().isin(['CODIGO','NAN'])]
    df['Monto']           = pd.to_numeric(df['Monto'], errors='coerce').fillna(0)
    df['Antigüedad_Dias'] = pd.to_numeric(df['Antigüedad_Dias'], errors='coerce').fillna(0)
    df['fecha_str']       = pd.to_datetime(df['Fecha_Envio'], errors='coerce').dt.strftime('%Y-%m-%d')

    t1 = []
    for _, r in df.iterrows():
        ref  = safe_str(r['Referencia'])
        tipo = safe_str(r['Tipo_Venta'])
        t1.append({
            'Codigo':            safe_str(r['Codigo']),
            'Cliente':           safe_str(r['Cliente']),
            'Referencia':        ref,
            'Tipo_Venta':        tipo,
            'Monto':             round(float(r['Monto']), 2),
            'fecha_str':         safe_str(r['fecha_str']).replace('NaT',''),
            'Antigüedad_Dias':   int(r['Antigüedad_Dias']),
            'Estado':            safe_str(r['Estado']),
            'En_Territorio':     safe_str(r['En_Territorio']),
            'Cuenta_Territorio': safe_str(r['Cuenta_Territorio']),
            'Segmento':          safe_str(r['Segmento']),
            'Segmento_BVS':      safe_str(r['Segmento_BVS']),
            'Q_Cierre':          safe_str(r['Q_Cierre']),
            'Mes':               safe_str(r['Mes']),
            'isCisco':           'CISCO' in ref.upper(),
            'isRenovacion':      tipo == 'Renovación',
        })
    log.info(f"   T1: {len(t1)} cotizaciones")

    # ── T5_Oportunidades_CRM ─────────────────────────────────────
    df5 = pd.read_excel(xl, sheet_name="T5_Oportunidades_CRM", header=3)
    df5 = df5[df5['OP ID'].astype(str).str.contains('OP-', na=False)].copy()

    t5 = []
    for _, r in df5.iterrows():
        tema  = safe_str(r['Tema / Oportunidad'])
        fab   = safe_str(r['Fabricante'])
        fecha = ''
        try:
            fecha = pd.to_datetime(r['Fecha Est. Cierre']).strftime('%Y-%m-%d')
        except Exception:
            pass
        # Semaforo column name may have newline
        sem = ''
        for col in df5.columns:
            if 'Sem' in col and 'foro' in col.replace('á','a').replace('á','a'):
                sem = safe_str(r[col])
                break

        t5.append({
            'op_id':           safe_str(r['OP ID']),
            'cuenta':          safe_str(r['Cuenta']),
            'tema':            tema,
            'tecnologia':      safe_str(r.get('Tecnología', '')),
            'fabricante':      fab,
            'fase':            safe_str(r['Fase Pipeline']),
            'fcst_status':     safe_str(r['FCST Status']),
            'ingresos':        safe_float(r['Ingresos Potenc.']),
            'ganancia_abs':    safe_float(r['Ganancia Abs.']),
            'pct_ganancia':    round(safe_float(r.get('% Ganancia', 0)) * 100, 1),
            'pct_exito':       safe_float(r.get('% Éxito CRM', 0)),
            'fecha_cierre_str': fecha,
            'semaforo':        sem,
            'isCisco':         'CISCO' in tema.upper() or 'CISCO' in fab.upper(),
            'logrado':         safe_float(r.get('Logrado\nen T1', 0)),
            'por_conf':        safe_float(r.get('Por Conf.\nen T1', 0)),
            'perdido':         safe_float(r.get('Perdido\nen T1', 0)),
        })
    log.info(f"   T5: {len(t5)} oportunidades CRM")

    # ── Históricos 2024 / 2025 / 2026 ───────────────────────────
    hist = {}
    for yr, sheet in [('2024','Logradas 2024'), ('2025','Logradas 2025'), ('2026','Logradas 2026')]:
        try:
            dfh = pd.read_excel(xl, sheet_name=sheet, header=0)
            rows = []
            for _, r in dfh.iterrows():
                cuenta = safe_str(r.get('Cuenta', ''))
                tema   = safe_str(r.get('Tema', ''))
                fab    = safe_str(r.get('Marca / Fabricante', ''))
                ingr   = safe_float(r.get('Ingresos reales', 0))
                if cuenta and cuenta not in ('Cuenta', 'nan') and ingr > 0:
                    rows.append({'cuenta': cuenta, 'tema': tema,
                                 'fabricante': fab, 'ingresos': ingr})
            hist[yr] = rows
            total = sum(r['ingresos'] for r in rows)
            log.info(f"   Hist {yr}: {len(rows)} ops = ${total:,.0f}")
        except Exception as e:
            log.warning(f"   Hist {yr}: no disponible ({e})")
            hist[yr] = []

    # ── KPIs desde T1 raw ────────────────────────────────────────
    df_raw = pd.read_excel(xl, sheet_name="T1_Cotizaciones", header=None)
    kpis = {
        'cuota':         3000000.0,
        'lograda':       safe_float(df_raw.iloc[2, 4]),
        'por_confirmar': safe_float(df_raw.iloc[2, 5]),
        'perdido':       safe_float(df_raw.iloc[2, 6]),
        'avance_ytd':    safe_float(df_raw.iloc[2, 11]),
    }
    log.info(f"   KPIs: logrado=${kpis['lograda']:,.0f} / cuota=${kpis['cuota']:,.0f}")

    return {'T1_RAW': t1, 'T5_RAW': t5, 'KPIS': kpis, 'HIST': hist}


# ──────────────────────────────────────────────
# INYECCIÓN EN HTML
# ──────────────────────────────────────────────

def update_html(data: dict, html_path: str):
    html = Path(html_path).read_text(encoding='utf-8')

    replacements = {
        r'const T1_RAW\s*=\s*\[.*?\];':
            f"const T1_RAW = {json.dumps(data['T1_RAW'], ensure_ascii=False)};",
        r'const T5_RAW\s*=\s*\[.*?\];':
            f"const T5_RAW = {json.dumps(data['T5_RAW'], ensure_ascii=False)};",
        r'const KPIS\s*=\s*\{{.*?\}};':
            f"const KPIS   = {json.dumps(data['KPIS'], ensure_ascii=False)};",
        r'const HIST\s*=\s*\{{.*?\}};':
            f"const HIST   = {json.dumps(data['HIST'], ensure_ascii=False)};",
    }

    for pattern, replacement in replacements.items():
        html = re.sub(pattern, replacement, html, flags=re.DOTALL)

    Path(html_path).write_text(html, encoding='utf-8')
    log.info(f"   HTML actualizado ({len(html):,} chars)")


# ──────────────────────────────────────────────
# GIT PUSH
# ──────────────────────────────────────────────

def git_push(repo: str, html_file: str):
    def run(cmd):
        r = subprocess.run(cmd, capture_output=True, text=True, cwd=repo)
        return r.returncode == 0, r.stderr.strip()

    ts = datetime.now().strftime('%Y-%m-%d %H:%M')

    run(['git', 'add', html_file])

    status = subprocess.run(['git', 'status', '--porcelain'],
                            capture_output=True, text=True, cwd=repo)
    if not status.stdout.strip():
        log.info("   Sin cambios en git — push omitido.")
        return

    ok, err = run(['git', 'commit', '-m', f'data: auto-update {ts}'])
    if not ok:
        log.warning(f"   Commit falló: {err}")
        return

    ok, err = run(['git', 'push', 'origin', 'main'])
    if not ok:
        ok, err = run(['git', 'push', 'origin', 'master'])

    if ok:
        log.info("   ✅ Push exitoso → GitHub Pages actualizado")
    else:
        log.error(f"   ❌ Push falló: {err}")


# ──────────────────────────────────────────────
# PIPELINE COMPLETO
# ──────────────────────────────────────────────

def run_pipeline():
    html_path = str(Path(REPO_PATH) / HTML_FILE)

    log.info("[EXCEL] Leyendo Excel...")
    try:
        data = read_excel(EXCEL_PATH)
    except Exception as e:
        log.error(f"   Error leyendo Excel: {e}")
        return

    log.info("[HTML]  Actualizando HTML...")
    try:
        update_html(data, html_path)
    except Exception as e:
        log.error(f"   Error actualizando HTML: {e}")
        return

    log.info("[GIT]   Haciendo push...")
    try:
        git_push(REPO_PATH, HTML_FILE)
    except Exception as e:
        log.error(f"   Error en git: {e}")

    log.info("[OK]    Listo.\n")


# ──────────────────────────────────────────────
# WATCHER
# ──────────────────────────────────────────────

class ExcelHandler(FileSystemEventHandler):
    def on_modified(self, event):
        global _last_modified, _pending
        if event.src_path and Path(event.src_path).name == Path(EXCEL_PATH).name:
            _last_modified = time.time()
            _pending = True

    on_created = on_modified


def main():
    global _pending, _last_modified

    if not Path(EXCEL_PATH).exists():
        sys.exit(
            f"[ERROR]  No se encontró el Excel en:\n   {EXCEL_PATH}\n\n"
            "   Edita la variable EXCEL_PATH en este script."
        )
    if not Path(REPO_PATH).exists():
        sys.exit(
            f"[ERROR]  No se encontró el repo en:\n   {REPO_PATH}\n\n"
            "   Edita la variable REPO_PATH en este script."
        )
    if not (Path(REPO_PATH) / HTML_FILE).exists():
        sys.exit(f"[ERROR]  No se encontró {HTML_FILE} en {REPO_PATH}")

    watch_dir = str(Path(EXCEL_PATH).parent)
    observer  = Observer()
    observer.schedule(ExcelHandler(), path=watch_dir, recursive=False)
    observer.start()

    log.info("=" * 55)
    log.info("[WATCH]  Watcher v2.0 iniciado")
    log.info(f"   Excel : {EXCEL_PATH}")
    log.info(f"   Repo  : {REPO_PATH}")
    log.info(f"   HTML  : {HTML_FILE}")
    log.info("   Secciones: T1 · T5 · Renovaciones · Histórico · KPIs")
    log.info("   Esperando cambios... (Ctrl+C para detener)")
    log.info("=" * 55)

    try:
        while True:
            time.sleep(1)
            if _pending and (time.time() - _last_modified) >= DEBOUNCE_SECONDS:
                _pending = False
                log.info("\n[CAMBIO] Cambio detectado en el Excel...")

                # Espera a que Excel libere el archivo
                max_wait = 300
                waited   = 0
                while waited < max_wait:
                    try:
                        with open(EXCEL_PATH, 'rb'):
                            break
                    except (PermissionError, IOError):
                        log.info(f"   Excel abierto, esperando... ({waited}s)")
                        time.sleep(5)
                        waited += 5

                if waited >= max_wait:
                    log.warning("   Tiempo de espera agotado.")
                else:
                    log.info("   Archivo libre → procesando...")
                    run_pipeline()

    except KeyboardInterrupt:
        log.info("\n[STOP]  Watcher detenido.")
    finally:
        observer.stop()
        observer.join()


if __name__ == '__main__':
    main()
