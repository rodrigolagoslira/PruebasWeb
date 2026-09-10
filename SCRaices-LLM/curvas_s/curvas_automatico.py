"""
curvas_automatico.py  -  Automatizacion semanal Curvas S - Proyecto Nuke Mapu
========================================================================
- Fecha de control:  date.today() (automatica, no requiere entrada manual)
- Datos de avance:   leidos en tiempo real desde AppSheet (proyecto P119)
  Fallback:          si AppSheet no esta disponible, lee desde hoja
                     'Datos Control' del Gantt (entrada manual)
- Genera 7 graficos PNG, actualiza Drive, marca PENDIENTE para Apps Script

Ejecucion manual:  python curvas_automatico.py
Ejecucion auto:    Tarea programada en Windows Task Scheduler (ver README)

Primera ejecucion:
  Si AppSheet aun no tiene auth guardada, ejecutar primero:
    python leer_appsheet.py --setup
"""

import os
import sys
import logging
import traceback
from pathlib import Path
from datetime import date, timedelta, datetime

# AppSheet reader (Playwright-based)
try:
    from leer_appsheet import leer_avance_p119
    APPSHEET_DISPONIBLE = True
except ImportError:
    APPSHEET_DISPONIBLE = False

import numpy as np
import matplotlib
matplotlib.use("Agg")          # sin ventana
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.lines import Line2D
from scipy.special import expit
from PIL import Image
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
import curvas_cloud_utils as _ccu
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURACION
# ─────────────────────────────────────────────────────────────────────────────
SPREADSHEET_ID = "1t_1j62f_3l1nrlufmvhnV-o1WTplv0OnQL_JdVaWgKA"
GANTT_HOJA       = "Ñuke Mapu"
OBRA_FOLDER    = "Nuke Mapu"
TOKEN_FILE     = _ccu.TOKEN_FILE
OUTPUT_DIR     = _ccu.get_output_dir_obra(OBRA_FOLDER)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

DPI_EXPORT = 200           # resolucion de exportacion (10"x5.1" -> 2000x1020 px, optimo para celda 755x385)

# IDs estables en Google Drive (NO cambiar - se actualiza el contenido, no el ID)
DRIVE_IDS = {
    "CurvaS_GRUPO_1.png":         "1k2gpSr9Sk5zUZ3MAOxw3zGcHRa9ljFSA",
    "CurvaS_GRUPO_2.png":         "1Ciu52kLYT9NjgaqUBPif4mtWk6l0RtPm",
    "CurvaS_GRUPO_3.png":         "1mDdpyG5zGmSgOXZVMN4qzGbp_0BMrzT2",
    "CurvaS_GRUPO_4.png":         "1GHthHYK1bpoA7_2Q5qSm3VFJq4KPYQa_",
    "CurvaS_GRUPO_5.png":         "12yQw_tt3W4H91Q4KpiMafm_eLUNkwqpP",
    "CurvaS_TOTAL_Nuke_Mapu.png": "11L-TIagyTGZOh3yhbGvjdIruG0dJQGPQ",
    "CurvaS_Todos_Grupos.png":    "1_xxaqAay-UeB4POTcXMaBRD_Xo5ceb-1",
}

# Curva S programada semanal (35 semanas)
PCT_SEMANA = [0, 4, 7, 11, 14, 18, 21, 25, 29, 32, 36, 39, 43, 46, 50,
              54, 57, 61, 64, 68, 71, 75, 79, 82, 86, 89, 93, 96, 100, 100]
DURACION_DIAS = 245

COLORES = {
    "GRUPO 1": "#1a6eb5",
    "GRUPO 2": "#e07b00",
    "GRUPO 3": "#2ca02c",
    "GRUPO 4": "#d62728",
    "GRUPO 5": "#9467bd",
}

# Vencimiento del Programa de Obra del proyecto (contrato: 09/01/2026 + 402 días)
FIN_PROGRAMADO_PROYECTO = date(2027, 2, 15)

# ─────────────────────────────────────────────────────────────────────────────
# LOG
# ─────────────────────────────────────────────────────────────────────────────
log = _ccu.setup_logging("nuke_mapu")


# ─────────────────────────────────────────────────────────────────────────────
# 1. CREDENCIALES
# ─────────────────────────────────────────────────────────────────────────────
def get_credentials():
    return _ccu.get_credentials(SCOPES)


# ─────────────────────────────────────────────────────────────────────────────
# 2. LEER DATOS DESDE AppSheet + 'Datos Control'
# ─────────────────────────────────────────────────────────────────────────────
def _normalizar_nombre(nombre):
    """Normaliza nombre para comparacion: mayusculas, sin espacios dobles."""
    import unicodedata
    # Eliminar tildes para comparacion robusta
    nfkd = unicodedata.normalize("NFKD", nombre)
    sin_tildes = "".join(c for c in nfkd if not unicodedata.combining(c))
    return " ".join(sin_tildes.upper().split())


def leer_avance_appsheet():
    """
    Intenta leer % real desde AppSheet.
    Retorna dict {nombre_normalizado: pct_int} o None si falla.
    """
    if not APPSHEET_DISPONIBLE:
        log.warning("Modulo leer_appsheet no disponible.")
        return None
    try:
        log.info("Leyendo datos de avance desde AppSheet...")
        datos = leer_avance_p119()
        # Normalizar nombres para matching robusto
        result = {_normalizar_nombre(k): v for k, v in datos.items()}
        log.info(f"  AppSheet: {len(result)} beneficiarios P119 leidos.")
        return result
    except Exception as e:
        log.warning(f"AppSheet no disponible ({e}). Usando datos de 'Datos Control'.")
        return None


def leer_datos_control(sheets_svc):
    """
    Lee la hoja 'Datos Control' para obtener la estructura de beneficiarios
    (grupos, nombres, fechas de inicio).

    La FECHA DE CONTROL se toma de date.today() automaticamente.
    El % REAL se lee desde AppSheet (si disponible) o desde col D de la hoja.

    Retorna:
      - control_date: date  (= hoy)
      - grupos: dict {nombre_grupo: [(nombre, inicio, pct_real), ...]}
    """
    # ── Fecha de control = HOY (automatica) ──────────────────────────────────
    control_date = date.today()
    log.info(f"Fecha de control (hoy): {control_date.strftime('%d/%m/%Y')}")

    # ── Leer estructura desde 'Datos Control' ────────────────────────────────
    rng = "Datos Control!A1:D100"
    result = sheets_svc.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=rng,
        valueRenderOption="UNFORMATTED_VALUE",
        dateTimeRenderOption="FORMATTED_STRING",
    ).execute()
    rows = result.get("values", [])

    # ── Leer TERMINO individual por beneficiario desde Gantt ─────────────────
    termino_map = _leer_terminos_desde_gantt(sheets_svc)

    # ── Leer % real desde AppSheet (preferido) ────────────────────────────────
    avance_appsheet = leer_avance_appsheet()

    # ── Parsear filas de beneficiarios (fila 5 en adelante, indice 4) ─────────
    grupos = {}
    sin_match_appsheet = []

    for row in rows[4:]:
        if len(row) < 3:
            continue
        grupo_raw = str(row[0]).strip()
        nombre    = str(row[1]).strip()
        inicio_s  = str(row[2]).strip()
        pct_hoja  = row[3] if len(row) > 3 else 0

        if not grupo_raw or not nombre:
            continue

        # Parsear inicio
        try:
            inicio = datetime.strptime(inicio_s, "%d/%m/%Y").date()
        except ValueError:
            try:
                inicio = datetime.strptime(inicio_s, "%m/%d/%Y").date()
            except ValueError:
                log.warning(f"Fecha de inicio no reconocida para {nombre}: {inicio_s!r}, usando hoy")
                inicio = date.today()

        # Determinar % real: AppSheet tiene prioridad
        if avance_appsheet:
            nombre_norm = _normalizar_nombre(nombre)
            pct = avance_appsheet.get(nombre_norm)
            if pct is None:
                tokens = set(nombre_norm.split())
                best_key, best_score = None, 0
                for k, v in avance_appsheet.items():
                    score = len(tokens & set(k.split()))
                    if score > best_score:
                        best_score, best_key = score, k
                if best_key is not None and best_score >= max(2, len(tokens) - 1):
                    pct = avance_appsheet[best_key]
            if pct is None:
                sin_match_appsheet.append(nombre)
                # Fallback: valor de la hoja
                try:
                    pct = float(str(pct_hoja).replace("%", "").strip())
                except (ValueError, TypeError):
                    pct = 0.0
        else:
            # Sin AppSheet: usar valor de la hoja
            try:
                pct = float(str(pct_hoja).replace("%", "").strip())
            except (ValueError, TypeError):
                pct = 0.0

        termino = termino_map.get(_normalizar_nombre(nombre),
                                  inicio + timedelta(days=DURACION_DIAS))
        grupos.setdefault(grupo_raw, []).append((nombre, inicio, float(pct), termino))

    if sin_match_appsheet:
        log.warning(f"Sin match en AppSheet ({len(sin_match_appsheet)}): "
                    + ", ".join(sin_match_appsheet[:5]))

    fuente = "AppSheet" if avance_appsheet else "hoja 'Datos Control'"
    log.info(f"Beneficiarios: {sum(len(v) for v in grupos.values())} "
             f"en {len(grupos)} grupos | avance desde: {fuente}")
    return control_date, grupos


# ─────────────────────────────────────────────────────────────────────────────
# 2b. LEER TERMINOS INDIVIDUALES DESDE GANTT
# ─────────────────────────────────────────────────────────────────────────────
def _leer_terminos_desde_gantt(sheets_svc):
    """Lee fechas TERMINO individuales de cada beneficiario desde el Gantt 'Nuke Mapu'.
    Columna J (idx 9) = TERMINO, columna D (idx 3) = BENEFICIARIO.
    Retorna dict {nombre_normalizado: termino_date}.
    """
    r = sheets_svc.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range="'Ñuke Mapu'!A1:J80",
        valueRenderOption="FORMATTED_VALUE",
    ).execute()
    prog_rows = r.get("values", [])

    termino_map = {}
    for row in prog_rows:
        nombre_raw  = str(row[3]).strip() if len(row) > 3 else ""
        termino_raw = str(row[9]).strip() if len(row) > 9 else ""
        if not nombre_raw or not termino_raw:
            continue
        if (nombre_raw.upper().startswith("GRUPO")
                or nombre_raw.upper() in ("BENEFICIARIO", "CARTA GANTT", "TERMINO", "")):
            continue
        try:
            float(nombre_raw.replace("%", "").replace(",", "."))
            continue
        except ValueError:
            pass
        termino_date = None
        for fmt in ("%d/%m/%Y", "%m/%d/%Y"):
            try:
                termino_date = datetime.strptime(termino_raw, fmt).date()
                break
            except ValueError:
                pass
        if termino_date:
            termino_map[_normalizar_nombre(nombre_raw)] = termino_date

    log.info(f"  Terminos leidos desde Gantt: {len(termino_map)} beneficiarios")
    return termino_map


# ─────────────────────────────────────────────────────────────────────────────
# 2c. SINCRONIZAR GRUPOS DESDE GANTT
# ─────────────────────────────────────────────────────────────────────────────
def _get_dc_gid(sheets_svc):
    meta = sheets_svc.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    for s in meta["sheets"]:
        if s["properties"]["title"] == "Datos Control":
            return s["properties"]["sheetId"]
    return 0


def _reordenar_datos_control(sheets_svc):
    dc = sheets_svc.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range="Datos Control!A1:D50",
        valueRenderOption="UNFORMATTED_VALUE",
        dateTimeRenderOption="FORMATTED_STRING",
    ).execute()
    rows = dc.get("values", [])
    header = rows[:4]
    datos = [r for r in rows[4:] if len(r) >= 2 and r[0].strip() and r[1].strip()]

    orden_grupos = ["GRUPO 1", "GRUPO 2", "GRUPO 3", "GRUPO 4", "GRUPO 5"]

    def sort_key(r):
        g = str(r[0]).strip().upper()
        try:
            return (orden_grupos.index(g), r[1])
        except ValueError:
            return (99, r[1])

    datos_ordenados = sorted(datos, key=sort_key)
    nuevas_filas = header + datos_ordenados

    sheets_svc.spreadsheets().values().clear(
        spreadsheetId=SPREADSHEET_ID, range="Datos Control!A1:D50"
    ).execute()
    sheets_svc.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range="Datos Control!A1",
        valueInputOption="USER_ENTERED",
        body={"values": nuevas_filas},
    ).execute()
    log.info("  Datos Control reordenado por grupo.")


def sincronizar_grupos_desde_gantt(sheets_svc):
    log.info("Verificando movimientos de grupo en 'Ñuke Mapu'...")

    # Ñuke Mapu: GRUPO headers at col B (idx 1), BENEFICIARIO at col D (idx 3), INICIO at col I (idx 8)
    r = sheets_svc.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range="'Ñuke Mapu'!A1:J60",
        valueRenderOption="FORMATTED_VALUE",
    ).execute()
    prog_rows = r.get("values", [])

    grupo_actual = None
    gantt_map = {}
    for row in prog_rows:
        if not row:
            continue
        # Group headers are in col B (idx 1)
        for col_idx in [1, 0, 3]:
            if col_idx < len(row):
                cell = str(row[col_idx]).strip().upper()
                if cell.startswith("GRUPO") or cell in ("REZAGADOS",):
                    grupo_actual = cell if cell.startswith("GRUPO") else "GRUPO " + cell
                    break
        nombre_raw = str(row[3]).strip() if len(row) > 3 else ""
        inicio_raw = str(row[8]).strip() if len(row) > 8 else ""
        if (not nombre_raw or nombre_raw.upper().startswith("GRUPO")
                or nombre_raw.upper() in ("BENEFICIARIO", "CARTA GANTT", "")
                or not grupo_actual):
            continue
        try:
            float(nombre_raw.replace("%", "").replace(",", "."))
            continue
        except ValueError:
            pass
        # Convert US date format (M/D/YYYY) to DD/MM/YYYY if needed
        if inicio_raw:
            from datetime import datetime as _dt
            for fmt in ("%m/%d/%Y", "%d/%m/%Y", "%-m/%-d/%Y"):
                try:
                    inicio_raw = _dt.strptime(inicio_raw, fmt).strftime("%d/%m/%Y")
                    break
                except ValueError:
                    pass
        nombre_norm = _normalizar_nombre(nombre_raw)
        if nombre_norm:
            gantt_map[nombre_norm] = {"grupo": grupo_actual, "inicio": inicio_raw}

    if not gantt_map:
        log.warning("  No se pudo leer estructura de grupos desde 'Ñuke Mapu'.")
        return

    dc = sheets_svc.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range="Datos Control!A1:D50",
        valueRenderOption="UNFORMATTED_VALUE",
        dateTimeRenderOption="FORMATTED_STRING",
    ).execute()
    dc_rows = dc.get("values", [])

    def _parse_fecha(s):
        for fmt in ("%d/%m/%Y", "%m/%d/%Y"):
            try:
                from datetime import datetime as _dt
                return _dt.strptime(str(s).strip(), fmt).date()
            except (ValueError, TypeError):
                pass
        return None

    cambios = []
    for i, row in enumerate(dc_rows[4:], start=5):
        if len(row) < 2 or not row[0].strip() or not row[1].strip():
            continue
        grupo_dc = str(row[0]).strip().upper()
        nombre_dc = str(row[1]).strip()
        nombre_norm = _normalizar_nombre(nombre_dc)
        inicio_dc_s = str(row[2]).strip() if len(row) > 2 else ""
        if nombre_norm in gantt_map:
            grupo_gantt = gantt_map[nombre_norm]["grupo"]
            inicio_gantt = gantt_map[nombre_norm]["inicio"]
            grupo_cambio = grupo_dc != grupo_gantt
            fecha_cambio = (bool(inicio_dc_s) and bool(inicio_gantt)
                            and _parse_fecha(inicio_dc_s) is not None
                            and _parse_fecha(inicio_gantt) is not None
                            and _parse_fecha(inicio_dc_s) != _parse_fecha(inicio_gantt))
            if grupo_cambio or fecha_cambio:
                cambios.append({
                    "fila_sheet": i,
                    "nombre": nombre_dc,
                    "grupo_anterior": grupo_dc,
                    "grupo_nuevo": grupo_gantt,
                    "inicio_nuevo": inicio_gantt,
                    "pct": row[3] if len(row) > 3 else 0,
                })

    if not cambios:
        log.info("  Sin cambios de grupo ni de fecha detectados.")
        return

    log.info(f"  CAMBIOS DETECTADOS ({len(cambios)}):")
    requests = []
    dc_gid = _get_dc_gid(sheets_svc)
    for c in cambios:
        log.info(f"    >> {c['nombre']}: {c['grupo_anterior']} → {c['grupo_nuevo']} "
                 f"(inicio: {c['inicio_nuevo']})")
        fila_idx = c["fila_sheet"] - 1
        if c.get("grupo_cambio", True):
            requests.append({"updateCells": {
                "range": {"sheetId": dc_gid,
                          "startRowIndex": fila_idx, "endRowIndex": fila_idx + 1,
                          "startColumnIndex": 0, "endColumnIndex": 1},
                "rows": [{"values": [{"userEnteredValue": {"stringValue": c["grupo_nuevo"]}}]}],
                "fields": "userEnteredValue"
            }})
        if c["inicio_nuevo"]:
            requests.append({"updateCells": {
                "range": {"sheetId": dc_gid,
                          "startRowIndex": fila_idx, "endRowIndex": fila_idx + 1,
                          "startColumnIndex": 2, "endColumnIndex": 3},
                "rows": [{"values": [{"userEnteredValue": {"stringValue": c["inicio_nuevo"]}}]}],
                "fields": "userEnteredValue"
            }})

    if requests:
        sheets_svc.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"requests": requests},
        ).execute()
        _reordenar_datos_control(sheets_svc)
        log.info(f"  Datos Control actualizado con {len(cambios)} movimiento(s).")

    return cambios


# ─────────────────────────────────────────────────────────────────────────────
# 3. LOGICA DE CURVAS S
# ─────────────────────────────────────────────────────────────────────────────
def pct_programada(dia):
    if dia <= 0:
        return 0.0
    semana = dia / 7.0
    idx = int(semana)
    frac = semana - idx
    if idx >= len(PCT_SEMANA) - 1:
        return 100.0
    return PCT_SEMANA[idx] + frac * (PCT_SEMANA[idx + 1] - PCT_SEMANA[idx])


def s_curve_real(t_days, pct_real, t_control):
    if t_control <= 0 or pct_real == 0:
        return 0.0
    t_norm = t_days / t_control
    k = 8.0
    val      = expit(k * (t_norm - 0.5))
    val_max  = expit(k * 0.5)
    val_min  = expit(k * (-0.5))
    return float(np.clip((val - val_min) / (val_max - val_min) * pct_real, 0, pct_real))


def _estimar_inicio_efectivo(control, pct_real):
    """
    Cuando un beneficiario tiene inicio >= control pero pct_real > 0,
    estima retroactivamente la fecha en que debio haber empezado para
    alcanzar ese porcentaje segun la curva programada PCT_SEMANA.
    Ej: pct_real=18% → busca semana 5 en PCT_SEMANA → inicio_eff = control - 35 dias
    """
    if pct_real <= 0:
        return control
    pct_capped = min(float(pct_real), 99.9)
    for semana in range(len(PCT_SEMANA) - 1):
        p0, p1 = PCT_SEMANA[semana], PCT_SEMANA[semana + 1]
        if p0 <= pct_capped <= p1:
            frac = (pct_capped - p0) / max(0.1, p1 - p0)
            dias_estimados = max(7, int((semana + frac) * 7))
            return control - timedelta(days=dias_estimados)
    # pct_real muy alto (>= PCT_SEMANA[-1]) → casi al final
    return control - timedelta(days=int(DURACION_DIAS * 0.9))


def proyectar_fin(inicio, pct_real, control):
    if pct_real <= 10:
        return inicio + timedelta(days=DURACION_DIAS)
    inicio_ef = min(inicio, control)
    dias_trans = max(1, (control - inicio_ef).days)
    tasa = pct_real / dias_trans
    dias_rest = (100 - pct_real) / tasa
    return control + timedelta(days=int(dias_rest))


def build_group_curves(beneficiarios, control):
    # Opción B: si inicio >= control pero pct_real > 0, estimar inicio efectivo
    beneficiarios_adj = []
    for nombre, inicio, pct, termino in beneficiarios:
        if inicio >= control and pct > 0:
            inicio_eff = _estimar_inicio_efectivo(control, pct)
            log.info(f"    [inicio_eff] {nombre}: planif={_fmt_date(inicio)} "
                     f"→ eff={_fmt_date(inicio_eff)} ({pct:.0f}% real)")
            beneficiarios_adj.append((nombre, inicio_eff, pct, termino))
        else:
            beneficiarios_adj.append((nombre, inicio, pct, termino))
    beneficiarios = beneficiarios_adj

    inicio_grupo = min(b[1] for b in beneficiarios)
    fines_prog   = [b[3] for b in beneficiarios]   # termino individual desde Gantt
    fin_proyecto = max(fines_prog)
    fines_real   = [proyectar_fin(b[1], b[2], control) for b in beneficiarios]
    fin_proyectado = max(fines_real)

    fecha_fin = max(fin_proyecto, fin_proyectado) + timedelta(days=14)
    fechas = [inicio_grupo + timedelta(days=d)
              for d in range((fecha_fin - inicio_grupo).days + 1)]

    prog_list, real_list, proj_list = [], [], []
    for fecha in fechas:
        vp, vr, vproj = [], [], []
        for _, inicio, pct, termino in beneficiarios:
            dias = (fecha - inicio).days
            # Escalar dias a la calibración de PCT_SEMANA (245 días)
            duracion = max(1, (termino - inicio).days)
            dias_esc = dias * DURACION_DIAS / duracion
            vp.append(pct_programada(dias_esc))
            if fecha <= control:
                t_ctrl = max(1, (control - inicio).days)
                vr.append(s_curve_real(max(0, dias), pct, t_ctrl))
            if fecha >= control:
                if pct <= 0:
                    vproj.append(pct_programada(dias_esc))
                else:
                    dias_ctrl = max(1, (control - inicio).days)
                    vproj.append(min(100, (pct / dias_ctrl) * max(0, dias)))
        prog_list.append(np.mean(vp))
        real_list.append(np.mean(vr) if vr else None)
        proj_list.append(np.mean(vproj) if vproj else None)

    return fechas, prog_list, real_list, proj_list, fin_proyectado


# ─────────────────────────────────────────────────────────────────────────────
# 4. GENERAR GRAFICOS
# ─────────────────────────────────────────────────────────────────────────────
def _fmt_date(d):
    return d.strftime("%d/%m/%Y")


def generar_grafico_grupo(nombre_grupo, beneficiarios, control, outdir, pct_prog_gantt=None):
    color = COLORES.get(nombre_grupo, "#333333")
    fechas, prog, real_hist, proj, fin_proy = build_group_curves(beneficiarios, control)

    pct_real_avg   = np.mean([b[2] for b in beneficiarios])
    idx_ctrl       = (control - fechas[0]).days
    pct_prog_ctrl  = prog[idx_ctrl] if 0 <= idx_ctrl < len(prog) else 0.0

    fig, ax = plt.subplots(figsize=(10, 5.1))  # 2000x1020 px @ 200 DPI, optimo para celda 755x385
    ax.set_facecolor("#f8f9fa")
    fig.patch.set_facecolor("#ffffff")

    ax.plot(fechas, prog, color="#1f77b4", linewidth=2.5, linestyle="--",
            label="% Programado", zorder=3)

    rf = [f for f, v in zip(fechas, real_hist) if v is not None]
    rv = [v for v in real_hist if v is not None]
    ax.plot(rf, rv, color="#2ca02c", linewidth=3, label="% Real (suavizado)", zorder=4)

    pf = [f for f, v in zip(fechas, proj) if v is not None]
    pv = [v for v in proj if v is not None]
    ax.plot(pf, pv, color="#d62728", linewidth=2, linestyle=":",
            label="% Proyectado", zorder=3)

    ax.axvline(control, color="#ff7f0e", linewidth=2, linestyle="-.",
               label=f"Fecha Control ({_fmt_date(control)})", zorder=5)

    if fin_proy <= fechas[-1]:
        ax.axvline(fin_proy, color="#9467bd", linewidth=1.5, linestyle=":",
                   label=f"Termino Proyectado ({_fmt_date(fin_proy)})", zorder=4)

    if 0 <= idx_ctrl < len(prog):
        _pct_p = pct_prog_gantt if pct_prog_gantt is not None else prog[idx_ctrl]
        _diff = pct_real_avg - _pct_p
        _signo = "+" if _diff >= 0 else ""
        ax.annotate(
            f"Prog: {_pct_p:.2f}%\nReal: {pct_real_avg:.2f}%\nDesv: {_signo}{_diff:.2f}%",
            xy=(control, prog[idx_ctrl]),
            xycoords='data',
            xytext=(0.97, 0.06),
            textcoords='axes fraction',
            fontsize=12, color="#111111", fontweight="bold",
            ha='right', va='bottom',
            arrowprops=dict(arrowstyle="->", color="#ff7f0e", lw=1.0),
            bbox=dict(boxstyle="round,pad=0.4", facecolor="#fffbe6",
                      edgecolor="#ff7f0e", linewidth=1.0, alpha=0.95),
        )

    fin_prog_grupo = max(b[3] for b in beneficiarios)   # termino real del Gantt
    ax.set_xlim(fechas[0] - timedelta(days=7), fechas[-1] + timedelta(days=7))
    ax.set_ylim(-2, 108)
    ax.set_ylabel("Avance (%)", fontsize=12)
    ax.set_xlabel("Fecha", fontsize=12)
    ax.set_title(
        f"Curva S - {nombre_grupo} ({len(beneficiarios)} viviendas)\n"
        f"Inicio: {_fmt_date(min(b[1] for b in beneficiarios))}  |  "
        f"Fin Prog.: {_fmt_date(fin_prog_grupo)}  |  "
        f"Fin Proy.: {_fmt_date(fin_proy)}",
        fontsize=13, fontweight="bold", pad=10,
    )
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%b %Y"))
    ax.xaxis.set_major_locator(mdates.MonthLocator())
    plt.setp(ax.xaxis.get_majorticklabels(), rotation=45, ha="right")
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f"{x:.0f}%"))
    ax.grid(axis="y", alpha=0.4, linestyle="--")
    ax.grid(axis="x", alpha=0.2, linestyle=":")
    ax.legend(loc="upper left", fontsize=9, framealpha=0.9)
    ax.axvspan(fechas[0], control, alpha=0.04, color="green")
    ax.axvspan(control, fechas[-1], alpha=0.04, color="red")

    plt.tight_layout()
    fname = Path(outdir) / f"CurvaS_{nombre_grupo.replace(' ', '_')}.png"
    plt.savefig(fname, dpi=DPI_EXPORT, bbox_inches="tight")
    plt.close()
    log.info(f"  Grafico guardado: {fname.name}")
    return pct_real_avg, pct_prog_ctrl, min(b[1] for b in beneficiarios), fin_proy


def generar_grafico_total(grupos, control, fines_proy_global, outdir, pct_prog_gantt=None):
    todos = [b for bens in grupos.values() for b in bens]
    inicio_total   = min(b[1] for b in todos)
    fin_prog_total = FIN_PROGRAMADO_PROYECTO     # vencimiento programa de obra del proyecto
    fin_proy_total = max(fines_proy_global)

    fecha_fin_total = max(fin_prog_total, fin_proy_total) + timedelta(days=14)
    fechas_total = [inicio_total + timedelta(days=d)
                    for d in range((fecha_fin_total - inicio_total).days + 1)]

    prog_total, real_total_hist, proj_total = [], [], []
    for fecha in fechas_total:
        vp, vr, vproj = [], [], []
        for _, inicio, pct, termino in todos:
            dias = (fecha - inicio).days
            duracion = max(1, (termino - inicio).days)
            dias_esc = dias * DURACION_DIAS / duracion
            vp.append(pct_programada(dias_esc))
            if fecha <= control:
                t_ctrl = max(1, (control - inicio).days)
                vr.append(s_curve_real(max(0, dias), pct, t_ctrl))
            if fecha >= control:
                if pct <= 0:
                    vproj.append(pct_programada(dias_esc))
                else:
                    dias_ctrl = max(1, (control - inicio).days)
                    vproj.append(min(100, (pct / dias_ctrl) * max(0, dias)))
        prog_total.append(np.mean(vp))
        real_total_hist.append(np.mean(vr) if vr else None)
        proj_total.append(np.mean(vproj) if vproj else None)

    pct_real_total = np.mean([b[2] for b in todos])
    idx_ctrl_total = (control - inicio_total).days

    fig2, ax2 = plt.subplots(figsize=(10, 5.1))  # 2000x1020 px @ 200 DPI, optimo para celda 755x385
    ax2.set_facecolor("#f8f9fa")
    fig2.patch.set_facecolor("#ffffff")

    ax2.plot(fechas_total, prog_total, color="#1f77b4", linewidth=3,
             linestyle="--", label="% Programado Total", zorder=3)
    rt_f = [f for f, v in zip(fechas_total, real_total_hist) if v is not None]
    rt_v = [v for v in real_total_hist if v is not None]
    ax2.plot(rt_f, rt_v, color="#2ca02c", linewidth=3.5,
             label="% Real Total (suavizado)", zorder=4)
    pt_f = [f for f, v in zip(fechas_total, proj_total) if v is not None]
    pt_v = [v for v in proj_total if v is not None]
    ax2.plot(pt_f, pt_v, color="#d62728", linewidth=2.5, linestyle=":",
             label="% Proyectado Total", zorder=3)

    ax2.axvline(control, color="#ff7f0e", linewidth=2.5, linestyle="-.",
                label=f"Fecha Control ({_fmt_date(control)})", zorder=5)
    ax2.axvline(fin_proy_total, color="#9467bd", linewidth=2, linestyle=":",
                label=f"Termino Proyectado ({_fmt_date(fin_proy_total)})", zorder=4)
    ax2.axvline(fin_prog_total, color="#1f77b4", linewidth=1.5, linestyle=":",
                alpha=0.6, label=f"Termino Programado ({_fmt_date(fin_prog_total)})", zorder=3)

    if 0 <= idx_ctrl_total < len(prog_total):
        prog_en_ctrl = prog_total[idx_ctrl_total]
        _pct_p_t = pct_prog_gantt if pct_prog_gantt is not None else prog_en_ctrl
        _diff2 = pct_real_total - _pct_p_t
        _signo2 = "+" if _diff2 >= 0 else ""
        ax2.annotate(
            f"Prog:  {_pct_p_t:.2f}%\nReal:  {pct_real_total:.2f}%\n"
            f"Desv: {_signo2}{_diff2:.2f}%",
            xy=(control, prog_en_ctrl),
            xycoords='data',
            xytext=(0.97, 0.06),
            textcoords='axes fraction',
            fontsize=12, color="#111111", fontweight="bold",
            ha='right', va='bottom',
            arrowprops=dict(arrowstyle="->", color="#ff7f0e", lw=1.0),
            bbox=dict(boxstyle="round,pad=0.4", facecolor="#fffbe6",
                      edgecolor="#ff7f0e", linewidth=1.0, alpha=0.95),
        )

    ax2.set_xlim(fechas_total[0] - timedelta(days=7),
                 fechas_total[-1] + timedelta(days=7))
    ax2.set_ylim(-2, 108)
    ax2.set_ylabel("Avance promedio (%)", fontsize=12)
    ax2.set_xlabel("Fecha", fontsize=12)
    ax2.set_title(
        f"Curva S Total - Proyecto Nuke Mapu ({len(todos)} viviendas)\n"
        f"Inicio: {_fmt_date(inicio_total)}  |  "
        f"Fin Programado: {_fmt_date(fin_prog_total)}  |  "
        f"Fin Proyectado: {_fmt_date(fin_proy_total)}",
        fontsize=13, fontweight="bold", pad=10,
    )
    ax2.xaxis.set_major_formatter(mdates.DateFormatter("%b %Y"))
    ax2.xaxis.set_major_locator(mdates.MonthLocator())
    plt.setp(ax2.xaxis.get_majorticklabels(), rotation=45, ha="right")
    ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f"{x:.0f}%"))
    ax2.grid(axis="y", alpha=0.4, linestyle="--")
    ax2.grid(axis="x", alpha=0.2, linestyle=":")
    ax2.legend(loc="upper left", fontsize=9, framealpha=0.95)
    ax2.axvspan(fechas_total[0], control, alpha=0.03, color="green")
    ax2.axvspan(control, fechas_total[-1], alpha=0.03, color="red")

    plt.tight_layout()
    fname_total = Path(outdir) / "CurvaS_TOTAL_Nuke_Mapu.png"
    plt.savefig(fname_total, dpi=DPI_EXPORT, bbox_inches="tight")
    plt.close()
    log.info(f"  Grafico guardado: {fname_total.name}")


def generar_grafico_todos(grupos, control, fines_proy_global, outdir):
    fig3, ax3 = plt.subplots(figsize=(10, 5.1))  # 2000x1020 px @ 200 DPI, optimo para celda 755x385
    ax3.set_facecolor("#f8f9fa")
    fig3.patch.set_facecolor("#ffffff")

    for nombre_grupo, beneficiarios in grupos.items():
        color = COLORES.get(nombre_grupo, "#333333")
        fechas, prog, real_hist, proj, _ = build_group_curves(beneficiarios, control)
        ax3.plot(fechas, prog, color=color, linewidth=1.5, linestyle="--", alpha=0.6)
        rf = [f for f, v in zip(fechas, real_hist) if v is not None]
        rv = [v for v in real_hist if v is not None]
        ax3.plot(rf, rv, color=color, linewidth=2.5, label=nombre_grupo)
        pf = [f for f, v in zip(fechas, proj) if v is not None]
        pv = [v for v in proj if v is not None]
        ax3.plot(pf, pv, color=color, linewidth=1.8, linestyle=":", alpha=0.8)

    ax3.axvline(control, color="#ff7f0e", linewidth=2.5, linestyle="-.",
                label="Fecha Control", zorder=5)

    legend_extra = [
        Line2D([0], [0], color="gray", linestyle="--", linewidth=1.5, label="Programado"),
        Line2D([0], [0], color="gray", linestyle="-",  linewidth=2.5, label="Real (suavizado)"),
        Line2D([0], [0], color="gray", linestyle=":",  linewidth=1.8, label="Proyectado"),
    ]
    handles, labels = ax3.get_legend_handles_labels()
    ax3.legend(handles=handles + legend_extra, loc="upper left", fontsize=9, framealpha=0.95)

    ax3.set_xlim(date(2026, 1, 1), max(fines_proy_global) + timedelta(days=30))
    ax3.set_ylim(-2, 108)
    ax3.set_ylabel("Avance (%)", fontsize=12)
    ax3.set_xlabel("Fecha", fontsize=12)
    ax3.set_title(
        "Curva S por Grupo - Proyecto Nuke Mapu\n"
        "Linea continua = Real  |  Discontinua = Programado  |  Punteado = Proyectado",
        fontsize=13, fontweight="bold", pad=10,
    )
    ax3.xaxis.set_major_formatter(mdates.DateFormatter("%b %Y"))
    ax3.xaxis.set_major_locator(mdates.MonthLocator())
    plt.setp(ax3.xaxis.get_majorticklabels(), rotation=45, ha="right")
    ax3.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f"{x:.0f}%"))
    ax3.grid(axis="y", alpha=0.4, linestyle="--")
    ax3.grid(axis="x", alpha=0.2, linestyle=":")
    ax3.axvspan(date(2026, 1, 1), control, alpha=0.03, color="green")

    plt.tight_layout()
    fname = Path(outdir) / "CurvaS_Todos_Grupos.png"
    plt.savefig(fname, dpi=DPI_EXPORT, bbox_inches="tight")
    plt.close()
    log.info(f"  Grafico guardado: {fname.name}")


# ─────────────────────────────────────────────────────────────────────────────
# 5. REDIMENSIONAR
# ─────────────────────────────────────────────────────────────────────────────
def redimensionar(outdir):
    # Redimensionado desactivado: se sube en resolucion nativa para maxima calidad
    log.info("Resolucion nativa: sin redimensionado.")


# ─────────────────────────────────────────────────────────────────────────────
# 6. INSERTAR IMAGENES EN HOJA 'Curvas S' via formula IMAGE()
# ─────────────────────────────────────────────────────────────────────────────
CURVAS_GID = 1131401009   # gid de la hoja 'Curvas S' en Nuke Mapu

IMAGENES_SHEETS = [
    # (drive_id, col, row)
    ("1k2gpSr9Sk5zUZ3MAOxw3zGcHRa9ljFSA", 0, 0),  # GRUPO 1
    ("1Ciu52kLYT9NjgaqUBPif4mtWk6l0RtPm", 1, 0),  # GRUPO 2
    ("1mDdpyG5zGmSgOXZVMN4qzGbp_0BMrzT2", 0, 1),  # GRUPO 3
    ("1GHthHYK1bpoA7_2Q5qSm3VFJq4KPYQa_", 1, 1),  # GRUPO 4
    ("12yQw_tt3W4H91Q4KpiMafm_eLUNkwqpP", 0, 2),  # GRUPO 5
    ("11L-TIagyTGZOh3yhbGvjdIruG0dJQGPQ", 1, 2),  # TOTAL
    ("1_xxaqAay-UeB4POTcXMaBRD_Xo5ceb-1", 0, 3),  # Todos Grupos
]


def insertar_imagenes_en_sheets(sheets_svc):
    """Actualiza las formulas IMAGE() en la hoja 'Curvas S' de Nuke Mapu."""
    log.info("Actualizando formulas IMAGE() en hoja 'Curvas S'...")
    W_PX, H_PX = 755, 385
    n_cols = max(c for _, c, _ in IMAGENES_SHEETS) + 1
    n_rows = max(r for _, _, r in IMAGENES_SHEETS) + 1

    requests = [
        {"updateDimensionProperties": {
            "range": {"sheetId": CURVAS_GID, "dimension": "COLUMNS",
                      "startIndex": 0, "endIndex": n_cols},
            "properties": {"pixelSize": W_PX}, "fields": "pixelSize"
        }},
        {"updateDimensionProperties": {
            "range": {"sheetId": CURVAS_GID, "dimension": "ROWS",
                      "startIndex": 0, "endIndex": n_rows},
            "properties": {"pixelSize": H_PX}, "fields": "pixelSize"
        }},
    ]
    # Timestamp unico por ejecucion -> fuerza re-fetch en Sheets (evita cache)
    ts = datetime.now().strftime("%Y%m%d%H%M%S")
    for drive_id, col, row in IMAGENES_SHEETS:
        url = f"https://drive.google.com/uc?id={drive_id}&export=view&v={ts}"
        requests.append({"updateCells": {
            "range": {"sheetId": CURVAS_GID,
                      "startRowIndex": row, "endRowIndex": row + 1,
                      "startColumnIndex": col, "endColumnIndex": col + 1},
            "rows": [{"values": [{"userEnteredValue": {
                "formulaValue": f'=IMAGE("{url}",1)'
            }}]}],
            "fields": "userEnteredValue"
        }})

    sheets_svc.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"requests": requests}
    ).execute()
    log.info(f"  {len(IMAGENES_SHEETS)} formulas IMAGE() actualizadas (v={ts}).")


# ─────────────────────────────────────────────────────────────────────────────
# 7. ACTUALIZAR ARCHIVOS EN DRIVE (mismo ID = mismos links en Sheets)
# ─────────────────────────────────────────────────────────────────────────────
def actualizar_drive(drive_svc, outdir):
    global DRIVE_IDS
    # Seed migration: si Firebase/JSON aún no tiene IDs para este proyecto,
    # sembrar con los IDs hardcodeados para que el update in-place funcione.
    existing = _ccu.load_drive_ids("nuke_mapu")
    if not existing:
        _ccu.save_drive_ids("nuke_mapu", DRIVE_IDS)
    DRIVE_IDS = _ccu.actualizar_drive_organizado(
        drive_svc, outdir, list(DRIVE_IDS.keys()), OBRA_FOLDER, "nuke_mapu"
    )


# ─────────────────────────────────────────────────────────────────────────────
# 7. MARCAR HOJA COMO PENDIENTE (Apps Script detectara el flag)
# ─────────────────────────────────────────────────────────────────────────────
def marcar_pendiente(sheets_svc, timestamp_str):
    control_str = date.today().strftime("%d/%m/%Y")
    sheets_svc.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range="Datos Control!A1:B2",
        valueInputOption="USER_ENTERED",
        body={"values": [
            ["ESTADO",         f"PENDIENTE - {timestamp_str}"],
            ["FECHA_CONTROL",  control_str],
        ]},
    ).execute()
    log.info(f"Flag PENDIENTE escrito | FECHA_CONTROL={control_str}")


def actualizar_pct_en_hoja(sheets_svc, grupos):
    """Escribe % real de vuelta a 'Datos Control' col D para mantener fallback fresco."""
    valores = []
    for bens in grupos.values():
        for b in bens:
            valores.append([round(b[2], 2)])
    if not valores:
        return
    n = len(valores)
    sheets_svc.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'Datos Control'!D5:D{4 + n}",
        valueInputOption="USER_ENTERED",
        body={"values": valores},
    ).execute()
    log.info(f"  % real guardados en 'Datos Control' ({n} filas)")


def marcar_ok(sheets_svc, timestamp_str):
    sheets_svc.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range="Datos Control!A1:B1",
        valueInputOption="USER_ENTERED",
        body={"values": [["ESTADO", f"OK - {timestamp_str}"]]},
    ).execute()


# ─────────────────────────────────────────────────────────────────────────────
# 8. MAIN
# ─────────────────────────────────────────────────────────────────────────────
def main():
    ts = datetime.now().strftime("%d/%m/%Y %H:%M")
    log.info("=" * 65)
    log.info(f"INICIO ACTUALIZACION CURVAS S  -  {ts}")
    log.info("=" * 65)

    try:
        # Credenciales
        creds      = get_credentials()
        sheets_svc = build("sheets", "v4", credentials=creds)
        drive_svc  = build("drive",  "v3", credentials=creds)

        # Sincronizar grupos desde Gantt (detecta movimientos semanales)
        sincronizar_grupos_desde_gantt(sheets_svc)

        # Leer datos de control
        control_date, grupos = leer_datos_control(sheets_svc)

        # Guardar % actualizados de vuelta a Datos Control (fallback fresco para la próxima semana)
        actualizar_pct_en_hoja(sheets_svc, grupos)

        # Crear carpeta de salida si no existe
        Path(OUTPUT_DIR).mkdir(parents=True, exist_ok=True)

        # Generar graficos por grupo
        log.info("Generando graficos por grupo...")
        pct_prog_gantt = _ccu.leer_pct_programa_gantt(sheets_svc, SPREADSHEET_ID, GANTT_HOJA)
        log.info(f"% Prog Gantt leido: {pct_prog_gantt}%")
        pct_prog_por_grupo = _ccu.leer_pct_prog_por_grupo(sheets_svc, SPREADSHEET_ID, GANTT_HOJA, control_date)
        log.info(f"% Prog Gantt por grupo: {pct_prog_por_grupo}")
        fines_proy_global = []
        for nombre_grupo, beneficiarios in grupos.items():
            _, _, _, fin_proy = generar_grafico_grupo(
                nombre_grupo, beneficiarios, control_date, OUTPUT_DIR, pct_prog_gantt=pct_prog_por_grupo.get(nombre_grupo.upper())
            )
            fines_proy_global.append(fin_proy)

        # Grafico total y todos grupos
        log.info("Generando graficos consolidados...")
        generar_grafico_total(grupos, control_date, fines_proy_global, OUTPUT_DIR, pct_prog_gantt=pct_prog_gantt)
        generar_grafico_todos(grupos, control_date, fines_proy_global, OUTPUT_DIR)

        # Redimensionar
        redimensionar(OUTPUT_DIR)

        # Actualizar Drive
        actualizar_drive(drive_svc, OUTPUT_DIR)

        # Insertar/actualizar imagenes directamente en hoja 'Curvas S'
        _ccu.retry_sheets(insertar_imagenes_en_sheets, sheets_svc)

        # Marcar pendiente para que Apps Script refresque Sheets
        marcar_pendiente(sheets_svc, ts)

        log.info("=" * 65)
        log.info("COMPLETADO. Apps Script actualizara las imagenes en Sheets.")
        log.info("=" * 65)

    except Exception:
        log.error("ERROR durante la actualizacion:\n" + traceback.format_exc())
        sys.exit(1)


if __name__ == "__main__":
    main()
