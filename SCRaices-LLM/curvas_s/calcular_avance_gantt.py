"""
calcular_avance_gantt.py
========================
Lee la pestaña 'Datos Control' de cada Gantt de control (Google Sheets),
calcula el promedio de avance real (%col D) y lee el avance programado
directamente de la fila "Programa" de la hoja Gantt principal, y escribe
el resultado en Firebase RTDB:

  avance_gantt/{pid} = {
      pct:       <float, 1 decimal>,   # promedio avance real (%)
      pct_prog:  <float, 1 decimal>,   # avance programado leido del Gantt a hoy
      n:         <int>,                # beneficiarios con valor > 0
      total:     <int>,                # total beneficiarios leidos
      fuente:    "Datos Control",
      actualizado: "YYYY-MM-DD",
  }

Ejecutar: python calcular_avance_gantt.py
Integrado en ejecutar_curvas_cloud.py (POST_SCRIPTS).
"""

import sys
import time
import logging
import requests
from datetime import date, datetime, timedelta
from googleapiclient.discovery import build
import curvas_cloud_utils as _ccu

# ─────────────────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────────────────
FIREBASE_URL       = "https://scraices-dashboard-default-rtdb.firebaseio.com/avance_gantt.json"
FIREBASE_GANTT_URL = "https://scraices-dashboard-default-rtdb.firebaseio.com/gantt_programa.json"
FIREBASE_BENEF_URL = "https://scraices-dashboard-default-rtdb.firebaseio.com/avance_benef.json"

# Spreadsheet ID de cada Gantt de control
PROYECTOS = {
    "P31":  "1kgroktRIto3gGGnvmXGMgoetXT-Rv8JFue1K9eNlNMg",
    "P38":  "151wIDnZn8_b7egJKLQKUcD6QDCflEWF5OlAU9KWgc-M",
    "P126": "1IwBN7CpDvKVAvaRuYHNUfkQ0e98cJMVFDKFrntv3oNw",
    "P39":  "1GiaZ1i3BN3mbgFmg16Ze5E25R0jEmYULtCWRy6cKaHo",
    "P127": "1bgT-83Aea0DlyeQ6OvitGDZfm3jLRQUV0GooVP-G8EI",
    "P12":  "1SLu5lQTAzhHOUuorM3jbxSBes7vMyIB9mMXCihQ6A40",
    "P14":  "1B4wO-UkIDDyFvwRYjMAGksJvKqtNblwl_IirLl3qA6E",
    "P116": "1z9kNq9uo363NrWqCojGfGpP326FDj3V60irWxtMMbU8",
    "P119": "1t_1j62f_3l1nrlufmvhnV-o1WTplv0OnQL_JdVaWgKA",
    "P131": "1n5F-P5cy8Wj5BujllzdnwCrKwGsfIkdvHxcyd6YwscU",
    "P28":  "18XkRb7RAF52Aqj4immGebME-d9sODY0HgxIKWcBGrlg",
    # ── Proyectos futuros — descomentar y reemplazar TODO con el ID de la planilla Gantt ──
    # "P118": "TODO",  # El Canelo             — Rural Araucanía 2024
    # "P123": "TODO",  # Peumayen 2023          — Rural Araucanía 2024
    # "P128": "TODO",  # Com. José Carvajal     — Rural Araucanía 2024
    # "P129": "TODO",  # Nuevo Gorbea           — Rural Araucanía 2024
    # "P132": "TODO",  # Com. Fermín Manquilef  — Rural Araucanía 2024
    # "P145": "TODO",  # Perkenko 2025          — Rural Araucanía 2025
    # "P146": "TODO",  # Demanda Villarrica 2025
    # "P147": "TODO",  # Ruka Antu              — Rural Araucanía 2025
    # "P150": "TODO",  # Llaima Antu            — Rural Araucanía 2025
    # "P152": "TODO",  # Ayun Ruka              — Rural Araucanía 2025
    # "P153": "TODO",  # Vilcun Mapu            — Rural Araucanía 2025
    # "P154": "TODO",  # Com. José Carvajal 2   — Rural Araucanía 2025
    # "P155": "TODO",  # Los Arrayanes          — Rural Araucanía 2025
    # "P156": "TODO",  # Poyen Ruka             — Rural Araucanía 2025
    # "P164": "TODO",  # Conun Huenu            — Rural Araucanía 2026
    # "P166": "TODO",  # Malalhue               — Rural Araucanía 2026
    # "P167": "TODO",  # Witran Donguil         — Rural Araucanía 2026
    # "P168": "TODO",  # Raíces de Perquenco    — Rural Araucanía 2026
    # "P170": "TODO",  # Los Copihues de Cunco  — Rural Araucanía 2026
    # "P171": "TODO",  # Raíces de Trovolhue    — Rural Araucanía 2026
    # "P172": "TODO",  # Raíces Costeras        — Rural Araucanía 2026
}

# Nombre de la hoja Gantt principal por proyecto (donde está la fila "Programa")
GANTT_HOJAS = {
    "P31":  "Programa de obra",
    "P38":  "Programa de obra",
    "P126": "Programa de obra",
    "P39":  "Programa de obra",
    "P127": "Programa de obra",
    "P12":  "Programa de obra",
    "P14":  "Programa de obra",
    "P116": "Programa de obra",
    "P119": "Programa de obra",
    "P131": "Programa de obra",
    "P28":  "Programa de obra",
}

# Curva S de fallback (solo se usa si el Gantt no puede leerse)
PROYECTO_CURVAS = {
    "P14": {
        "pct_semana": [0, 2, 7, 11, 16, 20, 25, 30, 34, 39, 43, 48, 52, 57, 61,
                       66, 70, 75, 80, 84, 89, 93, 98, 100, 100, 100],
        "duracion": 161,
    },
    "P116": {
        "pct_semana": [0, 5, 9, 14, 18, 23, 27, 32, 36, 41, 45, 50,
                       55, 59, 64, 68, 73, 77, 82, 86, 91, 95, 100, 100],
        "duracion": 147,
    },
    "P38": {
        "pct_semana": [0, 2, 5, 9, 14, 18, 23, 27, 32, 36, 41, 45, 50, 55, 59, 64,
                       68, 73, 77, 82, 86, 91, 95, 100, 100, 100, 100, 100, 100, 100,
                       100, 100, 100],
        "duracion": 224,
    },
    "P31": {
        "pct_semana": [0, 2, 3, 6, 9, 12, 15, 18, 21, 24, 27, 30, 33, 36, 39, 42, 45,
                       48, 52, 55, 58, 61, 64, 67, 70, 73, 76, 79, 82, 85, 88, 91, 94,
                       97, 98, 100, 100],
        "duracion": 221,
    },
    "P12": {
        "pct_semana": [0, 4, 7, 11, 14, 18, 21, 25, 29, 32, 36, 39, 43, 46, 50,
                       54, 57, 61, 64, 68, 71, 75, 79, 82, 86, 89, 93, 96, 100, 100],
        "duracion": 245,
    },
    "P119": {
        "pct_semana": [0, 4, 7, 11, 14, 18, 21, 25, 29, 32, 36, 39, 43, 46, 50,
                       54, 57, 61, 64, 68, 71, 75, 79, 82, 86, 89, 93, 96, 100, 100],
        "duracion": 245,
    },
    "P131": {
        "pct_semana": [0, 3, 6, 8, 11, 14, 17, 19, 22, 25, 28, 31, 33, 36, 39, 42,
                       44, 47, 50, 53, 56, 58, 61, 64, 67, 69, 72, 75, 78, 81, 83,
                       86, 89, 92, 94, 97, 100, 100, 100, 100, 100, 100, 100],
        "duracion": 294,
    },
    "P28": {
        "pct_semana": [0,  2,  4,  7, 10, 13, 17, 21, 25, 30, 35, 40, 45, 50, 55,
                       60, 65, 70, 74, 78, 82, 85, 88, 90, 92, 94, 96, 97, 98, 99,
                       100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100],
        "duracion": 154,
    },
    "P127": {
        "pct_semana": [0, 4, 7, 11, 14, 18, 21, 25, 29, 32, 36, 39, 43, 46, 50,
                       54, 57, 61, 64, 68, 71, 75, 79, 82, 86, 89, 93, 96, 100, 100,
                       100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100,
                       100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100,
                       100, 100],
        "duracion": 385,
    },
    "P39": {
        "pct_semana": [0,  2,  3,  5,  7,  9, 12, 14, 17, 20, 23, 27, 30, 34, 37,
                       41, 44, 48, 51, 55, 58, 62, 65, 68, 71, 75, 78, 81, 84, 87,
                       89, 91, 93, 95, 96, 97, 98, 99, 100, 100, 100, 100, 100, 100,
                       100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100,
                       100, 100],
        "duracion": 399,
    },
    "P126": {
        "pct_semana": [0, 4, 9, 13, 17, 22, 26, 30, 35, 39, 43, 48, 52, 57, 61, 65,
                       70, 74, 78, 83, 87, 91, 96, 100, 100, 100, 100, 100, 100, 100,
                       100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100,
                       100, 100, 100, 100, 100, 100, 100, 100, 100],
        "duracion": 350,
    },
}

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger(__name__)


def _normalizar_clave(nombre: str) -> str:
    """Normaliza un nombre para usarlo como clave Firebase: sin tildes, mayúsculas, espacios→guión bajo."""
    import unicodedata
    nfkd = unicodedata.normalize("NFKD", nombre)
    sin_t = "".join(c for c in nfkd if not unicodedata.combining(c))
    return "_".join(sin_t.upper().split())


# ─────────────────────────────────────────────────────────────────────────────
# Lectura directa del Gantt — fila "Programa"
# ─────────────────────────────────────────────────────────────────────────────

_SHEETS_BASE = date(1899, 12, 30)
_PROGRAMA_LABELS = {"programa", "progra", "prog.", "prog"}


def _leer_gantt_datos(sheets_svc, spreadsheet_id, hoja) -> tuple:
    """
    Lee la hoja Gantt en una sola llamada y retorna:
      (pct_actual, mensual)
    donde:
      pct_actual : float % programado acumulado a la columna más reciente <= hoy,
                   o None si no se puede determinar.
      mensual    : dict {YYYY-MM: float} con el % acumulado al último día de
                   cada mes presente en el Gantt, o None si no hay datos.

    Lógica de columna actual:
      - Prefiere la columna pasada más reciente (serial <= hoy).
      - Si todas son futuras, usa la más próxima.
      - Si hoy supera el Gantt → pct_actual=100; si es anterior al inicio → 0.
    """
    hoy = date.today()
    hoy_ser = (hoy - _SHEETS_BASE).days

    # ── Leer toda la hoja en una sola llamada ─────────────────────────────────
    try:
        r = sheets_svc.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"'{hoja}'!A1:EZ120",
            valueRenderOption="UNFORMATTED_VALUE",
        ).execute()
    except Exception as e:
        log.warning(f"    Gantt '{hoja}': error leyendo hoja → {e}")
        return None, None

    rows = r.get("values", [])
    if not rows:
        return None, None

    # ── Columnas de fecha desde la primera fila ────────────────────────────────
    row1 = rows[0]
    date_cols = []  # [(col_idx, date, serial)]
    for ci, v in enumerate(row1):
        if isinstance(v, (int, float)) and 44000 < v < 52000:
            try:
                d = _SHEETS_BASE + timedelta(days=int(v))
                date_cols.append((ci, d, int(v)))
            except Exception:
                pass

    if not date_cols:
        log.warning(f"    Gantt '{hoja}': no se encontraron columnas de fecha")
        return None, None

    serials = [s for _, _, s in date_cols]
    min_ser, max_ser = min(serials), max(serials)

    # ── Casos fuera de rango ───────────────────────────────────────────────────
    pct_override = None
    if hoy_ser >= max_ser:
        log.info(f"    Gantt '{hoja}': hoy >= última columna Gantt → pct_prog=100%")
        pct_override = 100.0
    elif hoy_ser < min_ser - 7:
        log.info(f"    Gantt '{hoja}': hoy < inicio → pct_prog=0%")
        return 0.0, None

    # ── Columna actual: más reciente que no supere hoy ────────────────────────
    past = [(ci, d, s) for ci, d, s in date_cols if s <= hoy_ser]
    if past:
        cur_ci, _, cur_s = max(past, key=lambda x: x[2])
        best_dist = hoy_ser - cur_s
    else:
        cur_ci, _, cur_s = min(date_cols, key=lambda x: abs(x[2] - hoy_ser))
        best_dist = min(abs(s - hoy_ser) for _, _, s in date_cols)

    if best_dist > 14 and pct_override is None:
        log.warning(f"    Gantt '{hoja}': columna actual no encontrada (dist={best_dist})")
        return None, None

    # ── Buscar fila "Programa" ────────────────────────────────────────────────
    prog_row = None
    prog_label = "?"
    for row in rows:
        for ci in range(min(25, len(row))):
            label = str(row[ci]).strip().lower()
            if (label in _PROGRAMA_LABELS or "rograma" in label) and "grupo" not in label:
                prog_row = row
                prog_label = row[ci]
                break
        if prog_row is not None:
            break

    if prog_row is None:
        log.warning(f"    Gantt '{hoja}': fila 'Programa' no encontrada")
        return None, None

    def _parse_val(v):
        if not isinstance(v, (int, float)) or v < 0:
            return None
        return round(float(v) * (100.0 if v <= 1.5 else 1.0), 2)

    # ── pct_actual ────────────────────────────────────────────────────────────
    pct_actual = None
    if cur_ci < len(prog_row):
        pct_actual = _parse_val(prog_row[cur_ci])
        if pct_actual is not None:
            log.info(
                f"    Gantt '{hoja}': pct_prog={pct_actual}%  "
                f"(dist={best_dist}d label='{prog_label}')"
            )

    # Si pct_actual >=90% y el ultimo valor del Gantt es 100% a <=7 dias,
    # el proyecto esta en su etapa final → 100%
    if pct_actual is not None and 90.0 <= pct_actual < 100.0:
        last_ser, last_pct = None, None
        for ci, _, s in sorted(date_cols, key=lambda x: x[2], reverse=True):
            if ci < len(prog_row):
                v = _parse_val(prog_row[ci])
                if v is not None and v > 0:
                    last_ser, last_pct = s, v
                    break
        if last_pct is not None and last_pct >= 100.0 and last_ser - hoy_ser <= 7:
            log.info(
                f"    Gantt '{hoja}': pct_actual={pct_actual}% >=90% y "
                f"100% a {last_ser - hoy_ser}d → 100%"
            )
            pct_actual = 100.0

    # ── Serie mensual completa ────────────────────────────────────────────────
    # Para cada mes toma el valor de la columna más tardía dentro del mes.
    mensual: dict[str, float] = {}
    for ci, d, _ in date_cols:
        if ci >= len(prog_row):
            continue
        pct = _parse_val(prog_row[ci])
        if pct is None:
            continue
        ym = d.strftime("%Y-%m")
        mensual[ym] = pct  # última columna del mes sobreescribe las anteriores

    if mensual:
        log.info(
            f"    Gantt '{hoja}': mensual {len(mensual)} meses "
            f"({min(mensual)} → {max(mensual)})"
        )

    pct_final = pct_override if pct_override is not None else pct_actual
    return pct_final, (mensual if mensual else None)


# ─────────────────────────────────────────────────────────────────────────────
# Fallback: curva S (cuando el Gantt no puede leerse)
# ─────────────────────────────────────────────────────────────────────────────

def _parse_inicio(val) -> date | None:
    if val is None or val == "":
        return None
    if isinstance(val, (int, float)):
        try:
            return _SHEETS_BASE + timedelta(days=int(val))
        except Exception:
            return None
    s = str(val).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%m/%d/%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def _pct_programada(dia: int, pct_semana: list) -> float:
    if dia <= 0:
        return 0.0
    semana = dia / 7.0
    idx = int(semana)
    frac = semana - idx
    if idx >= len(pct_semana) - 1:
        return 100.0
    return pct_semana[idx] + frac * (pct_semana[idx + 1] - pct_semana[idx])


def _calc_pct_prog_curvas(inicios: list, hoy: date, pct_semana: list) -> float | None:
    if not inicios:
        return None
    vals = [_pct_programada((hoy - ini).days, pct_semana) for ini in inicios]
    return round(sum(vals) / len(vals), 2)


# ─────────────────────────────────────────────────────────────────────────────
# Lectura Datos Control (avance real)
# ─────────────────────────────────────────────────────────────────────────────

def _leer_datos_control(sheets_svc, spreadsheet_id, pid) -> dict | None:
    """
    Lee 'Datos Control' y retorna pct (avance real promedio) sin pct_prog.
    pct_prog se obtiene por separado de la hoja Gantt.
    """
    try:
        meta = sheets_svc.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheet_names = [s["properties"]["title"] for s in meta["sheets"]]
    except Exception as e:
        log.error(f"  {pid}: error obteniendo hojas → {e}")
        return None

    hoja = None
    for nombre in ["Datos Control", "datos control", "DatosControl"]:
        if nombre in sheet_names:
            hoja = nombre
            break
    if not hoja:
        log.warning(f"  {pid}: no se encontro 'Datos Control' en {sheet_names[:5]}")
        return None

    try:
        result = sheets_svc.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"'{hoja}'!A1:D200",
            valueRenderOption="UNFORMATTED_VALUE",
            dateTimeRenderOption="FORMATTED_STRING",
        ).execute()
    except Exception as e:
        log.error(f"  {pid}: error leyendo '{hoja}': {e}")
        return None

    rows = result.get("values", [])
    if len(rows) < 5:
        log.warning(f"  {pid}: muy pocas filas en '{hoja}' ({len(rows)})")
        return None

    hoy = date.today()
    curvas_cfg = PROYECTO_CURVAS.get(pid)
    valores_real   = []
    inicios_curvas = []
    benef_dict     = {}  # {clave_normalizada: pct} para avance_benef Firebase

    for row in rows[4:]:
        if len(row) < 2:
            continue
        grupo  = str(row[0]).strip() if row[0] else ""
        nombre = str(row[1]).strip() if len(row) > 1 and row[1] else ""
        if not grupo or not nombre:
            continue

        pct_raw = row[3] if len(row) > 3 else 0
        try:
            pct = float(str(pct_raw).replace("%", "").strip())
        except (ValueError, TypeError):
            pct = 0.0
        valores_real.append(pct)
        benef_dict[_normalizar_clave(nombre)] = round(pct, 2)

        if curvas_cfg and len(row) > 2:
            ini = _parse_inicio(row[2])
            if ini:
                inicios_curvas.append(ini)

    if not valores_real:
        log.warning(f"  {pid}: sin beneficiarios en '{hoja}'")
        return None

    n_con_valor = sum(1 for v in valores_real if v > 0)
    promedio    = round(sum(valores_real) / len(valores_real), 2)

    return {
        "pct":            promedio,
        "n":              n_con_valor,
        "total":          len(valores_real),
        "fuente":         "Datos Control",
        "actualizado":    hoy.isoformat(),
        "_inicios_curvas": inicios_curvas,  # solo para fallback, no se escribe a Firebase
        "_benef":         benef_dict,       # por beneficiario, se escribe a avance_benef
    }


# ─────────────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────────────

def main():
    log.info("=== calcular_avance_gantt ===")
    creds = _ccu.get_credentials()
    sheets_svc = build("sheets", "v4", credentials=creds)

    # Fechas de inicio de gantt_programa (fallback de ultimo recurso)
    try:
        gp_raw = requests.get(FIREBASE_GANTT_URL, timeout=15).json() or {}
    except Exception as e:
        log.warning(f"No se pudo leer gantt_programa: {e}")
        gp_raw = {}

    def _gantt_inicio_fb(pid) -> date | None:
        node = gp_raw.get(pid, {})
        ini = node.get("inicio") if isinstance(node, dict) else None
        if not ini:
            return None
        try:
            return datetime.strptime(str(ini)[:10], "%Y-%m-%d").date()
        except Exception:
            return None

    resultado = {}
    resultado_benef = {}  # {pid: {clave_benef: pct}}
    for pid, sid in PROYECTOS.items():
        log.info(f"Leyendo {pid}...")
        time.sleep(2)  # evita saturar cuota 60 reads/min de Sheets API

        # ── Avance real desde 'Datos Control' ─────────────────────────────────
        datos = _leer_datos_control(sheets_svc, sid, pid)
        if not datos:
            continue

        inicios_curvas = datos.pop("_inicios_curvas", [])
        benef_dict     = datos.pop("_benef", {})
        if benef_dict:
            resultado_benef[pid] = benef_dict

        # ── Avance programado: Gantt directo → curva S → fallback Firebase ────
        hoja_gantt = GANTT_HOJAS.get(pid)
        pct_prog = None

        if hoja_gantt:
            pct_prog, mensual = _leer_gantt_datos(sheets_svc, sid, hoja_gantt)
            if mensual:
                datos["mensual"] = mensual
        else:
            mensual = None

        if pct_prog is None:
            # Fallback 1: curva S con fechas de col C
            curvas_cfg = PROYECTO_CURVAS.get(pid)
            hoy = date.today()
            if curvas_cfg and inicios_curvas:
                pct_prog = _calc_pct_prog_curvas(inicios_curvas, hoy, curvas_cfg["pct_semana"])
                log.info(
                    f"  {pid}: pct_prog={pct_prog}% [fallback curva S, "
                    f"{len(inicios_curvas)}/{datos['total']} benef.]"
                )

        if pct_prog is None:
            # Fallback 2: fecha inicio del proyecto desde Firebase gantt_programa
            curvas_cfg = PROYECTO_CURVAS.get(pid)
            gi = _gantt_inicio_fb(pid)
            hoy = date.today()
            if curvas_cfg and gi:
                pct_prog = round(_pct_programada((hoy - gi).days, curvas_cfg["pct_semana"]), 2)
                log.info(f"  {pid}: pct_prog={pct_prog}% [fallback inicio proy {gi}]")

        if pct_prog is not None:
            datos["pct_prog"] = pct_prog

        log.info(
            f"  {pid}: real={datos['pct']}%"
            + (f"  prog={datos['pct_prog']}%" if "pct_prog" in datos else "")
            + f"  ({datos['n']}/{datos['total']} benef.)"
        )
        resultado[pid] = datos

    if not resultado:
        log.error("Sin datos para ningun proyecto. Abortando escritura en Firebase.")
        sys.exit(1)

    log.info(f"Escribiendo {len(resultado)} proyectos en Firebase (avance_gantt)...")
    resp = requests.put(FIREBASE_URL, json=resultado, timeout=30)
    if resp.status_code == 200:
        log.info("Firebase avance_gantt actualizado OK.")
    else:
        log.error(f"Error Firebase avance_gantt: {resp.status_code} {resp.text[:200]}")
        sys.exit(1)

    if resultado_benef:
        log.info(f"Escribiendo avance por beneficiario ({sum(len(v) for v in resultado_benef.values())} benef. en {len(resultado_benef)} proyectos)...")
        resp2 = requests.put(FIREBASE_BENEF_URL, json=resultado_benef, timeout=30)
        if resp2.status_code == 200:
            log.info("Firebase avance_benef actualizado OK.")
        else:
            log.error(f"Error Firebase avance_benef: {resp2.status_code} {resp2.text[:200]}")

    log.info("=== Listo ===")


if __name__ == "__main__":
    main()
