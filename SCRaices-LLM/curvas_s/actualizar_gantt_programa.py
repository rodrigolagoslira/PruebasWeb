"""
actualizar_gantt_programa.py
============================
Lee la fila de resumen INICIO / TERMINO / DIAS de cada Gantt de control
(Google Sheets) y escribe en Firebase RTDB:

  gantt_programa/{pid} = {
      inicio:   "YYYY-MM-DD",
      finProg:  "YYYY-MM-DD",
      plazo:    NNN,
  }

La fila de resumen es la que está inmediatamente debajo de los headers
"INICIO | TERMINO | DIAS" en el bloque FECHAS del Gantt. En todos los
proyectos se usa el mismo esquema visual; sólo varía el nombre de la
hoja y la columna de inicio del bloque.

Ejecutar: python actualizar_gantt_programa.py
Integrado en ejecutar_curvas_semanal.bat.
"""

import sys, logging, requests
from datetime import datetime
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import curvas_cloud_utils as _ccu

# ─────────────────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────────────────
TOKEN_FILE   = _ccu.TOKEN_FILE
FIREBASE_URL = "https://scraices-dashboard-default-rtdb.firebaseio.com/gantt_programa.json"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# Spreadsheet ID por proyecto.
# Orden de búsqueda de hojas: se prueba la lista hasta encontrar el bloque FECHAS.
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
    # "P147": "TODO",  # Ruka Antü              — Rural Araucanía 2025
    # "P150": "TODO",  # Llaima Antu            — Rural Araucanía 2025
    # "P152": "TODO",  # Ayün Ruka              — Rural Araucanía 2025
    # "P153": "TODO",  # Vilcún Mapu            — Rural Araucanía 2025
    # "P154": "TODO",  # Com. José Carvajal 2   — Rural Araucanía 2025
    # "P155": "TODO",  # Los Arrayanes          — Rural Araucanía 2025
    # "P156": "TODO",  # Poyen Ruka             — Rural Araucanía 2025
    # "P164": "TODO",  # Conún Huenu            — Rural Araucanía 2026
    # "P166": "TODO",  # Malalhue               — Rural Araucanía 2026
    # "P167": "TODO",  # Witran Donguil         — Rural Araucanía 2026
    # "P168": "TODO",  # Raíces de Perquenco    — Rural Araucanía 2026
    # "P170": "TODO",  # Los Copihues de Cunco  — Rural Araucanía 2026
    # "P171": "TODO",  # Raíces de Trovolhue    — Rural Araucanía 2026
    # "P172": "TODO",  # Raíces Costeras        — Rural Araucanía 2026
}

# ─────────────────────────────────────────────────────────────────────────────
# LOG
# ─────────────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger(__name__)


def get_credentials():
    return _ccu.get_credentials(SCOPES)


def _parse_fecha(s):
    """
    Parsea una fecha en formato DD/MM/YYYY, MM/DD/YYYY, DD-MM-YYYY, MM-DD-YYYY o YYYY-MM-DD.
    Retorna datetime o None.
    """
    s = str(s).strip()
    for fmt in ("%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y", "%m-%d-%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    return None


def _resolver_fechas(s1, s2, dias_val):
    """
    Intenta todos los formatos de fecha para encontrar la combinación
    que sea consistente con dias_val (TERMINO - INICIO ≈ dias_val).
    Retorna (inicio, termino, dias) o (None, None, None).
    """
    formatos = ["%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y", "%m-%d-%Y"]
    try:
        dias_val = int(float(str(dias_val)))
    except (ValueError, TypeError):
        return None, None, None

    for fmt1 in formatos:
        try:
            d1 = datetime.strptime(str(s1).strip(), fmt1)
        except ValueError:
            continue
        for fmt2 in formatos:
            try:
                d2 = datetime.strptime(str(s2).strip(), fmt2)
            except ValueError:
                continue
            if d2 > d1:
                diferencia = (d2 - d1).days
                if abs(diferencia - dias_val) <= 5:  # tolerancia 5 días
                    return d1, d2, dias_val

    return None, None, None


def buscar_resumen_en_hoja(sheets_svc, spreadsheet_id, sheet_name):
    """
    Busca la fila de resumen INICIO/TERMINO/DIAS en una hoja del Gantt.
    Escanea A1:Q12 buscando el header "FECHAS" y luego "INICIO | TERMINO | DIAS".
    Retorna (inicio_date, termino_date, dias_int) o None.
    """
    try:
        result = sheets_svc.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"'{sheet_name}'!A1:Q12",
            valueRenderOption="UNFORMATTED_VALUE",
            dateTimeRenderOption="FORMATTED_STRING",
        ).execute()
    except Exception as e:
        log.debug(f"    Error leyendo hoja '{sheet_name}': {e}")
        return None

    rows = result.get("values", [])
    if not rows:
        return None

    # Encontrar la fila con headers INICIO / TERMINO / DIAS
    header_row_idx = None
    inicio_col = termino_col = dias_col = None

    for ri, row in enumerate(rows):
        row_upper = [str(c).strip().upper() for c in row]
        has_inicio = any("INICIO" in c for c in row_upper)
        has_termino = any("TERMINO" in c or "TÉRMINO" in c for c in row_upper)
        has_dias = any(c == "DIAS" or c == "DÍAS" for c in row_upper)
        if has_inicio and has_termino:
            header_row_idx = ri
            # Encontrar columnas exactas
            for ci, c in enumerate(row_upper):
                if "INICIO" in c and inicio_col is None:
                    inicio_col = ci
                if ("TERMINO" in c or "TÉRMINO" in c) and termino_col is None:
                    termino_col = ci
                if (c == "DIAS" or c == "DÍAS") and dias_col is None:
                    dias_col = ci
            break

    if header_row_idx is None or inicio_col is None or termino_col is None:
        return None

    # Buscar la fila de resumen: las primeras 1-5 filas tras el header
    # que tengan 2 fechas válidas y un número entero (DIAS)
    for ri in range(header_row_idx + 1, min(header_row_idx + 6, len(rows))):
        row = rows[ri]
        if len(row) <= max(inicio_col, termino_col):
            continue

        s_inicio  = row[inicio_col]  if inicio_col  < len(row) else ""
        s_termino = row[termino_col] if termino_col < len(row) else ""
        s_dias    = row[dias_col]    if (dias_col is not None and dias_col < len(row)) else ""

        if not s_inicio or not s_termino:
            continue

        # Si no tenemos DIAS en la hoja, calcular después
        if s_dias == "":
            d1 = _parse_fecha(str(s_inicio))
            d2 = _parse_fecha(str(s_termino))
            if d1 and d2 and d2 > d1:
                dias_calc = (d2 - d1).days
                if 30 < dias_calc < 900:
                    return d1, d2, dias_calc
            continue

        # Intentar resolver formato con validación cruzada contra DIAS
        try:
            dias_val = int(float(str(s_dias)))
        except (ValueError, TypeError):
            continue

        if not (30 < dias_val < 900):
            continue

        d1, d2, dias = _resolver_fechas(s_inicio, s_termino, dias_val)
        if d1 and d2:
            return d1, d2, dias

    return None


def leer_programa_proyecto(sheets_svc, pid, spreadsheet_id):
    """
    Encuentra la fila de resumen en el Gantt del proyecto.
    Prueba todas las hojas y retorna el primer resultado válido.
    """
    # Obtener lista de hojas del spreadsheet
    try:
        meta = sheets_svc.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheet_names = [s["properties"]["title"] for s in meta["sheets"]]
    except Exception as e:
        log.error(f"  {pid}: error obteniendo hojas → {e}")
        return None

    # Orden preferido: primero las hojas con nombres típicos del Gantt principal
    preferred = ["Programa de obra", "% Avance"]
    ordered = [n for n in preferred if n in sheet_names]
    ordered += [n for n in sheet_names if n not in preferred]

    for sheet_name in ordered:
        log.debug(f"    {pid}: buscando en hoja '{sheet_name}'...")
        result = buscar_resumen_en_hoja(sheets_svc, spreadsheet_id, sheet_name)
        if result:
            d1, d2, dias = result
            log.info(
                f"  {pid}: [{sheet_name}] "
                f"inicio={d1.strftime('%Y-%m-%d')}  "
                f"finProg={d2.strftime('%Y-%m-%d')}  "
                f"plazo={dias}d"
            )
            return {
                "inicio":  d1.strftime("%Y-%m-%d"),
                "finProg": d2.strftime("%Y-%m-%d"),
                "plazo":   dias,
            }

    log.warning(f"  {pid}: no se encontró fila de resumen en ninguna hoja")
    return None


def main():
    log.info("=== actualizar_gantt_programa ===")
    creds = get_credentials()
    sheets_svc = build("sheets", "v4", credentials=creds)

    resultado = {}
    for pid, sid in PROYECTOS.items():
        log.info(f"Leyendo {pid}...")
        datos = leer_programa_proyecto(sheets_svc, pid, sid)
        if datos:
            resultado[pid] = datos

    if not resultado:
        log.error("Sin datos para ningún proyecto. Abortando escritura en Firebase.")
        sys.exit(1)

    log.info(f"Escribiendo {len(resultado)} proyectos en Firebase...")
    resp = requests.put(FIREBASE_URL, json=resultado, timeout=30)
    if resp.status_code == 200:
        log.info("Firebase actualizado OK.")
    else:
        log.error(f"Error Firebase: {resp.status_code} {resp.text[:200]}")
        sys.exit(1)

    log.info("=== Listo ===")


if __name__ == "__main__":
    main()
