"""
Exporta a Excel el detalle de gastos de pensión, alimentación y alojamiento.

Fuentes revisadas:
  1. Reg_pago_ex + Reg_pago_ex_det + Reg_pago_proveed  (facturas / pago directo)
  2. rend_caja + rend_caja_det                          (rendiciones / reembolsos)
  3. Solpago                                            (familia_pago = Viatico)

Filtro: búsqueda de keywords en todos los campos de texto relevantes.
IVA:    se calcula columna adicional asumiendo que el monto almacenado es NETO (sin IVA).
        Si el monto ya incluye IVA, ignorar la columna "Monto_con_IVA_19pct".
"""

import sys
import os
import pandas as pd
from datetime import datetime

sys.path.insert(0, os.path.dirname(__file__))
from sheets_connection import SheetsConnection

# ─── Configuración ───────────────────────────────────────────────────────────

KEYWORDS = [
    'pension', 'pensión',
    'alojamiento', 'hospedaje', 'hospederia', 'hostal', 'hotel', 'residencial', 'posada',
    'alimentacion', 'alimentación', 'colacion', 'colación',
    'comida', 'desayuno', 'almuerzo', 'cena',
    'viatico', 'viático',
    'lunch',
]

IVA = 0.19

# Directorio de salida: OneDrive local (Windows) o carpeta output/ del proyecto (Linux/remoto)
_WINDOWS_OUTPUT_DIR = r'C:\Users\rodri\OneDrive\Documentos Claude Code'
_FALLBACK_OUTPUT_DIR = os.path.join(os.path.dirname(__file__), '..', 'output')

_OUTPUT_DIR = _WINDOWS_OUTPUT_DIR if os.path.isdir(_WINDOWS_OUTPUT_DIR) else _FALLBACK_OUTPUT_DIR
os.makedirs(_OUTPUT_DIR, exist_ok=True)

OUTPUT_FILE = os.path.join(_OUTPUT_DIR, f'Pension_Alojamiento_Alimentacion_{datetime.now():%Y%m%d_%H%M}.xlsx')

# ─── Helpers ─────────────────────────────────────────────────────────────────

def contiene_keyword(texto):
    if pd.isna(texto):
        return False
    t = str(texto).lower()
    return any(kw in t for kw in KEYWORDS)


def parse_monto(serie):
    """Convierte montos en formato chileno (1.234,56 o 1234) a float."""
    return (
        serie.astype(str)
             .str.strip()
             .str.replace(r'[\$\s]', '', regex=True)
             .str.replace('.', '', regex=False)
             .str.replace(',', '.', regex=False)
             .pipe(pd.to_numeric, errors='coerce')
             .fillna(0)
    )


def detectar_campo_ndoc(df):
    """Busca el nombre de columna más probable para número de documento."""
    candidatos = ['N_doc', 'num_doc', 'N_factura', 'N_boleta', 'folio',
                  'numero_doc', 'N_Fact', 'ndoc', 'n_documento',
                  'num_factura', 'factura', 'boleta', 'N_Boleta', 'N_Folio']
    for c in candidatos:
        for col in df.columns:
            if col.lower() == c.lower():
                return col
    for col in df.columns:
        cl = col.lower()
        if any(x in cl for x in ['ndoc', 'n_doc', 'folio', 'factura', 'boleta']):
            return col
    return None


def mask_cualquier_campo(df, campos):
    """Devuelve máscara booleana: True si alguno de los campos contiene keyword."""
    mask = pd.Series(False, index=df.index)
    for c in campos:
        if c in df.columns:
            mask |= df[c].apply(contiene_keyword)
    return mask


def agregar_iva(df, col_monto='Monto_Neto'):
    df['Monto_con_IVA_19pct'] = (df[col_monto] * (1 + IVA)).round(0).astype(int)
    return df


# ─── Carga de datos ───────────────────────────────────────────────────────────

print("Conectando a Google Sheets...")
conn = SheetsConnection()

print("Cargando tablas (puede tomar 1-2 min)...")

tablas = {}
nombres = ['Reg_pago_ex', 'Reg_pago_ex_det', 'Reg_pago_proveed',
           'rend_caja', 'rend_caja_det', 'Solpago', 'Proyectos']

for nombre in nombres:
    try:
        tablas[nombre] = conn.get_sheet_data(nombre)
        print(f"  OK  {nombre}: {len(tablas[nombre])} filas, {len(tablas[nombre].columns)} cols")
    except Exception as e:
        print(f"  WARN  {nombre}: {e}")
        tablas[nombre] = pd.DataFrame()

df_reg_ex   = tablas['Reg_pago_ex']
df_reg_det  = tablas['Reg_pago_ex_det']
df_proveed  = tablas['Reg_pago_proveed']
df_rc       = tablas['rend_caja']
df_rc_det   = tablas['rend_caja_det']
df_sol      = tablas['Solpago']
df_proy     = tablas['Proyectos']

# ─── FUENTE 1: Reg_pago_ex / Reg_pago_ex_det ─────────────────────────────────

print("\nProcesando Gastos Generales (facturas)...")

filas_gg = pd.DataFrame()

if not df_reg_det.empty and not df_reg_ex.empty:
    campos_busqueda_det = ['Tipo_pago', 'detalle_pago', 'observaciones']
    campos_busqueda_det += [c for c in df_reg_det.columns
                            if c not in campos_busqueda_det and df_reg_det[c].dtype == object]

    campos_busqueda_cab = ['descripcion', 'estado']
    campos_busqueda_cab += [c for c in df_reg_ex.columns
                            if c not in campos_busqueda_cab and df_reg_ex[c].dtype == object]

    mask_det = mask_cualquier_campo(df_reg_det, campos_busqueda_det)
    det_filtrado = df_reg_det[mask_det].copy()

    if not det_filtrado.empty:
        # Join con cabecera
        joined = det_filtrado.merge(
            df_reg_ex[['IDU_regpx'] + [c for c in df_reg_ex.columns if c != 'IDU_regpx']],
            on='IDU_regpx', how='left'
        )

        # Join con proveedores
        ndoc_prov = detectar_campo_ndoc(df_proveed)
        prov_cols = ['IDU_proveed', 'Nombre', 'rut']
        if ndoc_prov:
            prov_cols.append(ndoc_prov)

        if not df_proveed.empty:
            joined = joined.merge(
                df_proveed[prov_cols].rename(columns={'Nombre': 'Proveedor', 'rut': 'RUT_Proveedor'}),
                left_on='proveedor', right_on='IDU_proveed', how='left'
            )
        else:
            joined['Proveedor'] = ''
            joined['RUT_Proveedor'] = ''

        # Join con proyectos
        if not df_proy.empty and 'ID_proy' in joined.columns:
            joined = joined.merge(
                df_proy[['ID_proy', 'NOMBRE_PROYECTO']].rename(columns={'NOMBRE_PROYECTO': 'Proyecto'}),
                on='ID_proy', how='left'
            )

        # Detectar N° documento
        ndoc_col = detectar_campo_ndoc(joined)
        ndoc_vals = joined[ndoc_col] if ndoc_col else ''

        # Parsear monto
        col_monto = 'monto' if 'monto' in joined.columns else joined.columns[0]
        joined['Monto_Neto'] = parse_monto(joined[col_monto])

        filas_gg = pd.DataFrame({
            'Fuente':          'Gastos Generales (Factura)',
            'Proyecto':        joined.get('Proyecto', joined.get('ID_proy', '')),
            'Proveedor':       joined.get('Proveedor', ''),
            'RUT_Proveedor':   joined.get('RUT_Proveedor', ''),
            'Fecha':           joined.get('fecha', ''),
            'Categoria':       joined.get('Tipo_pago', ''),
            'Descripcion':     joined.get('descripcion', '').fillna('') if 'descripcion' in joined.columns else '',
            'Detalle_Pago':    joined.get('detalle_pago', '').fillna('') if 'detalle_pago' in joined.columns else '',
            'Observaciones':   joined.get('observaciones', '').fillna('') if 'observaciones' in joined.columns else '',
            'N_Documento':     ndoc_vals if isinstance(ndoc_vals, pd.Series) else '',
            'Estado':          joined.get('estado', ''),
            'Monto_Neto':      joined['Monto_Neto'],
        })
        agregar_iva(filas_gg)
        print(f"  Gastos GG filtrados: {len(filas_gg)} filas")
    else:
        print("  Sin resultados en Reg_pago_ex_det con los keywords dados")

    # También buscar en cabecera (el keyword puede estar en descripción del pago global)
    mask_cab = mask_cualquier_campo(df_reg_ex, campos_busqueda_cab)
    cab_extra = df_reg_ex[mask_cab].copy()
    ids_ya_incluidos = set(det_filtrado['IDU_regpx'].dropna().unique()) if not det_filtrado.empty else set()
    cab_extra = cab_extra[~cab_extra['IDU_regpx'].isin(ids_ya_incluidos)]
    if not cab_extra.empty:
        print(f"  + {len(cab_extra)} cabeceras adicionales con keywords (sin detalle filtrado previo)")


# ─── FUENTE 2: rend_caja / rend_caja_det ────────────────────────────────────

print("\nProcesando Rendiciones...")

filas_rend = pd.DataFrame()

if not df_rc_det.empty and not df_rc.empty:
    campos_det = ['clase', 'Detalle', 'tipo_doc']
    campos_det += [c for c in df_rc_det.columns
                   if c not in campos_det and df_rc_det[c].dtype == object]

    mask_rend = mask_cualquier_campo(df_rc_det, campos_det)
    det_rend = df_rc_det[mask_rend].copy()

    if not det_rend.empty:
        cab_cols = ['IDU_Rendcaja', 'usuario', 'Estado', 'Ccosto']
        cab_cols += [c for c in df_rc.columns if c not in cab_cols]
        joined_r = det_rend.merge(
            df_rc[cab_cols].rename(columns={'Ccosto': 'ID_proy'}),
            on='IDU_Rendcaja', how='left'
        )

        # Join con proyectos por centro de costo
        if not df_proy.empty and 'ID_proy' in joined_r.columns:
            joined_r = joined_r.merge(
                df_proy[['ID_proy', 'NOMBRE_PROYECTO']].rename(columns={'NOMBRE_PROYECTO': 'Proyecto'}),
                on='ID_proy', how='left'
            )

        ndoc_col_r = detectar_campo_ndoc(joined_r)
        ndoc_vals_r = joined_r[ndoc_col_r] if ndoc_col_r else ''

        col_monto_r = 'Monto' if 'Monto' in joined_r.columns else (
                      'monto' if 'monto' in joined_r.columns else joined_r.columns[0])
        joined_r['Monto_Neto'] = parse_monto(joined_r[col_monto_r])

        filas_rend = pd.DataFrame({
            'Fuente':          'Rendicion (Reembolso)',
            'Proyecto':        joined_r.get('Proyecto', joined_r.get('ID_proy', '')),
            'Proveedor':       joined_r.get('usuario', ''),
            'RUT_Proveedor':   '',
            'Fecha':           joined_r.get('fecha', ''),
            'Categoria':       joined_r.get('clase', ''),
            'Descripcion':     '',
            'Detalle_Pago':    joined_r.get('Detalle', '').fillna('') if 'Detalle' in joined_r.columns else '',
            'Observaciones':   joined_r.get('tipo_doc', '').fillna('') if 'tipo_doc' in joined_r.columns else '',
            'N_Documento':     ndoc_vals_r if isinstance(ndoc_vals_r, pd.Series) else '',
            'Estado':          joined_r.get('Estado', '').fillna('') if 'Estado' in joined_r.columns else '',
            'Monto_Neto':      joined_r['Monto_Neto'],
        })
        agregar_iva(filas_rend)
        print(f"  Rendiciones filtradas: {len(filas_rend)} filas")
    else:
        print("  Sin resultados en rend_caja_det con los keywords dados")


# ─── FUENTE 3: Solpago – Familia Viatico ────────────────────────────────────

print("\nProcesando Solpago (Familia Viatico)...")

filas_sol = pd.DataFrame()

if not df_sol.empty:
    campos_sol = ['Familia_pago', 'Tipo_pago', 'Observacion']
    campos_sol += [c for c in df_sol.columns
                   if c not in campos_sol and df_sol[c].dtype == object]

    mask_sol = mask_cualquier_campo(df_sol, campos_sol)
    sol_filtrado = df_sol[mask_sol].copy()

    if not sol_filtrado.empty:
        if not df_proy.empty:
            pk_proy = 'ID_proy' if 'ID_proy' in sol_filtrado.columns else (
                      'ID_Proy' if 'ID_Proy' in sol_filtrado.columns else None)
            if pk_proy:
                sol_filtrado = sol_filtrado.merge(
                    df_proy[['ID_proy', 'NOMBRE_PROYECTO']].rename(columns={'NOMBRE_PROYECTO': 'Proyecto'}),
                    left_on=pk_proy, right_on='ID_proy', how='left'
                )

        ndoc_col_s = detectar_campo_ndoc(sol_filtrado)
        ndoc_vals_s = sol_filtrado[ndoc_col_s] if ndoc_col_s else ''

        col_monto_s = 'monto' if 'monto' in sol_filtrado.columns else sol_filtrado.columns[0]
        sol_filtrado['Monto_Neto'] = parse_monto(sol_filtrado[col_monto_s])

        maestro_col = next((c for c in ['maestro', 'Maestro', 'ID_maestro'] if c in sol_filtrado.columns), None)

        filas_sol = pd.DataFrame({
            'Fuente':          'Solpago (Viatico MO)',
            'Proyecto':        sol_filtrado.get('Proyecto', sol_filtrado.get('ID_proy', sol_filtrado.get('ID_Proy', ''))),
            'Proveedor':       sol_filtrado[maestro_col] if maestro_col else '',
            'RUT_Proveedor':   '',
            'Fecha':           sol_filtrado.get('fecha', sol_filtrado.get('Fecha', '')),
            'Categoria':       sol_filtrado.get('Familia_pago', ''),
            'Descripcion':     sol_filtrado.get('Tipo_pago', '').fillna('') if 'Tipo_pago' in sol_filtrado.columns else '',
            'Detalle_Pago':    sol_filtrado.get('Observacion', '').fillna('') if 'Observacion' in sol_filtrado.columns else '',
            'Observaciones':   '',
            'N_Documento':     ndoc_vals_s if isinstance(ndoc_vals_s, pd.Series) else '',
            'Estado':          sol_filtrado.get('Estado', sol_filtrado.get('estado', '')),
            'Monto_Neto':      sol_filtrado['Monto_Neto'],
        })
        agregar_iva(filas_sol)
        print(f"  Solpago filtradas: {len(filas_sol)} filas")
    else:
        print("  Sin resultados en Solpago con los keywords dados")


# ─── Consolidado y resumen por proveedor ─────────────────────────────────────

print("\nConsolidando...")

df_total = pd.concat(
    [df for df in [filas_gg, filas_rend, filas_sol] if not df.empty],
    ignore_index=True
)

if df_total.empty:
    print("\nATENCION: No se encontraron registros con los keywords dados.")
    print("Generando Excel de diagnóstico con categorías disponibles...")

# Resumen por proveedor
if not df_total.empty:
    resumen = (
        df_total.groupby(['Fuente', 'Proveedor', 'RUT_Proveedor'])
                .agg(
                    Transacciones=('Monto_Neto', 'count'),
                    Monto_Neto_Total=('Monto_Neto', 'sum'),
                    Monto_con_IVA_Total=('Monto_con_IVA_19pct', 'sum'),
                )
                .reset_index()
                .sort_values('Monto_con_IVA_Total', ascending=False)
    )
    resumen['Monto_Neto_Total'] = resumen['Monto_Neto_Total'].round(0).astype(int)
    resumen['Monto_con_IVA_Total'] = resumen['Monto_con_IVA_Total'].round(0).astype(int)


# ─── Hoja diagnóstico: valores únicos de categorías ─────────────────────────

diag_rows = []

if not df_reg_det.empty and 'Tipo_pago' in df_reg_det.columns:
    for v in df_reg_det['Tipo_pago'].dropna().unique():
        diag_rows.append({'Tabla': 'Reg_pago_ex_det', 'Campo': 'Tipo_pago', 'Valor': v,
                          'Conteo': int((df_reg_det['Tipo_pago'] == v).sum()),
                          'Keyword_match': contiene_keyword(v)})

if not df_rc_det.empty and 'clase' in df_rc_det.columns:
    for v in df_rc_det['clase'].dropna().unique():
        diag_rows.append({'Tabla': 'rend_caja_det', 'Campo': 'clase', 'Valor': v,
                          'Conteo': int((df_rc_det['clase'] == v).sum()),
                          'Keyword_match': contiene_keyword(v)})

if not df_sol.empty and 'Familia_pago' in df_sol.columns:
    for v in df_sol['Familia_pago'].dropna().unique():
        diag_rows.append({'Tabla': 'Solpago', 'Campo': 'Familia_pago', 'Valor': v,
                          'Conteo': int((df_sol['Familia_pago'] == v).sum()),
                          'Keyword_match': contiene_keyword(v)})

df_diag = pd.DataFrame(diag_rows).sort_values(['Tabla', 'Campo', 'Keyword_match'],
                                               ascending=[True, True, False])


# ─── Exportar Excel ──────────────────────────────────────────────────────────

os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)

print(f"\nGenerando Excel: {OUTPUT_FILE}")

with pd.ExcelWriter(OUTPUT_FILE, engine='xlsxwriter') as writer:
    wb = writer.book

    # Formatos
    fmt_titulo  = wb.add_format({'bold': True, 'font_size': 12, 'bg_color': '#2E4053',
                                  'font_color': 'white', 'border': 1})
    fmt_header  = wb.add_format({'bold': True, 'bg_color': '#AEB6BF', 'border': 1,
                                  'text_wrap': True, 'align': 'center', 'valign': 'vcenter'})
    fmt_money   = wb.add_format({'num_format': '$ #,##0', 'border': 1})
    fmt_moneyh  = wb.add_format({'bold': True, 'num_format': '$ #,##0', 'bg_color': '#D5E8D4', 'border': 1})
    fmt_match   = wb.add_format({'bg_color': '#D5E8D4', 'border': 1})
    fmt_cell    = wb.add_format({'border': 1, 'text_wrap': False})
    fmt_warn    = wb.add_format({'bg_color': '#FFF2CC', 'border': 1, 'italic': True})
    fmt_nota    = wb.add_format({'italic': True, 'font_color': '#555555', 'font_size': 9})

    def escribir_hoja(df_hoja, nombre_hoja, titulo):
        if df_hoja.empty:
            ws = wb.add_worksheet(nombre_hoja)
            ws.write(0, 0, f'{titulo} — Sin datos con los keywords actuales', fmt_warn)
            return
        df_hoja.to_excel(writer, sheet_name=nombre_hoja, index=False, startrow=2)
        ws = writer.sheets[nombre_hoja]
        ws.merge_range(0, 0, 0, len(df_hoja.columns) - 1, titulo, fmt_titulo)
        ws.set_row(1, 5)
        for col_num, col_name in enumerate(df_hoja.columns):
            ws.write(2, col_num, col_name, fmt_header)
        # Formato monto
        for col_num, col_name in enumerate(df_hoja.columns):
            if 'Monto' in col_name:
                ws.set_column(col_num, col_num, 16, fmt_money)
            elif 'Fecha' in col_name:
                ws.set_column(col_num, col_num, 12)
            elif col_name in ('Proveedor', 'Proyecto', 'Detalle_Pago', 'Descripcion'):
                ws.set_column(col_num, col_num, 30)
            else:
                ws.set_column(col_num, col_num, 18)
        # Nota IVA
        fila_nota = len(df_hoja) + 4
        ws.write(fila_nota, 0,
                 'NOTA: Monto_Neto = valor almacenado en BD. '
                 'Monto_con_IVA_19pct = Monto_Neto × 1.19. '
                 'Verificar si la BD almacena montos netos o brutos.', fmt_nota)

    # ── Hoja 1: Detalle completo
    escribir_hoja(
        df_total,
        'Detalle_Completo',
        'Pensión / Alojamiento / Alimentación — Detalle completo (todas las fuentes)'
    )

    # ── Hoja 2: Gastos Generales (facturas)
    escribir_hoja(
        filas_gg,
        'GG_Facturas',
        'Gastos Generales — Facturas pago directo (Reg_pago_ex + Reg_pago_ex_det)'
    )

    # ── Hoja 3: Rendiciones
    escribir_hoja(
        filas_rend,
        'Rendiciones',
        'Rendiciones de Caja — Reembolsos (rend_caja + rend_caja_det)'
    )

    # ── Hoja 4: Solpago Viáticos
    escribir_hoja(
        filas_sol,
        'Solpago_Viaticos',
        'Solicitudes de Pago — Familia Viático/Pensión (Solpago)'
    )

    # ── Hoja 5: Resumen por proveedor
    if not df_total.empty:
        resumen.to_excel(writer, sheet_name='Resumen_Proveedor', index=False, startrow=2)
        ws_r = writer.sheets['Resumen_Proveedor']
        ws_r.merge_range(0, 0, 0, len(resumen.columns) - 1,
                         'Resumen por Proveedor — Pensión / Alojamiento / Alimentación', fmt_titulo)
        ws_r.set_row(1, 5)
        for col_num, col_name in enumerate(resumen.columns):
            ws_r.write(2, col_num, col_name, fmt_header)
            if 'Monto' in col_name or 'Total' in col_name:
                ws_r.set_column(col_num, col_num, 18, fmt_moneyh)
            else:
                ws_r.set_column(col_num, col_num, 22)

    # ── Hoja 6: Diagnóstico de categorías
    if not df_diag.empty:
        df_diag.to_excel(writer, sheet_name='Diagnostico_Categorias', index=False, startrow=2)
        ws_d = writer.sheets['Diagnostico_Categorias']
        ws_d.merge_range(0, 0, 0, len(df_diag.columns) - 1,
                         'Diagnóstico: todas las categorías disponibles en la BD (resaltadas = match con keywords)',
                         fmt_titulo)
        ws_d.set_row(1, 5)
        for col_num, col_name in enumerate(df_diag.columns):
            ws_d.write(2, col_num, col_name, fmt_header)
            ws_d.set_column(col_num, col_num, 20)
        # Resaltar filas con match
        for row_num, (_, row) in enumerate(df_diag.iterrows(), start=3):
            fmt_row = fmt_match if row['Keyword_match'] else fmt_cell
            for col_num, val in enumerate(row):
                ws_d.write(row_num, col_num, val, fmt_row)

print(f"\n{'=' * 60}")
print(f"Excel generado exitosamente:")
print(f"  {os.path.abspath(OUTPUT_FILE)}")
print(f"{'=' * 60}")
print(f"\nHojas incluidas:")
print(f"  1. Detalle_Completo   — {len(df_total)} filas totales")
print(f"  2. GG_Facturas        — {len(filas_gg)} filas")
print(f"  3. Rendiciones        — {len(filas_rend)} filas")
print(f"  4. Solpago_Viaticos   — {len(filas_sol)} filas")
print(f"  5. Resumen_Proveedor")
print(f"  6. Diagnostico_Categorias — {len(df_diag)} categorías únicas")
print(f"\nKeywords usados: {', '.join(KEYWORDS)}")
print(f"\nIMPORTANTE: Revisar hoja 'Diagnostico_Categorias' para ver TODAS las")
print(f"categorías de la BD y confirmar que el filtro está capturando lo correcto.")
