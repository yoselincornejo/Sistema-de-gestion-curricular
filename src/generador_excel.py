"""
generador_excel.py — Matriz de Competencias plan 2025
Estrategia template-as-data: toma la plantilla original, borra las filas
de asignaturas y las reconstruye desde la BD. Todo el formato queda intacto.
"""
import sqlite3, argparse, copy
from datetime import datetime
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

DB_PATH       = Path("data/sistema.db")
OUT_DIR       = Path("data/output")
PLANTILLA_PATH = Path("data/plantilla_matriz.xlsx")

# Filas en la plantilla original
F_DATOS = 8    # primera fila de asignaturas
F_ULTIMA_ORIG = 62  # última fila de datos en el original (para borrar)

def conectar():
    if not DB_PATH.exists():
        raise FileNotFoundError(f"BD no encontrada: {DB_PATH}")
    conn = sqlite3.connect(str(DB_PATH)); conn.row_factory = None
    return conn

def obtener_datos():
    conn = conectar()
    # RAs válidos — solo los que existen en la plantilla original
    ras = conn.execute("""
        SELECT ra.id, ra.codigo_completo, c.codigo, c.tipo
        FROM resultados_aprendizaje ra
        JOIN competencias c ON c.id = ra.competencia_id
        WHERE c.tipo != 'desconocido'
        ORDER BY
            CASE c.tipo WHEN 'licenciatura' THEN 0 WHEN 'sello_uv' THEN 1 WHEN 'titulo' THEN 2 ELSE 3 END,
            c.codigo, COALESCE(ra.nivel_dominio,''), ra.codigo
    """).fetchall()
    # (ra_id, codigo_completo, comp_codigo, tipo)

    asigs = conn.execute("""
        SELECT id, codigo, nombre, semestre FROM asignaturas ORDER BY semestre, codigo
    """).fetchall()
    # (asig_id, codigo, nombre, semestre)

    tribs = set(conn.execute(
        "SELECT asignatura_id, ra_id FROM asignatura_ra"
    ).fetchall())

    conn.close()
    return ras, asigs, tribs

def nivel_academico(semestre):
    if semestre <= 4: return "N1"
    if semestre <= 8: return "N2"
    return "N3"

def _copy_cell_format(src_cell, dst_cell):
    """Copia formato completo de src a dst."""
    if src_cell.has_style:
        dst_cell.font      = copy.copy(src_cell.font)
        dst_cell.fill      = copy.copy(src_cell.fill)
        dst_cell.border    = copy.copy(src_cell.border)
        dst_cell.alignment = copy.copy(src_cell.alignment)
        dst_cell.number_format = src_cell.number_format

def _set_cell(ws, row, col, value, font=None, alignment=None):
    """Escribe en celda si no es MergedCell."""
    try:
        c = ws.cell(row=row, column=col)
        c.value = value
        if font:      c.font      = font
        if alignment: c.alignment = alignment
    except AttributeError:
        pass  # MergedCell — ignorar

def generar_matriz(salida=None):
    if not PLANTILLA_PATH.exists():
        raise FileNotFoundError(
            f"Plantilla no encontrada: {PLANTILLA_PATH}\n"
            f"Copia la plantilla original a data/plantilla_matriz.xlsx"
        )

    ras, asigs, tribs = obtener_datos()
    if not ras:
        print("⚠ Sin RAs válidos en la BD."); return

    # ── Cargar plantilla original ─────────────────────────────────
    wb = openpyxl.load_workbook(str(PLANTILLA_PATH))
    ws = wb['mapa (v2 2018)']

    # ── Mapear RA codigo_completo → columna desde la plantilla ────
    # La fila 6 tiene los códigos de RA como texto en cada columna
    ra_to_col = {}
    for col in range(1, ws.max_column + 1):
        val = ws.cell(6, col).value
        if val and isinstance(val, str):
            # Normalizar: quitar descripción larga, dejar solo el código
            # Ej: "CL1, N1, RA1: Relaciona..." → "CL1, N1, RA1"
            codigo = val.split(":")[0].strip()
            ra_to_col[codigo] = col

    # También mapear por código_completo exacto de la BD
    # Construir lookup normalizado
    def normalizar(s):
        return s.strip().replace(" ", "").upper()

    ra_col_norm = {normalizar(k): v for k, v in ra_to_col.items()}

    # Mapear cada RA de la BD a su columna en la plantilla
    ra_id_to_col = {}
    for ra_id, ccomp, comp, tipo in ras:
        norm = normalizar(ccomp)
        if norm in ra_col_norm:
            ra_id_to_col[ra_id] = ra_col_norm[norm]

    print(f"  RAs mapeados a columnas: {len(ra_id_to_col)}/{len(ras)}")

    # ── Limpiar filas de datos existentes ─────────────────────────
    # Deshacer TODOS los merges que toquen filas de datos
    merges_a_eliminar = [str(m) for m in ws.merged_cells.ranges
                         if m.max_row >= F_DATOS]
    for m in merges_a_eliminar:
        try: ws.unmerge_cells(m)
        except: pass

    # Borrar contenido y formato de filas de datos
    for row in range(F_DATOS, F_ULTIMA_ORIG + 10):  # margen extra
        for col in range(1, ws.max_column + 1):
            try:
                c = ws.cell(row, col)
                c.value = None
                c.font = Font(name="Calibri", size=12)
                c.fill = PatternFill(fill_type=None)
            except AttributeError:
                pass

    # ── Tomar fila plantilla de referencia (formato de fila 8 original)
    # Guardamos el formato de cada columna de la fila 8 original ANTES de borrar
    # → ya fue borrado, así que usamos valores conocidos del análisis
    font_x      = Font(name="Calibri", size=12, bold=True, color="FF0000")
    font_normal = Font(name="Calibri", size=12)
    font_sem    = Font(name="Calibri", size=18, bold=True)
    font_nivel  = Font(name="Calibri", size=36, bold=True)
    font_codigo = Font(name="Calibri", size=11, bold=True)
    font_nombre = Font(name="Calibri", size=11, bold=True)
    font_cred   = Font(name="Calibri", size=12, bold=True)
    fill_cred   = PatternFill("solid", fgColor="D8D8D8")

    align_center = Alignment(horizontal="center", vertical="center")
    align_left   = Alignment(horizontal="left",   vertical="center")
    align_rot90  = Alignment(horizontal="center", vertical="center", text_rotation=90)

    # ── Calcular agrupaciones por semestre ────────────────────────
    from collections import OrderedDict
    sem_filas = OrderedDict()
    for idx, (asig_id, codigo, nombre, semestre) in enumerate(asigs):
        fila = F_DATOS + idx
        sem_filas.setdefault(semestre, []).append(fila)

    # ── Columnas de COUNTIF por bloque (fiel al original) ─────────
    # En el original: G=comp.lic (L:AG), H=comp.gen (AI:BZ), I=comp.esp (CB:DC)
    # Para el plan 2025, calculamos los rangos reales desde ra_to_col
    col_lic_min = col_lic_max = col_gen_min = col_gen_max = col_esp_min = col_esp_max = None
    for ra_id, ccomp, comp, tipo in ras:
        col = ra_id_to_col.get(ra_id)
        if col is None: continue
        if tipo == "licenciatura":
            col_lic_min = min(col, col_lic_min or col)
            col_lic_max = max(col, col_lic_max or col)
        elif tipo == "sello_uv":
            col_gen_min = min(col, col_gen_min or col)
            col_gen_max = max(col, col_gen_max or col)
        elif tipo == "titulo":
            col_esp_min = min(col, col_esp_min or col)
            col_esp_max = max(col, col_esp_max or col)

    COL_TOT = 6   # F
    COL_LIC = 7   # G
    COL_GEN = 8   # H
    COL_ESP = 9   # I

    # Actualizar fila 7 (contadores por columna de RA)
    UF = F_DATOS + len(asigs) - 1
    for ra_id, ccomp, comp, tipo in ras:
        col = ra_id_to_col.get(ra_id)
        if col is None: continue
        cl = get_column_letter(col)
        try:
            ws.cell(7, col).value = f'=COUNTIF({cl}{F_DATOS}:{cl}{UF},"X")'
        except AttributeError:
            pass

    # ── Escribir filas de asignaturas ─────────────────────────────
    sem_actual = None
    for idx, (asig_id, codigo, nombre, semestre) in enumerate(asigs):
        fila = F_DATOS + idx

        # Altura de fila
        ws.row_dimensions[fila].height = 35.25

        # ── Col A: Nivel académico (merge vertical, fuente 36) ────
        if semestre != sem_actual:
            sem_actual = semestre
            filas_sem = sem_filas[semestre]
            f1, f2 = filas_sem[0], filas_sem[-1]

            if f1 != f2:
                ws.merge_cells(start_row=f1, start_column=1,
                               end_row=f2, end_column=1)
            ws.cell(f1, 1).value = nivel_academico(semestre)
            ws.cell(f1, 1).font = font_nivel
            ws.cell(f1, 1).alignment = align_rot90

            # ── Col B: Número semestre (merge vertical) ───────────
            if f1 != f2:
                ws.merge_cells(start_row=f1, start_column=2,
                               end_row=f2, end_column=2)
            ws.cell(f1, 2).value = semestre
            ws.cell(f1, 2).font = font_sem
            ws.cell(f1, 2).alignment = align_center

        # ── Col C: Código ─────────────────────────────────────────
        _set_cell(ws, fila, 3, codigo, font=font_codigo, alignment=align_center)

        # ── Col D: Nombre ─────────────────────────────────────────
        _set_cell(ws, fila, 4, nombre, font=font_nombre, alignment=align_left)

        # ── Col E: Créditos (fondo gris) ──────────────────────────
        try:
            c = ws.cell(fila, 5)
            c.value = None
            c.font = font_cred
            c.fill = fill_cred
        except AttributeError: pass

        # ── Col F: Total SUM ──────────────────────────────────────
        lc = get_column_letter(COL_LIC)
        le = get_column_letter(COL_ESP)
        _set_cell(ws, fila, COL_TOT,
                  f"=SUM({lc}{fila}:{le}{fila})",
                  font=font_normal, alignment=align_center)

        # ── Col G,H,I: COUNTIF por bloque ────────────────────────
        if col_lic_min and col_lic_max:
            rng = f"{get_column_letter(col_lic_min)}{fila}:{get_column_letter(col_lic_max)}{fila}"
            _set_cell(ws, fila, COL_LIC, f'=COUNTIF({rng},"X")',
                      font=font_normal, alignment=align_center)
        if col_gen_min and col_gen_max:
            rng = f"{get_column_letter(col_gen_min)}{fila}:{get_column_letter(col_gen_max)}{fila}"
            _set_cell(ws, fila, COL_GEN, f'=COUNTIF({rng},"X")',
                      font=font_normal, alignment=align_center)
        if col_esp_min and col_esp_max:
            rng = f"{get_column_letter(col_esp_min)}{fila}:{get_column_letter(col_esp_max)}{fila}"
            _set_cell(ws, fila, COL_ESP, f'=COUNTIF({rng},"X")',
                      font=font_normal, alignment=align_center)

        # ── X en columnas de RA ───────────────────────────────────
        for ra_id, ccomp, comp, tipo in ras:
            col = ra_id_to_col.get(ra_id)
            if col is None: continue
            if (asig_id, ra_id) in tribs:
                _set_cell(ws, fila, col, "X",
                          font=font_x, alignment=align_center)

    # ── Guardar ───────────────────────────────────────────────────
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    if salida is None:
        salida = str(OUT_DIR / f"matriz_competencias_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")
    wb.save(salida)
    print(f"✓ Matriz generada: {salida}")
    print(f"  {len(asigs)} asignaturas × {len(ra_id_to_col)} RAs mapeados")
    return salida

if __name__ == "__main__":
    p = argparse.ArgumentParser()
    p.add_argument("--salida", default=None)
    a = p.parse_args()
    generar_matriz(salida=a.salida)