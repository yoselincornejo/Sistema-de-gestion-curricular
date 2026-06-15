"""
generador_excel.py
Genera la Matriz de Competencias en formato Excel desde la base de datos.

Uso:
    python3 src/generador_excel.py
    python3 src/generador_excel.py --salida data/output/mi_matriz.xlsx
"""

import sqlite3
import argparse
import os
from datetime import datetime
from pathlib import Path

try:
    import openpyxl
    from openpyxl.styles import (
        PatternFill, Font, Alignment, Border, Side
    )
    from openpyxl.utils import get_column_letter
except ImportError:
    print("ERROR: openpyxl no instalado. Ejecuta: pip install openpyxl")
    raise

# ──────────────────────────────────────────────
# CONFIGURACIÓN DE COLORES Y ESTILOS
# ──────────────────────────────────────────────
COLORES = {
    "licenciatura":  "1F4E79",   # azul oscuro
    "titulo":        "375623",   # verde oscuro
    "sello_uv":      "7B2C2C",   # rojo oscuro
    "desconocido":   "808080",   # gris
    "header_fila":   "D6DCE4",   # gris azulado claro para filas de asignatura
    "marca_x":       "FFD966",   # amarillo para la X de tributación
    "blanco":        "FFFFFF",
    "fila_par":      "F2F2F2",
}

TIPOS_ORDEN = {"licenciatura": 0, "titulo": 1, "sello_uv": 2, "desconocido": 3}
TIPOS_ETIQUETA = {
    "licenciatura": "Competencias de Licenciatura",
    "titulo":       "Competencias Específicas del Título",
    "sello_uv":     "Competencias Genéricas Sello UV",
}

DB_PATH = "data/sistema.db"


def conectar():
    if not os.path.exists(DB_PATH):
        raise FileNotFoundError(f"No se encontró la BD en {DB_PATH}. "
                                f"Ejecuta desde la raíz del proyecto.")
    return sqlite3.connect(DB_PATH)


def obtener_columnas(cur):
    """Retorna lista de RAs válidos ordenados para las columnas de la matriz."""
    cur.execute("""
        SELECT ra.id, c.codigo AS comp, c.tipo, c.descripcion AS comp_desc,
               ra.nivel_dominio, ra.codigo AS ra_codigo,
               ra.codigo_completo, ra.descripcion
        FROM resultados_aprendizaje ra
        JOIN competencias c ON ra.competencia_id = c.id
        WHERE c.tipo != 'desconocido'
        ORDER BY
            CASE c.tipo
                WHEN 'licenciatura' THEN 0
                WHEN 'titulo'       THEN 1
                WHEN 'sello_uv'     THEN 2
                ELSE 3
            END,
            c.codigo,
            COALESCE(ra.nivel_dominio, ''),
            ra.codigo
    """)
    return cur.fetchall()


def obtener_asignaturas(cur):
    """Retorna asignaturas ordenadas por semestre y código."""
    cur.execute("""
        SELECT id, codigo, nombre, semestre
        FROM asignaturas
        ORDER BY semestre, codigo
    """)
    return cur.fetchall()


def obtener_tributaciones(cur):
    """Retorna set de (asignatura_id, ra_id) para lookup O(1)."""
    cur.execute("SELECT asignatura_id, ra_id FROM asignatura_ra")
    return set(cur.fetchall())


def fill(color_hex):
    return PatternFill("solid", fgColor=color_hex)


def font(bold=False, color="000000", size=10, italic=False):
    return Font(bold=bold, color=color, size=size, italic=italic,
                name="Calibri")


def border_thin():
    lado = Side(style="thin", color="AAAAAA")
    return Border(left=lado, right=lado, top=lado, bottom=lado)


def border_medium():
    lado = Side(style="medium", color="444444")
    return Border(left=lado, right=lado, top=lado, bottom=lado)


def centrado(wrap=False):
    return Alignment(horizontal="center", vertical="center", wrap_text=wrap)


def aplicar_celda(ws, row, col, value, fill_color=None, bold=False,
                  font_color="000000", font_size=10, align_center=True,
                  wrap=True, border=True, italic=False):
    cell = ws.cell(row=row, column=col, value=value)
    if fill_color:
        cell.fill = fill(fill_color)
    cell.font = font(bold=bold, color=font_color, size=font_size, italic=italic)
    if align_center:
        cell.alignment = centrado(wrap=wrap)
    else:
        cell.alignment = Alignment(vertical="center", wrap_text=wrap)
    if border:
        cell.border = border_thin()
    return cell


def generar_matriz(salida: str = None):
    conn = conectar()
    cur = conn.cursor()

    columnas = obtener_columnas(cur)      # lista de RAs
    asignaturas = obtener_asignaturas(cur)
    tributaciones = obtener_tributaciones(cur)

    conn.close()

    if not columnas:
        print("⚠ No hay RAs válidos en la BD. Verifica competencias.")
        return

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Matriz de Competencias"

    # ──────────────────────────────────────────
    # FILA 1: Título principal
    # ──────────────────────────────────────────
    COLS_FIJAS = 4   # semestre | código | asignatura | créditos
    N_COLS = COLS_FIJAS + len(columnas)

    ws.merge_cells(start_row=1, start_column=1,
                   end_row=1, end_column=N_COLS)
    titulo_cell = ws.cell(row=1, column=1,
                          value="MATRIZ DE COMPETENCIAS — INGENIERÍA CIVIL MATEMÁTICA — PLAN 2025")
    titulo_cell.font = font(bold=True, color="FFFFFF", size=13)
    titulo_cell.fill = fill("1A3A5C")
    titulo_cell.alignment = centrado(wrap=False)

    # ──────────────────────────────────────────
    # FILA 2: Grupos de tipo de competencia
    # ──────────────────────────────────────────
    # Calcular rangos por tipo
    grupos = {}
    for idx, (ra_id, comp, tipo, comp_desc, nivel, ra_cod, cod_completo, ra_desc) in enumerate(columnas):
        grupos.setdefault(tipo, []).append(COLS_FIJAS + 1 + idx)

    # Celdas fijas vacías fila 2
    for col in range(1, COLS_FIJAS + 1):
        aplicar_celda(ws, 2, col, "", fill_color="1A3A5C", border=True)

    for tipo, col_indices in grupos.items():
        col_ini = min(col_indices)
        col_fin = max(col_indices)
        if col_ini < col_fin:
            ws.merge_cells(start_row=2, start_column=col_ini,
                           end_row=2, end_column=col_fin)
        etiqueta = TIPOS_ETIQUETA.get(tipo, tipo.upper())
        color = COLORES.get(tipo, "808080")
        aplicar_celda(ws, 2, col_ini, etiqueta,
                      fill_color=color, bold=True,
                      font_color="FFFFFF", font_size=10)

    # ──────────────────────────────────────────
    # FILA 3: Grupos de competencia individual (CL1, CL2, CE1...)
    # ──────────────────────────────────────────
    grupos_comp = {}
    for idx, (ra_id, comp, tipo, comp_desc, nivel, ra_cod, cod_completo, ra_desc) in enumerate(columnas):
        key = (comp, tipo)
        grupos_comp.setdefault(key, []).append(COLS_FIJAS + 1 + idx)

    for col in range(1, COLS_FIJAS + 1):
        aplicar_celda(ws, 3, col, "", fill_color="D6DCE4", border=True)

    for (comp, tipo), col_indices in grupos_comp.items():
        col_ini = min(col_indices)
        col_fin = max(col_indices)
        if col_ini < col_fin:
            ws.merge_cells(start_row=3, start_column=col_ini,
                           end_row=3, end_column=col_fin)
        color = COLORES.get(tipo, "808080")
        # Color más claro para este nivel
        color_claro = _aclarar_color(color)
        aplicar_celda(ws, 3, col_ini, comp,
                      fill_color=color_claro, bold=True,
                      font_color=color, font_size=10)

    # ──────────────────────────────────────────
    # FILA 4: Nivel de dominio (N1, N2, N3...)
    # ──────────────────────────────────────────
    grupos_nivel = {}
    for idx, (ra_id, comp, tipo, comp_desc, nivel, ra_cod, cod_completo, ra_desc) in enumerate(columnas):
        nivel_key = (comp, tipo, nivel or "—")
        grupos_nivel.setdefault(nivel_key, []).append(COLS_FIJAS + 1 + idx)

    for col in range(1, COLS_FIJAS + 1):
        aplicar_celda(ws, 4, col, "", fill_color="EBF3FB", border=True)

    for (comp, tipo, nivel_key), col_indices in grupos_nivel.items():
        col_ini = min(col_indices)
        col_fin = max(col_indices)
        if col_ini < col_fin:
            ws.merge_cells(start_row=4, start_column=col_ini,
                           end_row=4, end_column=col_fin)
        color = COLORES.get(tipo, "808080")
        aplicar_celda(ws, 4, col_ini, nivel_key,
                      fill_color="EBF3FB", bold=False,
                      font_color=color, font_size=9)

    # ──────────────────────────────────────────
    # FILA 5: Código completo del RA (encabezado de columna)
    # ──────────────────────────────────────────
    HEADER_FIJOS = [("Sem.", 6), ("Código", 10), ("Asignatura", 32), ("Cred.", 6)]
    for col, (texto, ancho) in enumerate(HEADER_FIJOS, start=1):
        aplicar_celda(ws, 5, col, texto, fill_color="2E5E8E",
                      bold=True, font_color="FFFFFF", font_size=10)
        ws.column_dimensions[get_column_letter(col)].width = ancho

    for idx, (ra_id, comp, tipo, comp_desc, nivel, ra_cod, cod_completo, ra_desc) in enumerate(columnas):
        col = COLS_FIJAS + 1 + idx
        color = COLORES.get(tipo, "808080")
        color_claro = _aclarar_color(color)
        aplicar_celda(ws, 5, col, cod_completo,
                      fill_color=color_claro,
                      font_color=color, font_size=8,
                      bold=False, wrap=True)
        ws.column_dimensions[get_column_letter(col)].width = 14

    # Altura de la fila de headers de RA
    ws.row_dimensions[5].height = 80

    # ──────────────────────────────────────────
    # FILAS DE DATOS: una por asignatura
    # ──────────────────────────────────────────
    sem_actual = None
    for fila_idx, (asig_id, codigo, nombre, semestre) in enumerate(asignaturas):
        fila = 6 + fila_idx
        es_par = fila_idx % 2 == 0
        color_fila = "F7FAFD" if es_par else "FFFFFF"

        # Separador visual de semestre
        if semestre != sem_actual:
            sem_actual = semestre

        # Créditos (si están disponibles — campo no en BD actual, ponemos vacío)
        datos_fila = [semestre, codigo, nombre, ""]

        for col, valor in enumerate(datos_fila, start=1):
            bold = col == 3  # nombre en negrita
            aplicar_celda(ws, fila, col, valor,
                          fill_color=color_fila,
                          bold=bold, align_center=(col != 3),
                          wrap=(col == 3))

        # Columnas de RAs
        for idx, (ra_id, comp, tipo, comp_desc, nivel, ra_cod, cod_completo, ra_desc) in enumerate(columnas):
            col = COLS_FIJAS + 1 + idx
            tributa = (asig_id, ra_id) in tributaciones
            valor = "X" if tributa else ""
            color_celda = COLORES["marca_x"] if tributa else color_fila
            font_color = "7B4F00" if tributa else "AAAAAA"
            aplicar_celda(ws, fila, col, valor,
                          fill_color=color_celda,
                          bold=tributa, font_color=font_color,
                          font_size=10)

        ws.row_dimensions[fila].height = 28

    # ──────────────────────────────────────────
    # FILA FINAL: leyenda y metadatos
    # ──────────────────────────────────────────
    fila_leyenda = 6 + len(asignaturas) + 1
    ws.merge_cells(start_row=fila_leyenda, start_column=1,
                   end_row=fila_leyenda, end_column=N_COLS)
    leyenda = (f"Generado automáticamente desde BD — {datetime.now().strftime('%d/%m/%Y %H:%M')} "
               f"| {len(asignaturas)} asignaturas | {len(columnas)} RAs/Desempeños")
    cell = ws.cell(row=fila_leyenda, column=1, value=leyenda)
    cell.font = font(italic=True, color="666666", size=9)
    cell.fill = fill("F2F2F2")
    cell.alignment = Alignment(horizontal="left", vertical="center")

    # ──────────────────────────────────────────
    # FREEZE PANES y configuración de hoja
    # ──────────────────────────────────────────
    ws.freeze_panes = ws.cell(row=6, column=COLS_FIJAS + 1)
    ws.sheet_properties.tabColor = "1F4E79"

    # Fila 1 altura
    ws.row_dimensions[1].height = 30
    for r in [2, 3, 4]:
        ws.row_dimensions[r].height = 22
    ws.row_dimensions[5].height = 80

    # ──────────────────────────────────────────
    # GUARDAR
    # ──────────────────────────────────────────
    if salida is None:
        os.makedirs("data/output", exist_ok=True)
        fecha = datetime.now().strftime("%Y%m%d_%H%M")
        salida = f"data/output/matriz_competencias_{fecha}.xlsx"

    wb.save(salida)
    print(f"✓ Matriz generada: {salida}")
    print(f"  {len(asignaturas)} asignaturas × {len(columnas)} RAs/Desempeños")
    return salida


def _aclarar_color(hex_color: str) -> str:
    """Convierte un color hex oscuro en una versión muy clara (mezcla con blanco 85%)."""
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    factor = 0.82
    r2 = int(r + (255 - r) * factor)
    g2 = int(g + (255 - g) * factor)
    b2 = int(b + (255 - b) * factor)
    return f"{r2:02X}{g2:02X}{b2:02X}"


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Genera la Matriz de Competencias en Excel")
    parser.add_argument("--salida", default=None,
                        help="Ruta del archivo de salida (default: data/output/matriz_FECHA.xlsx)")
    args = parser.parse_args()
    generar_matriz(salida=args.salida)