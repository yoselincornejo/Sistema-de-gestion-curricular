"""
generador_word.py — Generador de programas desde cero
Construye el .docx completo a partir de la BD, sin leer ningún archivo original.
"""
import sqlite3
from pathlib import Path
from datetime import datetime

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

DB_PATH   = Path("data/sistema.db")
OUT_DIR   = Path("data/output")
LOGO_PATH = Path("data/logo_uv.jpg")

UV_BLUE  = RGBColor(0x1F, 0x4E, 0x79)
UV_GREEN = RGBColor(0x37, 0x56, 0x23)
UV_RED   = RGBColor(0x7B, 0x2C, 0x2C)

TIPO_COLOR = {"licenciatura": UV_BLUE, "titulo": UV_GREEN, "sello_uv": UV_RED}
TIPO_HEX   = {"licenciatura": "1F4E79", "titulo": "375623", "sello_uv": "843C0C"}

# ── Helpers de nivel ─────────────────────────────────────────────────

def nivel_desde_semestre(semestre):
    if semestre is None:
        return "N1"
    s = int(semestre)
    if s <= 4:
        return "N1"
    if s <= 8:
        return "N2"
    return "N3"

# ── BD ───────────────────────────────────────────────────────────────

def conectar():
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    return conn

def get_programa(asig_id):
    conn = conectar()
    asig = dict(conn.execute("""
        SELECT id, codigo, nombre, semestre, duracion, requisitos, version
        FROM asignaturas WHERE id=?""", (asig_id,)
    ).fetchone())
    unidades = [dict(r) for r in conn.execute(
        "SELECT * FROM unidades WHERE asignatura_id=? ORDER BY orden", (asig_id,)
    ).fetchall()]
    metodologias = [dict(r) for r in conn.execute(
        "SELECT * FROM metodologias WHERE asignatura_id=?", (asig_id,)
    ).fetchall()]
    evaluaciones = [dict(r) for r in conn.execute(
        "SELECT * FROM evaluaciones WHERE asignatura_id=? ORDER BY id", (asig_id,)
    ).fetchall()]
    conn.row_factory = None
    ras = conn.execute("""
        SELECT ra.codigo_completo, ra.descripcion, c.codigo, c.tipo
        FROM asignatura_ra ar
        JOIN resultados_aprendizaje ra ON ra.id = ar.ra_id
        JOIN competencias c ON c.id = ra.competencia_id
        WHERE ar.asignatura_id = ?
        ORDER BY c.codigo, ra.codigo_completo
    """, (asig_id,)).fetchall()
    conn.close()
    ras = [{"codigo_completo": r[0], "descripcion": r[1],
            "comp": r[2], "tipo": r[3]} for r in ras]
    return asig, unidades, metodologias, evaluaciones, ras

# ── Helpers de formato ───────────────────────────────────────────────

def _set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)

def _cell_para(cell, text, bold=False, size=10,
               color=None, align=WD_ALIGN_PARAGRAPH.LEFT,
               italic=False):
    p = cell.paragraphs[0]
    p.clear()
    p.alignment = align
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size)
    if color:
        run.font.color.rgb = color

def _section_heading(doc, text):
    """Párrafo de encabezado de sección con línea de fondo."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after  = Pt(4)
    run = p.add_run(text.upper())
    run.bold = True
    run.font.size = Pt(11)
    run.font.color.rgb = UV_BLUE
    return p

def _ident_row(table, label, value):
    """Agrega una fila label | value a la tabla de identificación."""
    row = table.add_row()
    _cell_para(row.cells[0], label, bold=True, size=10)
    _set_cell_bg(row.cells[0], "DEEAF1")
    _cell_para(row.cells[1], value or "", size=10)

# ── Generador principal ──────────────────────────────────────────────

def generar_programa_individual(asig_id, salida=None):
    asig, unidades, metodologias, evaluaciones, ras = get_programa(asig_id)
    codigo  = asig.get("codigo", "")
    nombre  = asig.get("nombre", "")
    semestre = asig.get("semestre")
    nivel   = nivel_desde_semestre(semestre)
    duracion   = asig.get("duracion") or ""
    requisitos = asig.get("requisitos") or "Sin requisito"

    doc = Document()

    # Márgenes
    for section in doc.sections:
        section.top_margin    = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin   = Cm(3)
        section.right_margin  = Cm(2.5)

    # ── Encabezado con logo ────────────────────────────────────────
    t_hdr = doc.add_table(rows=1, cols=2)
    t_hdr.style = "Table Grid"
    t_hdr.columns[0].width = Cm(4.5)

    c_logo = t_hdr.cell(0, 0)
    if LOGO_PATH.exists():
        p = c_logo.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(str(LOGO_PATH), width=Cm(4))
    else:
        _cell_para(c_logo, "UV", bold=True, size=14, align=WD_ALIGN_PARAGRAPH.CENTER)

    c_tit = t_hdr.cell(0, 1)
    _set_cell_bg(c_tit, "1F4E79")
    _cell_para(
        c_tit,
        f"PROGRAMA DE ASIGNATURA\n{nombre}\n{codigo}",
        bold=True, size=12,
        color=RGBColor(255, 255, 255),
        align=WD_ALIGN_PARAGRAPH.CENTER,
    )

    doc.add_paragraph()

    # ── Tabla de identificación ────────────────────────────────────
    _section_heading(doc, "Identificación")

    t_id = doc.add_table(rows=0, cols=2)
    t_id.style = "Table Grid"
    t_id.columns[0].width = Cm(4.5)

    _ident_row(t_id, "Código",    codigo)
    _ident_row(t_id, "Nombre",    nombre)
    _ident_row(t_id, "Semestre",  str(semestre) if semestre else "")
    _ident_row(t_id, "Nivel",     nivel)
    _ident_row(t_id, "Duración",  duracion)
    _ident_row(t_id, "Requisitos", requisitos)

    doc.add_paragraph()

    # ── Resultados de aprendizaje ──────────────────────────────────
    _section_heading(doc, "Resultados de Aprendizaje y Desempeños")

    if ras:
        for ra in ras:
            p = doc.add_paragraph(style="List Bullet")
            run = p.add_run(ra["codigo_completo"])
            run.bold = True
            run.font.size = Pt(10)
            comp_tipo = ra.get("tipo", "licenciatura")
            color = TIPO_COLOR.get(comp_tipo, UV_BLUE)
            run.font.color.rgb = color
            if ra.get("descripcion"):
                run2 = p.add_run(f": {ra['descripcion']}")
                run2.font.size = Pt(10)
    else:
        p = doc.add_paragraph()
        p.add_run("Sin resultados de aprendizaje declarados.").italic = True

    doc.add_paragraph()

    # ── Unidades y contenidos ──────────────────────────────────────
    _section_heading(doc, "Unidades de Aprendizaje y Contenidos")

    if unidades:
        t_uni = doc.add_table(rows=0, cols=3)
        t_uni.style = "Table Grid"

        # Header
        hrow = t_uni.add_row()
        for ci, (txt, w) in enumerate([
            ("N°", Cm(1.2)),
            ("Nombre de Unidad", Cm(5.5)),
            ("Contenidos", Cm(9)),
        ]):
            c = hrow.cells[ci]
            c.width = w
            _set_cell_bg(c, "1F4E79")
            _cell_para(c, txt, bold=True, size=10,
                       color=RGBColor(255, 255, 255),
                       align=WD_ALIGN_PARAGRAPH.CENTER)

        for u in unidades:
            row = t_uni.add_row()
            _cell_para(row.cells[0], str(u.get("orden", "")), size=10,
                       align=WD_ALIGN_PARAGRAPH.CENTER)
            nombre_u = u.get("nombre", "") or ""
            _cell_para(row.cells[1], nombre_u, size=9)
            _cell_para(row.cells[2], u.get("contenidos", "") or "", size=9)
    else:
        doc.add_paragraph("Sin unidades declaradas.").runs[0].italic = True

    doc.add_paragraph()

    # ── Estrategia metodológica ────────────────────────────────────
    _section_heading(doc, "Estrategia Metodológica")

    if metodologias:
        for m in metodologias:
            desc = m.get("descripcion", "") or ""
            for linea in desc.split("\n"):
                linea = linea.strip()
                if linea:
                    p = doc.add_paragraph(style="List Bullet")
                    p.add_run(linea).font.size = Pt(10)
    else:
        doc.add_paragraph("Sin metodologías declaradas.").runs[0].italic = True

    doc.add_paragraph()

    # ── Estrategia de evaluación ───────────────────────────────────
    _section_heading(doc, "Estrategia de Evaluación")

    evaluaciones_validas = [e for e in evaluaciones if (e.get("tipo") or "").strip()]
    if evaluaciones_validas:
        t_ev = doc.add_table(rows=0, cols=2)
        t_ev.style = "Table Grid"

        hrow = t_ev.add_row()
        for ci, (txt, w) in enumerate([
            ("Instrumento de Evaluación", Cm(12)),
            ("Ponderación", Cm(3.5)),
        ]):
            c = hrow.cells[ci]
            c.width = w
            _set_cell_bg(c, "1F4E79")
            _cell_para(c, txt, bold=True, size=10,
                       color=RGBColor(255, 255, 255),
                       align=WD_ALIGN_PARAGRAPH.CENTER)

        for ev in evaluaciones_validas:
            row = t_ev.add_row()
            _cell_para(row.cells[0], ev.get("tipo", ""), size=9)
            _cell_para(row.cells[1], ev.get("porcentaje", ""), size=9,
                       align=WD_ALIGN_PARAGRAPH.CENTER)
    else:
        doc.add_paragraph("Sin evaluaciones declaradas.").runs[0].italic = True

    doc.add_paragraph()

    # ── Pie de documento ──────────────────────────────────────────
    p_pie = doc.add_paragraph()
    p_pie.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_pie = p_pie.add_run(
        f"Instituto de Matemática — Universidad de Valparaíso · "
        f"Generado el {datetime.now().strftime('%d/%m/%Y')}"
    )
    run_pie.font.size = Pt(8)
    run_pie.font.color.rgb = RGBColor(0x94, 0xA3, 0xB8)
    run_pie.italic = True

    # ── Guardar ───────────────────────────────────────────────────
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    if salida is None:
        codigo_safe = codigo.replace(" ", "_")
        salida = str(OUT_DIR / f"programa_{codigo_safe}.docx")

    doc.save(salida)
    print(f"✓ Programa generado: {salida}")
    return salida


# ── Mapa de Progreso (sin cambios) ───────────────────────────────────

def generar_mapa_progreso(salida=None):
    """Genera el Mapa de Progreso en Word desde la BD."""
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = None

    competencias = conn.execute("""
        SELECT id, codigo, tipo, descripcion FROM competencias
        WHERE tipo != 'desconocido'
        ORDER BY
            CASE tipo
                WHEN 'licenciatura' THEN 0
                WHEN 'titulo'       THEN 1
                WHEN 'sello_uv'     THEN 2
            END, codigo
    """).fetchall()

    doc = Document()
    for section in doc.sections:
        section.top_margin    = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin   = Cm(2.5)
        section.right_margin  = Cm(2.5)

    def _set_bg(cell, hex_color):
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), hex_color)
        tcPr.append(shd)

    def _para(cell, text, bold=False, size=10, color=None,
              align=WD_ALIGN_PARAGRAPH.LEFT):
        cell.paragraphs[0].clear()
        p = cell.paragraphs[0]
        p.alignment = align
        run = p.add_run(text)
        run.bold = bold
        run.font.size = Pt(size)
        if color:
            run.font.color.rgb = color

    # Header
    t_hdr = doc.add_table(rows=1, cols=2)
    t_hdr.style = "Table Grid"
    c_logo = t_hdr.cell(0, 0)
    c_logo.width = Cm(5)
    if LOGO_PATH.exists():
        p = c_logo.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(str(LOGO_PATH), width=Cm(4.5))
    c_tit = t_hdr.cell(0, 1)
    _set_bg(c_tit, "1F4E79")
    _para(c_tit,
          "MAPA DE PROGRESO\nINGENIERÍA CIVIL MATEMÁTICA — PLAN 2025",
          bold=True, color=RGBColor(255, 255, 255), size=13,
          align=WD_ALIGN_PARAGRAPH.CENTER)

    doc.add_paragraph()

    TIPO_HDR = {
        "licenciatura": "COMPETENCIAS DE LICENCIATURA",
        "titulo":       "COMPETENCIAS ESPECÍFICAS DEL TÍTULO PROFESIONAL",
        "sello_uv":     "COMPETENCIAS GENÉRICAS SELLO UV",
    }

    tipo_actual = None
    for comp_id, comp_cod, comp_tipo, comp_desc in competencias:
        if comp_tipo != tipo_actual:
            tipo_actual = comp_tipo
            p = doc.add_paragraph()
            r = p.add_run(TIPO_HDR.get(tipo_actual, tipo_actual.upper()))
            r.bold = True
            r.font.size = Pt(13)
            hx = TIPO_HEX.get(tipo_actual, "1F4E79")
            r.font.color.rgb = RGBColor(int(hx[:2], 16), int(hx[2:4], 16), int(hx[4:], 16))

        p2 = doc.add_paragraph()
        r2 = p2.add_run(f"{comp_cod}: {comp_desc or ''}")
        r2.bold = True
        r2.font.size = Pt(11)
        hx = TIPO_HEX.get(comp_tipo, "1F4E79")
        r2.font.color.rgb = RGBColor(int(hx[:2], 16), int(hx[2:4], 16), int(hx[4:], 16))

        ras_raw = conn.execute("""
            SELECT nivel_dominio, codigo_completo, descripcion
            FROM resultados_aprendizaje WHERE competencia_id=?
            ORDER BY COALESCE(nivel_dominio,''), codigo
        """, (comp_id,)).fetchall()

        niveles = {}
        for nivel, ccomp, desc in ras_raw:
            niveles.setdefault(nivel or "Sin nivel", []).append((ccomp, desc or ""))

        if not niveles:
            doc.add_paragraph(
                "   (Sin resultados de aprendizaje declarados)"
            ).runs[0].italic = True
            doc.add_paragraph()
            continue

        cols_niv = sorted(niveles.keys())
        t = doc.add_table(rows=0, cols=len(cols_niv))
        t.style = "Table Grid"

        row_h = t.add_row()
        hx = TIPO_HEX.get(comp_tipo, "1F4E79")
        for ci, niv in enumerate(cols_niv):
            c = row_h.cells[ci]
            _set_bg(c, hx)
            _para(c, niv, bold=True, color=RGBColor(255, 255, 255),
                  align=WD_ALIGN_PARAGRAPH.CENTER)

        max_n = max(len(v) for v in niveles.values())
        for ri in range(max_n):
            row_d = t.add_row()
            for ci, niv in enumerate(cols_niv):
                items = niveles[niv]
                if ri < len(items):
                    ccomp, desc = items[ri]
                    txt = f"• {ccomp}"
                    if desc:
                        txt += f": {desc}"
                    _para(row_d.cells[ci], txt, size=9)

        row_a = t.add_row()
        bg_claro = {
            "licenciatura": "DEEAF1",
            "titulo":       "E2EFDA",
            "sello_uv":     "FCE4D6",
        }.get(comp_tipo, "F2F2F2")
        for ci, niv in enumerate(cols_niv):
            ra_codes = [r[0] for r in niveles[niv]]
            if ra_codes:
                placeholders = ",".join("?" * len(ra_codes))
                asigs_niv = conn.execute(f"""
                    SELECT DISTINCT a.codigo, a.nombre FROM asignaturas a
                    JOIN asignatura_ra ar ON ar.asignatura_id = a.id
                    JOIN resultados_aprendizaje ra ON ra.id = ar.ra_id
                    WHERE ra.codigo_completo IN ({placeholders})
                    ORDER BY a.semestre, a.codigo
                """, ra_codes).fetchall()
            else:
                asigs_niv = []
            c = row_a.cells[ci]
            _set_bg(c, bg_claro)
            txt = "\n".join(f"• {a[0]} {a[1]}" for a in asigs_niv) or "(ninguna)"
            _para(c, txt, size=8)

        doc.add_paragraph()

    conn.close()
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    if salida is None:
        salida = str(OUT_DIR / f"mapa_progreso_{datetime.now().strftime('%Y%m%d_%H%M')}.docx")
    doc.save(salida)
    print(f"✓ Mapa de Progreso generado: {salida}")
    return salida


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1 and sys.argv[1] == "mapa":
        generar_mapa_progreso()
    elif len(sys.argv) > 1:
        try:
            asig_id = int(sys.argv[1])
            generar_programa_individual(asig_id)
        except ValueError:
            print("Uso: python3 src/generador_word.py <asig_id>  |  mapa")
    else:
        print("Uso: python3 src/generador_word.py <asig_id>  |  mapa")
