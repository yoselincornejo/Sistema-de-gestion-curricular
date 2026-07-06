"""
app.py — Sistema de Gestión Curricular ICM
panel serve src/app.py --show --autoreload
"""
import sqlite3, os, sys, io
import panel as pn
import param
from pathlib import Path

sys.path.insert(0, os.path.dirname(__file__))
from generador_excel import generar_matriz
from generador_word import generar_programa_individual, generar_mapa_progreso

pn.extension(notifications=True)
RUTA_DB     = Path("data/sistema.db")
RUTA_OUTPUT = Path("data/output")
RUTA_OUTPUT.mkdir(parents=True, exist_ok=True)

# ── Helpers ───────────────────────────────────────────────────────

def nivel_desde_semestre(semestre):
    if semestre is None: return "N1"
    s = int(semestre)
    if s <= 4:  return "N1"
    if s <= 8:  return "N2"
    return "N3"

# ── BD ────────────────────────────────────────────────────────────

def conexion():
    conn = sqlite3.connect(str(RUTA_DB))
    conn.execute("PRAGMA foreign_keys = ON")
    conn.row_factory = sqlite3.Row
    return conn

def get_dashboard_data():
    conn = conexion()
    filas = conn.execute("""
        SELECT c.codigo, c.tipo, c.descripcion,
               COUNT(DISTINCT ar.asignatura_id) as n_asig
        FROM competencias c
        LEFT JOIN niveles_dominio nd ON nd.competencia_id = c.id
        LEFT JOIN resultados_aprendizaje ra ON ra.nivel_dominio_id = nd.id
        LEFT JOIN tributaciones ar ON ar.ra_id = ra.id
        WHERE c.tipo != 'desconocido'
        GROUP BY c.id
        ORDER BY
            CASE c.tipo
                WHEN 'licenciatura' THEN 0
                WHEN 'titulo'       THEN 1
                WHEN 'sello_uv'     THEN 2
            END, c.codigo
    """).fetchall()
    conn.close()
    return [dict(r) for r in filas]

def get_asignaturas_por_competencia(comp_codigo):
    conn = conexion()
    filas = conn.execute("""
        SELECT DISTINCT a.codigo, a.nombre, a.semestre
        FROM asignaturas a
        JOIN tributaciones ar ON ar.asignatura_id = a.id
        JOIN resultados_aprendizaje ra ON ra.id = ar.ra_id
        JOIN niveles_dominio nd ON nd.id = ra.nivel_dominio_id
        JOIN competencias c ON c.id = nd.competencia_id
        WHERE c.codigo = ?
        ORDER BY a.semestre, a.codigo
    """, (comp_codigo,)).fetchall()
    conn.close()
    return [dict(r) for r in filas]

def get_ras_con_asignaturas(comp_codigo):
    """Retorna lista de RAs de la competencia, cada uno con sus asignaturas tributantes."""
    conn = conexion()
    ras = conn.execute("""
        SELECT ra.id, ra.codigo_completo, ra.descripcion
        FROM resultados_aprendizaje ra
        JOIN niveles_dominio nd ON nd.id = ra.nivel_dominio_id
        JOIN competencias c ON c.id = nd.competencia_id
        WHERE c.codigo = ?
        ORDER BY ra.codigo_completo
    """, (comp_codigo,)).fetchall()
    resultado = []
    for ra in ras:
        asigs = conn.execute("""
            SELECT a.codigo, a.nombre, a.semestre
            FROM asignaturas a
            JOIN tributaciones ar ON ar.asignatura_id = a.id
            WHERE ar.ra_id = ?
            ORDER BY a.semestre, a.codigo
        """, (ra["id"],)).fetchall()
        resultado.append({
            "codigo_completo": ra["codigo_completo"],
            "descripcion":     ra["descripcion"] or "",
            "asignaturas":     [dict(a) for a in asigs],
        })
    conn.close()
    return resultado

def get_discrepancias():
    """
    Retorna dos listas para el panel de discrepancias del dashboard:
    - sugerencias: RAs con 0 tributaciones de tipo 'titulo', con asignaturas propuestas.
    - todas: todos los RAs con 0 tributaciones (cualquier tipo) + asignaturas sin tributaciones.
    """
    conn = conexion()

    # RAs sin tributaciones
    ras_vacios = conn.execute("""
        SELECT ra.codigo_completo, ra.descripcion, c.codigo as comp, c.tipo,
               nd.codigo_nivel,
               (SELECT COUNT(*) FROM tributaciones t WHERE t.ra_id = ra.id) as n_trib
        FROM resultados_aprendizaje ra
        JOIN niveles_dominio nd ON nd.id = ra.nivel_dominio_id
        JOIN competencias c ON c.id = nd.competencia_id
        WHERE (SELECT COUNT(*) FROM tributaciones t WHERE t.ra_id = ra.id) = 0
        ORDER BY c.tipo, c.codigo, nd.codigo_nivel, ra.codigo_ra
    """).fetchall()

    # Asignaturas sin ninguna tributación
    asigs_sin = conn.execute("""
        SELECT a.codigo, a.nombre
        FROM asignaturas a
        WHERE (SELECT COUNT(*) FROM tributaciones t WHERE t.asignatura_id = a.id) = 0
        ORDER BY a.codigo
    """).fetchall()

    conn.close()
    return [dict(r) for r in ras_vacios], [dict(r) for r in asigs_sin]

# Sugerencias basadas en equivalencia con plan ICM
_SUGERENCIAS_ICM = {
    "CE1, ND2, RA2": {
        "desc": "Plantea ecuaciones diferenciales (ordinarias o en derivadas parciales).",
        "asigs": ["MAT 222 — Ecuaciones Diferenciales Ordinarias",
                  "IMAT522 — Ecuaciones en Derivadas Parciales"],
        "ref":  "ICM 324, ICM 414",
    },
    "CE1, ND2, RA7": {
        "desc": "Maneja herramientas avanzadas de la Física, en contextos generales.",
        "asigs": ["IMAT424 — TIPE ICM"],
        "ref":  "ICM 425",
    },
    "CE2, ND1, RA1": {
        "desc": "Maneja software matemáticos, estadísticos y lenguaje de programación.",
        "asigs": ["PRO 111 — Fundamentos de Programación",
                  "PRO 121 — Programación",
                  "IMAT211 — Programación para Ingeniería"],
        "ref":  "ICM 114, ICM 124, ICM 214",
    },
    "CE2, ND1, RA3": {
        "desc": "Plantea problemas simplificados de ingeniería, matemáticamente.",
        "asigs": ["ING 111 — Desafíos de Ingeniería"],
        "ref":  "ICM 114",
    },
    "CE2, ND2, RA4": {
        "desc": "Implementa métodos numéricos, computacionalmente, para simulación.",
        "asigs": ["IMAT 312 — Métodos Numéricos",
                  "IMAT424 — TIPE ICM"],
        "ref":  "ICM 425",
    },
}


def get_asignaturas_lista():
    conn = conexion()
    filas = conn.execute(
        "SELECT id, codigo, nombre, semestre FROM asignaturas ORDER BY semestre, codigo"
    ).fetchall()
    conn.close()
    return [dict(r) for r in filas]

def get_programa_completo(asig_id):
    conn = conexion()
    asig = dict(conn.execute("""
        SELECT id, codigo, nombre, semestre, duracion, requisitos, version
        FROM asignaturas WHERE id=?""", (asig_id,)).fetchone())
    unidades = [dict(r) for r in conn.execute(
        "SELECT * FROM unidades WHERE asignatura_id=? ORDER BY orden",
        (asig_id,)).fetchall()]
    metodologias = [dict(r) for r in conn.execute(
        "SELECT * FROM metodologias WHERE asignatura_id=?",
        (asig_id,)).fetchall()]
    evaluaciones = [dict(r) for r in conn.execute(
        "SELECT * FROM evaluaciones WHERE asignatura_id=? ORDER BY id",
        (asig_id,)).fetchall()]
    conn.row_factory = None
    ras_raw = conn.execute("""
        SELECT ra.id, ra.codigo_completo, c.codigo, c.tipo
        FROM tributaciones ar
        JOIN resultados_aprendizaje ra ON ra.id = ar.ra_id
        JOIN niveles_dominio nd ON nd.id = ra.nivel_dominio_id
        JOIN competencias c ON c.id = nd.competencia_id
        WHERE ar.asignatura_id = ?
        ORDER BY c.codigo, ra.codigo_completo
    """, (asig_id,)).fetchall()
    todos_ras_raw = conn.execute("""
        SELECT ra.id, ra.codigo_completo, c.codigo, c.tipo
        FROM resultados_aprendizaje ra
        JOIN niveles_dominio nd ON nd.id = ra.nivel_dominio_id
        JOIN competencias c ON c.id = nd.competencia_id
        WHERE c.tipo != 'desconocido'
        ORDER BY
            CASE c.tipo
                WHEN 'licenciatura' THEN 0
                WHEN 'titulo'       THEN 1
                WHEN 'sello_uv'     THEN 2
                ELSE 3
            END, c.codigo, ra.codigo_completo
    """).fetchall()
    conn.close()
    ras      = [{"id": r[0], "codigo_completo": r[1], "comp": r[2], "tipo": r[3]}
                for r in ras_raw]
    todos_ras = [{"id": r[0], "codigo_completo": r[1], "comp": r[2], "tipo": r[3]}
                 for r in todos_ras_raw]
    return asig, unidades, metodologias, evaluaciones, ras, todos_ras

def guardar_programa(asig_id, datos):
    conn = conexion()
    try:
        conn.execute("""
            UPDATE asignaturas
            SET nombre=?, semestre=?, nivel=?, duracion=?, requisitos=?
            WHERE id=?
        """, (datos["nombre"], datos["semestre"], datos["nivel"],
              datos["duracion"], datos["requisitos"], asig_id))
        conn.execute("DELETE FROM unidades WHERE asignatura_id=?", (asig_id,))
        for i, u in enumerate(datos["unidades"]):
            conn.execute("""
                INSERT INTO unidades
                (asignatura_id, orden, nombre, contenidos, indicador_logro)
                VALUES (?,?,?,?,?)
            """, (asig_id, i+1, u["nombre"], u["contenidos"],
                  u.get("indicador_logro", "")))
        conn.execute("DELETE FROM metodologias WHERE asignatura_id=?", (asig_id,))
        for m in datos["metodologias"]:
            if m.strip():
                conn.execute(
                    "INSERT INTO metodologias (asignatura_id, descripcion) VALUES (?,?)",
                    (asig_id, m))
        conn.execute("DELETE FROM evaluaciones WHERE asignatura_id=?", (asig_id,))
        for ev in datos["evaluaciones"]:
            if ev["tipo"].strip():
                conn.execute(
                    "INSERT INTO evaluaciones (asignatura_id, tipo, porcentaje) VALUES (?,?,?)",
                    (asig_id, ev["tipo"], ev["porcentaje"]))
        conn.execute("DELETE FROM tributaciones WHERE asignatura_id=?", (asig_id,))
        for ra_id in datos["ra_ids"]:
            conn.execute(
                "INSERT OR IGNORE INTO tributaciones (asignatura_id, ra_id) VALUES (?,?)",
                (asig_id, ra_id))
        conn.commit()
        return True, "Cambios guardados correctamente"
    except Exception as e:
        conn.rollback()
        return False, str(e)
    finally:
        conn.close()

# ── CSS ───────────────────────────────────────────────────────────

CSS = """
:root {
    --uv-blue: #1F4E79; --uv-hover: #163A5C;
    --verde:   #375623; --rojo:    #7B2C2C;
    --bg:      #F4F7F9; --border:  #E2E8F0;
    --txt:     #1E293B; --muted:   #64748B;
    --orange:  #BA7517;
}
body { background:var(--bg); font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif; color:var(--txt); }
.bk-root { background:var(--bg) !important; }
.main-wrap { max-width:1400px; margin:0 auto; padding:0 32px; }

/* Header */
.app-header { background:white; padding:20px 28px; border-radius:12px;
    box-shadow:0 2px 8px rgba(0,0,0,0.06); display:flex; align-items:center;
    gap:16px; border-left:6px solid var(--uv-blue); margin-bottom:24px; }
.app-icon { width:44px; height:44px; background:var(--uv-blue); border-radius:8px;
    display:flex; align-items:center; justify-content:center;
    color:white; font-size:18px; font-weight:700; flex-shrink:0; }
.app-title { margin:0; font-size:20px; font-weight:700; }
.app-sub   { margin:3px 0 0; font-size:13px; color:var(--muted); }

/* Cards */
.card { background:white; border-radius:12px; padding:24px;
    box-shadow:0 2px 6px rgba(0,0,0,0.04); border:1px solid var(--border);
    margin-bottom:20px; }
.card-title { font-size:11px; font-weight:700; text-transform:uppercase;
    letter-spacing:1.2px; color:var(--uv-blue);
    border-bottom:2px solid #F1F5F9; padding-bottom:10px; margin-bottom:16px; }

/* Bloques dashboard */
.bloque { background:white; border-radius:12px; padding:22px;
    box-shadow:0 2px 6px rgba(0,0,0,0.04); border:1px solid var(--border);
    border-top:4px solid var(--uv-blue); flex:1; min-width:280px; }
.bloque-title { font-size:11px; font-weight:700; text-transform:uppercase;
    letter-spacing:1px; margin-bottom:14px; }

/* Fila de competencia clickeable */
.comp-row { display:flex; align-items:center; padding:10px 12px;
    border-radius:8px; margin-bottom:6px; border:1px solid #F1F5F9;
    background:#F8FAFC; transition:all 0.15s; }
.comp-row:hover { border-color:#CBD5E1; background:white;
    box-shadow:0 2px 6px rgba(0,0,0,0.06); }
.comp-code  { font-weight:700; font-size:14px; min-width:42px; }
.comp-desc  { font-size:12px; color:var(--muted); flex:1; margin:0 10px; line-height:1.4; }
.comp-badge { color:white; border-radius:20px; padding:3px 11px;
    font-size:12px; font-weight:700; white-space:nowrap; }

/* Popup asignaturas */
.popup-box { background:white; border-radius:12px; padding:20px;
    box-shadow:0 4px 20px rgba(0,0,0,0.10); border:1px solid var(--border);
    margin-top:12px; }
.popup-header { display:flex; justify-content:space-between; align-items:center;
    margin-bottom:14px; padding-bottom:10px; border-bottom:1px solid var(--border); }
.popup-title  { font-weight:700; font-size:16px; }
.popup-count  { font-size:12px; color:var(--muted); background:#F1F5F9;
    padding:3px 10px; border-radius:10px; font-weight:600; }
.asig-chip { display:inline-block; background:#F8FAFC; color:var(--txt);
    padding:5px 11px; border-radius:6px; margin:3px; font-size:12px;
    border:1px solid var(--border); }
.sem-tag   { color:var(--uv-blue); font-weight:700; margin-right:5px; }

/* Botones */
.btn-p { background:var(--uv-blue) !important; color:white !important;
    font-weight:600 !important; border-radius:8px !important; border:none !important; }
.btn-s { background:white !important; color:var(--uv-blue) !important;
    font-weight:600 !important; border-radius:8px !important;
    border:1.5px solid var(--uv-blue) !important; }
.btn-add { background:#FFFBEB !important; color:var(--orange) !important;
    border:1.5px dashed var(--orange) !important; border-radius:8px !important;
    font-weight:600 !important; }

/* Tributaciones checkboxes — checked state highlighted in green */
.bk-input-group input[type="checkbox"] {
    width: 16px; height: 16px; cursor: pointer;
    accent-color: #16a34a;
    margin-right: 5px;
}
.bk-input-group input[type="checkbox"]:checked + span {
    color: #16a34a !important;
    font-weight: 700 !important;
}
.bk-input-group input[type="checkbox"]:not(:checked) + span {
    color: #94a3b8;
}
"""

COLOR = {"licenciatura": "#1F4E79", "titulo": "#375623", "sello_uv": "#7B2C2C"}
LABEL = {"licenciatura": "Licenciatura", "titulo": "Título Profesional", "sello_uv": "Sello UV"}
DESC  = {
    "CL1": "Aplica ciencias e ingeniería",
    "CL2": "Valida soluciones ingenieriles",
    "CE1": "Modelos matemáticos avanzados",
    "CE2": "Implementa soluciones computacionales",
    "CG1": "Responsabilidad ética y profesional",
    "CG2": "Comunicación escrita",
    "CG3": "Comunicación oral",
    "CG4": "Comunicación en inglés",
}

# ── DASHBOARD ─────────────────────────────────────────────────────

def build_popup_html(comp_codigo, color):
    ras = get_ras_con_asignaturas(comp_codigo)
    if not ras:
        cuerpo = "<p style='color:#94A3B8;padding:8px 0'>Sin resultados de aprendizaje declarados.</p>"
        total_asigs = 0
    else:
        total_asigs = sum(len(r["asignaturas"]) for r in ras)
        bloques = []
        for ra in ras:
            sin_asigs = len(ra["asignaturas"]) == 0
            # Cabecera del RA
            if sin_asigs:
                ra_hdr = (
                    f'<div style="display:flex;align-items:center;gap:8px;'
                    f'margin-bottom:4px;">'
                    f'<span style="font-weight:700;font-size:12px;color:#DC2626">'
                    f'⚠ {ra["codigo_completo"]}</span>'
                    f'<span style="font-size:11px;color:#DC2626;background:#FEE2E2;'
                    f'padding:1px 8px;border-radius:10px;font-weight:600">sin cobertura</span>'
                    f'</div>'
                )
                if ra["descripcion"]:
                    ra_hdr += (
                        f'<div style="font-size:11px;color:#94A3B8;'
                        f'margin-bottom:6px;padding-left:4px">{ra["descripcion"]}</div>'
                    )
                chips = ""
                border = "#FCA5A5"
                bg = "#FFF5F5"
            else:
                ra_hdr = (
                    f'<div style="display:flex;align-items:center;gap:8px;'
                    f'margin-bottom:4px;">'
                    f'<span style="font-weight:700;font-size:12px;color:{color}">'
                    f'{ra["codigo_completo"]}</span>'
                    f'<span style="font-size:11px;color:#64748B;background:#F1F5F9;'
                    f'padding:1px 8px;border-radius:10px;font-weight:600">'
                    f'{len(ra["asignaturas"])} asignatura(s)</span>'
                    f'</div>'
                )
                if ra["descripcion"]:
                    ra_hdr += (
                        f'<div style="font-size:11px;color:#94A3B8;'
                        f'margin-bottom:6px;padding-left:4px">{ra["descripcion"]}</div>'
                    )
                chips = "".join(
                    f'<span class="asig-chip">'
                    f'<span class="sem-tag" style="color:{color}">S{a["semestre"]}</span>'
                    f'{a["codigo"]} — {a["nombre"]}</span>'
                    for a in ra["asignaturas"]
                )
                border = color
                bg = "#F8FAFC"

            bloques.append(
                f'<div style="margin-bottom:10px;padding:10px 12px;'
                f'background:{bg};border-radius:8px;border-left:3px solid {border};">'
                f'{ra_hdr}'
                f'<div>{chips}</div>'
                f'</div>'
            )
        cuerpo = "".join(bloques)

    sin_cobertura = sum(1 for r in ras if len(r["asignaturas"]) == 0)
    badge_extra = (
        f'<span style="font-size:11px;color:#DC2626;background:#FEE2E2;'
        f'padding:2px 9px;border-radius:10px;font-weight:600;margin-left:6px">'
        f'⚠ {sin_cobertura} RA sin cobertura</span>'
        if sin_cobertura else ""
    )

    return f"""
    <div style="margin-top:8px;padding:14px 16px;background:#F8FAFC;
                border-radius:10px;border-left:3px solid {color};">
        <div style="display:flex;justify-content:space-between;align-items:center;
                    margin-bottom:12px;flex-wrap:wrap;gap:6px;">
            <span style="font-weight:700;font-size:13px;color:{color}">
                {comp_codigo} — {DESC.get(comp_codigo,"")}
            </span>
            <div>
                <span style="font-size:11px;color:#64748B;background:white;
                             padding:2px 9px;border-radius:10px;font-weight:600;
                             border:1px solid #E2E8F0;">{total_asigs} asignatura(s)</span>
                {badge_extra}
            </div>
        </div>
        {cuerpo}
    </div>"""


class Dashboard(param.Parameterized):

    def _build_bloque(self, tipo, rows, popup_pane):
        """Construye un bloque de competencias con filas desplegables."""
        color = COLOR[tipo]
        label = LABEL[tipo]

        # Por cada competencia: fila con toggle + panel colapsable
        items = []
        for r in rows:
            codigo = r["codigo"]
            n_asig = r["n_asig"]

            # Panel de asignaturas (inicialmente oculto)
            asig_panel = pn.pane.HTML("", visible=False, sizing_mode="stretch_width")
            estado = {"abierto": False}

            # Botón toggle — toda la fila es clickeable
            toggle = pn.widgets.Toggle(
                name="",
                value=False,
                stylesheets=[f"""
                    :host {{ display:block; width:100%; margin-bottom:6px; }}
                    button {{
                        width:100% !important;
                        display:flex !important;
                        align-items:center !important;
                        padding:10px 14px !important;
                        border-radius:8px !important;
                        border:1px solid #F1F5F9 !important;
                        background:#F8FAFC !important;
                        cursor:pointer !important;
                        font-family:inherit !important;
                        font-size:13px !important;
                        color:#475569 !important;
                        transition:all 0.15s !important;
                        text-align:left !important;
                    }}
                    button.bk-active {{
                        border-color:{color} !important;
                        background:white !important;
                        box-shadow:0 2px 6px rgba(0,0,0,0.06) !important;
                    }}
                    button:hover {{
                        border-color:{color} !important;
                        background:white !important;
                    }}
                """],
                sizing_mode="stretch_width",
                height=44,
            )

            # Overlay HTML encima del toggle para el contenido visual
            fila_html = pn.pane.HTML(f"""
                <div style="display:flex;align-items:center;pointer-events:none;
                            padding:0 2px;position:relative;top:-44px;margin-bottom:-44px;">
                    <span style="font-weight:700;font-size:14px;color:{color};min-width:42px">{codigo}</span>
                    <span style="font-size:12px;color:#64748B;flex:1;margin:0 10px">{DESC.get(codigo,"")}</span>
                    <span style="background:{color};color:white;border-radius:20px;
                                 padding:3px 11px;font-size:12px;font-weight:700;margin-right:8px">{n_asig}</span>
                    <span style="color:{color};font-size:11px;font-weight:600">▶</span>
                </div>
            """, sizing_mode="stretch_width")

            def _on_toggle(event, c=codigo, col=color, panel=asig_panel):
                if event.new:
                    panel.object = build_popup_html(c, col)
                    panel.visible = True
                    popup_pane.object = ""  # cerrar popup global si hubiera
                else:
                    panel.visible = False
                    panel.object = ""

            toggle.param.watch(_on_toggle, "value")

            items.append(pn.Column(
                toggle,
                fila_html,
                asig_panel,
                sizing_mode="stretch_width",
                margin=(0,0,2,0)
            ))

        return pn.Column(
            pn.pane.HTML(f"""
                <div style="font-size:11px;font-weight:700;text-transform:uppercase;
                            letter-spacing:1px;color:{color};margin-bottom:12px;
                            padding-bottom:8px;border-bottom:2px solid #F1F5F9;">
                    {label}
                </div>
            """),
            *items,
            css_classes=["bloque"],
            sizing_mode="stretch_width"
        )

    def view(self):
        data   = get_dashboard_data()
        grupos = {}
        for r in data:
            grupos.setdefault(r["tipo"], []).append(r)

        popup_pane = pn.pane.HTML("", sizing_mode="stretch_width")

        bloques = []
        for tipo in ["licenciatura", "titulo", "sello_uv"]:
            if tipo not in grupos:
                continue
            bloques.append(self._build_bloque(tipo, grupos[tipo], popup_pane))

        # ── Botones generación ────────────────────────────────────
        status = pn.pane.HTML("", align="center")

        def _hacer_matriz():
            try:
                ruta = generar_matriz()
                with open(ruta, 'rb') as f:
                    return io.BytesIO(f.read())
            except Exception as e:
                status.object = f'<p style="color:#EF4444;text-align:center">✗ {e}</p>'
                return io.BytesIO(b"")

        def _hacer_mapa():
            try:
                ruta = generar_mapa_progreso()
                with open(ruta, 'rb') as f:
                    return io.BytesIO(f.read())
            except Exception as e:
                status.object = f'<p style="color:#EF4444;text-align:center">✗ {e}</p>'
                return io.BytesIO(b"")

        btn_excel = pn.widgets.FileDownload(
            callback=_hacer_matriz,
            filename="matriz_competencias.xlsx",
            label="📊 Descargar Matriz de Competencias",
            button_type="primary",
            width=320, height=48,
            embed=False,
        )
        btn_word = pn.widgets.FileDownload(
            callback=_hacer_mapa,
            filename="mapa_progreso.docx",
            label="🗺️ Descargar Mapa de Progreso",
            button_type="default",
            width=320, height=48,
            embed=False,
        )

        # ── Panel de discrepancias ────────────────────────────────
        discr_panel = self._build_discrepancias()

        return pn.Column(
            pn.Row(*bloques, sizing_mode="stretch_width"),
            popup_pane,
            discr_panel,
            pn.layout.Divider(),
            pn.Column(
                pn.pane.HTML('<div class="card-title" style="text-align:center">Documentos Oficiales</div>'),
                pn.Row(btn_excel, btn_word, align="center"),
                status,
                css_classes=["card"],
                sizing_mode="stretch_width",
                align="center"
            ),
            sizing_mode="stretch_width"
        )

    def _build_discrepancias(self):
        ras_vacios, asigs_sin = get_discrepancias()

        # ── Tabla 1: Sugerencias para RAs sin tributación (tipo título) ──
        filas_suger = ""
        for ra in ras_vacios:
            if ra["codigo_completo"] not in _SUGERENCIAS_ICM:
                continue
            s = _SUGERENCIAS_ICM[ra["codigo_completo"]]
            asigs_html = "".join(
                f'<span style="display:inline-block;background:#F0FDF4;color:#166534;'
                f'border:1px solid #BBF7D0;border-radius:5px;padding:2px 8px;'
                f'font-size:11px;margin:2px">{a}</span>'
                for a in s["asigs"]
            )
            filas_suger += f"""
            <tr>
              <td style="font-weight:700;color:#375623;white-space:nowrap;padding:8px 12px;
                         border-bottom:1px solid #F1F5F9">{ra["codigo_completo"]}</td>
              <td style="padding:8px 12px;border-bottom:1px solid #F1F5F9;
                         font-size:12px;color:#475569">{s["desc"]}</td>
              <td style="padding:8px 12px;border-bottom:1px solid #F1F5F9">{asigs_html}</td>
              <td style="padding:8px 12px;border-bottom:1px solid #F1F5F9;
                         font-size:11px;color:#94A3B8;white-space:nowrap">{s["ref"]}</td>
            </tr>"""

        tabla_suger = f"""
        <div style="margin-bottom:6px">
          <div class="card-title">Tributaciones faltantes sugeridas — referencia plan ICM</div>
          <p style="font-size:12px;color:#64748B;margin-bottom:12px">
            Los siguientes RAs no tienen ninguna asignatura tributante. Se propone vincularlos
            a las asignaturas indicadas, por equivalencia con el plan original ICM.
            Puedes confirmarlos desde <strong>Detalle de Asignatura → Tributación</strong>.
          </p>
          <table style="width:100%;border-collapse:collapse;font-size:13px">
            <thead>
              <tr style="background:#F8FAFC">
                <th style="text-align:left;padding:8px 12px;font-size:11px;font-weight:700;
                           text-transform:uppercase;letter-spacing:1px;color:#1F4E79;
                           border-bottom:2px solid #E2E8F0">RA</th>
                <th style="text-align:left;padding:8px 12px;font-size:11px;font-weight:700;
                           text-transform:uppercase;letter-spacing:1px;color:#1F4E79;
                           border-bottom:2px solid #E2E8F0">Descripción</th>
                <th style="text-align:left;padding:8px 12px;font-size:11px;font-weight:700;
                           text-transform:uppercase;letter-spacing:1px;color:#1F4E79;
                           border-bottom:2px solid #E2E8F0">Asignaturas propuestas</th>
                <th style="text-align:left;padding:8px 12px;font-size:11px;font-weight:700;
                           text-transform:uppercase;letter-spacing:1px;color:#1F4E79;
                           border-bottom:2px solid #E2E8F0">Ref. plan ICM</th>
              </tr>
            </thead>
            <tbody>{filas_suger}</tbody>
          </table>
        </div>"""

        # ── Tabla 2: Todas las discrepancias ──
        COLOR_TIPO = {"titulo": "#375623", "licenciatura": "#1F4E79", "sello_uv": "#7B2C2C"}
        BADGE_BG   = {"titulo": "#F0FDF4", "licenciatura": "#EFF6FF", "sello_uv": "#FFF1F2"}
        BADGE_BDR  = {"titulo": "#BBF7D0", "licenciatura": "#BFDBFE", "sello_uv": "#FECDD3"}

        filas_all = ""
        for ra in ras_vacios:
            tipo = ra["tipo"]
            color = COLOR_TIPO.get(tipo, "#475569")
            bg    = BADGE_BG.get(tipo, "#F8FAFC")
            bdr   = BADGE_BDR.get(tipo, "#E2E8F0")
            tipo_label = LABEL.get(tipo, tipo)
            filas_all += f"""
            <tr>
              <td style="padding:7px 12px;border-bottom:1px solid #F1F5F9">
                <span style="background:{bg};color:{color};border:1px solid {bdr};
                  border-radius:5px;padding:2px 8px;font-size:11px;font-weight:700;
                  white-space:nowrap">{tipo_label}</span>
              </td>
              <td style="font-weight:700;color:{color};padding:7px 12px;
                         border-bottom:1px solid #F1F5F9;white-space:nowrap;
                         font-size:13px">{ra["codigo_completo"]}</td>
              <td style="padding:7px 12px;border-bottom:1px solid #F1F5F9;
                         font-size:12px;color:#475569">{ra["descripcion"][:80]}{'…' if len(ra["descripcion"])>80 else ''}</td>
              <td style="padding:7px 12px;border-bottom:1px solid #F1F5F9;
                         text-align:center">
                <span style="background:#FEF2F2;color:#991B1B;border:1px solid #FECACA;
                  border-radius:10px;padding:2px 10px;font-size:11px;font-weight:700">0</span>
              </td>
              <td style="padding:7px 12px;border-bottom:1px solid #F1F5F9;
                         font-size:11px;color:#94A3B8">
                {"✅ Sugerencia disponible" if ra["codigo_completo"] in _SUGERENCIAS_ICM else "⚠ Sin sugerencia — revisar manualmente"}
              </td>
            </tr>"""

        for asig in asigs_sin:
            filas_all += f"""
            <tr>
              <td style="padding:7px 12px;border-bottom:1px solid #F1F5F9">
                <span style="background:#FFFBEB;color:#92400E;border:1px solid #FDE68A;
                  border-radius:5px;padding:2px 8px;font-size:11px;font-weight:700">
                  Asignatura</span>
              </td>
              <td style="font-weight:700;color:#92400E;padding:7px 12px;
                         border-bottom:1px solid #F1F5F9;font-size:13px">{asig["codigo"]}</td>
              <td style="padding:7px 12px;border-bottom:1px solid #F1F5F9;
                         font-size:12px;color:#475569">{asig["nombre"]}</td>
              <td style="padding:7px 12px;border-bottom:1px solid #F1F5F9;
                         text-align:center">
                <span style="background:#FEF2F2;color:#991B1B;border:1px solid #FECACA;
                  border-radius:10px;padding:2px 10px;font-size:11px;font-weight:700">0</span>
              </td>
              <td style="padding:7px 12px;border-bottom:1px solid #F1F5F9;
                         font-size:11px;color:#64748B">
                Sin tributaciones registradas (asignatura electiva / idioma)
              </td>
            </tr>"""

        tabla_all = f"""
        <div style="margin-top:20px">
          <div class="card-title">Resumen de discrepancias — tributaciones nulas</div>
          <table style="width:100%;border-collapse:collapse;font-size:13px">
            <thead>
              <tr style="background:#F8FAFC">
                <th style="text-align:left;padding:8px 12px;font-size:11px;font-weight:700;
                           text-transform:uppercase;letter-spacing:1px;color:#1F4E79;
                           border-bottom:2px solid #E2E8F0">Tipo</th>
                <th style="text-align:left;padding:8px 12px;font-size:11px;font-weight:700;
                           text-transform:uppercase;letter-spacing:1px;color:#1F4E79;
                           border-bottom:2px solid #E2E8F0">Código</th>
                <th style="text-align:left;padding:8px 12px;font-size:11px;font-weight:700;
                           text-transform:uppercase;letter-spacing:1px;color:#1F4E79;
                           border-bottom:2px solid #E2E8F0">Descripción</th>
                <th style="text-align:center;padding:8px 12px;font-size:11px;font-weight:700;
                           text-transform:uppercase;letter-spacing:1px;color:#1F4E79;
                           border-bottom:2px solid #E2E8F0">Tributaciones</th>
                <th style="text-align:left;padding:8px 12px;font-size:11px;font-weight:700;
                           text-transform:uppercase;letter-spacing:1px;color:#1F4E79;
                           border-bottom:2px solid #E2E8F0">Estado</th>
              </tr>
            </thead>
            <tbody>{filas_all}</tbody>
          </table>
        </div>"""

        return pn.Column(
            pn.pane.HTML(tabla_suger + tabla_all),
            css_classes=["card"],
            sizing_mode="stretch_width",
            margin=(16, 0, 0, 0),
        )


# ── EDITOR DE PROGRAMAS ───────────────────────────────────────────

class EditorProgramas(param.Parameterized):
    asignatura_id = param.Integer(default=0)

    def __init__(self, **params):
        super().__init__(**params)
        self._widgets          = {}
        self._unidades_widgets = []
        self._eval_widgets     = []
        self._ra_checks        = {}
        self._contenido_editor = pn.Column(sizing_mode="stretch_width")
        self._status           = pn.pane.HTML("")

    def _build_selector(self):
        asigs   = get_asignaturas_lista()
        opciones = {"— Elige una asignatura —": 0}
        for a in asigs:
            nivel = nivel_desde_semestre(a["semestre"])
            opciones[f"{a['codigo']} -- {a['nombre']}"] = a["id"]

        sel = pn.widgets.Select(
            name="Selecciona una asignatura",
            options=opciones,
            width=640, margin=(0,0,20,0)
        )

        def on_change(event):
            if event.new:
                self.asignatura_id = event.new
                self._cargar_editor()

        sel.param.watch(on_change, "value")
        return sel

    def _cargar_editor(self):
        if not self.asignatura_id:
            self._contenido_editor.objects = []
            return

        asig, unidades, metodologias, evaluaciones, ras_actuales, todos_ras = \
            get_programa_completo(self.asignatura_id)

        ra_ids_actuales = {r["id"] for r in ras_actuales}

        # Nivel derivado del semestre
        nivel_calc = nivel_desde_semestre(asig.get("semestre"))

        # ── Identificación ────────────────────────────────────────
        w_nombre = pn.widgets.TextInput(
            name="Nombre", value=asig.get("nombre",""), width=460)
        w_sem = pn.widgets.IntInput(
            name="Semestre", value=asig.get("semestre") or 1, width=110)
        w_nivel = pn.widgets.Select(
            name="Nivel", options=["N1","N2","N3"],
            value=nivel_calc, width=90)
        w_dur = pn.widgets.TextInput(
            name="Duración", value=asig.get("duracion",""), width=200)
        w_req = pn.widgets.TextInput(
            name="Requisitos", value=asig.get("requisitos",""), width=460)

        # El nivel se recalcula si cambia el semestre
        def on_sem_change(event):
            w_nivel.value = nivel_desde_semestre(event.new)
        w_sem.param.watch(on_sem_change, "value")

        sec_ident = pn.Column(
            pn.pane.HTML('<div class="card-title">Identificación</div>'),
            pn.Row(w_nombre, w_sem, w_nivel),
            pn.Row(w_dur, w_req),
            css_classes=["card"], sizing_mode="stretch_width"
        )

        # ── Tributación ───────────────────────────────────────────
        grupos_ra = {}
        for ra in todos_ras:
            grupos_ra.setdefault(ra["comp"], []).append(ra)

        ra_checks = {}
        ra_cols   = []
        for comp, ras in grupos_ra.items():
            tipo_comp = ras[0]["tipo"]
            color_comp = COLOR.get(tipo_comp, "#1F4E79")
            checks = []
            for ra in ras:
                cb = pn.widgets.Checkbox(
                    name=ra["codigo_completo"],
                    value=(ra["id"] in ra_ids_actuales)
                )
                ra_checks[ra["id"]] = cb
                checks.append(cb)
            ra_cols.append(pn.Column(
                pn.pane.HTML(
                    f'<div style="font-weight:700;color:{color_comp};'
                    f'font-size:13px;margin-bottom:6px">{comp}</div>'
                ),
                *checks, width=300
            ))

        self._ra_checks = ra_checks
        ra_filas = [pn.Row(*ra_cols[i:i+3]) for i in range(0, len(ra_cols), 3)]
        sec_ras = pn.Column(
            pn.pane.HTML('<div class="card-title">Tributación — Resultados de Aprendizaje</div>'),
            *ra_filas,
            css_classes=["card"], sizing_mode="stretch_width"
        )

        # ── Unidades ─────────────────────────────────────────────
        self._unidades_widgets = []
        col_unidades = pn.Column(sizing_mode="stretch_width")

        def crear_unidad(orden, nombre="", contenidos="", indicador=""):
            wn = pn.widgets.TextInput(
                name=f"Unidad {orden}", value=str(nombre)[:100], width=460)
            wc = pn.widgets.TextAreaInput(
                name="Contenidos", value=str(contenidos),
                height=80, sizing_mode="stretch_width")
            wi = pn.widgets.TextAreaInput(
                name="Indicador de logro", value=str(indicador),
                height=60, sizing_mode="stretch_width")
            panel = pn.Column(
                wn, wc, wi,
                css_classes=["card"], margin=(0,0,10,0),
                sizing_mode="stretch_width"
            )
            return {"nombre": wn, "contenidos": wc, "indicador_logro": wi, "panel": panel}

        for i, u in enumerate(unidades):
            ud = crear_unidad(
                u.get("orden", i+1), u.get("nombre",""),
                u.get("contenidos",""), u.get("indicador_logro",""))
            self._unidades_widgets.append(ud)
            col_unidades.append(ud["panel"])

        btn_add = pn.widgets.Button(
            name="➕ Añadir unidad", css_classes=["btn-add"], width=200, height=42)

        def on_add(event):
            ud = crear_unidad(len(self._unidades_widgets) + 1)
            self._unidades_widgets.append(ud)
            col_unidades.append(ud["panel"])

        btn_add.on_click(on_add)

        sec_unidades = pn.Column(
            pn.pane.HTML('<div class="card-title">Unidades y Contenidos</div>'),
            col_unidades,
            pn.Row(pn.layout.HSpacer(), btn_add, pn.layout.HSpacer()),
            css_classes=["card"], sizing_mode="stretch_width"
        )

        # ── Metodología ───────────────────────────────────────────
        w_metod = pn.widgets.TextAreaInput(
            name="Metodología",
            value="\n".join(m.get("descripcion","") for m in metodologias),
            height=100, sizing_mode="stretch_width")
        sec_metod = pn.Column(
            pn.pane.HTML('<div class="card-title">Estrategia Metodológica</div>'),
            w_metod, css_classes=["card"], sizing_mode="stretch_width"
        )

        # ── Evaluaciones ──────────────────────────────────────────
        self._eval_widgets = []
        for ev in evaluaciones:
            self._eval_widgets.append({
                "tipo":       pn.widgets.TextInput(
                    name="Instrumento", value=ev.get("tipo",""), width=360),
                "porcentaje": pn.widgets.TextInput(
                    name="Ponderación (%)", value=ev.get("porcentaje",""), width=130)
            })
        self._eval_widgets.append({
            "tipo":       pn.widgets.TextInput(
                name="Instrumento", placeholder="Nueva evaluación...", width=360),
            "porcentaje": pn.widgets.TextInput(
                name="Ponderación (%)", placeholder="0%", width=130)
        })
        sec_eval = pn.Column(
            pn.pane.HTML('<div class="card-title">Evaluaciones</div>'),
            *[pn.Row(e["tipo"], e["porcentaje"]) for e in self._eval_widgets],
            css_classes=["card"], sizing_mode="stretch_width"
        )

        # ── Acciones ──────────────────────────────────────────────
        self._widgets = {
            "nombre": w_nombre, "semestre": w_sem, "nivel": w_nivel,
            "duracion": w_dur, "requisitos": w_req, "metodologia": w_metod
        }

        btn_guardar = pn.widgets.Button(
            name="💾 Guardar cambios", css_classes=["btn-p"], width=270, height=48)

        def _hacer_word():
            try:
                ruta = generar_programa_individual(self.asignatura_id)
                with open(ruta, 'rb') as f:
                    return io.BytesIO(f.read())
            except Exception as e:
                self._status.object = f'<p style="color:#EF4444">✗ {e}</p>'
                return io.BytesIO(b"")

        codigo_asig = (asig.get("codigo", "asig") or "asig").replace(" ", "_")
        btn_word = pn.widgets.FileDownload(
            callback=_hacer_word,
            filename=f"programa_{codigo_asig}.docx",
            label="📥 Descargar Word",
            button_type="primary",
            width=220, height=48,
            embed=False,
        )

        def on_guardar(event):
            datos = {
                "nombre":       self._widgets["nombre"].value,
                "semestre":     self._widgets["semestre"].value,
                "nivel":        self._widgets["nivel"].value,
                "duracion":     self._widgets["duracion"].value,
                "requisitos":   self._widgets["requisitos"].value,
                "unidades":     [{"nombre": u["nombre"].value,
                                  "contenidos": u["contenidos"].value,
                                  "indicador_logro": u["indicador_logro"].value}
                                 for u in self._unidades_widgets],
                "metodologias": [self._widgets["metodologia"].value],
                "evaluaciones": [{"tipo": e["tipo"].value,
                                  "porcentaje": e["porcentaje"].value}
                                 for e in self._eval_widgets],
                "ra_ids": [rid for rid, cb in self._ra_checks.items() if cb.value]
            }
            ok, msg = guardar_programa(self.asignatura_id, datos)
            self._status.object = (
                f'<p style="color:#10B981;font-weight:600;font-size:14px">✓ {msg}</p>'
                if ok else
                f'<p style="color:#EF4444">✗ {msg}</p>'
            )

        btn_guardar.on_click(on_guardar)

        sec_acciones = pn.Column(
            pn.Row(btn_guardar, btn_word, align="center",
                   sizing_mode="stretch_width"),
            self._status,
            css_classes=["card"], sizing_mode="stretch_width",
            align="center", margin=(16,0,40,0)
        )

        self._contenido_editor.objects = [
            sec_ident, sec_ras,
            # sec_unidades, sec_metod, sec_eval,  # temporalmente ocultas
            sec_acciones
        ]

    def view(self):
        return pn.Column(
            self._build_selector(),
            self._contenido_editor,
            sizing_mode="stretch_width"
        )


# ── APP ───────────────────────────────────────────────────────────

def crear_app():
    pn.config.raw_css.append(CSS)

    header = pn.pane.HTML("""
        <div class="app-header">
            <div class="app-icon">ICM</div>
            <div>
                <h1 class="app-title">Sistema de Gestión Curricular</h1>
                <p class="app-sub">Instituto de Matemática · Universidad de Valparaíso · Plan 2025</p>
            </div>
        </div>
    """)

    tabs = pn.Tabs(
        ("📊 Dashboard de Cobertura", Dashboard().view()),
        ("✏️ Editor de Programas",    EditorProgramas().view()),
        dynamic=False
    )

    return pn.Column(
        header, tabs,
        sizing_mode="stretch_width",
        margin=(24, 40)
    )

app = crear_app()
app.servable()