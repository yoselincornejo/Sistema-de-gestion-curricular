"""
app.py — Sistema de Gestión Curricular ICM
Aplicación web local con Panel.

Ejecutar:
    panel serve src/app.py --show --autoreload
"""

import sqlite3
import os
import panel as pn
import param
from pathlib import Path
from datetime import datetime

pn.extension(sizing_mode="stretch_width", notifications=True)

RUTA_DB = Path("data/sistema.db")
RUTA_OUTPUT = Path("data/output")
RUTA_OUTPUT.mkdir(parents=True, exist_ok=True)

# ──────────────────────────────────────────────────────────────
# CAPA DE DATOS
# ──────────────────────────────────────────────────────────────

def conexion():
    conn = sqlite3.connect(str(RUTA_DB))
    conn.execute("PRAGMA foreign_keys = ON")
    conn.row_factory = sqlite3.Row
    return conn


def get_dashboard_data():
    """Retorna datos para el dashboard: competencias con conteo de asignaturas."""
    conn = conexion()
    filas = conn.execute("""
        SELECT c.codigo, c.tipo, c.descripcion,
               COUNT(DISTINCT ar.asignatura_id) as n_asig
        FROM competencias c
        LEFT JOIN resultados_aprendizaje ra ON ra.competencia_id = c.id
        LEFT JOIN asignatura_ra ar ON ar.ra_id = ra.id
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
    return filas


def get_asignaturas_por_competencia(comp_codigo):
    """Asignaturas que tributan a una competencia dada."""
    conn = conexion()
    filas = conn.execute("""
        SELECT DISTINCT a.codigo, a.nombre, a.semestre,
               GROUP_CONCAT(ra.codigo_completo, ' | ') as ras
        FROM asignaturas a
        JOIN asignatura_ra ar ON ar.asignatura_id = a.id
        JOIN resultados_aprendizaje ra ON ra.id = ar.ra_id
        JOIN competencias c ON c.id = ra.competencia_id
        WHERE c.codigo = ?
        GROUP BY a.id
        ORDER BY a.semestre, a.codigo
    """, (comp_codigo,)).fetchall()
    conn.close()
    return filas


def get_asignaturas_lista():
    """Lista simple de asignaturas para el selector."""
    conn = conexion()
    filas = conn.execute("""
        SELECT id, codigo, nombre, semestre
        FROM asignaturas ORDER BY semestre, codigo
    """).fetchall()
    conn.close()
    return filas


def get_programa_completo(asig_id):
    """Retorna todos los datos de un programa para el editor."""
    conn = conexion()
    asig = conn.execute(
        "SELECT * FROM asignaturas WHERE id = ?", (asig_id,)
    ).fetchone()
    unidades = conn.execute(
        "SELECT * FROM unidades WHERE asignatura_id = ? ORDER BY orden",
        (asig_id,)
    ).fetchall()
    metodologias = conn.execute(
        "SELECT * FROM metodologias WHERE asignatura_id = ?", (asig_id,)
    ).fetchall()
    evaluaciones = conn.execute(
        "SELECT * FROM evaluaciones WHERE asignatura_id = ? ORDER BY id",
        (asig_id,)
    ).fetchall()
    ras = conn.execute("""
        SELECT ra.id, ra.codigo_completo, c.codigo as comp, c.tipo
        FROM asignatura_ra ar
        JOIN resultados_aprendizaje ra ON ra.id = ar.ra_id
        JOIN competencias c ON c.id = ra.competencia_id
        ORDER BY c.codigo, ra.codigo_completo
    """).fetchall() if False else conn.execute("""
        SELECT ra.id, ra.codigo_completo, c.codigo as comp, c.tipo
        FROM asignatura_ra ar
        JOIN resultados_aprendizaje ra ON ra.id = ar.ra_id
        JOIN competencias c ON c.id = ra.competencia_id
        WHERE ar.asignatura_id = ?
        ORDER BY c.codigo, ra.codigo_completo
    """, (asig_id,)).fetchall()

    todos_ras = conn.execute("""
        SELECT ra.id, ra.codigo_completo, c.codigo, c.tipo
        FROM resultados_aprendizaje ra
        JOIN competencias c ON c.id = ra.competencia_id
        WHERE c.tipo != 'desconocido'
        ORDER BY c.codigo, ra.codigo_completo
    """).fetchall()

    conn.close()
    return dict(asig), list(unidades), list(metodologias), list(evaluaciones), list(ras), list(todos_ras)


def guardar_programa(asig_id, datos):
    """Persiste cambios de un programa en la BD."""
    conn = conexion()
    try:
        # Actualizar asignatura
        conn.execute("""
            UPDATE asignaturas SET
                nombre = ?, semestre = ?, nivel = ?,
                duracion = ?, requisitos = ?
            WHERE id = ?
        """, (
            datos["nombre"], datos["semestre"], datos["nivel"],
            datos["duracion"], datos["requisitos"], asig_id
        ))

        # Reemplazar unidades
        conn.execute("DELETE FROM unidades WHERE asignatura_id = ?", (asig_id,))
        for i, u in enumerate(datos["unidades"]):
            conn.execute("""
                INSERT INTO unidades (asignatura_id, orden, nombre, contenidos, indicador_logro)
                VALUES (?, ?, ?, ?, ?)
            """, (asig_id, i+1, u["nombre"], u["contenidos"], u.get("indicador_logro", "")))

        # Reemplazar metodologías
        conn.execute("DELETE FROM metodologias WHERE asignatura_id = ?", (asig_id,))
        for m in datos["metodologias"]:
            if m.strip():
                conn.execute(
                    "INSERT INTO metodologias (asignatura_id, descripcion) VALUES (?, ?)",
                    (asig_id, m)
                )

        # Reemplazar evaluaciones
        conn.execute("DELETE FROM evaluaciones WHERE asignatura_id = ?", (asig_id,))
        for ev in datos["evaluaciones"]:
            if ev["tipo"].strip():
                conn.execute(
                    "INSERT INTO evaluaciones (asignatura_id, tipo, porcentaje) VALUES (?, ?, ?)",
                    (asig_id, ev["tipo"], ev["porcentaje"])
                )

        # Reemplazar tributaciones
        conn.execute("DELETE FROM asignatura_ra WHERE asignatura_id = ?", (asig_id,))
        for ra_id in datos["ra_ids"]:
            conn.execute(
                "INSERT OR IGNORE INTO asignatura_ra (asignatura_id, ra_id) VALUES (?, ?)",
                (asig_id, ra_id)
            )

        conn.commit()
        return True, "Cambios guardados correctamente"
    except Exception as e:
        conn.rollback()
        return False, str(e)
    finally:
        conn.close()


# ──────────────────────────────────────────────────────────────
# ESTILOS CSS
# ──────────────────────────────────────────────────────────────

CSS = """
:host { font-family: 'Segoe UI', Arial, sans-serif; }

.app-header {
    background: linear-gradient(135deg, #1a3a5c 0%, #2e6da4 100%);
    color: white;
    padding: 16px 24px;
    border-radius: 8px;
    margin-bottom: 16px;
}
.app-header h1 {
    margin: 0;
    font-size: 20px;
    font-weight: 600;
    letter-spacing: 0.5px;
}
.app-header p {
    margin: 4px 0 0;
    font-size: 12px;
    opacity: 0.8;
}

.bloque-card {
    background: white;
    border-radius: 10px;
    padding: 16px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.08);
    border-top: 4px solid #2e6da4;
    height: 100%;
}
.bloque-titulo {
    font-size: 13px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 1px;
    margin-bottom: 12px;
    color: #1a3a5c;
}
.comp-fila {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 8px 10px;
    border-radius: 6px;
    margin-bottom: 6px;
    cursor: pointer;
    transition: background 0.15s;
    border: 1px solid #e8ecf0;
}
.comp-fila:hover { background: #f0f4f8; border-color: #b0c4d8; }
.comp-codigo {
    font-weight: 700;
    font-size: 13px;
    color: #1a3a5c;
    min-width: 40px;
}
.comp-badge {
    background: #2e6da4;
    color: white;
    border-radius: 12px;
    padding: 2px 10px;
    font-size: 12px;
    font-weight: 600;
}
.popup-overlay {
    background: white;
    border-radius: 10px;
    padding: 20px;
    box-shadow: 0 8px 32px rgba(0,0,0,0.15);
    border: 1px solid #dde3ea;
}
.asig-chip {
    display: inline-block;
    background: #e8f0fb;
    color: #1a3a5c;
    border-radius: 4px;
    padding: 3px 8px;
    margin: 3px;
    font-size: 12px;
}
.sem-label {
    font-size: 11px;
    color: #888;
    font-weight: 600;
    margin-right: 4px;
}
.editor-section {
    background: #f8fafc;
    border-radius: 8px;
    padding: 14px;
    margin-bottom: 12px;
    border: 1px solid #e2e8f0;
}
.editor-section-title {
    font-size: 12px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.8px;
    color: #4a6785;
    margin-bottom: 10px;
}
.ra-chip {
    display: inline-flex;
    align-items: center;
    background: #e8f4e8;
    border: 1px solid #b8d8b8;
    border-radius: 4px;
    padding: 3px 8px;
    margin: 3px;
    font-size: 11px;
    color: #2d5a2d;
}
.btn-generar {
    background: #1a3a5c !important;
    color: white !important;
    border-radius: 6px !important;
    font-weight: 600 !important;
}
"""

# ──────────────────────────────────────────────────────────────
# COLORES POR TIPO DE COMPETENCIA
# ──────────────────────────────────────────────────────────────

COLOR_TIPO = {
    "licenciatura": "#1F4E79",
    "titulo":       "#375623",
    "sello_uv":     "#7B2C2C",
}

LABEL_TIPO = {
    "licenciatura": "Licenciatura",
    "titulo":       "Título Profesional",
    "sello_uv":     "Sello UV",
}


# ──────────────────────────────────────────────────────────────
# PESTAÑA 1: DASHBOARD
# ──────────────────────────────────────────────────────────────

class Dashboard(param.Parameterized):
    competencia_seleccionada = param.String(default="")

    def __init__(self, **params):
        super().__init__(**params)
        self._popup_panel = pn.pane.HTML("", width=420)
        self._popup_visible = pn.Column(visible=False)

    def _comp_html(self, codigo, tipo, n_asig):
        color = COLOR_TIPO.get(tipo, "#666")
        return f"""
        <div class="comp-fila" onclick="
            document.dispatchEvent(new CustomEvent('comp-click', {{detail: '{codigo}'}}))
        ">
            <span class="comp-codigo" style="color:{color}">{codigo}</span>
            <span style="font-size:11px;color:#666;flex:1;margin-left:8px;line-height:1.3">
                {self._desc_corta(codigo)}
            </span>
            <span class="comp-badge" style="background:{color}">{n_asig}</span>
        </div>
        """

    def _desc_corta(self, codigo):
        descripciones = {
            "CL1": "Resuelve problemas usando ciencias e ingeniería",
            "CL2": "Valida soluciones ingenieriles con evidencias",
            "CE1": "Modelos matemáticos avanzados",
            "CE2": "Implementa soluciones computacionales",
            "CG1": "Responsabilidad ética y profesional",
            "CG2": "Comunicación escrita",
            "CG3": "Comunicación oral",
            "CG4": "Comunicación en inglés",
        }
        return descripciones.get(codigo, "")

    def construir_bloques(self):
        data = get_dashboard_data()
        grupos = {}
        for row in data:
            tipo = row["tipo"]
            grupos.setdefault(tipo, []).append(row)

        bloques = []
        for tipo in ["licenciatura", "titulo", "sello_uv"]:
            if tipo not in grupos:
                continue
            color = COLOR_TIPO[tipo]
            label = LABEL_TIPO[tipo]
            comps_html = "".join(
                self._comp_html(r["codigo"], r["tipo"], r["n_asig"])
                for r in grupos[tipo]
            )
            bloque = pn.pane.HTML(f"""
                <div class="bloque-card" style="border-top-color:{color}">
                    <div class="bloque-titulo" style="color:{color}">{label}</div>
                    {comps_html}
                </div>
            """, min_height=280)
            bloques.append(bloque)

        return bloques

    def construir_popup(self, comp_codigo):
        if not comp_codigo:
            return pn.pane.HTML("")
        asigs = get_asignaturas_por_competencia(comp_codigo)
        color = "#1a3a5c"
        for row in get_dashboard_data():
            if row["codigo"] == comp_codigo:
                color = COLOR_TIPO.get(row["tipo"], "#1a3a5c")
                break

        if not asigs:
            contenido = "<p style='color:#888;font-style:italic'>Sin asignaturas tributando</p>"
        else:
            sems = {}
            for a in asigs:
                sems.setdefault(a["semestre"], []).append(a)
            filas = ""
            for sem in sorted(sems.keys()):
                chips = "".join(
                    f'<span class="asig-chip"><span class="sem-label">S{sem}</span>'
                    f'<b>{a["codigo"]}</b> {a["nombre"]}</span>'
                    for a in sems[sem]
                )
                filas += f'<div style="margin-bottom:8px">{chips}</div>'
            contenido = filas

        return pn.pane.HTML(f"""
            <div class="popup-overlay">
                <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px">
                    <span style="font-weight:700;font-size:15px;color:{color}">{comp_codigo}</span>
                    <span style="font-size:12px;color:#888">{len(asigs)} asignatura(s)</span>
                </div>
                {contenido}
            </div>
        """, width=440)

    def view(self):
        bloques = self.construir_bloques()

        # Selector de competencia para popup
        data = get_dashboard_data()
        opciones = ["— seleccionar competencia —"] + [r["codigo"] for r in data]
        selector = pn.widgets.Select(
            options=opciones, value="— seleccionar competencia —",
            name="Ver asignaturas de:", width=240
        )
        popup_col = pn.Column(pn.pane.HTML(""))

        def on_select(event):
            val = event.new
            if val and val != "— seleccionar competencia —":
                popup_col.objects = [self.construir_popup(val)]
            else:
                popup_col.objects = [pn.pane.HTML("")]

        selector.param.watch(on_select, "value")

        # Botones de generación
        btn_excel = pn.widgets.Button(
            name="⬇ Generar Matriz de Competencias (.xlsx)",
            button_type="primary", width=300
        )
        btn_word = pn.widgets.Button(
            name="⬇ Generar Mapa de Progreso (.docx)",
            button_type="success", width=300
        )
        status = pn.pane.HTML("")

        def generar_excel(event):
            try:
                from generador_excel import generar_matriz
                ruta = generar_matriz()
                status.object = f'<p style="color:green;font-weight:600">✓ Generado: {ruta}</p>'
            except Exception as e:
                status.object = f'<p style="color:red">✗ Error: {e}</p>'

        def generar_word(event):
            try:
                from generador_word import generar_mapa_progreso
                ruta = generar_mapa_progreso()
                status.object = f'<p style="color:green;font-weight:600">✓ Generado: {ruta}</p>'
            except Exception as e:
                status.object = f'<p style="color:#e67e00">⚠ Mapa de Progreso aún no implementado</p>'

        btn_excel.on_click(generar_excel)
        btn_word.on_click(generar_word)

        return pn.Column(
            pn.Row(*bloques, sizing_mode="stretch_width"),
            pn.layout.Divider(),
            pn.Row(
                pn.Column(
                    pn.pane.HTML('<div class="editor-section-title">Ver cobertura por competencia</div>'),
                    selector,
                    popup_col,
                ),
                pn.layout.HSpacer(),
                pn.Column(
                    pn.pane.HTML('<div class="editor-section-title">Generar documentos</div>'),
                    btn_excel,
                    btn_word,
                    status,
                    align="end"
                ),
            ),
        )


# ──────────────────────────────────────────────────────────────
# PESTAÑA 2: EDITOR DE PROGRAMAS
# ──────────────────────────────────────────────────────────────

class EditorProgramas(param.Parameterized):
    asignatura_id = param.Integer(default=0)

    def __init__(self, **params):
        super().__init__(**params)
        self._widgets = {}
        self._unidades_widgets = []
        self._eval_widgets = []
        self._ra_checks = {}
        self._contenido_editor = pn.Column()
        self._status = pn.pane.HTML("")

    def _build_selector(self):
        asigs = get_asignaturas_lista()
        opciones = {f"S{a['semestre']} — {a['codigo']} — {a['nombre']}": a["id"] for a in asigs}
        sel = pn.widgets.Select(
            name="Seleccionar asignatura",
            options={"— elige una asignatura —": 0, **opciones},
            width=500
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

        # ── SECCIÓN: Identificación ──
        w_nombre = pn.widgets.TextInput(name="Nombre", value=asig.get("nombre", ""), width=400)
        w_sem = pn.widgets.IntInput(name="Semestre", value=asig.get("semestre") or 1, width=100)
        w_nivel = pn.widgets.TextInput(name="Nivel", value=asig.get("nivel", ""), width=200)
        w_dur = pn.widgets.TextInput(name="Duración", value=asig.get("duracion", ""), width=200)
        w_req = pn.widgets.TextInput(name="Requisitos", value=asig.get("requisitos", ""), width=400)

        sec_ident = pn.Column(
            pn.pane.HTML('<div class="editor-section-title">Identificación</div>'),
            pn.Row(w_nombre, w_sem),
            pn.Row(w_nivel, w_dur),
            w_req,
            css_classes=["editor-section"]
        )

        # ── SECCIÓN: RAs / Tributación ──
        grupos_ra = {}
        for ra in todos_ras:
            grupos_ra.setdefault(ra["comp"], []).append(ra)

        ra_checks = {}
        ra_cols = []
        for comp, ras in grupos_ra.items():
            checks = []
            for ra in ras:
                cb = pn.widgets.Checkbox(
                    name=ra["codigo_completo"],
                    value=(ra["id"] in ra_ids_actuales),
                    width=380
                )
                ra_checks[ra["id"]] = cb
                checks.append(cb)
            ra_cols.append(pn.Column(
                pn.pane.HTML(f'<div style="font-weight:700;color:#1a3a5c;margin-bottom:4px">{comp}</div>'),
                *checks,
                width=400
            ))

        self._ra_checks = ra_checks
        # Agrupar en filas de 2
        ra_filas = []
        for i in range(0, len(ra_cols), 2):
            ra_filas.append(pn.Row(*ra_cols[i:i+2]))

        sec_ras = pn.Column(
            pn.pane.HTML('<div class="editor-section-title">Tributación — Resultados de Aprendizaje</div>'),
            *ra_filas,
            css_classes=["editor-section"]
        )

        # ── SECCIÓN: Unidades ──
        unidad_widgets = []
        for u in unidades:
            w_u_nombre = pn.widgets.TextInput(
                name="Nombre de unidad", value=u.get("nombre", "")[:100], width=380
            )
            w_u_cont = pn.widgets.TextAreaInput(
                name="Contenidos", value=u.get("contenidos", ""),
                height=100, width=560
            )
            w_u_il = pn.widgets.TextAreaInput(
                name="Indicador de logro", value=u.get("indicador_logro", ""),
                height=60, width=560
            )
            unidad_widgets.append({
                "nombre": w_u_nombre,
                "contenidos": w_u_cont,
                "indicador_logro": w_u_il,
                "panel": pn.Column(
                    pn.pane.HTML(f'<div style="font-size:12px;font-weight:600;color:#4a6785">Unidad {u.get("orden","")}</div>'),
                    w_u_nombre, w_u_cont, w_u_il,
                    css_classes=["editor-section"]
                )
            })

        self._unidades_widgets = unidad_widgets
        sec_unidades = pn.Column(
            pn.pane.HTML('<div class="editor-section-title">Unidades y Contenidos</div>'),
            *[u["panel"] for u in unidad_widgets],
        )

        # ── SECCIÓN: Metodología ──
        metod_text = "\n".join(m.get("descripcion", "") for m in metodologias)
        w_metod = pn.widgets.TextAreaInput(
            name="Metodología / Estrategias de enseñanza",
            value=metod_text, height=120, width=620
        )
        sec_metod = pn.Column(
            pn.pane.HTML('<div class="editor-section-title">Metodología</div>'),
            w_metod,
            css_classes=["editor-section"]
        )

        # ── SECCIÓN: Evaluaciones ──
        eval_widgets = []
        for ev in evaluaciones:
            w_tipo = pn.widgets.TextInput(
                name="Tipo", value=ev.get("tipo", ""), width=280
            )
            w_porc = pn.widgets.TextInput(
                name="%", value=ev.get("porcentaje", ""), width=100
            )
            eval_widgets.append({"tipo": w_tipo, "porcentaje": w_porc})

        # Agregar fila vacía para nueva evaluación
        eval_widgets.append({
            "tipo": pn.widgets.TextInput(name="Tipo", placeholder="Nueva evaluación...", width=280),
            "porcentaje": pn.widgets.TextInput(name="%", placeholder="0%", width=100)
        })
        self._eval_widgets = eval_widgets

        sec_eval = pn.Column(
            pn.pane.HTML('<div class="editor-section-title">Evaluaciones</div>'),
            *[pn.Row(e["tipo"], e["porcentaje"]) for e in eval_widgets],
            css_classes=["editor-section"]
        )

        # ── BOTÓN GUARDAR ──
        btn_guardar = pn.widgets.Button(
            name="💾 Guardar cambios", button_type="primary", width=200
        )
        btn_generar_word = pn.widgets.Button(
            name="⬇ Generar programa Word", button_type="success", width=220
        )

        self._widgets = {
            "nombre": w_nombre, "semestre": w_sem, "nivel": w_nivel,
            "duracion": w_dur, "requisitos": w_req, "metodologia": w_metod
        }

        def on_guardar(event):
            datos = {
                "nombre": self._widgets["nombre"].value,
                "semestre": self._widgets["semestre"].value,
                "nivel": self._widgets["nivel"].value,
                "duracion": self._widgets["duracion"].value,
                "requisitos": self._widgets["requisitos"].value,
                "unidades": [
                    {
                        "nombre": u["nombre"].value,
                        "contenidos": u["contenidos"].value,
                        "indicador_logro": u["indicador_logro"].value,
                    }
                    for u in self._unidades_widgets
                ],
                "metodologias": [self._widgets["metodologia"].value],
                "evaluaciones": [
                    {"tipo": e["tipo"].value, "porcentaje": e["porcentaje"].value}
                    for e in self._eval_widgets
                ],
                "ra_ids": [
                    ra_id for ra_id, cb in self._ra_checks.items() if cb.value
                ]
            }
            ok, msg = guardar_programa(self.asignatura_id, datos)
            if ok:
                self._status.object = f'<p style="color:green;font-weight:600">✓ {msg}</p>'
            else:
                self._status.object = f'<p style="color:red">✗ {msg}</p>'

        def on_generar_word(event):
            try:
                from generador_word import generar_programa_individual
                ruta = generar_programa_individual(self.asignatura_id)
                self._status.object = f'<p style="color:green;font-weight:600">✓ Word generado: {ruta}</p>'
            except Exception as e:
                self._status.object = f'<p style="color:#e67e00">⚠ Generador Word pendiente de implementar</p>'

        btn_guardar.on_click(on_guardar)
        btn_generar_word.on_click(on_generar_word)

        self._contenido_editor.objects = [
            sec_ident,
            sec_ras,
            sec_unidades,
            sec_metod,
            sec_eval,
            pn.Row(btn_guardar, btn_generar_word, self._status),
        ]

    def view(self):
        selector = self._build_selector()
        return pn.Column(
            pn.pane.HTML('<div class="editor-section-title" style="font-size:14px">Editar programa de asignatura</div>'),
            selector,
            pn.layout.Divider(),
            self._contenido_editor,
            scroll=True,
        )


# ──────────────────────────────────────────────────────────────
# APP PRINCIPAL
# ──────────────────────────────────────────────────────────────

def crear_app():
    pn.config.raw_css.append(CSS)

    dashboard = Dashboard()
    editor = EditorProgramas()

    header = pn.pane.HTML("""
        <div class="app-header">
            <h1>Sistema de Gestión Curricular — ICM</h1>
            <p>Instituto de Matemática · Universidad de Valparaíso · Plan 2025</p>
        </div>
    """)

    tabs = pn.Tabs(
        ("📊 Dashboard", dashboard.view()),
        ("✏️ Editar programas", editor.view()),
        dynamic=False,
    )

    return pn.Column(header, tabs, sizing_mode="stretch_width", margin=(0, 20))


app = crear_app()
app.servable()