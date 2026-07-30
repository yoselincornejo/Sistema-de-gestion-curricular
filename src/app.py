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

# Inconsistencias documentadas en documentos oficiales: discrepancia entre
# tributaciones en 'Perfil de Egreso' vs 'Resultados de Aprendizaje'.
_INCONSISTENCIAS_DOCX = {
    "QUI 121": {
        "asignatura": "Química para Ingeniería",
        "seccion": "RESULTADOS DE APRENDIZAJE / UNIDADES DE APRENDIZAJE Y CONTENIDOS",
        "desc_general": "CL1 ND1, CL2 ND1, CG2 ND1 D2, CG4 ND1 D1",
        "programa":     "CL1 ND1, CL2 ND1, CG2 ND1, CG4 ND1 (sección Aporte) · CL1 ND1 RA1, RA2 · CL2 ND1 RA1 · CG2 ND1 D2 · CG4 ND1 D1 (sección RA)",
        "nota": "En el documento oficial de la asignatura se encontró una discrepancia: "
                "el desempeño CG2 ND1 D3 no aparece en la sección 'Resultados de Aprendizaje', "
                "pero sí se menciona en la sección 'Unidades de Aprendizaje y Contenidos' "
                "como una tributación presente en los resultados de aprendizaje de la Unidad 2. "
                "La sección 'Aporte al Perfil de Egreso' y la sección 'Resultados de Aprendizaje' "
                "son consistentes entre sí (CL1 ND1, CL2 ND1, CG2 ND1, CG4 ND1).",
    },
    "IMAT 222": {
        "asignatura": "Evaluación de Proyectos",
        "seccion": "RESULTADOS DE APRENDIZAJE / UNIDADES DE APRENDIZAJE Y CONTENIDOS",
        "desc_general": "CL1 ND2, CG3 ND2 D2, CG4 ND2 D1, CG4 ND2 D2",
        "programa":     "CL1 ND2 RA2 · CG3 ND2 D2 · CG4 ND2 D1 · CG4 ND2 D2 (sección RA)",
        "nota": "En el documento oficial de la asignatura se encontró una discrepancia: "
                "el desempeño CG4 ND2 D3 no aparece en la sección 'Resultados de Aprendizaje', "
                "pero sí se menciona en la sección 'Unidades de Aprendizaje y Contenidos' "
                "como una tributación presente en los resultados de aprendizaje de la Unidad 1. "
                "Las tributaciones declaradas en la sección 'Resultados de Aprendizaje' son: "
                "CL1 ND2 RA2, CG3 ND2 D2, CG4 ND2 D1, CG4 ND2 D2.",
    },
    "IMAT 223": {
        "asignatura": "Idioma I",
        "seccion": "RESULTADOS DE APRENDIZAJE Y DESEMPEÑOS",
        "desc_general": "Nomenclatura antigua del Sello UV — códigos CG9.C, CG10, CG11 no reconocidos",
        "programa":     "CG9.C D1, CG9.C D2, CG10 D4, CG10 D5 (Unidades 1-2) · CG9.C D2, CG11 D2 (Unidad 3)",
        "nota": "El documento de IMAT 223 usa la nomenclatura anterior del Sello UV (CG9.C, CG10, CG11) "
                "con descriptores numerados globalmente (D1–D5), incompatible con el sistema actual (CG1–CG4, "
                "descriptores por nivel de dominio). "
                "El mapeo tentativo es: CG9.C → CG4 (Comunicación), CG10 → CG1 (Aprendizaje Continuo), "
                "CG11 → CG2 (Trabajo en Equipo). Sin embargo, los descriptores D4 y D5 referenciados para "
                "CG10 no existen en el plan de egreso actual (CG1 tiene máximo D2 por nivel de dominio). "
                "Las tributaciones no pueden cargarse automáticamente sin una tabla de equivalencia oficial "
                "entre el Sello UV anterior y el vigente.",
    },
    "IMAT 414": {
        "asignatura": "Taller de Integración al Perfil de Egreso Sello UV II (TIPE Sello UV II)",
        "seccion": "RESULTADOS DE APRENDIZAJE / UNIDADES DE APRENDIZAJE Y CONTENIDOS",
        "desc_general": "CG1 ND2 D1, CG2 ND2 D1-D3, CG3 ND2 D1-D3, CG4 ND2 DC3 (no existe en el plan)",
        "programa":     "CG1-ND2, CG2-ND2-DC1/DC2/DC3, CG3-ND2-DC1/DC2/DC3, CG4-ND2-DC3",
        "nota": "El documento de IMAT 414 tributa el desempeño CG4-ND2-DC3, cuya descripción es: "
                "'Expresa con claridad y respeto las propias necesidades y requerimientos en el trabajo "
                "colaborativo en contextos académicos y socioculturales.' "
                "Sin embargo, CG4 ND2 D3 no está definido en el plan de egreso vigente "
                "(CG4 ND2 solo contempla D1 y D2). "
                "Esta misma situación ocurre en IMAT 324 — la descripción es estructuralmente similar "
                "a CG4 ND3 D3 ('...en el trabajo profesional en contextos nacionales e internacionales'), "
                "lo que sugiere que el documento hace referencia a un desempeño de nivel ND2 "
                "que no existe formalmente en el plan. "
                "La tributación CG4-ND2-DC3 no se carga en la BD al no tener equivalente válido.",
    },
    "IMAT 523": {
        "asignatura": "Proyecto de Ingeniería",
        "seccion": "UNIDADES DE APRENDIZAJE Y CONTENIDOS",
        "desc_general": "CG1 N3 D3, CG5 N3 D1/D2, CG6 N3 D1/D2/D3 — no declarados en Resultados de Aprendizaje ni en el perfil de egreso vigente",
        "programa":     "CE1 ND3 RA1-RA3 · CE2 ND3 RA1-RA6 · CG1 ND3 D1/D2 · CG2 ND3 D1/D2 · CG3 ND3 D1/D2 · CG4 ND3 D1/D2/D3",
        "nota": "En el documento oficial de la asignatura se encontró una discrepancia: "
                "la columna 'Resultado de aprendizaje y/o desempeños' de la tabla de Unidades de "
                "Aprendizaje y Contenidos menciona los siguientes códigos que NO aparecen en la "
                "sección oficial 'Resultados de Aprendizaje': "
                "CG1 N3 D3, CG5 N3 D1, CG5 N3 D2, CG6 N3 D1, CG6 N3 D2, CG6 N3 D3. "
                "Adicionalmente, ninguno de estos códigos existe en el perfil de egreso vigente: "
                "CG1 ND3 solo contempla D1 y D2 (no D3), y las competencias CG5 y CG6 no están "
                "definidas en el plan de egreso actual (el Sello UV vigente tiene CG1–CG4). "
                "Estos códigos no se cargan en la BD al no tener equivalente válido.",
    },
    "IMAT 415": {
        "asignatura": "Idioma IV",
        "seccion": "UNIDADES DE APRENDIZAJE Y CONTENIDOS",
        "desc_general": "CG4 ND2 D1 y CG4 ND2 D2 mencionados en Unidades pero no en Resultados de Aprendizaje",
        "programa":     "CG4 ND3 D1 · CG4 ND3 D2 · CG4 ND3 D3",
        "nota": "En el documento oficial de la asignatura se encontró una discrepancia: "
                "los desempeños CG4 ND2 D1 ('Escucha atentamente a los distintos actores involucrados "
                "en el trabajo académico y en diversos contextos socioculturales') y CG4 ND2 D2 "
                "('Empatiza ante las necesidades y requerimientos que se expresan en el trabajo académico "
                "y en diversos contextos socioculturales') aparecen en la columna 'Desempeños clave' de "
                "la tabla de Unidades de Aprendizaje y Contenidos, junto con CG4 N3 D3. "
                "Sin embargo, la sección 'Resultados de Aprendizaje' declara únicamente CG4 ND3 D1, "
                "CG4 ND3 D2 y CG4 ND3 D3 — sin mencionar CG4 ND2 D1 ni CG4 ND2 D2. "
                "Las tributaciones cargadas en la BD corresponden a las de la sección oficial "
                "'Resultados de Aprendizaje'; CG4 ND2 D1 y CG4 ND2 D2 no se incluyen.",
    },
    "IMAT 515": {
        "asignatura": "Práctica Básica",
        "seccion": "UNIDADES DE APRENDIZAJE Y CONTENIDOS",
        "desc_general": "CG4 ND2 D2 mencionado en Actividad Curricular I pero no en Resultados de Aprendizaje",
        "programa":     "CE1 ND3 RA1 · CG2 ND2 D3 · CG3 ND3 D1 · CG4 ND1 D1 · CG4 ND1 D3 · CG4 ND3 D1 · CG4 ND3 D2 · CG4 ND3 D3",
        "nota": "En el documento oficial de la asignatura se encontró una discrepancia: "
                "el desempeño CG4 ND2 D2 ('Escucha atentamente a los distintos actores involucrados en el "
                "trabajo académico y en diversos contextos socioculturales') aparece en la columna "
                "'Resultado de aprendizaje y/o desempeños' de la Actividad Curricular I (sección "
                "'Unidades de Aprendizaje y Contenidos'), pero no está declarado en la sección oficial "
                "'Resultados de Aprendizaje'. "
                "Las tributaciones cargadas en la BD corresponden únicamente a las declaradas en la "
                "sección 'Resultados de Aprendizaje'; CG4 ND2 D2 no se incluye.",
    },
    "IMAT 324": {
        "asignatura": "TIPE I",
        "seccion": "RESULTADOS DE APRENDIZAJE",
        "desc_general": "CG4-ND2-DC3 declarado en Resultados de Aprendizaje pero no existe en el plan de egreso vigente",
        "programa":     "CG1 ND2 D1 · CG2 ND2 D1 · CG2 ND2 D2 · CG2 ND2 D3 · CG3 ND2 D1 · CG3 ND2 D2 · CG3 ND2 D3 (sin CG4 ND2 D3)",
        "nota": "El documento oficial de la asignatura declara CG4-ND2-DC3 como desempeño clave en la sección "
                "'Resultados de Aprendizaje', junto con CG1-ND2, CG2-ND2-DC1/DC2/DC3 y CG3-ND2-DC1/DC2/DC3. "
                "Sin embargo, CG4-ND2-D3 no está definido en el perfil de egreso vigente "
                "(CG4 ND2 solo contempla D1 y D2). "
                "La descripción indicada en el documento para CG4 ND2 D3 es: "
                "'Expresa con claridad y respeto las propias necesidades y requerimientos en el trabajo colaborativo en contextos académicos y socioculturales.' "
                "Esta descripción es estructuralmente similar a CG4 ND3 D3 del perfil de egreso "
                "('Expresa con claridad y respeto las propias necesidades y requerimientos en el trabajo profesional en contextos nacionales e internacionales'), "
                "lo que sugiere que el documento referencia un desempeño de nivel ND2 que no existe formalmente en el plan. "
                "Por esta razón, CG4-ND2-D3 no puede cargarse en la BD ni marcarse en las casillas de la aplicación.",
    },
    "IMAT 612": {
        "asignatura": "Proyecto de Título",
        "seccion": "UNIDADES DE APRENDIZAJE Y CONTENIDOS",
        "desc_general": "CG4 N2 D2 mencionado en Unidades de Aprendizaje pero no en Resultados de Aprendizaje",
        "programa":     "Tributaciones declaradas en sección Resultados de Aprendizaje (sin CG4 N2 D2)",
        "nota": "En el documento oficial de la asignatura se encontró una discrepancia: "
                "el desempeño CG4 N2 D2 aparece en la columna 'Resultado de aprendizaje y/o desempeños' "
                "de la sección 'Unidades de Aprendizaje y Contenidos', pero no está declarado en la "
                "sección anterior 'Resultados de Aprendizaje'. "
                "Las tributaciones cargadas en la BD corresponden únicamente a las declaradas en la "
                "sección 'Resultados de Aprendizaje'; CG4 N2 D2 no se incluye.",
    },
}

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
    bibliografia = [dict(r) for r in conn.execute(
        "SELECT id, tipo, numero, autor, titulo, editorial, anio, isbn, ejemplares FROM bibliografia WHERE asignatura_id=? ORDER BY tipo, id",
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
    return asig, unidades, metodologias, evaluaciones, ras, todos_ras, bibliografia

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
        conn.execute("DELETE FROM bibliografia WHERE asignatura_id=?", (asig_id,))
        for b in datos.get("bibliografia", []):
            if b["titulo"].strip():
                conn.execute("""
                    INSERT INTO bibliografia
                    (asignatura_id, tipo, numero, autor, titulo, editorial, anio, isbn, ejemplares)
                    VALUES (?,?,?,?,?,?,?,?,?)
                """, (asig_id, b["tipo"], b.get("numero",""), b.get("autor",""),
                      b["titulo"], b.get("editorial",""), b.get("anio",""),
                      b.get("isbn",""), b.get("ejemplares","")))
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

/* Tributaciones toggle buttons */
.bk-btn-light { color: #94a3b8 !important; background: #f8fafc !important;
    border: 1px solid #e2e8f0 !important; font-size: 12px !important; }
.bk-btn-success { font-size: 12px !important; font-weight: 700 !important; }
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

        # ── Tabla 1: Todas las discrepancias ──
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
            </tr>"""

        tabla_all = f"""
        <div style="margin-top:20px">
          <div class="card-title">Inventario completo de brechas de cobertura</div>
          <p style="font-size:12px;color:#64748B;margin-bottom:12px">
            Lista de <strong>todos</strong> los resultados de aprendizaje y asignaturas que actualmente
            tienen cobertura cero — es decir, ninguna asignatura del plan los trabaja (o la asignatura
            no tributa a ningún resultado). Incluye tanto los casos con sugerencia de la tabla anterior
            como los que requieren revisión manual. Las asignaturas electivas e idiomas aparecen aquí
            porque su naturaleza variable no permite asignarles tributaciones fijas.
          </p>
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
              </tr>
            </thead>
            <tbody>{filas_all}</tbody>
          </table>
        </div>"""

        # ── Tabla 3: Inconsistencias documentadas ──
        filas_incons = ""
        for codigo, inc in _INCONSISTENCIAS_DOCX.items():
            filas_incons += f"""
            <tr>
              <td style="font-weight:700;color:#1F4E79;padding:8px 12px;
                         border-bottom:1px solid #F1F5F9;white-space:nowrap;font-size:13px">
                {codigo}</td>
              <td style="padding:8px 12px;border-bottom:1px solid #F1F5F9;
                         font-size:12px;color:#475569">{inc['asignatura']}</td>
              <td style="padding:8px 12px;border-bottom:1px solid #F1F5F9;
                         font-size:12px;color:#475569">{inc['seccion']}</td>
              <td style="padding:8px 12px;border-bottom:1px solid #F1F5F9;font-size:12px">
                <span style="color:#7B2C2C;font-weight:600">Descripción:</span>
                <span style="color:#475569"> {inc['desc_general']}</span><br>
                <span style="color:#375623;font-weight:600">Programa:</span>
                <span style="color:#475569"> {inc['programa']}</span>
              </td>
              <td style="padding:8px 12px;border-bottom:1px solid #F1F5F9;
                         font-size:11px;color:#92400E;background:#FFFBEB">{inc['nota']}</td>
            </tr>"""

        tabla_incons = f"""
        <div style="margin-top:20px">
          <div class="card-title">Inconsistencias documentadas — conflicto entre secciones del .docx</div>
          <p style="font-size:12px;color:#64748B;margin-bottom:12px">
            En los documentos oficiales de estas asignaturas se encontró una discrepancia entre
            las tributaciones indicadas en la sección <strong>Perfil de Egreso</strong>
            y las tributaciones en la sección <strong>Resultados de Aprendizaje</strong>.
            Requieren revisión y corrección en el documento oficial de la asignatura.
          </p>
          <table style="width:100%;border-collapse:collapse;font-size:13px">
            <thead>
              <tr style="background:#FFFBEB">
                <th style="text-align:left;padding:8px 12px;font-size:11px;font-weight:700;
                           text-transform:uppercase;letter-spacing:1px;color:#92400E;
                           border-bottom:2px solid #FDE68A">Código</th>
                <th style="text-align:left;padding:8px 12px;font-size:11px;font-weight:700;
                           text-transform:uppercase;letter-spacing:1px;color:#92400E;
                           border-bottom:2px solid #FDE68A">Asignatura</th>
                <th style="text-align:left;padding:8px 12px;font-size:11px;font-weight:700;
                           text-transform:uppercase;letter-spacing:1px;color:#92400E;
                           border-bottom:2px solid #FDE68A">Sección</th>
                <th style="text-align:left;padding:8px 12px;font-size:11px;font-weight:700;
                           text-transform:uppercase;letter-spacing:1px;color:#92400E;
                           border-bottom:2px solid #FDE68A">Datos en conflicto</th>
                <th style="text-align:left;padding:8px 12px;font-size:11px;font-weight:700;
                           text-transform:uppercase;letter-spacing:1px;color:#92400E;
                           border-bottom:2px solid #FDE68A">Nota</th>
              </tr>
            </thead>
            <tbody>{filas_incons}</tbody>
          </table>
        </div>"""

        return pn.Column(
            pn.pane.HTML(tabla_all + tabla_incons),
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
        import unicodedata, re as _re

        def _norm(s):
            # Quita espacios, pasa a minúsculas y elimina acentos para comparación flexible
            s = unicodedata.normalize("NFD", s).encode("ascii", "ignore").decode()
            return _re.sub(r"\s+", "", s).lower()

        asigs = get_asignaturas_lista()
        etiqueta_a_id = {f"{a['codigo']} -- {a['nombre']}": a["id"] for a in asigs}
        todas_opciones = {"— Elige una asignatura —": 0, **etiqueta_a_id}

        buscador = pn.widgets.TextInput(
            placeholder="Buscar por código o nombre (ej: IMAT413, imat 413, Inferencia)…",
            width=460, margin=(0, 8, 0, 0),
        )

        sel = pn.widgets.Select(
            name="",
            options=dict(todas_opciones),
            width=640, margin=(0, 0, 20, 0),
        )

        _updating = [False]

        def on_buscador(event):
            if _updating[0]:
                return
            q = _norm(event.new or "")
            if not q:
                _updating[0] = True
                sel.options = dict(todas_opciones)
                sel.value = 0
                _updating[0] = False
                return
            filtradas = {k: v for k, v in etiqueta_a_id.items() if q in _norm(k)}
            if not filtradas:
                return
            _updating[0] = True
            sel.options = {"— Elige una asignatura —": 0, **filtradas}
            if len(filtradas) == 1:
                único_id = list(filtradas.values())[0]
                sel.value = único_id
                self.asignatura_id = único_id
                self._cargar_editor()
            _updating[0] = False

        def on_sel(event):
            if _updating[0] or not event.new:
                return
            _updating[0] = True
            etiqueta = next((k for k, v in etiqueta_a_id.items() if v == event.new), None)
            if etiqueta:
                buscador.value = etiqueta
            sel.options = dict(todas_opciones)
            sel.value = event.new
            _updating[0] = False
            self.asignatura_id = event.new
            self._cargar_editor()

        buscador.param.watch(on_buscador, "value_input")
        sel.param.watch(on_sel, "value")

        return pn.Column(
            pn.pane.HTML('<label style="font-size:13px;font-weight:600;color:#475569;">'
                         'Selecciona una asignatura</label>'),
            pn.Row(buscador, sel, align="end"),
            margin=(0, 0, 8, 0)
        )

    def _cargar_editor(self):
        if not self.asignatura_id:
            self._contenido_editor.objects = []
            return

        asig, unidades, metodologias, evaluaciones, ras_actuales, todos_ras, bibliografia = \
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
                checked = (ra["id"] in ra_ids_actuales)
                cb = pn.widgets.Toggle(
                    name=ra["codigo_completo"],
                    value=checked,
                    button_type="success" if checked else "light",
                    width=200, height=30, margin=(2, 4)
                )
                def _make_watcher(toggle):
                    def _on_toggle(event):
                        toggle.button_type = "success" if event.new else "light"
                    return _on_toggle
                cb.param.watch(_make_watcher(cb), "value")
                ra_checks[ra["id"]] = cb
                checks.append(cb)
            ra_cols.append(pn.Column(
                pn.pane.HTML(
                    f'<div style="font-weight:700;color:{color_comp};'
                    f'font-size:13px;margin-bottom:6px">{comp}</div>'
                ),
                *checks, width=220
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

        # ── Bibliografía ──────────────────────────────────────────
        self._biblio_widgets = []
        col_biblio = pn.Column(sizing_mode="stretch_width")

        CAMPOS_BIBLIO = [
            ("autor",      "Autor(es)",       400, False),
            ("titulo",     "Título",          400, False),
            ("editorial",  "Editorial",       200, False),
            ("anio",       "Año",              80, False),
            ("isbn",       "ISBN",            160, False),
            ("ejemplares", "Ejemplares/Acceso",140, False),
        ]

        def crear_entrada_biblio(tipo, autor="", titulo="", editorial="",
                                  anio="", isbn="", ejemplares=""):
            ws = {
                "tipo":       tipo,
                "autor":      pn.widgets.TextInput(name="Autor(es)",        value=str(autor),      width=400),
                "titulo":     pn.widgets.TextInput(name="Título",           value=str(titulo),     sizing_mode="stretch_width"),
                "editorial":  pn.widgets.TextInput(name="Editorial",        value=str(editorial),  width=200),
                "anio":       pn.widgets.TextInput(name="Año",              value=str(anio),       width=80),
                "isbn":       pn.widgets.TextInput(name="ISBN",             value=str(isbn),       width=160),
                "ejemplares": pn.widgets.TextInput(name="Ejemplares/Acceso",value=str(ejemplares), width=140),
            }
            label_color = "#1F4E79" if tipo == "basica" else "#375623"
            label_txt = "Básica Obligatoria" if tipo == "basica" else "Complementaria"
            panel = pn.Column(
                pn.pane.HTML(f'<span style="font-size:11px;font-weight:600;color:{label_color};text-transform:uppercase">{label_txt}</span>'),
                pn.Row(ws["autor"], ws["titulo"]),
                pn.Row(ws["editorial"], ws["anio"], ws["isbn"], ws["ejemplares"]),
                css_classes=["card"], margin=(0, 0, 8, 0), sizing_mode="stretch_width"
            )
            ws["panel"] = panel
            return ws

        for b in bibliografia:
            w = crear_entrada_biblio(
                b.get("tipo","basica"), b.get("autor",""), b.get("titulo",""),
                b.get("editorial",""), b.get("anio",""), b.get("isbn",""),
                b.get("ejemplares",""))
            self._biblio_widgets.append(w)
            col_biblio.append(w["panel"])

        btn_add_basica = pn.widgets.Button(
            name="➕ Básica", css_classes=["btn-add"], width=140, height=38)
        btn_add_compl = pn.widgets.Button(
            name="➕ Complementaria", css_classes=["btn-add"], width=180, height=38)

        def on_add_basica(event):
            w = crear_entrada_biblio("basica")
            self._biblio_widgets.append(w)
            col_biblio.append(w["panel"])

        def on_add_compl(event):
            w = crear_entrada_biblio("complementaria")
            self._biblio_widgets.append(w)
            col_biblio.append(w["panel"])

        btn_add_basica.on_click(on_add_basica)
        btn_add_compl.on_click(on_add_compl)

        sec_biblio = pn.Column(
            pn.pane.HTML('<div class="card-title">Bibliografía</div>'),
            col_biblio,
            pn.Row(pn.layout.HSpacer(), btn_add_basica, btn_add_compl, pn.layout.HSpacer()),
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
                "ra_ids": [rid for rid, cb in self._ra_checks.items() if cb.value],
                "bibliografia": [{"tipo": b["tipo"], "autor": b["autor"].value,
                                   "titulo": b["titulo"].value, "editorial": b["editorial"].value,
                                   "anio": b["anio"].value, "isbn": b["isbn"].value,
                                   "ejemplares": b["ejemplares"].value}
                                  for b in self._biblio_widgets]
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
            sec_biblio,
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