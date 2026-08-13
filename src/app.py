"""
app.py — Sistema de Gestión Curricular ICM
panel serve src/app.py --show --autoreload
"""
import sqlite3, os, sys, io, re
import pandas as pd
import panel as pn
import param
from auth import verificar_credenciales
import auth as _auth
from login_ui import crear_login
from pathlib import Path

sys.path.insert(0, os.path.dirname(__file__))
from generador_excel import generar_matriz
from generador_word import generar_programa_individual, generar_mapa_progreso, generar_mapa_ra
from auditoria import registrar_accion

pn.extension("tabulator", notifications=True)
RUTA_DB     = Path("data/sistema.db")
RUTA_OUTPUT = Path("data/output")
RUTA_OUTPUT.mkdir(parents=True, exist_ok=True)

# ── Helpers ───────────────────────────────────────────────────────

def nivel_desde_semestre(semestre):
    if semestre is None: return "1° Semestre del 1° Ciclo"
    s = int(semestre)
    ciclo = 1 if s <= 4 else (2 if s <= 8 else 3)
    return f"{s}° Semestre del {ciclo}° Ciclo"

# ── BD ────────────────────────────────────────────────────────────

def conexion():
    conn = sqlite3.connect(str(RUTA_DB))
    conn.execute("PRAGMA foreign_keys = ON")
    conn.row_factory = sqlite3.Row
    return conn

def _init_db_migration():
    """Añade columnas nuevas a tablas existentes (idempotente)."""
    conn = conexion()
    try:
        conn.execute("ALTER TABLE asignaturas ADD COLUMN docente_a_cargo TEXT DEFAULT ''")
        conn.commit()
    except Exception:
        pass  # ya existe
    finally:
        conn.close()

_init_db_migration()


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
    "IMAT 612": {
        "asignatura": "Práctica Profesional",
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
        SELECT id, codigo, nombre, semestre, duracion, requisitos, version,
               COALESCE((
                   SELECT nombre FROM responsables
                   WHERE asignatura_id=a.id AND rol='docente_a_cargo'
                   LIMIT 1
               ), '') as docente_a_cargo
        FROM asignaturas a WHERE a.id=?""", (asig_id,)).fetchone())
    req_rows = conn.execute(
        "SELECT a2.codigo FROM requisitos r JOIN asignaturas a2 ON a2.id = r.requisito_id WHERE r.asignatura_id=?",
        (asig_id,)
    ).fetchall()
    asig["requisitos_codigos"] = [r[0] for r in req_rows]
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

def guardar_tributaciones(asig_id, ra_ids):
    conn = conexion()
    try:
        conn.execute("DELETE FROM tributaciones WHERE asignatura_id=?", (asig_id,))
        for ra_id in ra_ids:
            conn.execute(
                "INSERT OR IGNORE INTO tributaciones (asignatura_id, ra_id) VALUES (?,?)",
                (asig_id, ra_id))
        conn.commit()
        return True, "Tributaciones guardadas correctamente"
    except Exception as e:
        conn.rollback()
        return False, str(e)
    finally:
        conn.close()

def guardar_bibliografia(asig_id, bibliografia):
    conn = conexion()
    try:
        conn.execute("DELETE FROM bibliografia WHERE asignatura_id=?", (asig_id,))
        for b in bibliografia:
            if b["titulo"].strip():
                conn.execute("""
                    INSERT INTO bibliografia
                    (asignatura_id, tipo, numero, autor, titulo, editorial, anio, isbn, ejemplares)
                    VALUES (?,?,?,?,?,?,?,?,?)
                """, (asig_id, b["tipo"], b.get("numero",""), b.get("autor",""),
                      b["titulo"], b.get("editorial",""), b.get("anio",""),
                      b.get("isbn",""), b.get("ejemplares","")))
        conn.commit()
        return True, "Bibliografía guardada correctamente"
    except Exception as e:
        conn.rollback()
        return False, str(e)
    finally:
        conn.close()

def guardar_programa(asig_id, datos):
    conn = conexion()
    try:
        conn.execute("""
            UPDATE asignaturas
            SET nombre=?, semestre=?, nivel=?, duracion=?
            WHERE id=?
        """, (datos["nombre"], datos["semestre"],
              nivel_desde_semestre(datos["semestre"]),
              datos["duracion"], asig_id))
        conn.execute("DELETE FROM requisitos WHERE asignatura_id=?", (asig_id,))
        for req_cod in datos.get("requisitos_ids", []):
            req_asig = conn.execute("SELECT id FROM asignaturas WHERE codigo=?", (req_cod,)).fetchone()
            if req_asig:
                conn.execute(
                    "INSERT OR IGNORE INTO requisitos (asignatura_id, requisito_id) VALUES (?,?)",
                    (asig_id, req_asig[0])
                )
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

# ── RBAC ─────────────────────────────────────────────────────────

def _puede_editar(rol: str, nombre_completo: str, docente_asig: str) -> bool:
    """Retorna True si el usuario puede editar la asignatura.

    nombre_completo: nombre real del usuario (e.g. 'Gerardo Honorato')
    docente_asig: valor del campo docente_a_cargo en responsables (puede
                  contener varios nombres separados por coma)
    """
    if rol in ("admin", "superuser"):
        return True
    if not nombre_completo or not docente_asig:
        return False
    nombre_norm = nombre_completo.strip().lower()
    # El campo puede tener varios docentes: "Gerardo Honorato, Rodrigo Meneses"
    docentes = [d.strip().lower() for d in docente_asig.split(",")]
    return any(nombre_norm in d or d in nombre_norm for d in docentes)

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

        def _hacer_mapa_ra():
            try:
                ruta = generar_mapa_ra()
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
        btn_mapa_ra = pn.widgets.FileDownload(
            callback=_hacer_mapa_ra,
            filename="mapa_ra.docx",
            label="📋 Descargar Mapa de R.A.",
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
                pn.Row(btn_excel, btn_word, btn_mapa_ra, align="center"),
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
        # Cargar todas las asignaturas para el selector de requisitos
        _conn_asigs = conexion()
        _todas_asigs = _conn_asigs.execute(
            "SELECT codigo, nombre, semestre FROM asignaturas ORDER BY semestre, codigo"
        ).fetchall()
        _conn_asigs.close()
        # {semestre: ["CODIGO -- Nombre", ...]}
        _asigs_por_sem = {}
        for _cod, _nom, _sem in _todas_asigs:
            _asigs_por_sem.setdefault(_sem, []).append(f"{_cod} -- {_nom}")

        def _opciones_requisitos(semestre):
            opts = []
            for s in sorted(_asigs_por_sem):
                if s < semestre:
                    opts.extend(_asigs_por_sem[s])
            return opts

        w_nombre = pn.widgets.TextInput(
            name="Nombre", value=asig.get("nombre",""), width=460)
        w_sem = pn.widgets.IntInput(
            name="Semestre", value=asig.get("semestre") or 1, width=110)
        w_nivel = pn.widgets.TextInput(
            name="Nivel", value=nivel_calc, width=280, disabled=True)
        w_dur = pn.widgets.TextInput(
            name="Duración", value=asig.get("duracion",""), width=200)
        w_docente = pn.widgets.TextInput(
            name="Docente a cargo", value=asig.get("docente_a_cargo",""), width=300,
            placeholder="Sin docente asignado", disabled=True)

        _req_opts_ini = _opciones_requisitos(asig.get("semestre") or 1)

        # Pre-seleccionar usando los códigos de la tabla relacional `requisitos`
        _req_codigos = asig.get("requisitos_codigos", [])
        # Índice código -> opción completa del widget ("CODIGO -- Nombre")
        _req_idx = {opt.split(" -- ")[0]: opt for opt in _req_opts_ini}
        _req_val_ini = [_req_idx[c] for c in _req_codigos if c in _req_idx]

        w_req = pn.widgets.MultiChoice(
            name="Requisitos",
            options=_req_opts_ini,
            value=_req_val_ini,
            sizing_mode="stretch_width",
            placeholder="Selecciona asignaturas de semestres anteriores...",
        )

        # El nivel se recalcula si cambia el semestre; requisitos actualiza opciones
        def on_sem_change(event):
            w_nivel.value = nivel_desde_semestre(event.new)
            nuevas_opts = _opciones_requisitos(event.new)
            w_req.options = nuevas_opts
            w_req.value = [v for v in w_req.value if v in nuevas_opts]
        w_sem.param.watch(on_sem_change, "value")

        sec_ident = pn.Column(
            pn.pane.HTML('<div class="card-title">Identificación</div>'),
            pn.Row(w_nombre, w_sem, w_nivel),
            pn.Row(w_dur, w_docente, sizing_mode="stretch_width"),
            w_req,
            css_classes=["card"], sizing_mode="stretch_width"
        )

        # ── Tributación ───────────────────────────────────────────
        # Colores de fondo por categoría de competencia
        BG_CATEGORIA     = {"licenciatura": "#BBDEFB", "titulo": "#C8E6C9", "sello_uv": "#F8BBD0"}
        BORDER_CATEGORIA = {"licenciatura": "#1565C0", "titulo": "#2E7D32", "sello_uv": "#880E4F"}
        LABEL_CATEGORIA  = {"licenciatura": "Licenciatura (CL)", "titulo": "Título Profesional (CE)", "sello_uv": "Sello UV (CG)"}

        grupos_ra = {}
        for ra in todos_ras:
            grupos_ra.setdefault(ra["comp"], []).append(ra)

        # Agrupar competencias por categoría (CL, CE, CG, etc.)
        categorias = {}
        for comp, ras in grupos_ra.items():
            tipo_comp = ras[0]["tipo"]
            categorias.setdefault(tipo_comp, {})[comp] = ras

        ra_checks = {}
        _modal_trigger  = [None]   # se llena después de crear el modal
        _last_deselected = [None]  # id del último RA deseleccionado

        bloques_categoria = []
        for tipo_cat, comps in categorias.items():
            bg    = BG_CATEGORIA.get(tipo_cat, "#f5f5f5")
            bord  = BORDER_CATEGORIA.get(tipo_cat, "#999")
            label = LABEL_CATEGORIA.get(tipo_cat, tipo_cat.upper())

            comp_cols = []
            for comp, ras in comps.items():
                color_comp = COLOR.get(tipo_cat, "#1F4E79")
                checks = []
                for ra in ras:
                    checked = (ra["id"] in ra_ids_actuales)
                    cb = pn.widgets.Toggle(
                        name=ra["codigo_completo"],
                        value=checked,
                        button_type="success" if checked else "light",
                        sizing_mode="stretch_width", min_width=0,
                        height=30, margin=(2, 3)
                    )
                    def _make_watcher(toggle, rid):
                        def _on_toggle(event):
                            toggle.button_type = "success" if event.new else "light"
                            if not event.new:
                                _last_deselected[0] = rid
                                if _modal_trigger[0] and all(
                                    not cb2.value for cb2 in ra_checks.values()
                                ):
                                    _modal_trigger[0]()
                        return _on_toggle
                    cb.param.watch(_make_watcher(cb, ra["id"]), "value")
                    ra_checks[ra["id"]] = cb
                    checks.append(cb)
                comp_cols.append(pn.Column(
                    pn.pane.HTML(
                        f'<div style="font-weight:700;color:{color_comp};'
                        f'font-size:12px;margin-bottom:4px;text-align:center">{comp}</div>'
                    ),
                    *checks,
                    sizing_mode="stretch_width", min_width=0,
                    styles={"padding": "6px 4px"},
                ))

            bloque = pn.Column(
                pn.pane.HTML(
                    f'<div style="font-weight:600;font-size:12px;color:{bord};'
                    f'margin-bottom:6px;padding-bottom:3px;border-bottom:2px solid {bord}">'
                    f'{label}</div>'
                ),
                pn.Row(*comp_cols, sizing_mode="stretch_width", styles={"min-width": "0", "flex-wrap": "wrap"}),
                styles={
                    "background": bg,
                    "border": f"1px solid {bord}",
                    "border-radius": "6px",
                    "padding": "10px 12px",
                    "margin-bottom": "8px",
                    "overflow": "auto",
                    "min-width": "0",
                },
                sizing_mode="stretch_width",
            )
            bloques_categoria.append(bloque)

        self._ra_checks = ra_checks

        _BTN_COLORS_RAS = ["success", "primary", "warning", "danger", "light"]
        _btn_ras_color_idx = [0]
        btn_guardar_ras = pn.widgets.Button(
            name="💾 Guardar tributaciones",
            button_type=_BTN_COLORS_RAS[0], width=230, height=38)

        def on_guardar_ras(event):
            ra_ids = [rid for rid, cb in self._ra_checks.items() if cb.value]
            ok, msg = guardar_tributaciones(self.asignatura_id, ra_ids)
            _btn_ras_color_idx[0] = (_btn_ras_color_idx[0] + 1) % len(_BTN_COLORS_RAS)
            btn_guardar_ras.button_type = _BTN_COLORS_RAS[_btn_ras_color_idx[0]]
            if ok:
                _u = pn.state.cache.get("usuario_actual") or "desconocido"
                registrar_accion(_u, "MODIFICACION",
                                 f"Guardó tributaciones ({len(ra_ids)} RAs seleccionados)",
                                 entidad=asig.get("codigo", ""))
                pn.state.notifications.success(msg, duration=4000)
            else:
                pn.state.notifications.error(msg, duration=4000)

        btn_guardar_ras.on_click(on_guardar_ras)

        # ── Modal: advertencia 0 RAs seleccionados ────────────────
        modal_backdrop = pn.pane.HTML(
            '<div style="position:fixed;top:0;left:0;width:100vw;height:100vh;'
            'background:rgba(0,0,0,0.55);z-index:9998"></div>',
            visible=False,
        )

        btn_modal_si = pn.widgets.Button(
            name="Sí, mantén el cambio realizado.",
            button_type="warning", width=280, height=44)
        btn_modal_no = pn.widgets.Button(
            name="No, deshace el último cambio",
            button_type="success", width=260, height=44)

        modal_dialog = pn.Column(
            pn.pane.HTML(
                '<div style="font-size:36px;margin-bottom:8px">⚠️</div>'
                '<div style="font-size:17px;font-weight:700;color:#92400e;margin-bottom:12px">'
                'Deseleccionaste el último RA de este programa de asignatura.</div>'
                '<div style="font-size:14px;color:#374151;margin-bottom:20px">'
                'Esta asignatura actualmente tributa <strong>0 Resultados de Aprendizaje</strong>.'
                '<br>¿Estás seguro?</div>'
            ),
            pn.Row(btn_modal_si, pn.Spacer(width=16), btn_modal_no, align="center"),
            visible=False,
            styles={
                "position": "fixed",
                "top": "50%",
                "left": "50%",
                "transform": "translate(-50%, -50%)",
                "z-index": "9999",
                "background": "#fffbeb",
                "border": "2px solid #f59e0b",
                "border-radius": "14px",
                "padding": "36px 40px",
                "box-shadow": "0 12px 40px rgba(0,0,0,0.4)",
                "min-width": "500px",
                "text-align": "center",
            },
        )

        def _show_modal():
            modal_backdrop.visible = True
            modal_dialog.visible   = True

        def _hide_modal():
            modal_backdrop.visible = False
            modal_dialog.visible   = False

        def on_modal_si(event):
            _hide_modal()

        def on_modal_no(event):
            last_id = _last_deselected[0]
            if last_id and last_id in ra_checks:
                ra_checks[last_id].value = True
            _hide_modal()

        btn_modal_si.on_click(on_modal_si)
        btn_modal_no.on_click(on_modal_no)
        _modal_trigger[0] = _show_modal

        flex_bloques = pn.FlexBox(
            *bloques_categoria,
            flex_direction="row",
            flex_wrap="wrap",
            sizing_mode="stretch_width",
        )
        for bloque in bloques_categoria:
            bloque.min_width = 300
            bloque.sizing_mode = "stretch_width"
            bloque.styles = {**bloque.styles, "flex": "1 1 280px", "margin": "0 6px 8px 0"}
        sec_ras = pn.Column(
            pn.pane.HTML('<div class="card-title">Tributación — Resultados de Aprendizaje</div>'),
            flex_bloques,
            pn.Row(pn.layout.HSpacer(), btn_guardar_ras),
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

        def _fix_spacing(txt):
            txt = str(txt or "")
            txt = re.sub(r'([a-záéíóúñü])([A-ZÁÉÍÓÚÑÜ])', r'\1 \2', txt)
            return re.sub(r' {2,}', ' ', txt).strip()

        for b in bibliografia:
            w = crear_entrada_biblio(
                b.get("tipo","basica"), _fix_spacing(b.get("autor","")),
                _fix_spacing(b.get("titulo","")), b.get("editorial",""),
                b.get("anio",""), b.get("isbn",""), b.get("ejemplares",""))
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

        _BTN_COLORS_BIB = ["success", "primary", "warning", "danger", "light"]
        _btn_bib_color_idx = [0]
        btn_guardar_bib = pn.widgets.Button(
            name="💾 Guardar bibliografía",
            button_type=_BTN_COLORS_BIB[0], width=210, height=38)

        def on_guardar_bib(event):
            bib = [{"tipo": b["tipo"], "autor": b["autor"].value,
                    "titulo": b["titulo"].value, "editorial": b["editorial"].value,
                    "anio": b["anio"].value, "isbn": b["isbn"].value,
                    "ejemplares": b["ejemplares"].value}
                   for b in self._biblio_widgets]
            ok, msg = guardar_bibliografia(self.asignatura_id, bib)
            _btn_bib_color_idx[0] = (_btn_bib_color_idx[0] + 1) % len(_BTN_COLORS_BIB)
            btn_guardar_bib.button_type = _BTN_COLORS_BIB[_btn_bib_color_idx[0]]
            if ok:
                _u = pn.state.cache.get("usuario_actual") or "desconocido"
                registrar_accion(_u, "MODIFICACION",
                                 f"Guardó bibliografía ({len(bib)} entradas)",
                                 entidad=asig.get("codigo", ""))
                pn.state.notifications.success(msg, duration=4000)
            else:
                pn.state.notifications.error(msg, duration=4000)

        btn_guardar_bib.on_click(on_guardar_bib)

        sec_biblio = pn.Column(
            pn.pane.HTML('<div class="card-title">Bibliografía</div>'),
            col_biblio,
            pn.Row(pn.layout.HSpacer(), btn_add_basica, btn_add_compl, pn.layout.HSpacer()),
            pn.Row(pn.layout.HSpacer(), btn_guardar_bib),
            css_classes=["card"], sizing_mode="stretch_width"
        )

        # ── Acciones ──────────────────────────────────────────────
        self._widgets = {
            "nombre": w_nombre, "semestre": w_sem, "nivel": w_nivel,
            "duracion": w_dur, "requisitos": w_req, "metodologia": w_metod,
            "docente": w_docente,
        }

        _BTN_COLORS_ALL = ["success", "primary", "warning", "danger", "light"]
        _btn_all_color_idx = [0]
        btn_guardar = pn.widgets.Button(
            name="💾 Guardar todos los cambios",
            button_type=_BTN_COLORS_ALL[0], width=300, height=56)

        def _hacer_word():
            try:
                ruta = generar_programa_individual(self.asignatura_id)
                with open(ruta, 'rb') as f:
                    return io.BytesIO(f.read())
            except Exception as e:
                pn.state.notifications.error(str(e), duration=4000)
                return io.BytesIO(b"")

        codigo_asig = (asig.get("codigo", "asig") or "asig").replace(" ", "_")
        btn_word = pn.widgets.FileDownload(
            callback=_hacer_word,
            filename=f"programa_{codigo_asig}.docx",
            label="📥 Descargar Word",
            button_type="primary",
            width=260, height=56,
            embed=False,
        )

        def on_guardar(event):
            datos = {
                "nombre":           self._widgets["nombre"].value,
                "semestre":         self._widgets["semestre"].value,
                "nivel":            self._widgets["semestre"].value,
                "duracion":         self._widgets["duracion"].value,
                "requisitos_ids":   [v.split(" -- ")[0] for v in self._widgets["requisitos"].value],
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
            _btn_all_color_idx[0] = (_btn_all_color_idx[0] + 1) % len(_BTN_COLORS_ALL)
            btn_guardar.button_type = _BTN_COLORS_ALL[_btn_all_color_idx[0]]
            if ok:
                _u = pn.state.cache.get("usuario_actual") or "desconocido"
                registrar_accion(_u, "MODIFICACION",
                                 "Guardó todos los cambios del programa",
                                 entidad=asig.get("codigo", ""))
                pn.state.notifications.success(msg, duration=4000)
            else:
                pn.state.notifications.error(msg, duration=4000)

        btn_guardar.on_click(on_guardar)

        sec_acciones = pn.Column(
            pn.pane.HTML(
                '<div style="text-align:center;font-weight:700;font-size:15px;'
                'color:#374151;margin-bottom:12px">Acciones del programa</div>'
            ),
            pn.Row(
                pn.layout.HSpacer(),
                btn_guardar,
                pn.Spacer(width=20),
                btn_word,
                pn.layout.HSpacer(),
                align="center",
            ),
            css_classes=["card"],
            sizing_mode="stretch_width",
            align="center",
            margin=(16, 0, 40, 0),
            styles={
                "background": "linear-gradient(135deg,#1e3a5f 0%,#2d6a9f 100%)",
                "border-radius": "10px",
                "padding": "20px",
            },
        )

        # ── RBAC: aplicar permisos de edición ─────────────────────
        _rol_actual      = pn.state.cache.get("rol", "user")
        _nombre_completo = pn.state.cache.get("nombre_completo", "") or ""
        _docente_asig    = asig.get("docente_a_cargo", "") or ""
        _puede           = _puede_editar(_rol_actual, _nombre_completo, _docente_asig)

        _banner_ro = pn.pane.HTML(
            f'<div style="background:#FEF3C7;border:1px solid #F59E0B;border-radius:8px;'
            f'padding:12px 16px;font-size:13px;">'
            f'🔒 <strong>Modo lectura:</strong> Esta asignatura está asignada a '
            f'<strong>{_docente_asig or "(sin asignar)"}</strong>. '
            f'Solo puedes visualizar la información, no modificarla.</div>',
            sizing_mode="stretch_width", visible=not _puede,
        )

        if not _puede:
            _widgets_edit = (
                [w_nombre, w_sem, w_nivel, w_dur, w_req, w_docente,
                 btn_guardar_ras, btn_guardar, btn_add,
                 btn_add_basica, btn_add_compl, btn_guardar_bib,
                 btn_modal_si, btn_modal_no, w_metod]
                + list(ra_checks.values())
                + [u["nombre"] for u in self._unidades_widgets]
                + [u["contenidos"] for u in self._unidades_widgets]
                + [u["indicador_logro"] for u in self._unidades_widgets]
                + [e["tipo"] for e in self._eval_widgets]
                + [e["porcentaje"] for e in self._eval_widgets]
                + [b["autor"] for b in self._biblio_widgets]
                + [b["titulo"] for b in self._biblio_widgets]
                + [b["editorial"] for b in self._biblio_widgets]
                + [b["anio"] for b in self._biblio_widgets]
                + [b["isbn"] for b in self._biblio_widgets]
                + [b["ejemplares"] for b in self._biblio_widgets]
            )
            for _w in _widgets_edit:
                if hasattr(_w, "disabled"):
                    _w.disabled = True

        self._contenido_editor.objects = [
            _banner_ro,
            modal_backdrop, modal_dialog,
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


# ── GESTIÓN DE USUARIOS ───────────────────────────────────────────

def crear_gestion_usuarios():
    """Pestaña exclusiva para rol 'admin': crear, editar y eliminar usuarios."""

    ADMIN_SESION = pn.state.cache.get("usuario_actual", "")

    ROL_COLOR = {"admin": "#1F4E79", "superuser": "#375623", "user": "#7B2C2C"}

    def _tabla_html():
        filas = ""
        for nombre, datos in _auth.get_todos().items():
            badge = ROL_COLOR.get(datos["rol"], "#475569")
            es_yo = "⭐ " if nombre == ADMIN_SESION else ""
            filas += (
                f'<tr>'
                f'<td style="padding:8px 12px;border-bottom:1px solid #F1F5F9;font-weight:600">'
                f'{es_yo}{nombre}</td>'
                f'<td style="padding:8px 12px;border-bottom:1px solid #F1F5F9">'
                f'{datos.get("nombre_completo","")}</td>'
                f'<td style="padding:8px 12px;border-bottom:1px solid #F1F5F9">'
                f'<span style="background:{badge};color:white;border-radius:12px;'
                f'padding:2px 10px;font-size:11px;font-weight:700">{datos["rol"]}</span></td>'
                f'</tr>'
            )
        return (
            '<table style="width:100%;border-collapse:collapse;font-size:13px">'
            '<thead><tr style="background:#F8FAFC">'
            '<th style="text-align:left;padding:8px 12px;border-bottom:2px solid #E2E8F0;'
            'font-size:11px;text-transform:uppercase;letter-spacing:1px;color:#1F4E79">Usuario</th>'
            '<th style="text-align:left;padding:8px 12px;border-bottom:2px solid #E2E8F0;'
            'font-size:11px;text-transform:uppercase;letter-spacing:1px;color:#1F4E79">Nombre completo</th>'
            '<th style="text-align:left;padding:8px 12px;border-bottom:2px solid #E2E8F0;'
            'font-size:11px;text-transform:uppercase;letter-spacing:1px;color:#1F4E79">Rol</th>'
            f'</tr></thead><tbody>{filas}</tbody></table>'
        )

    tabla_pane = pn.pane.HTML(_tabla_html(), sizing_mode="stretch_width")
    msg_gral   = pn.pane.HTML("", sizing_mode="stretch_width")

    # ── Sección: Crear usuario ────────────────────────────────────
    w_nuevo_usuario  = pn.widgets.TextInput(name="Nombre de usuario", width=240,
                                             placeholder="Ej: Dr. García")
    w_nuevo_nombre_c = pn.widgets.TextInput(name="Nombre completo",   width=300,
                                             placeholder="Nombre para mostrar")
    w_nuevo_password = pn.widgets.PasswordInput(name="Contraseña",    width=200,
                                                 placeholder="Contraseña inicial")
    w_nuevo_rol      = pn.widgets.Select(name="Rol",
                                          options=["user", "superuser", "admin"],
                                          value="user", width=160)
    btn_crear = pn.widgets.Button(name="➕ Crear usuario",
                                   button_type="primary", width=200, height=42)

    def on_crear(event):
        nombre = w_nuevo_usuario.value.strip()
        pwd    = w_nuevo_password.value
        ncomp  = w_nuevo_nombre_c.value.strip()
        rol    = w_nuevo_rol.value
        if not nombre or not pwd:
            msg_gral.object = _msg_error("Completa el nombre de usuario y la contraseña.")
            return
        if nombre in _auth.get_todos():
            msg_gral.object = _msg_error("Ese usuario ya existe.")
            return
        try:
            _auth.crear_usuario(nombre, pwd, rol, ncomp or nombre)
        except Exception as e:
            msg_gral.object = _msg_error(f"Error al crear usuario: {e}")
            return
        registrar_accion(ADMIN_SESION, "CREACION",
                         f"Creó usuario '{nombre}' con rol '{rol}'")
        _refrescar()
        w_nuevo_usuario.value  = ""
        w_nuevo_nombre_c.value = ""
        w_nuevo_password.value = ""
        msg_gral.object = _msg_ok(f'Usuario "<strong>{nombre}</strong>" creado correctamente.')
        pn.state.notifications.success(f'Usuario "{nombre}" creado', duration=3000)

    btn_crear.on_click(on_crear)

    # ── Sección: Editar / Eliminar usuario ────────────────────────
    sel_usuario = pn.widgets.Select(
        name="Selecciona un usuario",
        options=list(_auth.get_todos().keys()),
        width=260,
    )
    w_edit_nombre_c = pn.widgets.TextInput(name="Nombre completo", width=300)
    w_edit_password = pn.widgets.PasswordInput(
        name="Nueva contraseña", width=200,
        placeholder="Dejar vacío para no cambiar")
    w_edit_rol = pn.widgets.Select(name="Rol",
                                    options=["user", "superuser", "admin"], width=160)
    btn_guardar_edit = pn.widgets.Button(name="💾 Guardar cambios",
                                          button_type="success", width=200, height=42)
    btn_eliminar     = pn.widgets.Button(name="🗑️ Eliminar usuario",
                                          button_type="danger",  width=200, height=42)

    # Confirmación de eliminación (doble clic)
    _confirm_delete = [False]
    _confirm_timer  = [None]

    def _cargar_usuario_en_form(nombre):
        datos = _auth.get_usuario_info(nombre) or {}
        w_edit_nombre_c.value = datos.get("nombre_completo", "")
        w_edit_password.value = ""
        w_edit_rol.value      = datos.get("rol", "user")
        btn_eliminar.name     = "🗑️ Eliminar usuario"
        btn_eliminar.button_type = "danger"
        _confirm_delete[0]    = False

    def on_sel_change(event):
        _cargar_usuario_en_form(event.new)

    sel_usuario.param.watch(on_sel_change, "value")
    if sel_usuario.value:
        _cargar_usuario_en_form(sel_usuario.value)

    def on_guardar_edit(event):
        nombre = sel_usuario.value
        if not nombre or nombre not in _auth.get_todos():
            msg_gral.object = _msg_error("Selecciona un usuario válido.")
            return
        nueva_pwd = w_edit_password.value.strip() or None
        _auth.actualizar_usuario(
            nombre,
            w_edit_nombre_c.value.strip() or nombre,
            w_edit_rol.value,
            nueva_pwd,
        )
        if nueva_pwd:
            w_edit_password.value = ""
        cambios = f"Editó usuario '{nombre}': rol={w_edit_rol.value}" + (", contraseña cambiada" if nueva_pwd else "")
        registrar_accion(ADMIN_SESION, "MODIFICACION", cambios)
        _refrescar()
        msg_gral.object = _msg_ok(f'Usuario "<strong>{nombre}</strong>" actualizado correctamente.')
        pn.state.notifications.success(f'Usuario "{nombre}" actualizado', duration=3000)

    def on_eliminar(event):
        nombre = sel_usuario.value
        if not nombre or nombre not in _auth.get_todos():
            msg_gral.object = _msg_error("Selecciona un usuario válido.")
            return
        if nombre == ADMIN_SESION:
            msg_gral.object = _msg_error("No puedes eliminar tu propia cuenta.")
            return
        if not _confirm_delete[0]:
            # Primer clic: pedir confirmación
            _confirm_delete[0]       = True
            btn_eliminar.name        = "⚠️ Confirmar eliminación"
            btn_eliminar.button_type = "warning"
            msg_gral.object = (
                f'<div style="background:#FEF3C7;border:1px solid #F59E0B;border-radius:6px;'
                f'padding:8px 12px;color:#92400E;font-size:13px">'
                f'⚠ ¿Seguro que deseas eliminar a <strong>{nombre}</strong>? '
                f'Haz clic de nuevo para confirmar.</div>'
            )
        else:
            # Segundo clic: eliminar
            _auth.eliminar_usuario(nombre)
            registrar_accion(ADMIN_SESION, "ELIMINACION",
                             f"Eliminó usuario '{nombre}'")
            _confirm_delete[0]       = False
            btn_eliminar.name        = "🗑️ Eliminar usuario"
            btn_eliminar.button_type = "danger"
            _refrescar()
            restantes = list(_auth.get_todos().keys())
            if restantes:
                sel_usuario.value = restantes[0]
            msg_gral.object = _msg_ok(f'Usuario "<strong>{nombre}</strong>" eliminado.')
            pn.state.notifications.warning(f'Usuario "{nombre}" eliminado', duration=3000)

    btn_guardar_edit.on_click(on_guardar_edit)
    btn_eliminar.on_click(on_eliminar)

    def _refrescar():
        tabla_pane.object   = _tabla_html()
        sel_usuario.options = list(_auth.get_todos().keys())

    def _msg_ok(txt):
        return (f'<div style="background:#DCFCE7;border:1px solid #86EFAC;border-radius:6px;'
                f'padding:8px 12px;color:#166534;font-size:13px">✓ {txt}</div>')

    def _msg_error(txt):
        return (f'<div style="background:#FEE2E2;border:1px solid #F87171;border-radius:6px;'
                f'padding:8px 12px;color:#991B1B;font-size:13px">⚠ {txt}</div>')

    return pn.Column(
        pn.pane.HTML('<div class="card-title">Usuarios registrados</div>'),
        tabla_pane,
        pn.layout.Divider(),
        # Editar / Eliminar
        pn.pane.HTML('<div class="card-title">Editar o eliminar usuario</div>'),
        pn.Row(sel_usuario, w_edit_nombre_c, w_edit_password, w_edit_rol, align="end"),
        pn.Row(btn_guardar_edit, pn.Spacer(width=12), btn_eliminar),
        pn.layout.Divider(),
        # Crear
        pn.pane.HTML('<div class="card-title">Crear nuevo usuario</div>'),
        pn.Row(w_nuevo_usuario, w_nuevo_nombre_c, w_nuevo_password, w_nuevo_rol, align="end"),
        pn.Row(btn_crear),
        msg_gral,
        css_classes=["card"],
        sizing_mode="stretch_width",
        margin=(0, 0, 40, 0),
    )


# ── REGISTRO DE ACCIONES ──────────────────────────────────────────

def get_registros_df() -> "pd.DataFrame":
    """Extrae registro_acciones de la BD como DataFrame, del más reciente al más antiguo."""
    conn = conexion()
    try:
        df = pd.read_sql_query(
            """
            SELECT id, timestamp, usuario, accion, detalles, entidad_afectada
            FROM registro_acciones
            ORDER BY id DESC
            """,
            conn,
        )
    finally:
        conn.close()
    df.columns = ["ID", "Timestamp", "Usuario", "Acción", "Detalles", "Entidad"]
    return df


def crear_vista_auditoria():
    """Pestaña exclusiva para rol 'admin': visor de registro de acciones."""

    # ── Carga inicial ─────────────────────────────────────────────
    df_orig = get_registros_df()

    # ── Filtros ───────────────────────────────────────────────────
    usuarios_opts  = ["Todos"] + sorted(df_orig["Usuario"].dropna().unique().tolist())
    acciones_opts  = ["Todas"] + sorted(df_orig["Acción"].dropna().unique().tolist())

    sel_usuario = pn.widgets.Select(
        name="Filtrar por usuario", options=usuarios_opts, value="Todos", width=220)
    sel_accion  = pn.widgets.Select(
        name="Filtrar por acción",  options=acciones_opts, value="Todas", width=220)
    txt_entidad = pn.widgets.TextInput(
        name="Filtrar por entidad", placeholder="Ej: IMAT 211", width=200)
    btn_refresh = pn.widgets.Button(
        name="🔄 Actualizar", button_type="light", width=130, height=38)

    badge_total = pn.pane.HTML("", width=160)

    # ── Tabla ─────────────────────────────────────────────────────
    ACCION_COLOR = {
        "LOGIN":            "#1F4E79",
        "LOGOUT":           "#475569",
        "MODIFICACION":     "#375623",
        "CREACION":         "#065F46",
        "CREACION_USUARIO": "#065F46",
        "ELIMINACION":      "#7B2C2C",
        "DESCARGA":         "#6B21A8",
    }

    tabla = pn.widgets.Tabulator(
        df_orig,
        pagination="local",
        page_size=20,
        sizing_mode="stretch_width",
        show_index=False,
        frozen_columns=["ID"],
        header_filters=False,
        stylesheets=["""
            .tabulator { font-size: 12px; }
            .tabulator-row:hover { background: #F8FAFC !important; }
            .tabulator-col-title { font-weight: 700; font-size: 11px;
                text-transform: uppercase; letter-spacing: 0.8px; color: #1F4E79; }
            .tabulator-cell { padding: 7px 10px !important; }
        """],
        configuration={
            "columnDefaults": {"resizable": True},
        },
    )

    # Ancho fijo para columnas cortas
    tabla.widths = {
        "ID":       60,
        "Timestamp": 160,
        "Usuario":  160,
        "Acción":   150,
        "Entidad":  130,
    }

    def _badge_html(n):
        return (
            f'<div style="background:#EFF6FF;border:1px solid #BFDBFE;border-radius:8px;'
            f'padding:5px 14px;font-size:12px;font-weight:700;color:#1F4E79;'
            f'display:inline-block">{n} registro(s)</div>'
        )

    def _aplicar_filtros(*_):
        nonlocal df_orig
        df = df_orig.copy()
        u = sel_usuario.value
        a = sel_accion.value
        e = txt_entidad.value.strip().lower()
        if u != "Todos":
            df = df[df["Usuario"] == u]
        if a != "Todas":
            df = df[df["Acción"] == a]
        if e:
            df = df[df["Entidad"].str.lower().str.contains(e, na=False)]
        tabla.value     = df
        badge_total.object = _badge_html(len(df))

    def on_refresh(event):
        nonlocal df_orig
        df_orig = get_registros_df()
        nuevos_u = ["Todos"] + sorted(df_orig["Usuario"].dropna().unique().tolist())
        nuevas_a = ["Todas"] + sorted(df_orig["Acción"].dropna().unique().tolist())
        sel_usuario.options = nuevos_u
        sel_accion.options  = nuevas_a
        _aplicar_filtros()

    sel_usuario.param.watch(_aplicar_filtros, "value")
    sel_accion.param.watch(_aplicar_filtros,  "value")
    txt_entidad.param.watch(_aplicar_filtros, "value")
    btn_refresh.on_click(on_refresh)

    badge_total.object = _badge_html(len(df_orig))

    return pn.Column(
        pn.pane.HTML('<div class="card-title">Registro de Acciones del Sistema</div>'),
        pn.pane.HTML(
            '<p style="font-size:12px;color:#64748B;margin-bottom:12px">'
            'Historial completo de acciones realizadas por todos los usuarios. '
            'Ordenado del más reciente al más antiguo.</p>'
        ),
        pn.Row(sel_usuario, sel_accion, txt_entidad,
               pn.layout.HSpacer(), badge_total, btn_refresh,
               align="end", sizing_mode="stretch_width"),
        tabla,
        css_classes=["card"],
        sizing_mode="stretch_width",
        margin=(0, 0, 40, 0),
    )


# ── APP ───────────────────────────────────────────────────────────

def crear_app():
    # CSS ya inyectado por iniciar() al arrancar

    usuario_sesion   = pn.state.cache.get("usuario_actual", "")
    rol_sesion       = pn.state.cache.get("rol", "user")
    nombre_completo  = pn.state.cache.get("nombre_completo", "") or usuario_sesion

    ROL_LABEL = {"admin": "Administrador", "superuser": "Super usuario", "user": "Docente"}
    rol_txt = ROL_LABEL.get(rol_sesion, rol_sesion)
    # Admin ve: "Nombre Completo (username)" · Resto ve solo su nombre completo
    if rol_sesion == "admin" and usuario_sesion != nombre_completo:
        nombre_display = f"{nombre_completo} ({usuario_sesion})"
    else:
        nombre_display = nombre_completo

    btn_logout = pn.widgets.Button(
        name="Cerrar sesión",
        button_type="light",
        width=150, height=36,
        stylesheets=["""
            :host button {
                border: 1.5px solid #CBD5E1 !important;
                border-radius: 8px !important;
                font-size: 13px !important;
                color: #475569 !important;
                background: white !important;
                cursor: pointer !important;
            }
            :host button:hover {
                background: #F1F5F9 !important;
                border-color: #94A3B8 !important;
            }
        """],
    )

    def on_logout(event):
        _u = pn.state.cache.get("usuario_actual") or "desconocido"
        registrar_accion(_u, "LOGOUT", "Cierre de sesión")
        pn.state.cache["usuario_actual"] = None
        pn.state.cache["rol"]            = None
        main_area.clear()
        main_area.append(crear_login(on_success=cargar_app_principal))

    btn_logout.on_click(on_logout)

    header = pn.Row(
        pn.pane.HTML(f"""
            <div style="background:white;padding:20px 28px;border-radius:12px;
                        box-shadow:0 2px 8px rgba(0,0,0,0.06);display:flex;
                        align-items:center;gap:16px;border-left:6px solid #1F4E79;
                        margin-bottom:0;flex:1">
                <div style="width:44px;height:44px;background:#1F4E79;border-radius:8px;
                            display:flex;align-items:center;justify-content:center;
                            color:white;font-size:18px;font-weight:700;flex-shrink:0">ICM</div>
                <div style="flex:1">
                    <h1 style="margin:0;font-size:20px;font-weight:700;color:#1E293B">
                        Sistema de Gestión Curricular</h1>
                    <p style="margin:3px 0 0;font-size:13px;color:#64748B">
                        Instituto de Matemática · Universidad de Valparaíso · Plan 2025</p>
                </div>
                <div style="text-align:right;margin-right:8px">
                    <div style="font-size:13px;font-weight:600;color:#1F4E79">{nombre_display}</div>
                    <div style="font-size:11px;color:#64748B">{rol_txt}</div>
                </div>
            </div>
        """, sizing_mode="stretch_width"),
        pn.Column(btn_logout, align="center", margin=(0, 0, 0, 8)),
        sizing_mode="stretch_width",
        margin=(0, 0, 24, 0),
    )

    tabs_items = [
        ("📊 Dashboard de Cobertura", Dashboard().view()),
        ("✏️ Editor de Programas",    EditorProgramas().view()),
    ]
    if rol_sesion == "admin":
        tabs_items.append(("⚙️ Gestión de Usuarios", crear_gestion_usuarios()))
        tabs_items.append(("📋 Registro de Acciones", crear_vista_auditoria()))

    tabs = pn.Tabs(*tabs_items, dynamic=False)

    return pn.Column(
        header, tabs,
        sizing_mode="stretch_width",
        margin=(24, 40)
    )

# ── Estado de sesión ──────────────────────────────────────────────
# usuario_actual y rol se almacenan por sesión usando pn.state.cache
# (cada conexión de navegador tiene su propia entrada en el cache).

# ── Contenedor raíz dinámico ──────────────────────────────────────
main_area = pn.Column(sizing_mode="stretch_both")


def cargar_app_principal(rol: str, usuario: str):
    """Limpia el login y carga la aplicación principal."""
    info = _auth.get_usuario_info(usuario) or {}
    nombre_completo = info.get("nombre_completo") or usuario
    pn.state.cache["usuario_actual"]   = usuario
    pn.state.cache["rol"]              = rol
    pn.state.cache["nombre_completo"]  = nombre_completo
    registrar_accion(usuario, "LOGIN", f"Inicio de sesión exitoso (rol: {rol})")
    main_area.clear()
    main_area.append(crear_app())
    pn.state.notifications.success(
        f"Bienvenido/a, {nombre_completo}", duration=3000
    )


def iniciar():
    """Punto de entrada: muestra el login al arrancar."""
    pn.config.raw_css.append(CSS)
    pn.state.cache.setdefault("usuario_actual", None)
    pn.state.cache.setdefault("rol", None)
    main_area.clear()
    main_area.append(crear_login(on_success=cargar_app_principal))


pn.state.onload(iniciar)
main_area.servable(title="SGC - Sistema de Gestión Curricular")