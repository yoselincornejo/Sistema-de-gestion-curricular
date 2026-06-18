"""
migrar_bd.py — Migra la BD al esquema del plan oficial 2025.

Cambios:
- Crea tabla niveles_dominio si no existe
- Actualiza descripciones de competencias con el texto oficial
- Migra nivel_dominio de formato N1/N2/N3 a ND1/ND2/ND3
- Actualiza/inserta los 65 RAs con descripciones completas
- Preserva asignatura_ra (actualiza en lugar de borrar/reinsertar)
"""

import sqlite3
from pathlib import Path

RUTA_DB = Path("data/sistema.db")

# ── Datos oficiales del plan 2025 ────────────────────────────────────

COMPETENCIAS = [
    ("CL1", "licenciatura",
     "Resuelve problemas de ciencias de la Ingeniería mediante la generación, "
     "análisis e interpretación de información incorporando los aspectos sociales, "
     "económicos, ambientales y de innovación."),
    ("CL2", "licenciatura",
     "Utiliza métodos y/o herramientas de investigación propios de las ciencias de "
     "la ingeniería que le permitan proponer soluciones a problemas disciplinares."),
    ("CE1", "titulo",
     "Desarrolla y utiliza modelos y herramientas avanzados de la Matemática Aplicada "
     "para formular y abordar problemas complejos de ingeniería, ciencias y tecnología."),
    ("CE2", "titulo",
     "Diseña e implementa soluciones a problemas complejos de ingeniería, ciencias y "
     "tecnología, basadas en el análisis matemático, la modelación y la simulación, "
     "usando software matemáticos, estadísticos y lenguajes de programación."),
    ("CG1", "sello_uv",
     "Mejora continuamente sus habilidades profesionales y de investigación a partir de "
     "un aprendizaje autorregulado y con pensamiento crítico, lo que le permite generar "
     "soluciones innovadoras pertinentes, según sus contextos de desempeño."),
    ("CG2", "sello_uv",
     "Colabora en equipos multidisciplinarios asumiendo diversos roles, liderando tareas "
     "y soluciones en entornos complejos en pos de un objetivo común."),
    ("CG3", "sello_uv",
     "Actúa en forma ética, demostrando un comportamiento inclusivo, con responsabilidad "
     "ciudadana, desde enfoque de género y derechos humanos, respetuoso de la diversidad, "
     "considerando el impacto social, económico y medioambiental de su desempeño profesional."),
    ("CG4", "sello_uv",
     "Maneja habilidades comunicativas que le permitan desempeñarse eficazmente en "
     "contextos profesionales a nivel nacional e internacional."),
]

NIVELES_DOMINIO = [
    ("CL1", "ND1", "Aplica el pensamiento lógico matemático para resolver problemas de nivel inicial usando técnicas de las ciencias básicas en el ámbito de la ingeniería."),
    ("CL1", "ND2", "Resuelve problemas de ciencias de la Ingeniería mediante la generación, análisis e interpretación de información incorporando los aspectos sociales, económicos, ambientales y de innovación."),
    ("CL2", "ND1", "Reconoce técnicas, métodos y/o herramientas de investigación propios de las ciencias de la ingeniería, que permitan resolver problemas disciplinares."),
    ("CL2", "ND2", "Utiliza métodos y/o herramientas de investigación propios de las ciencias de la ingeniería que le permitan proponer soluciones a problemas disciplinares."),
    ("CE1", "ND1", "Domina herramientas básicas de la Matemática Aplicada para resolver problemas simplificados de ciencias o ingeniería."),
    ("CE1", "ND2", "Aplica modelos y métodos avanzados de la Matemática Aplicada para resolver problemas delimitados de ingeniería, ciencias y tecnología."),
    ("CE1", "ND3", "Formula modelos y aplica métodos avanzados de la Matemática para resolver problemas complejos de ingeniería, ciencias y tecnología."),
    ("CE2", "ND1", "Utiliza técnicas de matemáticas básicas y computacionales para resolver problemas simplificados de ciencias o ingeniería."),
    ("CE2", "ND2", "Selecciona e implementa técnicas matemáticas y computacionales, para resolver problemas delimitados de ingeniería, ciencias y tecnología."),
    ("CE2", "ND3", "Diseña, planifica y valida soluciones a problemas complejos de ingeniería, ciencias y tecnología, en base a herramientas avanzadas de matemática aplicada y computación."),
    ("CG1", "ND1", "Emplea de forma autorregulada estrategias de aprendizaje y herramientas de búsqueda y gestión del conocimiento, según sus necesidades, para la solución de desafíos académicos."),
    ("CG1", "ND2", "Desarrolla procesos reflexivos en torno a prácticas propias y observadas, explorando nuevas áreas de conocimiento a partir de las necesidades, debilidades o problemáticas detectadas en su entorno sociocultural y profesional."),
    ("CG1", "ND3", "Mejora continuamente sus habilidades profesionales y de investigación a partir de un aprendizaje autorregulado y con pensamiento crítico, lo que le permite generar soluciones innovadoras pertinentes según sus contextos de desempeño."),
    ("CG2", "ND1", "Integra equipos activamente, ejecutando tareas académicas con responsabilidad y oportunidad, para el abordaje y resolución colaborativa de problemáticas y desafíos propios de la vida universitaria."),
    ("CG2", "ND2", "Desarrolla en equipo propuestas consensuadas para la resolución de problemas en diversos contextos académicos y socioculturales."),
    ("CG2", "ND3", "Colabora en equipos multidisciplinarios asumiendo diversos roles, liderando tareas y soluciones en entornos complejos en pos de un objetivo común."),
    ("CG3", "ND1", "Propone soluciones a dilemas del ámbito sociocultural y disciplinar, desde un enfoque de género, derechos humanos, diversidad e inclusión, considerando las consecuencias en la toma de decisiones, para la actuación ética en el entorno académico y social."),
    ("CG3", "ND2", "Demuestra comportamientos éticos asociados a la responsabilidad ciudadana en contextos académicos y de vinculación con el medio, para el desarrollo de proyectos que contemplan en su diseño y ejecución la perspectiva de género y los enfoques de derechos humanos, diversidad e inclusión."),
    ("CG3", "ND3", "Actúa en forma ética, demostrando un comportamiento inclusivo y con responsabilidad ciudadana, desde un enfoque de género y derechos humanos, respetuoso de la diversidad, para un desempeño profesional de excelencia que considera el impacto sociocultural, económico y medioambiental."),
    ("CG4", "ND1", "Utiliza herramientas de expresión oral y escrita para la comunicación efectiva de sus ideas, opiniones y emociones en contextos académicos."),
    ("CG4", "ND2", "Desarrolla habilidades de comunicación interpersonal en el trabajo académico y en distintos contextos socioculturales."),
    ("CG4", "ND3", "Actúa de manera eficaz en contextos comunicativos a nivel nacional e internacional."),
]

# (comp_codigo, nivel_dominio, codigo_ra, descripcion)
RESULTADOS_APRENDIZAJE = [
    ("CL1","ND1","RA1","Aplica conocimientos, métodos y herramientas de las ciencias básicas en situaciones simplificadas de ingeniería, para fortalecer el pensamiento lógico matemático."),
    ("CL1","ND1","RA2","Resuelve problemas de nivel inicial usando técnicas de las ciencias básicas para comprender fenómenos en el ámbito de la ingeniería."),
    ("CL1","ND2","RA1","Aplica conocimientos, métodos y herramientas de las ciencias de la ingeniería incorporando aspectos sociales, económicos, ambientales o de innovación para resolver problemas de ciencias de la ingeniería."),
    ("CL1","ND2","RA2","Propone alternativas de solución incorporando aspectos sociales, económicos, ambientales o de innovación para dar respuesta a problemáticas del área disciplinar."),
    ("CL2","ND1","RA1","Utiliza conocimientos basados en investigaciones o métodos de investigación, para alcanzar conclusiones válidas."),
    ("CL2","ND2","RA1","Analiza problemas complejos de ingeniería, utilizando métodos o herramientas de investigación para alcanzar conclusiones basadas en las ciencias de la ingeniería."),
    ("CL2","ND2","RA2","Aplica técnicas, métodos o herramientas de investigación propias de las ciencias de la ingeniería para resolver problemas disciplinares."),
    ("CE1","ND1","RA1","Maneja herramientas de Matemática, Física y Química, en un nivel básico, para resolver problemas de ciencias e ingeniería."),
    ("CE1","ND1","RA2","Utiliza lógica y argumentación matemática, en diferentes contextos, para resolver problemas simplificados de ciencias e ingeniería."),
    ("CE1","ND1","RA3","Reconoce estructuras matemáticas, considerando sus propiedades, para resolver problemas simplificados de ciencias e ingeniería."),
    ("CE1","ND1","RA4","Identifica técnicas de Matemática Aplicada en un nivel básico, para resolver problemas simplificados de ciencias e ingeniería."),
    ("CE1","ND2","RA1","Demuestra resultados más avanzados en matemáticas, en diferentes contextos, para resolver problemas delimitados de ingeniería, ciencias y tecnología."),
    ("CE1","ND2","RA2","Plantea ecuaciones diferenciales (ordinarias o en derivadas parciales) a partir de un problema físico delimitado, para lograr su resolución."),
    ("CE1","ND2","RA3","Formula problemas de optimización, en contextos generales y en aplicaciones reales, para ayudar a su solución."),
    ("CE1","ND2","RA4","Aplica conceptos y resultados fundamentales, de análisis avanzado, para resolver problemas delimitados de ingeniería, ciencias y tecnología."),
    ("CE1","ND2","RA5","Aplica métodos de Matemática Aplicada, básicos y avanzados, para resolver problemas combinatoriales teóricos y aplicados."),
    ("CE1","ND2","RA6","Analiza modelos matemáticos y estadísticos, involucrando incertidumbre y datos, para solucionar problemas o explicar fenómenos."),
    ("CE1","ND2","RA7","Maneja herramientas avanzadas de la Física, en contextos generales y en aplicaciones reales, para solucionar problemas delimitados de ingeniería, ciencias y tecnología."),
    ("CE1","ND3","RA1","Fundamenta los modelos y estrategias usados en la resolución del problema, para su comprensión y resolución."),
    ("CE1","ND3","RA2","Formula modelos y aplica métodos matemáticos, avanzados en la resolución de ecuaciones en derivadas parciales para problemas físicos complejos."),
    ("CE1","ND3","RA3","Formula modelos y aplica métodos probabilísticos y estadísticos, avanzados e involucrando incertidumbre, para el estudio de fenómenos."),
    ("CE2","ND1","RA1","Maneja software matemáticos, estadísticos y lenguaje de programación, en un nivel básico, para resolver problemas simplificados de ingeniería."),
    ("CE2","ND1","RA2","Diseña algoritmos y análisis de costo computacional, aplicando principios básicos, para resolver problemas simplificados de ingeniería."),
    ("CE2","ND1","RA3","Plantea problemas simplificados de ingeniería, matemáticamente, para facilitar su interpretación y solución."),
    ("CE2","ND1","RA4","Aplica técnicas matemáticas y computacionales, en un nivel básico, para resolver problemas simplificados de ingeniería."),
    ("CE2","ND2","RA1","Implementa modelos matemáticos y estadísticos, computacionalmente, involucrando incertidumbre y datos, para simular e investigar problemas y fenómenos."),
    ("CE2","ND2","RA2","Usa herramientas matemáticas en un nivel avanzado, para diseñar algoritmos y analizar el coste computacional de los mismos."),
    ("CE2","ND2","RA3","Aplica métodos matemáticos en un nivel avanzado, para la resolución numérica de problemas de álgebra lineal y de análisis."),
    ("CE2","ND2","RA4","Implementa métodos numéricos, computacionalmente, para simular e investigar soluciones de edp y problemas de optimización."),
    ("CE2","ND3","RA1","Genera soluciones, en contextos reales y multidisciplinarios, de manera autónoma y colaborativa, a problemas de ingeniería, ciencias y tecnología."),
    ("CE2","ND3","RA2","Planifica el proceso de implementación de la solución propuesta, en base a herramientas avanzadas de matemáticas, para solucionar problemas complejos."),
    ("CE2","ND3","RA3","Valida procesos y métodos optimizándolos, para lograr la eficiencia de la solución propuesta."),
    ("CE2","ND3","RA4","Fundamenta la factibilidad de la resolución del problema propuesto, indicando su complejidad, para lograr la eficiencia de la solución."),
    ("CE2","ND3","RA5","Simula métodos matemáticos avanzados en la resolución de ecuaciones en derivadas parciales, para problemas físicos complejos."),
    ("CE2","ND3","RA6","Implementa modelos probabilísticos y estadísticos, a través de su diseño y análisis, para la ingeniería financiera, la energía, biología y gestión de recursos, entre otros."),
    ("CG1","ND1","D1","Utiliza herramientas de búsqueda, actualización y gestión de la información, generando respuestas actualizadas y atingentes a problemáticas coyunturales."),
    ("CG1","ND1","D2","Explora y usa estrategias de aprendizaje y métodos de estudio pertinentes y eficientes para enfrentar los retos académicos de la vida universitaria."),
    ("CG1","ND2","D1","Reflexiona sobre su propio desempeño en el abordaje de las distintas problemáticas de su entorno, a fin de distinguir procedimientos idóneos para los requerimientos externos."),
    ("CG1","ND2","D2","Analiza diversos escenarios en contextos académicos prácticos a fin de dar respuestas oportunas y viables a retos que impone el medio en que se desenvuelve."),
    ("CG1","ND3","D1","Muestra un desempeño profesional innovador, integrando habilidades, conocimientos y experiencias de su área disciplinar."),
    ("CG1","ND3","D2","Genera nuevas ideas para la producción de soluciones prácticas y transformadoras en su área disciplinar."),
    ("CG2","ND1","D1","Participa en los espacios de encuentro acordados por el grupo, compartiendo información y sus experiencias, para el cumplimiento de los fines del equipo."),
    ("CG2","ND1","D2","Asume responsablemente las tareas asignadas dentro del equipo, para la concreción de metas grupales."),
    ("CG2","ND1","D3","Adapta su desempeño a diversos roles definidos por el equipo, para colaborar en el avance de las tareas asumidas por el equipo."),
    ("CG2","ND2","D1","Expone al equipo sus propuestas en forma organizada, considerando diversos escenarios y puntos de vista, para la definición de acuerdos grupales."),
    ("CG2","ND2","D2","Analiza en equipo la viabilidad de soluciones propuestas, para la selección grupal de las respuestas más idóneas."),
    ("CG2","ND2","D3","Colabora en la generación de climas grupales cooperativos, demostrando respeto por otras ideas y propuestas, para el cumplimiento de objetivos del equipo."),
    ("CG2","ND3","D1","Desarrolla tareas y soluciones en equipos inter y multidisciplinarios, en forma planificada, para favorecer el cumplimiento de objetivos en entornos complejos."),
    ("CG2","ND3","D2","Favorece la comunicación efectiva, el reparto equilibrado de tareas y la cohesión del equipo, para contribuir a la consolidación y desarrollo del grupo."),
    ("CG3","ND1","D1","Reconoce los principios y/o valores que están a la base de comportamientos ético-sociales y propios de su disciplina, desde un enfoque de género, derechos humanos, diversidad e inclusión."),
    ("CG3","ND1","D2","Distingue los comportamientos éticos, para visibilizar sus consecuencias en la esfera social y disciplinar, desde un enfoque de género, derechos humanos, diversidad e inclusión."),
    ("CG3","ND1","D3","Plantea soluciones a dilemas ético-sociales y disciplinares, para contribuir al debate grupal, considerando las consecuencias en la toma de decisiones, desde un enfoque de género, derechos humanos, diversidad e inclusión."),
    ("CG3","ND2","D1","Implementa actividades en vinculación con el medio, demostrando comportamientos éticos, contemplando las necesidades e intereses de la comunidad local y/o regional, desde un enfoque de género, derechos humanos, diversidad e inclusión."),
    ("CG3","ND2","D2","Reflexiona de manera permanente sobre su comportamiento ético y respetuoso de los derechos humanos, en el trabajo de vinculación con el medio, desde un enfoque de género, derechos humanos, diversidad e inclusión."),
    ("CG3","ND2","D3","Desarrolla proyectos de vinculación con el medio, contemplando en su diseño y ejecución un comportamiento ético, la perspectiva de género y el enfoque de derechos humanos, diversidad e inclusión."),
    ("CG3","ND3","D1","Ejecuta propuestas de intervención de manera colaborativa con su equipo de trabajo para solucionar problemáticas sociales, económicas y/o medioambientales, respondiendo a intereses de la comunidad local y/o regional."),
    ("CG3","ND3","D2","Evalúa procesos de intervención junto a los actores involucrados, considerando el impacto social, económico y medioambiental para contribuir al desarrollo integral de la comunidad local y/o regional."),
    ("CG4","ND1","D1","Emplea correctamente el lenguaje escrito para la construcción de documentos académicos, en distintos contextos académicos."),
    ("CG4","ND1","D2","Comprende textos relacionados con temáticas disciplinares, escritos en lengua materna y en un segundo idioma, extrayendo información relevante para la realización de actividades académicas."),
    ("CG4","ND1","D3","Expresa correctamente opiniones, ideas y experiencias para el abordaje de temas y tareas relacionadas con actividades propias de la vida universitaria."),
    ("CG4","ND2","D1","Manifiesta con claridad las propias necesidades y requerimientos en el trabajo colaborativo en contextos académicos y socioculturales."),
    ("CG4","ND2","D2","Escucha atentamente a los distintos actores involucrados en el trabajo académico y en diversos contextos socioculturales."),
    ("CG4","ND3","D1","Escucha atentamente a los distintos actores involucrados en el trabajo profesional en contextos nacionales e internacionales."),
    ("CG4","ND3","D2","Empatiza ante las necesidades y requerimientos que se expresan en el trabajo profesional en contextos nacionales e internacionales."),
    ("CG4","ND3","D3","Expresa con claridad y respeto las propias necesidades y requerimientos en el trabajo profesional en contextos nacionales e internacionales."),
]


def migrar(ruta_db=RUTA_DB):
    if not ruta_db.exists():
        print(f"BD no encontrada: {ruta_db}")
        return

    conn = sqlite3.connect(str(ruta_db))
    conn.execute("PRAGMA foreign_keys = ON")

    # 1. Crear tabla niveles_dominio si no existe
    conn.execute("""
        CREATE TABLE IF NOT EXISTS niveles_dominio (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            competencia_id INTEGER NOT NULL,
            codigo_nivel TEXT NOT NULL,
            descripcion TEXT,
            FOREIGN KEY (competencia_id) REFERENCES competencias(id) ON DELETE CASCADE,
            UNIQUE (competencia_id, codigo_nivel)
        )
    """)
    conn.execute("CREATE INDEX IF NOT EXISTS idx_nd_competencia ON niveles_dominio(competencia_id)")

    # 2. Actualizar descripciones de competencias
    for codigo, tipo, descripcion in COMPETENCIAS:
        conn.execute(
            "UPDATE competencias SET tipo=?, descripcion=? WHERE codigo=?",
            (tipo, descripcion, codigo)
        )
    print(f"  Competencias actualizadas: {len(COMPETENCIAS)}")

    # 3. Insertar/actualizar niveles_dominio
    nd_ok = 0
    for comp_cod, cod_nivel, descripcion in NIVELES_DOMINIO:
        row = conn.execute(
            "SELECT id FROM competencias WHERE codigo=?", (comp_cod,)
        ).fetchone()
        if not row:
            print(f"  ⚠ Competencia no encontrada para ND: {comp_cod}")
            continue
        comp_id = row[0]
        conn.execute("""
            INSERT INTO niveles_dominio (competencia_id, codigo_nivel, descripcion)
            VALUES (?, ?, ?)
            ON CONFLICT(competencia_id, codigo_nivel)
            DO UPDATE SET descripcion=excluded.descripcion
        """, (comp_id, cod_nivel, descripcion))
        nd_ok += 1
    print(f"  Niveles de dominio insertados/actualizados: {nd_ok}")

    # 4. Migrar resultados_aprendizaje
    # Estrategia: para cada RA del CSV, buscar el existente por:
    #   a) codigo_completo exacto con ND (ya migrado)
    #   b) codigo_completo con N en lugar de ND (formato viejo)
    # Si encontrado: UPDATE (preserva ID → asignatura_ra intacto)
    # Si no: INSERT nuevo

    ra_updated = ra_inserted = 0

    for comp_cod, nivel, cod_ra, descripcion in RESULTADOS_APRENDIZAJE:
        row = conn.execute(
            "SELECT id FROM competencias WHERE codigo=?", (comp_cod,)
        ).fetchone()
        if not row:
            print(f"  ⚠ Competencia no encontrada: {comp_cod}")
            continue
        comp_id = row[0]

        nuevo_codigo_completo = f"{comp_cod}, {nivel}, {cod_ra}"

        # Buscar por codigo_completo exacto (ND format)
        existing = conn.execute(
            "SELECT id FROM resultados_aprendizaje WHERE codigo_completo=?",
            (nuevo_codigo_completo,)
        ).fetchone()

        if not existing:
            # Buscar por formato viejo N en lugar de ND (ej: "CL1, N1, RA1")
            nivel_viejo = "N" + nivel[2:]  # ND1 → N1
            viejo_codigo = f"{comp_cod}, {nivel_viejo}, {cod_ra}"
            existing = conn.execute(
                "SELECT id FROM resultados_aprendizaje WHERE codigo_completo=?",
                (viejo_codigo,)
            ).fetchone()

        if existing:
            conn.execute("""
                UPDATE resultados_aprendizaje
                SET competencia_id=?, nivel_dominio=?, codigo=?,
                    codigo_completo=?, descripcion=?
                WHERE id=?
            """, (comp_id, nivel, cod_ra, nuevo_codigo_completo, descripcion, existing[0]))
            ra_updated += 1
        else:
            conn.execute("""
                INSERT INTO resultados_aprendizaje
                (competencia_id, nivel_dominio, codigo, codigo_completo, descripcion)
                VALUES (?, ?, ?, ?, ?)
            """, (comp_id, nivel, cod_ra, nuevo_codigo_completo, descripcion))
            ra_inserted += 1

    print(f"  RAs actualizados: {ra_updated}  |  RAs insertados nuevos: {ra_inserted}")

    # 5. Re-mapear tributaciones que apuntan a RAs en formato antiguo
    print("\n--- Remapeo de tributaciones a RAs en formato antiguo ---")

    def _nd_desde_semestre(sem):
        if sem is None: return "ND1"
        return "ND1" if sem <= 4 else "ND2" if sem <= 8 else "ND3"

    stale = conn.execute("""
        SELECT ar.id, ar.asignatura_id, ar.ra_id,
               a.semestre,
               ra.codigo_completo, ra.codigo, ra.nivel_dominio,
               c.codigo
        FROM asignatura_ra ar
        JOIN asignaturas a ON a.id = ar.asignatura_id
        JOIN resultados_aprendizaje ra ON ra.id = ar.ra_id
        JOIN competencias c ON c.id = ra.competencia_id
        WHERE ra.codigo_completo NOT LIKE '%, ND%, %'
    """).fetchall()

    for ar_id, asig_id, old_ra_id, semestre, old_cc, old_ra_cod, old_nivel, comp_cod in stale:
        # Determinar ND: preferir el que ya estaba en el registro viejo (N2→ND2)
        if old_nivel and old_nivel.startswith('N') and not old_nivel.startswith('ND'):
            nd = 'ND' + old_nivel[1:]
        else:
            nd = _nd_desde_semestre(semestre)

        # CG usa D-codes; mapear RA1→D1, RA2→D2, RA3→D3
        if comp_cod.startswith('CG') and old_ra_cod.upper().startswith('RA'):
            new_cod = 'D' + old_ra_cod[2:]
        else:
            new_cod = old_ra_cod

        new_cc = f"{comp_cod}, {nd}, {new_cod}"
        new_ra = conn.execute(
            "SELECT id FROM resultados_aprendizaje WHERE codigo_completo=?", (new_cc,)
        ).fetchone()

        if new_ra:
            dup = conn.execute(
                "SELECT id FROM asignatura_ra WHERE asignatura_id=? AND ra_id=?",
                (asig_id, new_ra[0])
            ).fetchone()
            if not dup:
                conn.execute("UPDATE asignatura_ra SET ra_id=? WHERE id=?", (new_ra[0], ar_id))
                print(f"  Remapeado: {old_cc} → {new_cc}")
            else:
                conn.execute("DELETE FROM asignatura_ra WHERE id=?", (ar_id,))
                print(f"  Duplicado eliminado: {old_cc} → {new_cc}")
        else:
            print(f"  ⚠ Sin equivalente para: {old_cc} → buscado {new_cc}")

    # Borrar RAs antiguos sin tributaciones
    old_ras = conn.execute("""
        SELECT ra.id, ra.codigo_completo FROM resultados_aprendizaje ra
        WHERE ra.codigo_completo NOT LIKE '%, ND%, %'
    """).fetchall()
    deleted = 0
    for old_id, cc in old_ras:
        if conn.execute("SELECT COUNT(*) FROM asignatura_ra WHERE ra_id=?", (old_id,)).fetchone()[0] == 0:
            conn.execute("DELETE FROM resultados_aprendizaje WHERE id=?", (old_id,))
            deleted += 1
    if deleted:
        print(f"  RAs antiguos sin tributaciones eliminados: {deleted}")

    # 6. Reporte final
    n_tribs = conn.execute("SELECT COUNT(*) FROM asignatura_ra").fetchone()[0]
    print(f"\n  Tributaciones asignatura_ra totales: {n_tribs}")

    # 7. Verificación final
    n_desconocidos = conn.execute(
        "SELECT COUNT(*) FROM competencias WHERE tipo='desconocido'"
    ).fetchone()[0]
    if n_desconocidos:
        print(f"  ⚠ Aún hay {n_desconocidos} competencias desconocidas en la BD")

    conn.commit()
    conn.close()
    print("\n✓ Migración completada.")


if __name__ == "__main__":
    migrar()
