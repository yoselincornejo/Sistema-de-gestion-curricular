"""
Carga los JSONs generados por procesar_todos.py dentro de la BD SQLite.
Es idempotente: si vuelves a correrlo, no duplica datos (usa INSERT OR REPLACE).
"""

import json
import re
import sqlite3
from datetime import datetime
from pathlib import Path

RUTA_DB = Path("data/sistema.db")
CARPETA_JSON = Path("data/programas_json")
LOG_NO_RESUELTOS = Path("data/output/ras_no_resueltos.log")

# Acumulador de RAs no resueltos durante la carga: {cod_comp: [archivo, ...]}
_ras_no_resueltos: list[dict] = []


def parse_ra_code(codigo_completo):
    """
    Parsea un código de RA tipo 'CL2, N2, RA1' o 'CL1, RA.1'.
    Devuelve (codigo_competencia, nivel_dominio, codigo_ra).
    Si no se puede parsear, devuelve (None, None, None).
    """
    # Normalizamos espacios
    s = re.sub(r"\s+", " ", codigo_completo).strip()
    # Patrón: COMPETENCIA(letras+digito), opcional N(digito), RA(opcional .)(digito)
    m = re.match(r"([A-Z]{1,3}\d+)(?:,\s*N(\d+))?,\s*RA\.?(\d+)", s)
    if not m:
        return None, None, None
    cod_comp = m.group(1)
    nivel = f"N{m.group(2)}" if m.group(2) else None
    cod_ra = f"RA{m.group(3)}"
    return cod_comp, nivel, cod_ra


def obtener_o_crear_ra(conn, codigo_completo, archivo_origen=""):
    """
    Dado un código tipo 'CL2, N2, RA1', devuelve el id en resultados_aprendizaje.
    Si la competencia no existe en la BD, NO la crea: registra el caso en
    _ras_no_resueltos y devuelve None para que la tributación se omita.
    """
    cod_comp, nivel, cod_ra = parse_ra_code(codigo_completo)
    if not cod_comp:
        return None

    cur = conn.execute("SELECT id FROM competencias WHERE codigo = ?", (cod_comp,))
    row = cur.fetchone()
    if not row:
        # Competencia no registrada — registrar y omitir sin crear nada
        _ras_no_resueltos.append({
            "codigo_completo": codigo_completo,
            "cod_comp": cod_comp,
            "archivo": archivo_origen,
        })
        return None
    competencia_id = row[0]

    # Buscar el RA por codigo_completo normalizado
    codigo_norm = f"{cod_comp}, {nivel + ', ' if nivel else ''}{cod_ra}"
    cur = conn.execute(
        "SELECT id FROM resultados_aprendizaje WHERE codigo_completo = ?",
        (codigo_norm,)
    )
    row = cur.fetchone()
    if row:
        return row[0]

    # Crear el RA (la competencia existe, solo falta el RA específico)
    cur = conn.execute(
        """INSERT INTO resultados_aprendizaje
           (competencia_id, nivel_dominio, codigo, codigo_completo, descripcion)
           VALUES (?, ?, ?, ?, ?)""",
        (competencia_id, nivel, cod_ra, codigo_norm, "")
    )
    return cur.lastrowid


def extraer_numero_unidad(texto):
    """Extrae el número de unidad desde un texto tipo 'Unidad 3: ...' o 'Unidad III: ...'."""
    m = re.match(r"Unidad\s+(\d+|[IVX]+)", texto, re.IGNORECASE)
    if not m:
        return None
    raw = m.group(1)
    if raw.isdigit():
        return int(raw)
    # Romanos básicos
    romanos = {"I": 1, "II": 2, "III": 3, "IV": 4, "V": 5, "VI": 6, "VII": 7, "VIII": 8, "IX": 9, "X": 10, "XI": 11, "XII": 12}
    return romanos.get(raw.upper())


def cargar_programa(conn, programa):
    """Inserta un programa completo en la BD."""
    ident = programa.get("identificacion", {})
    codigo = ident.get("codigo", "").strip()
    nombre = ident.get("nombre", "").strip()

    if not codigo:
        # Saltar programas sin código (no podemos identificarlos)
        return None, "sin_codigo"

    # 1. Insertar/reemplazar asignatura
    conn.execute(
        """INSERT OR REPLACE INTO asignaturas
           (codigo, nombre, semestre, nivel, duracion, tipo, facultad, carrera,
            requisitos, version, archivo_origen)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
        (
            codigo, nombre,
            programa.get("semestre"),
            ident.get("nivel", ""),
            ident.get("duracion", ""),
            programa.get("tipo", "disciplinar"),
            ident.get("facultad", ""),
            ident.get("carrera", ""),
            ident.get("requisitos", ""),
            programa.get("responsables", {}).get("version", ""),
            programa.get("archivo", "")
        )
    )
    asig_id = conn.execute("SELECT id FROM asignaturas WHERE codigo = ?", (codigo,)).fetchone()[0]

    # Limpiar datos previos del programa (para que sea idempotente)
    conn.execute("DELETE FROM responsables WHERE asignatura_id = ?", (asig_id,))
    conn.execute("DELETE FROM unidades WHERE asignatura_id = ?", (asig_id,))
    conn.execute("DELETE FROM metodologias WHERE asignatura_id = ?", (asig_id,))
    conn.execute("DELETE FROM evaluaciones WHERE asignatura_id = ?", (asig_id,))
    conn.execute("DELETE FROM asignatura_ra WHERE asignatura_id = ?", (asig_id,))

    # 2. Responsables
    resp = programa.get("responsables", {})
    if resp.get("responsable"):
        conn.execute(
            "INSERT INTO responsables (asignatura_id, rol, nombre) VALUES (?, ?, ?)",
            (asig_id, "responsable", resp["responsable"])
        )
    if resp.get("docente_a_cargo"):
        conn.execute(
            "INSERT INTO responsables (asignatura_id, rol, nombre) VALUES (?, ?, ?)",
            (asig_id, "docente_a_cargo", resp["docente_a_cargo"])
        )

    # 3. Metodologías
    for metod in programa.get("metodologias", []):
        if metod:
            conn.execute(
                "INSERT INTO metodologias (asignatura_id, descripcion) VALUES (?, ?)",
                (asig_id, metod)
            )

    # 4. Evaluaciones
    for ev in programa.get("evaluaciones", []):
        tipo = ev.get("tipo", "")
        porcentaje = ev.get("porcentaje", "")
        if tipo:
            conn.execute(
                "INSERT INTO evaluaciones (asignatura_id, tipo, porcentaje) VALUES (?, ?, ?)",
                (asig_id, tipo, porcentaje)
            )

    # 5. Unidades + tributación
    ras_de_la_asignatura = set()  # para evitar duplicados en asignatura_ra
    for u in programa.get("unidades", []):
        texto_unidad = u.get("unidad_y_contenidos", "")
        numero = extraer_numero_unidad(texto_unidad)
        conn.execute(
            """INSERT INTO unidades (asignatura_id, orden, nombre, contenidos, indicador_logro)
               VALUES (?, ?, ?, ?, ?)""",
            (asig_id, numero, texto_unidad[:100], texto_unidad, u.get("indicador_logro", ""))
        )

        # Acumular RAs únicos a tributar
        for cod_ra in u.get("ras", []):
            ras_de_la_asignatura.add(cod_ra)

    # 6. Tributación (asignatura_ra)
    archivo_origen = programa.get("archivo", "")
    for cod_ra in ras_de_la_asignatura:
        ra_id = obtener_o_crear_ra(conn, cod_ra, archivo_origen)
        if ra_id:
            conn.execute(
                "INSERT OR IGNORE INTO asignatura_ra (asignatura_id, ra_id) VALUES (?, ?)",
                (asig_id, ra_id)
            )

    return asig_id, "ok"


def cargar_todos():
    conn = sqlite3.connect(str(RUTA_DB))
    conn.execute("PRAGMA foreign_keys = ON")

    archivos = sorted(CARPETA_JSON.glob("*.json"))
    print(f"Encontrados {len(archivos)} JSONs en {CARPETA_JSON}\n")

    cargados = 0
    fallidos = []

    for archivo in archivos:
        try:
            programa = json.loads(archivo.read_text(encoding="utf-8"))
            asig_id, estado = cargar_programa(conn, programa)
            if estado == "ok":
                cargados += 1
                ident = programa.get("identificacion", {})
                print(f"  OK  {ident.get('codigo', '?'):10} {ident.get('nombre', '')[:50]}")
            else:
                fallidos.append((archivo.name, estado))
                print(f"  -- {archivo.name:55} ({estado})")
        except Exception as e:
            fallidos.append((archivo.name, str(e)))
            print(f"  XX {archivo.name:55} ERROR: {e}")

    conn.commit()
    conn.close()

    print(f"\n{'=' * 60}")
    print(f"Cargados a BD: {cargados}/{len(archivos)}")
    if fallidos:
        print(f"No cargados: {len(fallidos)}")
        for nombre, motivo in fallidos:
            print(f"  - {nombre}: {motivo}")

    # Resumen de RAs no resueltos (competencias fuera del plan)
    if _ras_no_resueltos:
        from collections import Counter
        por_comp = Counter(r["cod_comp"] for r in _ras_no_resueltos)
        por_archivo = {}
        for r in _ras_no_resueltos:
            por_archivo.setdefault(r["archivo"], set()).add(r["codigo_completo"])

        print(f"\n⚠ {len(_ras_no_resueltos)} tributación(es) omitidas por competencia desconocida:")
        for comp, n in por_comp.most_common():
            print(f"  {comp}: {n} RA(s)")
        print("  (La competencia no existe en la BD — no se creó ninguna entrada nueva)")
        print(f"  Detalle guardado en: {LOG_NO_RESUELTOS}")

        LOG_NO_RESUELTOS.parent.mkdir(parents=True, exist_ok=True)
        with open(LOG_NO_RESUELTOS, "w", encoding="utf-8") as f:
            f.write(f"RAs no resueltos — {datetime.now().strftime('%Y-%m-%d %H:%M')}\n")
            f.write("=" * 60 + "\n")
            for archivo, codigos in sorted(por_archivo.items()):
                f.write(f"\n{archivo}:\n")
                for cod in sorted(codigos):
                    f.write(f"  {cod}\n")


if __name__ == "__main__":
    cargar_todos()