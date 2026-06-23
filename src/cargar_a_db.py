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
    Parsea un código de RA. Acepta formatos:
      - 'CL2, N2, RA1'  (legado N-prefix)
      - 'CL1, ND1, RA1' (nuevo ND-prefix)
      - 'CG1, ND2, D1'  (CG competencias con D-codes)
      - 'CL1, RA.1'     (sin nivel)
    Devuelve (codigo_comp, nivel_nd, cod_ra) donde nivel_nd está siempre
    en formato ND (ND1, ND2, ND3) o None si no se especificó.
    """
    s = re.sub(r"\s+", " ", codigo_completo).strip()
    m = re.match(
        r"([A-Z]{1,3}\d+)"
        r"(?:,\s*(ND?\d+))?"
        r",\s*(RA\.?\d+|D\d+)",
        s
    )
    if not m:
        return None, None, None

    cod_comp = m.group(1)
    nivel_raw = m.group(2)

    if nivel_raw is None:
        nivel = None
    elif nivel_raw.startswith("ND"):
        nivel = nivel_raw
    else:
        nivel = "ND" + nivel_raw[1:]  # N1 → ND1

    cod_ra = re.sub(r"RA\.(\d+)", r"RA\1", m.group(3))  # RA.1 → RA1
    return cod_comp, nivel, cod_ra


def obtener_o_crear_ra(conn, codigo_completo, archivo_origen=""):
    """
    Dado un código de RA, devuelve su id en resultados_aprendizaje buscando
    en la BD sin crear registros nuevos. Intenta varios formatos para máxima
    compatibilidad con PDFs legados.
    Si no se puede resolver, registra en _ras_no_resueltos y devuelve None.
    """
    cod_comp, nivel, cod_ra = parse_ra_code(codigo_completo)
    if not cod_comp:
        return None

    cur = conn.execute("SELECT id FROM competencias WHERE codigo = ?", (cod_comp,))
    if not cur.fetchone():
        _ras_no_resueltos.append({
            "codigo_completo": codigo_completo,
            "cod_comp": cod_comp,
            "archivo": archivo_origen,
        })
        return None

    def _buscar(cc):
        r = conn.execute(
            "SELECT id FROM resultados_aprendizaje WHERE codigo_completo = ?", (cc,)
        ).fetchone()
        return r[0] if r else None

    # Intento 1: código canónico ND
    codigo_nd = f"{cod_comp}, {nivel + ', ' if nivel else ''}{cod_ra}"
    ra_id = _buscar(codigo_nd)
    if ra_id:
        return ra_id

    # Intento 2: para CG con RA-codes legados, probar D-code equivalente (RA1→D1)
    if cod_comp.startswith("CG") and cod_ra.startswith("RA"):
        d_cod = "D" + cod_ra[2:]
        codigo_d = f"{cod_comp}, {nivel + ', ' if nivel else ''}{d_cod}"
        ra_id = _buscar(codigo_d)
        if ra_id:
            return ra_id

    # Intento 3: sin nivel (bare format)
    ra_id = _buscar(f"{cod_comp}, {cod_ra}")
    if ra_id:
        return ra_id

    _ras_no_resueltos.append({
        "codigo_completo": codigo_completo,
        "cod_comp": cod_comp,
        "archivo": archivo_origen,
    })
    return None


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


def limpiar_asignatura_datos(conn, asig_id):
    """Borra todos los datos relacionados de una asignatura antes de recargar."""
    for tabla in ["unidades", "metodologias", "evaluaciones",
                  "bibliografia", "linkografia", "responsables"]:
        conn.execute(f"DELETE FROM {tabla} WHERE asignatura_id=?", (asig_id,))
    conn.execute("DELETE FROM tributaciones WHERE asignatura_id=?", (asig_id,))


def cargar_programa(conn, programa):
    """Inserta un programa completo (dict del parser parsear_programas) en la BD."""
    ident = programa.get("identificacion", {})
    codigo = (ident.get("codigo", "") or "").strip()
    nombre = (ident.get("nombre", "") or "").strip()

    if not codigo:
        # Saltar programas sin código (no podemos identificarlos)
        return None, "sin_codigo"

    # 1. Insertar/actualizar asignatura (incluye horas, créditos, descripción)
    conn.execute(
        """
        INSERT INTO asignaturas
        (codigo, nombre, semestre, nivel, duracion, tipo, facultad, carrera,
         requisitos, version, archivo_origen, horas_directa, horas_autonoma,
         semanas, creditos, descripcion)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        ON CONFLICT(codigo) DO UPDATE SET
            nombre=excluded.nombre, semestre=excluded.semestre,
            nivel=excluded.nivel, duracion=excluded.duracion,
            facultad=excluded.facultad, carrera=excluded.carrera,
            requisitos=excluded.requisitos, version=excluded.version,
            archivo_origen=excluded.archivo_origen,
            horas_directa=excluded.horas_directa,
            horas_autonoma=excluded.horas_autonoma,
            semanas=excluded.semanas, creditos=excluded.creditos,
            descripcion=excluded.descripcion
        """,
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
            programa.get("archivo", ""),
            ident.get("horas_directa"),
            ident.get("horas_autonoma"),
            ident.get("semanas"),
            ident.get("creditos"),
            programa.get("descripcion", ""),
        )
    )
    asig_id = conn.execute("SELECT id FROM asignaturas WHERE codigo = ?", (codigo,)).fetchone()[0]

    # Limpiar datos previos del programa (para que sea idempotente)
    limpiar_asignatura_datos(conn, asig_id)

    # otros_recursos (columna directa)
    if programa.get("otros_recursos"):
        conn.execute(
            "UPDATE asignaturas SET otros_recursos=? WHERE id=?",
            (programa["otros_recursos"], asig_id)
        )

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
    ras_de_la_asignatura = set()  # para evitar duplicados en tributaciones
    for i, u in enumerate(programa.get("unidades", [])):
        titulo = u.get("titulo", "")
        contenidos = u.get("contenidos", "")
        numero = extraer_numero_unidad(titulo) or extraer_numero_unidad(contenidos) or (i + 1)
        # En el schema, 'nombre' guarda los códigos RA de la unidad (col RA del doc)
        ra_str = "\n".join(u.get("ra_codes", []))
        conn.execute(
            """INSERT INTO unidades (asignatura_id, orden, nombre, contenidos, indicador_logro)
               VALUES (?, ?, ?, ?, ?)""",
            (asig_id, numero, ra_str, contenidos, u.get("indicador_logro", ""))
        )

        # Acumular RAs únicos a tributar
        for cod_ra in u.get("ra_codes", []):
            ras_de_la_asignatura.add(cod_ra)

    # 6. Tributación (tributaciones)
    archivo_origen = programa.get("archivo", "")
    for cod_ra in ras_de_la_asignatura:
        ra_id = obtener_o_crear_ra(conn, cod_ra, archivo_origen)
        if ra_id:
            conn.execute(
                "INSERT OR IGNORE INTO tributaciones (asignatura_id, ra_id) VALUES (?, ?)",
                (asig_id, ra_id)
            )

    # 7. Bibliografía
    for b in programa.get("bibliografia_basica", []):
        conn.execute(
            "INSERT INTO bibliografia (asignatura_id, tipo, numero, autor, titulo, editorial, anio, isbn, ejemplares) VALUES (?,?,?,?,?,?,?,?,?)",
            (asig_id, "basica", b.get("numero", ""), b.get("autor", ""), b.get("titulo", ""),
             b.get("editorial", ""), b.get("anio", ""), b.get("isbn", ""), b.get("ejemplares", ""))
        )
    for b in programa.get("bibliografia_complementaria", []):
        conn.execute(
            "INSERT INTO bibliografia (asignatura_id, tipo, numero, autor, titulo, editorial, anio, isbn, ejemplares) VALUES (?,?,?,?,?,?,?,?,?)",
            (asig_id, "complementaria", b.get("numero", ""), b.get("autor", ""), b.get("titulo", ""),
             b.get("editorial", ""), b.get("anio", ""), b.get("isbn", ""), b.get("ejemplares", ""))
        )

    # 8. Linkografía
    for l in programa.get("linkografia", []):
        conn.execute(
            "INSERT INTO linkografia (asignatura_id, url, titulo_articulo) VALUES (?,?,?)",
            (asig_id, l.get("url", ""), l.get("descripcion", ""))
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