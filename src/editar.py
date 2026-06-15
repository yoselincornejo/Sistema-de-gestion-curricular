"""
Editor de tributación: vincula y desvincula asignaturas con RAs.

Uso:
    python3 src/editar.py mostrar IMAT322
    python3 src/editar.py vincular IMAT322 "CL2, N2, RA1"
    python3 src/editar.py desvincular IMAT322 "CL2, N2, RA1"
    python3 src/editar.py listar-ras
    python3 src/editar.py listar-asignaturas
"""

import sqlite3
import sys
from pathlib import Path

RUTA_DB = Path("data/sistema.db")


def conexion():
    conn = sqlite3.connect(str(RUTA_DB))
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def buscar_asignatura(conn, codigo):
    """Busca asignatura por código. Tolera variantes con/sin espacios."""
    # Probar exacto, luego con espacio quitado, luego LIKE
    codigo_norm = codigo.strip()
    cur = conn.execute(
        "SELECT id, codigo, nombre FROM asignaturas WHERE codigo = ?",
        (codigo_norm,)
    )
    row = cur.fetchone()
    if row:
        return row

    # Intento sin espacios (IMAT322 vs IMAT 322)
    cur = conn.execute(
        "SELECT id, codigo, nombre FROM asignaturas WHERE REPLACE(codigo, ' ', '') = ?",
        (codigo_norm.replace(" ", ""),)
    )
    row = cur.fetchone()
    return row


def buscar_ra(conn, codigo_completo):
    """Busca RA por su codigo_completo normalizado."""
    cur = conn.execute(
        "SELECT id, codigo_completo, descripcion FROM resultados_aprendizaje WHERE codigo_completo = ?",
        (codigo_completo.strip(),)
    )
    return cur.fetchone()


def mostrar(codigo_asig):
    """Muestra la asignatura y todos los RAs que tributa actualmente."""
    conn = conexion()
    asig = buscar_asignatura(conn, codigo_asig)
    if not asig:
        print(f"No se encontró asignatura con código '{codigo_asig}'")
        return
    asig_id, codigo, nombre = asig
    print(f"\n=== {codigo} — {nombre} ===\n")

    filas = conn.execute("""
        SELECT ra.codigo_completo, c.codigo, c.tipo
        FROM asignatura_ra ar
        JOIN resultados_aprendizaje ra ON ra.id = ar.ra_id
        JOIN competencias c ON c.id = ra.competencia_id
        ORDER BY c.codigo, ra.codigo_completo
    """).fetchall() if False else conn.execute("""
        SELECT ra.codigo_completo, c.codigo, c.tipo
        FROM asignatura_ra ar
        JOIN resultados_aprendizaje ra ON ra.id = ar.ra_id
        JOIN competencias c ON c.id = ra.competencia_id
        WHERE ar.asignatura_id = ?
        ORDER BY c.codigo, ra.codigo_completo
    """, (asig_id,)).fetchall()

    if not filas:
        print("  (sin RAs vinculados)\n")
    else:
        print(f"  Tributa a {len(filas)} RAs:")
        for codigo_ra, codigo_comp, tipo in filas:
            print(f"    · {codigo_ra:25} (competencia {codigo_comp}, {tipo})")
        print()
    conn.close()


def vincular(codigo_asig, codigo_ra):
    """Crea un vínculo asignatura ↔ RA."""
    conn = conexion()
    asig = buscar_asignatura(conn, codigo_asig)
    if not asig:
        print(f"No se encontró asignatura '{codigo_asig}'")
        return
    ra = buscar_ra(conn, codigo_ra)
    if not ra:
        print(f"No se encontró RA '{codigo_ra}'")
        print("Pista: usa 'listar-ras' para ver los códigos disponibles")
        return

    asig_id, codigo_a, nombre_a = asig
    ra_id, codigo_r, _ = ra

    try:
        conn.execute(
            "INSERT INTO asignatura_ra (asignatura_id, ra_id) VALUES (?, ?)",
            (asig_id, ra_id)
        )
        conn.commit()
        print(f"OK: {codigo_a} ({nombre_a}) ahora tributa a {codigo_r}")
    except sqlite3.IntegrityError:
        print(f"El vínculo ya existía: {codigo_a} ↔ {codigo_r}")
    conn.close()


def desvincular(codigo_asig, codigo_ra):
    """Elimina un vínculo asignatura ↔ RA."""
    conn = conexion()
    asig = buscar_asignatura(conn, codigo_asig)
    if not asig:
        print(f"No se encontró asignatura '{codigo_asig}'")
        return
    ra = buscar_ra(conn, codigo_ra)
    if not ra:
        print(f"No se encontró RA '{codigo_ra}'")
        return

    asig_id, codigo_a, nombre_a = asig
    ra_id, codigo_r, _ = ra

    cur = conn.execute(
        "DELETE FROM asignatura_ra WHERE asignatura_id = ? AND ra_id = ?",
        (asig_id, ra_id)
    )
    conn.commit()
    if cur.rowcount > 0:
        print(f"OK: {codigo_a} ({nombre_a}) ya no tributa a {codigo_r}")
    else:
        print(f"No existía vínculo entre {codigo_a} y {codigo_r}")
    conn.close()


def listar_ras():
    """Lista todos los RAs catalogados en la BD."""
    conn = conexion()
    filas = conn.execute("""
        SELECT ra.codigo_completo, c.codigo, c.tipo, COUNT(ar.id) as n_asig
        FROM resultados_aprendizaje ra
        JOIN competencias c ON c.id = ra.competencia_id
        LEFT JOIN asignatura_ra ar ON ar.ra_id = ra.id
        GROUP BY ra.id
        ORDER BY c.codigo, ra.codigo_completo
    """).fetchall()
    print(f"\n{len(filas)} RAs catalogados:\n")
    comp_actual = None
    for codigo_ra, codigo_c, tipo, n_asig in filas:
        if codigo_c != comp_actual:
            print(f"\n  [{codigo_c} - {tipo}]")
            comp_actual = codigo_c
        print(f"    {codigo_ra:25} ({n_asig} asignaturas)")
    print()
    conn.close()


def listar_asignaturas():
    """Lista todas las asignaturas con su semestre."""
    conn = conexion()
    filas = conn.execute("""
        SELECT codigo, nombre, semestre, tipo
        FROM asignaturas
        ORDER BY semestre, codigo
    """).fetchall()
    print(f"\n{len(filas)} asignaturas:\n")
    sem_actual = None
    for codigo, nombre, sem, tipo in filas:
        if sem != sem_actual:
            print(f"\n  Semestre {sem}:")
            sem_actual = sem
        print(f"    {codigo:12} {nombre[:55]}")
    print()
    conn.close()


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return

    comando = sys.argv[1]

    if comando == "mostrar" and len(sys.argv) == 3:
        mostrar(sys.argv[2])
    elif comando == "vincular" and len(sys.argv) == 4:
        vincular(sys.argv[2], sys.argv[3])
    elif comando == "desvincular" and len(sys.argv) == 4:
        desvincular(sys.argv[2], sys.argv[3])
    elif comando == "listar-ras":
        listar_ras()
    elif comando == "listar-asignaturas":
        listar_asignaturas()
    else:
        print(__doc__)


if __name__ == "__main__":
    main()