"""
auditoria_db.py — Verifica la integridad de sistema.db tras la carga.
"""
import sqlite3
from pathlib import Path

RUTA_DB      = Path("data/sistema.db")
LOG_REVISION = Path("data/output/revision_manual.txt")

conn = sqlite3.connect(str(RUTA_DB))
conn.row_factory = sqlite3.Row

# ── 1. Total de asignaturas ───────────────────────────────────────────────────
total = conn.execute("SELECT COUNT(*) FROM asignaturas").fetchone()[0]
print(f"\n{'='*60}")
print(f"  AUDITORÍA DE INTEGRIDAD — sistema.db")
print(f"{'='*60}")
print(f"\n1. Total de asignaturas registradas : {total}")

# ── 2. Con y sin tributaciones ────────────────────────────────────────────────
con_trib = conn.execute("""
    SELECT COUNT(DISTINCT asignatura_id) FROM tributaciones
""").fetchone()[0]

sin_trib = total - con_trib
print(f"   Con ≥1 tributación              : {con_trib}")
print(f"   Con 0 tributaciones             : {sin_trib}")

# ── 3. Detalle de las asignaturas sin tributaciones ───────────────────────────
print(f"\n{'─'*60}")
print(f"3. Asignaturas con 0 tributaciones:")
print(f"{'─'*60}")

sin = conn.execute("""
    SELECT a.id, a.codigo, a.nombre, a.semestre
    FROM asignaturas a
    WHERE a.id NOT IN (SELECT DISTINCT asignatura_id FROM tributaciones)
    ORDER BY a.semestre, a.codigo
""").fetchall()

if sin:
    print(f"  {'ID':>4}  {'Código':<12}  {'Semestre':>8}  Nombre")
    print(f"  {'─'*4}  {'─'*12}  {'─'*8}  {'─'*38}")
    for r in sin:
        sem = f"S{r['semestre']}" if r['semestre'] else "  ?"
        print(f"  {r['id']:>4}  {r['codigo']:<12}  {sem:>8}  {r['nombre']}")
else:
    print("  ✓ Todas las asignaturas tienen al menos 1 tributación.")

# ── Extra: conteo general de todas las tablas ─────────────────────────────────
print(f"\n{'─'*60}")
print("   Estado general de la BD:")
print(f"{'─'*60}")
tablas = ["competencias","niveles_dominio","resultados_aprendizaje",
          "asignaturas","tributaciones","unidades","bibliografia",
          "linkografia","metodologias","evaluaciones","responsables"]
for t in tablas:
    n = conn.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0]
    print(f"  {t:<30} {n:>5} filas")

conn.close()

# ── 4. Contenido de revision_manual.txt ──────────────────────────────────────
print(f"\n{'─'*60}")
print("4. Archivo revision_manual.txt:")
print(f"{'─'*60}")
if LOG_REVISION.exists():
    contenido = LOG_REVISION.read_text(encoding="utf-8").strip()
    if contenido:
        print(contenido)
    else:
        print("  (archivo vacío — sin documentos en cuarentena)")
else:
    print("  (archivo no encontrado)")

print(f"\n{'='*60}\n")
