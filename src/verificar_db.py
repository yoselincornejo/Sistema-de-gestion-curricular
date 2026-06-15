"""Inspección rápida del estado de la BD."""
import sqlite3

conn = sqlite3.connect("data/sistema.db")

tablas = [
       "asignaturas", "responsables", "competencias",
       "resultados_aprendizaje", "asignatura_ra",
       "unidades", "metodologias", "evaluaciones"
   ]

print("Filas por tabla:")
for t in tablas:
       n = conn.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0]
       print(f"  {t:30} {n:5}")

print("\nCompetencias con cuántos RAs cada una:")
filas = conn.execute("""
       SELECT c.codigo, c.tipo, COUNT(ra.id)
       FROM competencias c
       LEFT JOIN resultados_aprendizaje ra ON ra.competencia_id = c.id
       GROUP BY c.id
       ORDER BY c.codigo
   """).fetchall()
for codigo, tipo, n in filas:
       print(f"  {codigo:6} ({tipo:15}) → {n} RAs")

print("\nTop 10 RAs más tributados:")
filas = conn.execute("""
       SELECT ra.codigo_completo, COUNT(ar.id) as n
       FROM resultados_aprendizaje ra
       JOIN asignatura_ra ar ON ar.ra_id = ra.id
       GROUP BY ra.id
       ORDER BY n DESC
       LIMIT 10
   """).fetchall()
for codigo, n in filas:
       print(f"  {codigo:25} → {n} asignaturas")

print("\nAsignaturas por tipo:")
filas = conn.execute("""
       SELECT tipo, COUNT(*) FROM asignaturas GROUP BY tipo ORDER BY tipo
   """).fetchall()
for tipo, n in filas:
       print(f"  {tipo:15} → {n}")