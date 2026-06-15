"""
    Procesa todos los archivos .docx que estén en data/programas/ y sus subcarpetas.
    Cada programa se extrae y se guarda como un JSON individual.
    El semestre se infiere del nombre de la carpeta contenedora (ej. "Semestre 3").
    """

import json
import re
from pathlib import Path
from extraer import extraer_programa

CARPETA_PROGRAMAS = Path("data/programas")
CARPETA_SALIDA = Path("data/programas_json")
ARCHIVO_RESUMEN = Path("data/resumen_carga.json")


def inferir_semestre(ruta):
        """
        Recibe la ruta de un archivo y busca en sus carpetas padre
        un nombre tipo 'Semestre N'. Devuelve el número como int o None.
        """
        for parte in ruta.parts:
            m = re.match(r"Semestre\s+(\d+)", parte, re.IGNORECASE)
            if m:
                return int(m.group(1))
        return None


def procesar_todos():
        CARPETA_SALIDA.mkdir(parents=True, exist_ok=True)

        archivos = sorted(CARPETA_PROGRAMAS.rglob("*.docx"))
        print(f"Encontrados {len(archivos)} archivos .docx (búsqueda recursiva)\n")

        exitos = []
        fallidos = []

        for archivo in archivos:
            # Saltar archivos temporales de Word/LibreOffice (~$nombre.docx)
            if archivo.name.startswith("~$"):
                continue

            try:
                programa = extraer_programa(str(archivo))

                # Agregar el semestre extraído del path
                semestre = inferir_semestre(archivo)
                programa["semestre"] = semestre

                codigo = programa["identificacion"].get("codigo", "")
                nombre_base = codigo.replace(" ", "_").replace("/", "_") if codigo else archivo.stem
                ruta_json = CARPETA_SALIDA / f"{nombre_base}.json"

                ruta_json.write_text(
                    json.dumps(programa, ensure_ascii=False, indent=2),
                    encoding="utf-8"
                )

                ident = programa["identificacion"]
                exitos.append({
                    "archivo": archivo.name,
                    "semestre": semestre,
                    "codigo": ident.get("codigo", ""),
                    "nombre": ident.get("nombre", ""),
                    "unidades": len(programa.get("unidades", [])),
                    "evaluaciones": len(programa.get("evaluaciones", [])),
                    "metodologias": len(programa.get("metodologias", [])),
                    "ras_totales": sum(len(u["ras"]) for u in programa.get("unidades", []))
                })
                marca_sem = f"S{semestre}" if semestre else "S?"
                print(f"  OK {marca_sem:4} {ident.get('codigo', '?'):10} {ident.get('nombre', '?')[:50]}")

            except Exception as e:
                fallidos.append({
                    "archivo": str(archivo.relative_to(CARPETA_PROGRAMAS)),
                    "error": str(e)
                })
                print(f"  XX      {archivo.name[:55]} -> ERROR: {e}")

        resumen = {
            "total_archivos": len(archivos),
            "exitos": len(exitos),
            "fallidos": len(fallidos),
            "detalle_exitos": exitos,
            "detalle_fallidos": fallidos
        }
        ARCHIVO_RESUMEN.write_text(
            json.dumps(resumen, ensure_ascii=False, indent=2),
            encoding="utf-8"
        )

        print(f"\n{'=' * 60}")
        print(f"Procesados: {len(exitos)} exitosos, {len(fallidos)} fallidos")
        print(f"JSONs individuales en: {CARPETA_SALIDA}")
        print(f"Resumen completo en: {ARCHIVO_RESUMEN}")


if __name__ == "__main__":
        procesar_todos()