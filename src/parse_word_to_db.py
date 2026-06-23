"""
parse_word_to_db.py — Parsea los 53 documentos Word e inserta en sistema.db.

Uso:
    python3 src/parse_word_to_db.py            # procesa todos los .docx
    python3 src/parse_word_to_db.py --dry-run  # parsea sin tocar la BD
    python3 src/parse_word_to_db.py --file "data/programas/Semestre 1/MAT 111.docx"

La BD debe existir previamente (ejecutar init_db.py primero).
"""

import logging
import re
import sqlite3
import sys
from pathlib import Path
from typing import Optional

from docx import Document
from docx.table import Table

# ── Configuración ─────────────────────────────────────────────────────────────

RUTA_DB       = Path("data/sistema.db")
CARPETA_DOCS  = Path("data/programas")
EXCLUIDOS     = {"IMAT 423"}   # Sistemas Dinámicos Continuos — excluir explícitamente

logging.basicConfig(
    level=logging.INFO,
    format="%(levelname)-8s %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler("data/output/parse_word_to_db.log", mode="w", encoding="utf-8"),
    ],
)
log = logging.getLogger(__name__)

# ── Helpers de texto ──────────────────────────────────────────────────────────

def _txt(cell) -> str:
    return cell.text.strip()

def _txt_t(tabla: Table, fila: int, col: int) -> str:
    try:
        return tabla.rows[fila].cells[col].text.strip()
    except IndexError:
        return ""

def _to_float(s: str) -> Optional[float]:
    s = s.replace(",", ".").strip()
    try:
        return float(s)
    except (ValueError, TypeError):
        return None

def _to_int(s: str) -> Optional[int]:
    try:
        return int(re.sub(r"[^\d]", "", s))
    except (ValueError, TypeError):
        return None

def _semestre_desde_nivel(nivel: str) -> Optional[int]:
    """'I Semestre del 1° Año' → 1, 'II Semestre del 2° Año' → 4, etc."""
    numeros = {"I": 1, "II": 2, "III": 3, "IV": 4, "V": 5,
               "VI": 6, "VII": 7, "VIII": 8, "IX": 9, "X": 10}
    m = re.search(r"(\d+)[°º]\s*Año", nivel)
    anio = int(m.group(1)) if m else None
    m2 = re.match(r"(I{1,3}V?|VI{0,3}|IX|X)\s+Semestre", nivel)
    sem_en_anio = numeros.get(m2.group(1), 1) if m2 else 1
    if anio:
        return (anio - 1) * 2 + sem_en_anio
    return None

# ── Parser principal ──────────────────────────────────────────────────────────

class ParseError(Exception):
    pass


def parsear_tabla_identificacion(doc: Document, nombre_archivo: str) -> dict:
    """
    Extrae datos de la Tabla 0 (siempre primera tabla, 8 filas x 6 cols).
    Estructura esperada:
        F0: Facultad | val | | Carreras | val
        F1: Nombre   | val | | Códigos  | val
        F2: Nivel    | val | | Duración | val
        F3: Requisito| val
        F4-F6: encabezados horas
        F7: docencia_directa | autonoma | total | semanas | total_hrs | creditos
    """
    if not doc.tables:
        raise ParseError("El documento no tiene tablas")

    t = doc.tables[0]
    rows = t.rows

    if len(rows) < 4:
        raise ParseError(f"Tabla 0 tiene solo {len(rows)} filas, esperadas ≥4")

    data = {}

    # Fila 0: Facultad / Carreras
    data["facultad"] = _txt_t(t, 0, 1)
    data["carrera"]  = _txt_t(t, 0, 4)

    # Fila 1: Nombre / Código
    data["nombre"] = _txt_t(t, 1, 1)
    data["codigo"] = _txt_t(t, 1, 4)

    # Fila 2: Nivel / Duración
    data["nivel"]    = _txt_t(t, 2, 1)
    data["duracion"] = _txt_t(t, 2, 4)
    data["semestre"] = _semestre_desde_nivel(data["nivel"])

    # Fila 3: Requisitos
    data["requisitos"] = _txt_t(t, 3, 1)

    # Última fila: valores numéricos de horas y créditos
    fila_vals = rows[-1]
    celdas = [_txt(c) for c in fila_vals.cells]
    # Columnas: 0=docencia_directa, 1=autonoma, 2=total, 3=semanas, 4=total_hrs, 5=creditos
    if len(celdas) >= 6:
        data["horas_directa"]  = _to_float(celdas[0])
        data["horas_autonoma"] = _to_float(celdas[1])
        data["semanas"]        = _to_int(celdas[3])
        data["creditos"]       = _to_int(celdas[5])
    else:
        log.warning("[%s] Fila de horas incompleta (%d celdas)", nombre_archivo, len(celdas))
        data["horas_directa"] = data["horas_autonoma"] = data["semanas"] = data["creditos"] = None

    return data


def _es_tabla_descripcion(t: Table) -> bool:
    """Una tabla de descripción tiene 1 col y el texto empieza con 'La asignatura'."""
    if len(t.columns) != 1:
        return False
    txt = _txt_t(t, 0, 0).lower()
    return txt.startswith("la asignatura") or txt.startswith("esta asignatura aporta")


def _es_tabla_ras_texto(t: Table) -> bool:
    """Tabla con texto largo de RAs: 1 col, comienza con 'Al final'."""
    if len(t.columns) != 1:
        return False
    return _txt_t(t, 0, 0).lower().startswith("al final")


def _es_tabla_unidades(t: Table) -> bool:
    """Tabla de unidades: 2 cols, cabecera contiene 'Resultado' o 'RA'."""
    if len(t.columns) < 2:
        return False
    h = _txt_t(t, 0, 0).lower()
    return "resultado" in h or "desempeño" in h or h.startswith("ra")


def _es_tabla_metod(t: Table) -> bool:
    """Tabla de metodologías: muchas celdas de checkboxes."""
    if len(t.rows) < 2:
        return False
    h = _txt_t(t, 0, 0).lower()
    return "basado" in h or "expositi" in h or "aprendizaje" in h or "taller" in h


def _es_tabla_eval(t: Table) -> bool:
    h = _txt_t(t, 0, 0).lower()
    return "evaluaci" in h and "tipo" in h


def _es_tabla_biblio(t: Table, tipo: str) -> bool:
    h = _txt_t(t, 0, 0).lower()
    if tipo == "basica":
        return "bibliografía básica" in h or "bibliografia basica" in h
    return "complementaria" in h


def _es_tabla_linkografia(t: Table) -> bool:
    h = _txt_t(t, 0, 0).lower()
    return "tipo de documento" in h or "linkografía" in h or "linkografia" in h


def _es_tabla_otros_recursos(t: Table) -> bool:
    txt = _txt_t(t, 0, 0).lower()
    return ("material" in txt or "recurso" in txt or "otro" in txt) and len(t.columns) == 1


def _es_tabla_responsables(t: Table) -> bool:
    h = _txt_t(t, 0, 0).lower()
    return "responsable" in h


# ── Extracción de secciones ───────────────────────────────────────────────────

def extraer_descripcion(doc: Document) -> str:
    """Busca tabla de descripción ('La asignatura…') y la concatena."""
    partes = []
    for t in doc.tables[1:]:
        if _es_tabla_descripcion(t):
            for row in t.rows:
                txt = _txt(row.cells[0])
                if txt:
                    partes.append(txt)
    return "\n".join(partes)


RA_PATTERN = re.compile(
    r"([A-Z]{1,3}\d+)"          # código competencia: CL1, CE2, CG3…
    r"(?:\s*,\s*(?:ND?)\d+)?"   # nivel opcional: N2 o ND2
    r"\s*,\s*"
    r"(RA\.?\s*\d+|D\s*\d+)",   # código RA o D
    re.IGNORECASE
)

def _normalizar_codigo_ra(raw: str) -> str:
    """'CL1,N2, RA.2' → 'CL1, ND2, RA2' | 'CG1, N2, RA1' → 'CG1, ND2, D1'"""
    raw = re.sub(r"\s+", " ", raw).strip()
    m = re.match(
        r"([A-Z]{1,3}\d+)"
        r"(?:\s*,\s*(ND?)\s*(\d+))?"
        r"\s*,\s*(RA\.?\s*\d+|D\s*\d+)",
        raw, re.IGNORECASE
    )
    if not m:
        return raw
    cod_comp = m.group(1).upper()
    nd_prefix = (m.group(2) or "").upper()
    nd_num    = m.group(3) or ""
    cod_ra    = re.sub(r"[\s\.]", "", m.group(4)).upper()  # RA.2 → RA2, D 1 → D1

    if nd_num:
        nivel = f"ND{nd_num}"
    else:
        nivel = None

    # CG competencias usan D-codes; si el doc dice RA convertir a D
    if cod_comp.startswith("CG") and cod_ra.startswith("RA"):
        cod_ra = "D" + cod_ra[2:]

    partes = [cod_comp]
    if nivel:
        partes.append(nivel)
    partes.append(cod_ra)
    return ", ".join(partes)


def extraer_codigos_ra(doc: Document) -> list[str]:
    """
    Recoge todos los códigos RA del documento.
    Estrategia: busca en tablas de texto libre (1 col con 'Al final…')
    y en la primera columna de tablas de unidades.
    """
    codigos = set()

    for t in doc.tables:
        texto_completo = "\n".join(
            c.text for row in t.rows for c in row.cells
        )
        for m in RA_PATTERN.finditer(texto_completo):
            raw = m.group(0)
            cod = _normalizar_codigo_ra(raw)
            codigos.add(cod)

    return sorted(codigos)


def extraer_unidades(doc: Document, nombre_archivo: str) -> list[dict]:
    """
    Extrae unidades desde la tabla de 2-3 columnas con cabecera 'Resultado…'.
    Columnas: col0=RA/Desempeño | col1=Contenidos | col2=Indicador (opcional)
    """
    unidades = []
    for t in doc.tables:
        if not _es_tabla_unidades(t):
            continue
        for i, row in enumerate(t.rows[1:], start=1):
            cells = row.cells
            if len(cells) < 2:
                continue
            contenidos = _txt(cells[1]) if len(cells) > 1 else ""
            indicador  = _txt(cells[2]) if len(cells) > 2 else ""

            if not contenidos:
                continue

            num = None
            m = re.search(r"Unidad\s+(\d+|[IVX]+)", contenidos, re.IGNORECASE)
            if m:
                raw = m.group(1)
                if raw.isdigit():
                    num = int(raw)
                else:
                    rom = {"I":1,"II":2,"III":3,"IV":4,"V":5,"VI":6,"VII":7,"VIII":8}
                    num = rom.get(raw.upper())

            nombre_u = contenidos[:100]
            unidades.append({
                "orden": num or i,
                "nombre": nombre_u,
                "contenidos": contenidos,
                "indicador_logro": indicador,
            })
    return unidades


def extraer_metodologias(doc: Document) -> list[str]:
    metods = []
    for t in doc.tables:
        if not _es_tabla_metod(t):
            continue
        for row in t.rows:
            for cell in row.cells:
                txt = _txt(cell)
                if txt and len(txt) > 4 and txt not in metods:
                    metods.append(txt)
    return metods


def extraer_evaluaciones(doc: Document) -> list[dict]:
    evals = []
    for t in doc.tables:
        if not _es_tabla_eval(t):
            continue
        for row in t.rows[1:]:
            if len(row.cells) >= 2:
                tipo = _txt(row.cells[0])
                porc = _txt(row.cells[1])
                if tipo:
                    evals.append({"tipo": tipo, "porcentaje": porc})
    return evals


def _parsear_biblio_tabla(t: Table) -> list[dict]:
    """
    Tabla de bibliografía con columnas:
    Nº | Autor | Título | Editorial | Año | ISBN | Ejemplares
    """
    entradas = []
    for row in t.rows[1:]:
        cells = [_txt(c) for c in row.cells]
        if len(cells) < 3:
            continue
        # Detectar si primera celda es número o autor
        if len(cells) >= 7:
            numero, autor, titulo, editorial, anio, isbn, ejemplares = cells[:7]
        elif len(cells) >= 4:
            numero, autor, titulo, editorial = cells[:4]
            anio = isbn = ejemplares = ""
        else:
            numero = ""
            autor  = cells[0] if len(cells) > 0 else ""
            titulo = cells[1] if len(cells) > 1 else ""
            editorial = anio = isbn = ejemplares = ""

        if not titulo and not autor:
            continue
        entradas.append({
            "numero": numero, "autor": autor, "titulo": titulo,
            "editorial": editorial, "anio": anio,
            "isbn": isbn, "ejemplares": ejemplares,
        })
    return entradas


def extraer_bibliografia(doc: Document) -> tuple[list[dict], list[dict]]:
    basica = []
    complementaria = []
    for t in doc.tables:
        h = _txt_t(t, 0, 0).lower()
        if "básica" in h or "basica" in h:
            basica.extend(_parsear_biblio_tabla(t))
        elif "complementaria" in h:
            complementaria.extend(_parsear_biblio_tabla(t))
    return basica, complementaria


def extraer_linkografia(doc: Document, nombre_archivo: str) -> list[dict]:
    links = []
    for t in doc.tables:
        if not _es_tabla_linkografia(t):
            continue
        for row in t.rows[1:]:
            cells = [_txt(c) for c in row.cells]
            if len(cells) >= 4:
                tipo_doc = cells[0]
                autor    = cells[1]
                titulo   = cells[2]
                anio     = cells[3]
                url      = cells[4] if len(cells) > 4 else ""
            elif len(cells) >= 1:
                # tabla libre con URL en texto
                texto = cells[0]
                urls  = re.findall(r"https?://\S+", texto)
                for u in urls:
                    links.append({"url": u, "titulo_articulo": "", "descripcion": texto[:200]})
                continue
            else:
                continue

            if url or titulo:
                links.append({
                    "url": url,
                    "titulo_articulo": titulo,
                    "descripcion": f"{tipo_doc} {autor} {anio}".strip(),
                })
    # Tabla "otros recursos" con URL en texto libre
    for t in doc.tables:
        if _es_tabla_otros_recursos(t):
            texto = "\n".join(c.text for row in t.rows for c in row.cells)
            for url in re.findall(r"https?://\S+", texto):
                if url not in [l["url"] for l in links]:
                    links.append({"url": url, "titulo_articulo": "", "descripcion": "otros recursos"})
    return links


def extraer_otros_recursos(doc: Document) -> str:
    for t in doc.tables:
        if _es_tabla_otros_recursos(t):
            return "\n".join(c.text.strip() for row in t.rows for c in row.cells if c.text.strip())
    return ""


def extraer_responsables(doc: Document) -> dict:
    resp = {}
    for t in doc.tables:
        if _es_tabla_responsables(t):
            for row in t.rows:
                cells = [_txt(c) for c in row.cells]
                if len(cells) >= 2:
                    clave = cells[0].lower()
                    valor = cells[1]
                    if "docente" in clave:
                        resp["docente_a_cargo"] = valor
                    elif "responsable" in clave:
                        resp["responsable"] = valor
                    elif "versión" in clave or "version" in clave:
                        resp["version"] = valor
    return resp


# ── Parseo completo de un documento ──────────────────────────────────────────

def parsear_docx(ruta: Path) -> dict:
    nombre = ruta.name
    doc = Document(str(ruta))

    if len(doc.tables) < 2:
        raise ParseError(f"Documento no estándar: solo {len(doc.tables)} tabla(s)")

    ident = parsear_tabla_identificacion(doc, nombre)

    if not ident.get("codigo"):
        raise ParseError("No se encontró código de asignatura")

    codigos_ra = extraer_codigos_ra(doc)
    if not codigos_ra:
        log.warning("[%s] No se encontraron códigos RA", nombre)

    basica, complementaria = extraer_bibliografia(doc)
    if not basica and not complementaria:
        log.warning("[%s] Bibliografía vacía", nombre)

    responsables = extraer_responsables(doc)
    if not responsables:
        log.warning("[%s] No se encontraron responsables", nombre)

    return {
        "archivo":          nombre,
        "identificacion":   ident,
        "descripcion":      extraer_descripcion(doc),
        "codigos_ra":       codigos_ra,
        "unidades":         extraer_unidades(doc, nombre),
        "metodologias":     extraer_metodologias(doc),
        "evaluaciones":     extraer_evaluaciones(doc),
        "biblio_basica":    basica,
        "biblio_compl":     complementaria,
        "linkografia":      extraer_linkografia(doc, nombre),
        "otros_recursos":   extraer_otros_recursos(doc),
        "responsables":     responsables,
    }


# ── Inserción en BD ───────────────────────────────────────────────────────────

def resolver_ra_id(conn: sqlite3.Connection, codigo: str) -> Optional[int]:
    """Busca el RA en la BD por codigo_completo con varios intentos de normalización."""
    r = conn.execute(
        "SELECT id FROM resultados_aprendizaje WHERE codigo_completo = ?", (codigo,)
    ).fetchone()
    if r:
        return r[0]

    # Intentar sin espacios extra
    codigo_limpio = re.sub(r"\s+", " ", codigo).strip()
    r = conn.execute(
        "SELECT id FROM resultados_aprendizaje WHERE TRIM(codigo_completo) = ?",
        (codigo_limpio,)
    ).fetchone()
    return r[0] if r else None


def insertar_programa(conn: sqlite3.Connection, data: dict, dry_run: bool = False) -> str:
    ident  = data["identificacion"]
    codigo = ident.get("codigo", "").strip()
    nombre = ident.get("nombre", "").strip()

    if not codigo:
        return "sin_codigo"

    if not dry_run:
        # Upsert asignatura
        conn.execute("""
            INSERT INTO asignaturas
              (codigo, nombre, semestre, nivel, duracion, tipo, facultad, carrera,
               requisitos, horas_directa, horas_autonoma, semanas, creditos,
               descripcion, otros_recursos, version, archivo_origen)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            ON CONFLICT(codigo) DO UPDATE SET
              nombre         = excluded.nombre,
              semestre       = excluded.semestre,
              nivel          = excluded.nivel,
              duracion       = excluded.duracion,
              facultad       = excluded.facultad,
              carrera        = excluded.carrera,
              requisitos     = excluded.requisitos,
              horas_directa  = excluded.horas_directa,
              horas_autonoma = excluded.horas_autonoma,
              semanas        = excluded.semanas,
              creditos       = excluded.creditos,
              descripcion    = excluded.descripcion,
              otros_recursos = excluded.otros_recursos,
              version        = excluded.version,
              archivo_origen = excluded.archivo_origen
        """, (
            codigo, nombre,
            ident.get("semestre"),
            ident.get("nivel", ""),
            ident.get("duracion", ""),
            "disciplinar",
            ident.get("facultad", ""),
            ident.get("carrera", ""),
            ident.get("requisitos", ""),
            ident.get("horas_directa"),
            ident.get("horas_autonoma"),
            ident.get("semanas"),
            ident.get("creditos"),
            data.get("descripcion", ""),
            data.get("otros_recursos", ""),
            data.get("responsables", {}).get("version", ""),
            data.get("archivo", ""),
        ))

        asig_id = conn.execute(
            "SELECT id FROM asignaturas WHERE codigo = ?", (codigo,)
        ).fetchone()[0]

        # Limpiar datos previos (idempotente)
        for tbl in ("responsables", "unidades", "metodologias", "evaluaciones",
                    "bibliografia", "linkografia", "tributaciones"):
            conn.execute(f"DELETE FROM {tbl} WHERE asignatura_id = ?", (asig_id,))

        # Responsables
        resp = data.get("responsables", {})
        for rol in ("responsable", "docente_a_cargo"):
            if resp.get(rol):
                conn.execute(
                    "INSERT INTO responsables (asignatura_id, rol, nombre) VALUES (?,?,?)",
                    (asig_id, rol, resp[rol])
                )

        # Unidades
        for u in data.get("unidades", []):
            conn.execute("""
                INSERT INTO unidades (asignatura_id, orden, nombre, contenidos, indicador_logro)
                VALUES (?,?,?,?,?)
            """, (asig_id, u["orden"], u["nombre"], u["contenidos"], u["indicador_logro"]))

        # Metodologías
        for m in data.get("metodologias", []):
            if m:
                conn.execute(
                    "INSERT INTO metodologias (asignatura_id, descripcion) VALUES (?,?)",
                    (asig_id, m)
                )

        # Evaluaciones
        for ev in data.get("evaluaciones", []):
            if ev.get("tipo"):
                conn.execute(
                    "INSERT INTO evaluaciones (asignatura_id, tipo, porcentaje) VALUES (?,?,?)",
                    (asig_id, ev["tipo"], ev.get("porcentaje", ""))
                )

        # Bibliografía básica
        for b in data.get("biblio_basica", []):
            conn.execute("""
                INSERT INTO bibliografia
                  (asignatura_id, tipo, numero, autor, titulo, editorial, anio, isbn, ejemplares)
                VALUES (?,?,?,?,?,?,?,?,?)
            """, (asig_id, "basica",
                  b.get("numero",""), b.get("autor",""), b.get("titulo",""),
                  b.get("editorial",""), b.get("anio",""),
                  b.get("isbn",""), b.get("ejemplares","")))

        # Bibliografía complementaria
        for b in data.get("biblio_compl", []):
            conn.execute("""
                INSERT INTO bibliografia
                  (asignatura_id, tipo, numero, autor, titulo, editorial, anio, isbn, ejemplares)
                VALUES (?,?,?,?,?,?,?,?,?)
            """, (asig_id, "complementaria",
                  b.get("numero",""), b.get("autor",""), b.get("titulo",""),
                  b.get("editorial",""), b.get("anio",""),
                  b.get("isbn",""), b.get("ejemplares","")))

        # Linkografía
        for lk in data.get("linkografia", []):
            conn.execute("""
                INSERT INTO linkografia (asignatura_id, url, titulo_articulo, descripcion)
                VALUES (?,?,?,?)
            """, (asig_id, lk.get("url",""), lk.get("titulo_articulo",""), lk.get("descripcion","")))

        # Tributaciones
        ra_no_resueltos = []
        for cod in data.get("codigos_ra", []):
            ra_id = resolver_ra_id(conn, cod)
            if ra_id:
                conn.execute("""
                    INSERT OR IGNORE INTO tributaciones (asignatura_id, ra_id)
                    VALUES (?,?)
                """, (asig_id, ra_id))
            else:
                ra_no_resueltos.append(cod)

        if ra_no_resueltos:
            log.warning("[%s] RAs no resueltos en BD: %s", codigo, ra_no_resueltos)

    return "ok"


# ── Procesamiento masivo ──────────────────────────────────────────────────────

def procesar_todos(dry_run: bool = False, archivo_unico: Optional[Path] = None):
    if not RUTA_DB.exists():
        log.error("BD no encontrada en %s. Ejecuta init_db.py primero.", RUTA_DB)
        sys.exit(1)

    Path("data/output").mkdir(parents=True, exist_ok=True)

    if archivo_unico:
        archivos = [archivo_unico]
    else:
        archivos = sorted(CARPETA_DOCS.rglob("*.docx"))

    log.info("=" * 60)
    log.info("parse_word_to_db.py — %s documentos", len(archivos))
    if dry_run:
        log.info("MODO DRY-RUN: no se modifica la BD")
    log.info("=" * 60)

    conn = sqlite3.connect(str(RUTA_DB)) if not dry_run else None
    if conn:
        conn.execute("PRAGMA foreign_keys = ON")

    ok = errores = omitidos = 0
    errores_detalle = []

    for ruta in archivos:
        nombre = ruta.name

        # Verificar exclusiones
        if any(ex in nombre for ex in EXCLUIDOS):
            log.info("OMITIDO  %s (excluido explícitamente)", nombre)
            omitidos += 1
            continue

        try:
            data = parsear_docx(ruta)
            codigo = data["identificacion"].get("codigo", "?")

            if not dry_run:
                estado = insertar_programa(conn, data)
            else:
                estado = "dry-run"
                # En dry-run, reportar lo que se extrajo
                n_ra   = len(data["codigos_ra"])
                n_u    = len(data["unidades"])
                n_bb   = len(data["biblio_basica"])
                n_bc   = len(data["biblio_compl"])
                log.info("DRY  %-10s %-45s | RA:%d U:%d BB:%d BC:%d",
                         codigo, data["identificacion"].get("nombre","")[:44],
                         n_ra, n_u, n_bb, n_bc)
                ok += 1
                continue

            if estado == "ok":
                n_ra = len(data["codigos_ra"])
                n_u  = len(data["unidades"])
                log.info("OK   %-10s %-45s | RA:%d U:%d", codigo,
                         data["identificacion"].get("nombre","")[:44], n_ra, n_u)
                ok += 1
            else:
                log.warning("SKIP %-10s %s (%s)", codigo, nombre, estado)
                omitidos += 1

        except ParseError as e:
            log.error("ERROR %-50s → %s", nombre, e)
            errores += 1
            errores_detalle.append((nombre, str(e)))
        except Exception as e:
            log.exception("EXCEPCION %-45s → %s", nombre, e)
            errores += 1
            errores_detalle.append((nombre, str(e)))

    if conn:
        conn.commit()
        conn.close()

    log.info("")
    log.info("=" * 60)
    log.info("RESULTADO: %d OK | %d omitidos | %d errores | total %d",
             ok, omitidos, errores, len(archivos))
    if errores_detalle:
        log.info("DOCUMENTOS CON ERROR:")
        for nombre, motivo in errores_detalle:
            log.info("  - %-50s %s", nombre, motivo)

    if not dry_run and conn is None is False:
        _verificar_bd()


def _verificar_bd():
    conn = sqlite3.connect(str(RUTA_DB))
    counts = {}
    for tbl in ("asignaturas", "tributaciones", "unidades", "bibliografia",
                "linkografia", "metodologias", "evaluaciones"):
        counts[tbl] = conn.execute(f"SELECT COUNT(*) FROM {tbl}").fetchone()[0]
    conn.close()
    log.info("")
    log.info("VERIFICACIÓN BD:")
    for tbl, n in counts.items():
        log.info("  %-20s %d filas", tbl, n)


# ── Punto de entrada ──────────────────────────────────────────────────────────

if __name__ == "__main__":
    dry_run       = "--dry-run" in sys.argv
    archivo_unico = None

    if "--file" in sys.argv:
        idx = sys.argv.index("--file")
        if idx + 1 < len(sys.argv):
            archivo_unico = Path(sys.argv[idx + 1])
            if not archivo_unico.exists():
                log.error("Archivo no encontrado: %s", archivo_unico)
                sys.exit(1)

    procesar_todos(dry_run=dry_run, archivo_unico=archivo_unico)
