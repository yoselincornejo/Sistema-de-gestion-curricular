"""
parsear_programas.py — Parser comprensivo de programas de asignatura (.docx).

Extrae TODA la información de un documento Word de programa de asignatura UV
y la devuelve como un diccionario estructurado.

Estrategia: se recorre el cuerpo del documento en orden (párrafos y tablas
intercalados), asociando cada tabla con el último encabezado de sección visto.
Esto permite distinguir tablas que tienen estructura idéntica (p. ej.
"DESCRIPCIÓN" vs "APORTE AL PERFIL") pero pertenecen a secciones distintas.
"""

import re
from pathlib import Path

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph


def _limpiar(t):
    if not t:
        return ""
    return re.sub(r"[ \t]+", " ", t.replace("\xa0", " ")).strip()


def _iter_bloques(doc):
    """Itera párrafos y tablas del cuerpo en orden de aparición."""
    body = doc.element.body
    for child in body.iterchildren():
        if child.tag.endswith("}p"):
            yield Paragraph(child, doc)
        elif child.tag.endswith("}tbl"):
            yield Table(child, doc)


def _es_encabezado(texto):
    """Determina si un párrafo es un encabezado de sección."""
    t = texto.strip()
    if not t or len(t) > 70:
        return False
    return t.isupper() or t.rstrip().endswith(":")


def _semestre_de_path(ruta):
    """Extrae el número de semestre del path (ej: '.../Semestre 5/...' -> 5)."""
    m = re.search(r"[Ss]emestre\s*(\d+)", str(ruta))
    return int(m.group(1)) if m else None


def _a_float(texto):
    """Convierte '4,5' (formato español) a float; None si no es numérico."""
    if not texto:
        return None
    t = texto.strip().replace(",", ".")
    try:
        return float(t)
    except ValueError:
        return None


def _a_int(texto):
    if not texto:
        return None
    t = texto.strip().replace(",", ".")
    try:
        f = float(t)
        return int(round(f))
    except ValueError:
        return None


def _normalizar_clave(texto):
    return texto.rstrip(":").strip().lower()


# ── Identificación + horas ────────────────────────────────────────────────

_ALIAS_IDENT = {
    "facultad": "facultad",
    "carrera": "carrera",
    "carreras": "carrera",
    "nombre": "nombre",
    "codigo": "codigo",
    "código": "codigo",
    "codigos": "codigo",
    "códigos": "codigo",
    "nivel": "nivel",
    "duracion": "duracion",
    "duración": "duracion",
    "requisito(s)": "requisitos",
    "requisitos": "requisitos",
}


def _extraer_identificacion(tabla):
    """Tabla 0: 8x6 con identificación + bloque de horas."""
    ident = {
        "codigo": "", "nombre": "", "nivel": "", "duracion": "",
        "facultad": "", "carrera": "", "requisitos": "",
        "horas_directa": None, "horas_autonoma": None,
        "semanas": None, "creditos": None,
    }
    if tabla is None:
        return ident

    filas = [[_limpiar(c.text) for c in fila.cells] for fila in tabla.rows]

    # Pares etiqueta:valor (tolerando celdas fusionadas duplicadas)
    for celdas in filas:
        i = 0
        while i < len(celdas) - 1:
            clave_raw, valor = celdas[i], celdas[i + 1]
            if clave_raw and valor and clave_raw.endswith(":"):
                clave = _normalizar_clave(clave_raw)
                if clave in _ALIAS_IDENT and not ident.get(_ALIAS_IDENT[clave]):
                    ident[_ALIAS_IDENT[clave]] = valor
                    i += 2
                    continue
            i += 1

    # Bloque de horas: localizar la fila de etiquetas "(A) (B) ..." y tomar la
    # fila siguiente con los valores numéricos.
    idx_letras = None
    for ri, celdas in enumerate(filas):
        unicos = [c for c in celdas]
        if any("(a)" in c.lower() for c in unicos) and any("(b)" in c.lower() for c in unicos):
            idx_letras = ri
            break

    if idx_letras is not None and idx_letras + 1 < len(filas):
        valores = filas[idx_letras + 1]
        # Reducir a columnas únicas (las celdas fusionadas repiten texto)
        col_vals = []
        prev = None
        for v in valores:
            if v != prev:
                col_vals.append(v)
            prev = v
        # Estructura esperada: A, B, C(total), D(semanas), E(total sem), F(creditos)
        # Pero usamos posición por columna de la tabla original (6 cols):
        # col0=A directa, col1=B autonoma, col3=D semanas, col5=F creditos
        directa = _a_float(valores[0]) if len(valores) > 0 else None
        autonoma = _a_float(valores[1]) if len(valores) > 1 else None
        semanas = _a_int(valores[3]) if len(valores) > 3 else None
        creditos = _a_int(valores[5]) if len(valores) > 5 else None
        ident["horas_directa"] = directa
        ident["horas_autonoma"] = autonoma
        ident["semanas"] = semanas
        ident["creditos"] = creditos

    return ident


# ── Texto narrativo (descripcion / aporte) ────────────────────────────────

def _texto_de_tabla_o_parrafos(tablas, parrafos):
    partes = []
    for t in tablas:
        for fila in t.rows:
            for celda in fila.cells:
                txt = _limpiar(celda.text)
                if txt and txt not in partes:
                    partes.append(txt)
    for p in parrafos:
        txt = _limpiar(p)
        if txt and txt not in partes:
            partes.append(txt)
    return "\n".join(partes).strip()


# ── Unidades ──────────────────────────────────────────────────────────────

def _parsear_ra_codes(texto):
    """Extrae códigos RA de una celda. Separados por ';' o '.' o saltos."""
    codigos = []
    # Patrón: COMP[, ND/N nivel], RA/D código
    patron = r"\b[A-Z]{1,3}\d+\s*,\s*(?:ND?\d+\s*,\s*)?(?:RA\.?\d+|D\d+)"
    for m in re.finditer(patron, texto):
        cod = re.sub(r"\s+", " ", m.group(0)).strip().rstrip(".").strip()
        codigos.append(cod)
    return codigos


def _es_tabla_unidades(tabla):
    if len(tabla.rows) < 1:
        return False
    celdas = [_limpiar(c.text).lower() for c in tabla.rows[0].cells]
    header = " | ".join(celdas)
    tiene_ra = (
        "resultado de aprendizaje" in header
        or "resultados de aprendizaje" in header
        or any(c.strip() == "ra" or c.startswith("ra ") or c.startswith("ra\n") for c in celdas)
    )
    tiene_unidad = "unidad" in header
    return tiene_ra and tiene_unidad


def _extraer_unidades(tabla):
    unidades = []
    n_cols = len(tabla.rows[0].cells)
    for fila in tabla.rows[1:]:
        celdas = [_limpiar(c.text) for c in fila.cells]
        if not celdas or (not celdas[0] and (len(celdas) < 2 or not celdas[1])):
            continue
        # Saltar repeticiones del header
        if "resultado de aprendizaje" in celdas[0].lower():
            continue
        ra_codes = _parsear_ra_codes(celdas[0])
        unidad_texto = celdas[1] if n_cols >= 2 and len(celdas) > 1 else ""
        indicador = celdas[2] if n_cols >= 3 and len(celdas) > 2 else ""
        # Separar titulo de contenidos (primera línea = titulo)
        lineas = [l.strip() for l in unidad_texto.split("\n") if l.strip()]
        titulo = lineas[0] if lineas else ""
        contenidos = unidad_texto.strip()
        unidades.append({
            "ra_codes": ra_codes,
            "titulo": titulo,
            "contenidos": contenidos,
            "indicador_logro": indicador,
        })
    return unidades


# ── Metodologías ──────────────────────────────────────────────────────────

def _extraer_metodologias(tablas, parrafos):
    items = []
    textos = []
    for t in tablas:
        # Tabla con checkboxes (4 cols con X)
        es_checkbox = False
        for fila in t.rows:
            celdas = [_limpiar(c.text) for c in fila.cells]
            if any(c.lower() in ("x", "☒", "[x]") for c in celdas):
                es_checkbox = True
                break
        if es_checkbox:
            for fila in t.rows:
                celdas = [_limpiar(c.text) for c in fila.cells]
                for i, c in enumerate(celdas):
                    if c.lower() in ("x", "☒", "[x]"):
                        # la etiqueta suele ser la celda contigua (anterior o siguiente)
                        etiqueta = ""
                        for j in (i - 1, i + 1):
                            if 0 <= j < len(celdas) and celdas[j] and celdas[j].lower() not in ("x", "☒", "[x]"):
                                etiqueta = celdas[j]
                                break
                        if etiqueta:
                            items.append(etiqueta)
        else:
            for fila in t.rows:
                for celda in fila.cells:
                    txt = _limpiar(celda.text)
                    if txt and txt not in textos:
                        textos.append(txt)
    for p in parrafos:
        txt = _limpiar(p)
        if txt and txt not in textos:
            textos.append(txt)

    if items:
        return _dedup(items)

    # Texto libre: separar por líneas y por puntos
    bloque = "\n".join(textos)
    bloque = re.sub(r"(?i)^.*?listado de metodolog[ií]as\s*:?", "", bloque).strip()
    crudos = re.split(r"[\n.;]", bloque)
    for c in crudos:
        c = c.strip()
        if c and len(c) > 4:
            items.append(c)
    return _dedup(items)


def _dedup(seq):
    visto = set()
    out = []
    for x in seq:
        if x not in visto:
            visto.add(x)
            out.append(x)
    return out


# ── Evaluaciones ──────────────────────────────────────────────────────────

def _extraer_evaluaciones(tablas):
    evals = []
    for t in tablas:
        if len(t.rows) < 1:
            continue
        header = " | ".join(_limpiar(c.text).lower() for c in t.rows[0].cells)
        if "tipo de evaluaci" in header or "porcentaje" in header:
            for fila in t.rows[1:]:
                celdas = [_limpiar(c.text) for c in fila.cells]
                if len(celdas) < 2 or not celdas[0]:
                    continue
                tipo = celdas[0]
                porc = celdas[1]
                if "resultado de aprendizaje" in tipo.lower():
                    continue
                evals.append({"tipo": tipo, "porcentaje": porc})
        else:
            # Tabla narrativa de evaluación
            for fila in t.rows:
                celdas = [_limpiar(c.text) for c in fila.cells]
                if celdas and celdas[0] and len(celdas[0]) > 15:
                    evals.append({"tipo": celdas[0], "porcentaje": celdas[1] if len(celdas) > 1 else ""})
    return evals


# ── Bibliografía ──────────────────────────────────────────────────────────

def _fila_biblio_vacia(d):
    return not any(d.get(k, "").strip() for k in ("autor", "titulo", "editorial", "isbn"))


_BIBLIO_CAMPOS = {
    "autor": "autor",
    "título": "titulo", "titulo": "titulo",
    "editorial": "editorial",
    "año": "anio", "ano": "anio",
    "isbn": "isbn",
}


def _mapear_columnas_biblio(celdas_header):
    """Dado el header de columnas, devuelve {indice_col: campo}.
    La col 0 (vacía o 'N°') es el numero; la última col es ejemplares."""
    mapa = {0: "numero"}
    n = len(celdas_header)
    for i, c in enumerate(celdas_header):
        cl = c.lower().strip()
        for clave, campo in _BIBLIO_CAMPOS.items():
            if cl == clave or cl.startswith(clave):
                mapa[i] = campo
                break
        if "ejemplar" in cl or "biblioteca" in cl:
            mapa[i] = "ejemplares"
    # Si la última columna no quedó mapeada, asumir ejemplares
    if (n - 1) not in mapa:
        mapa[n - 1] = "ejemplares"
    return mapa


def _extraer_bibliografia(tablas):
    """Procesa una o más tablas (7 u 8 cols). Distingue básica/complementaria
    por las filas de header internas. Mapea columnas por su encabezado."""
    basica, complementaria = [], []
    for t in tablas:
        destino = None
        mapa = None
        for fila in t.rows:
            celdas = [_limpiar(c.text) for c in fila.cells]
            texto_fila = " ".join(celdas).upper()
            up = texto_fila.replace("Á", "A")
            if "BIBLIOGRAF" in up and "BASICA" in up:
                destino = basica
                continue
            if "BIBLIOGRAF" in up and "COMPLEMENTARIA" in up:
                destino = complementaria
                continue
            # Header de columnas (contiene 'Autor')
            if any(c.lower() == "autor" for c in celdas):
                mapa = _mapear_columnas_biblio(celdas)
                continue
            if destino is None:
                destino = basica
            if mapa is None:
                mapa = {0: "numero", 1: "autor", 2: "titulo", 3: "editorial",
                        4: "anio", 5: "isbn", len(celdas) - 1: "ejemplares"}
            entrada = {k: "" for k in ("numero", "autor", "titulo", "editorial", "anio", "isbn", "ejemplares")}
            for i, campo in mapa.items():
                if i < len(celdas):
                    entrada[campo] = celdas[i]
            if _fila_biblio_vacia(entrada):
                continue
            destino.append(entrada)
    return basica, complementaria


# ── Linkografía ───────────────────────────────────────────────────────────

def _extraer_linkografia(tablas, parrafos):
    links = []
    textos = []
    url_re_busca = re.compile(r"https?://\S+|www\.\S+")
    for t in tablas:
        header = [_limpiar(c.text).lower() for c in t.rows[0].cells]
        idx_url = next((i for i, h in enumerate(header) if "direcci" in h or "electrón" in h or "electron" in h or "url" in h), None)
        idx_tit = next((i for i, h in enumerate(header) if "título" in h or "titulo" in h or "documento" in h), None)
        if idx_url is not None and len(t.rows) > 1:
            # Tabla estructurada de linkografía
            for fila in t.rows[1:]:
                celdas = [_limpiar(c.text) for c in fila.cells]
                url = celdas[idx_url] if idx_url < len(celdas) else ""
                desc = celdas[idx_tit] if idx_tit is not None and idx_tit < len(celdas) else ""
                if url.strip() or url_re_busca.search(" ".join(celdas)):
                    if not url.strip():
                        m = url_re_busca.search(" ".join(celdas))
                        url = m.group(0) if m else ""
                    if url.strip():
                        links.append({"url": url.strip().rstrip(".,);"), "descripcion": desc})
            continue
        for fila in t.rows:
            for celda in fila.cells:
                txt = _limpiar(celda.text)
                if txt:
                    textos.append(txt)
    for p in parrafos:
        txt = _limpiar(p)
        if txt:
            textos.append(txt)

    url_re = re.compile(r"https?://\S+|www\.\S+")
    for txt in _dedup(textos):
        urls = url_re.findall(txt)
        if urls:
            for u in urls:
                desc = url_re.sub("", txt).strip(" .-—:")
                links.append({"url": u.rstrip(".,);"), "descripcion": desc})
        # líneas sin URL se ignoran como linkografía
    return links


# ── Responsables ──────────────────────────────────────────────────────────

_ALIAS_RESP = {
    "responsable(s) del programa": "responsable",
    "responsables del programa": "responsable",
    "responsable del programa": "responsable",
    "docente(s) a cargo": "docente_a_cargo",
    "docentes a cargo": "docente_a_cargo",
    "docente a cargo": "docente_a_cargo",
    "versión / fecha de actualización": "version",
    "version / fecha de actualizacion": "version",
    "versión": "version",
    "version": "version",
}


def _extraer_responsables(tablas):
    datos = {"responsable": "", "docente_a_cargo": "", "version": ""}
    for t in tablas:
        for fila in t.rows:
            celdas = [_limpiar(c.text) for c in fila.cells]
            if len(celdas) >= 2 and celdas[0]:
                clave = celdas[0].rstrip(":").strip().lower()
                if clave in _ALIAS_RESP:
                    datos[_ALIAS_RESP[clave]] = celdas[1]
    return datos


# ── Función principal ─────────────────────────────────────────────────────

def parsear_docx(ruta_docx):
    """Parsea un programa .docx y devuelve un dict comprensivo."""
    doc = Document(ruta_docx)
    nombre_archivo = Path(ruta_docx).name

    # Recorrer cuerpo asociando tablas a su sección (último encabezado)
    seccion_actual = None
    parrafos_seccion = {}  # seccion -> [texto, ...]
    tablas_seccion = {}    # seccion -> [Table, ...]
    todas_tablas = []

    def _key(texto):
        u = texto.upper()
        if "DESCRIPCIÓN DE LA ASIGNATURA" in u or "DESCRIPCION DE LA ASIGNATURA" in u:
            return "descripcion"
        if "APORTE AL PERFIL" in u:
            return "aporte"
        if "RESULTADO" in u and "APRENDIZAJE" in u and "UNIDAD" not in u:
            return "ra"
        if "UNIDAD" in u and "APRENDIZAJE" in u:
            return "unidades"
        if "ESTRATEGIA" in u and ("ENSEÑANZA" in u or "ENSENANZA" in u):
            return "metodologia"
        if "EVALUACI" in u:
            return "evaluacion"
        if "BIBLIOGRAF" in u:
            return "bibliografia"
        if "LINKOGRAF" in u:
            return "linkografia"
        if "OTROS RECURSOS" in u:
            return "otros_recursos"
        if "DATOS ACTUALIZACI" in u or "ACTUALIZACIÓN" in u:
            return "actualizacion"
        if "IDENTIFICACIÓN" in u or "IDENTIFICACION" in u:
            return "identificacion"
        return None

    for bloque in _iter_bloques(doc):
        if isinstance(bloque, Paragraph):
            txt = _limpiar(bloque.text)
            if _es_encabezado(txt):
                k = _key(txt)
                if k:
                    seccion_actual = k
            elif txt and seccion_actual:
                parrafos_seccion.setdefault(seccion_actual, []).append(txt)
        else:  # Table
            todas_tablas.append(bloque)
            if seccion_actual:
                tablas_seccion.setdefault(seccion_actual, []).append(bloque)

    # Identificación: tabla 0
    tabla_ident = todas_tablas[0] if todas_tablas else None
    identificacion = _extraer_identificacion(tabla_ident)

    # Descripción
    descripcion = _texto_de_tabla_o_parrafos(
        tablas_seccion.get("descripcion", []),
        parrafos_seccion.get("descripcion", []),
    )

    # Aporte al perfil
    aporte = _texto_de_tabla_o_parrafos(
        tablas_seccion.get("aporte", []),
        parrafos_seccion.get("aporte", []),
    )

    # Unidades: usar tabla de la sección, o buscar por header en cualquier tabla
    tabla_unidades = None
    for t in tablas_seccion.get("unidades", []):
        if _es_tabla_unidades(t):
            tabla_unidades = t
            break
    if tabla_unidades is None:
        for t in todas_tablas:
            if _es_tabla_unidades(t):
                tabla_unidades = t
                break
    unidades = _extraer_unidades(tabla_unidades) if tabla_unidades else []

    # ra_codes globales
    ra_codes_global = _dedup([c for u in unidades for c in u["ra_codes"]])

    # Metodologías
    metodologias = _extraer_metodologias(
        tablas_seccion.get("metodologia", []),
        parrafos_seccion.get("metodologia", []),
    )

    # Evaluaciones
    evaluaciones = _extraer_evaluaciones(tablas_seccion.get("evaluacion", []))

    # Bibliografía
    biblio_basica, biblio_comp = _extraer_bibliografia(tablas_seccion.get("bibliografia", []))

    # Linkografía
    linkografia = _extraer_linkografia(
        tablas_seccion.get("linkografia", []),
        parrafos_seccion.get("linkografia", []),
    )

    # Otros recursos
    otros = _texto_de_tabla_o_parrafos(
        tablas_seccion.get("otros_recursos", []),
        parrafos_seccion.get("otros_recursos", []),
    )

    # Responsables
    responsables = _extraer_responsables(tablas_seccion.get("actualizacion", []))
    if not any(responsables.values()):
        # fallback: buscar en todas las tablas
        responsables = _extraer_responsables(todas_tablas)

    return {
        "archivo": nombre_archivo,
        "semestre": _semestre_de_path(ruta_docx),
        "identificacion": identificacion,
        "descripcion": descripcion,
        "aporte_perfil": aporte,
        "responsables": responsables,
        "ra_codes": ra_codes_global,
        "unidades": unidades,
        "metodologias": metodologias,
        "evaluaciones": evaluaciones,
        "bibliografia_basica": biblio_basica,
        "bibliografia_complementaria": biblio_comp,
        "linkografia": linkografia,
        "otros_recursos": otros,
    }


if __name__ == "__main__":
    import json
    import sys
    ruta = sys.argv[1] if len(sys.argv) > 1 else "data/programas/Semestre 5/IMAT 321 ANALISIS.docx"
    prog = parsear_docx(ruta)
    print(json.dumps(prog, ensure_ascii=False, indent=2))
