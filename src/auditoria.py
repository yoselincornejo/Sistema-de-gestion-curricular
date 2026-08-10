# auditoria.py — Sistema de registro de acciones de usuarios
#
# Uso:
#   from auditoria import registrar_accion
#   registrar_accion("Administradora", "LOGIN", "Inicio de sesión exitoso")
#   registrar_accion("Prof. X", "MODIFICACION", "Cambió bibliografía", entidad="IMAT 211")

import sqlite3
from pathlib import Path

RUTA_DB = Path(__file__).parent.parent / "data" / "sistema.db"

# Acciones estándar (referencia — no restrictivo, se puede pasar cualquier string)
ACCION_LOGIN         = "LOGIN"
ACCION_LOGOUT        = "LOGOUT"
ACCION_CREACION      = "CREACION"
ACCION_MODIFICACION  = "MODIFICACION"
ACCION_ELIMINACION   = "ELIMINACION"
ACCION_DESCARGA      = "DESCARGA"


def _conexion() -> sqlite3.Connection:
    conn = sqlite3.connect(str(RUTA_DB))
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def crear_tabla_auditoria() -> None:
    """Crea la tabla registro_acciones si no existe (idempotente)."""
    conn = _conexion()
    try:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS registro_acciones (
                id                INTEGER  PRIMARY KEY AUTOINCREMENT,
                timestamp         DATETIME DEFAULT CURRENT_TIMESTAMP,
                usuario           TEXT     NOT NULL,
                accion            TEXT     NOT NULL,
                detalles          TEXT,
                entidad_afectada  TEXT     DEFAULT ''
            )
        """)
        conn.commit()
    finally:
        conn.close()


def registrar_accion(
    usuario: str,
    accion: str,
    detalles: str = "",
    entidad: str = "",
) -> None:
    """
    Inserta un registro de auditoría en la tabla registro_acciones.

    Parámetros:
        usuario  — nombre del usuario que realizó la acción
        accion   — tipo de acción (LOGIN, MODIFICACION, CREACION, etc.)
        detalles — descripción libre o JSON con los cambios realizados
        entidad  — nombre de la asignatura u objeto afectado (opcional)
    """
    conn = _conexion()
    try:
        conn.execute(
            """
            INSERT INTO registro_acciones (usuario, accion, detalles, entidad_afectada)
            VALUES (?, ?, ?, ?)
            """,
            (usuario, accion, detalles, entidad),
        )
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def get_registros(
    usuario: str | None = None,
    accion: str | None = None,
    entidad: str | None = None,
    limit: int = 200,
) -> list[dict]:
    """
    Consulta el registro de auditoría con filtros opcionales.

    Retorna lista de dicts ordenada del más reciente al más antiguo.
    """
    conn = _conexion()
    conn.row_factory = sqlite3.Row
    try:
        condiciones = []
        params: list = []
        if usuario:
            condiciones.append("usuario = ?")
            params.append(usuario)
        if accion:
            condiciones.append("accion = ?")
            params.append(accion)
        if entidad:
            condiciones.append("entidad_afectada = ?")
            params.append(entidad)

        where = ("WHERE " + " AND ".join(condiciones)) if condiciones else ""
        filas = conn.execute(
            f"""
            SELECT id, timestamp, usuario, accion, detalles, entidad_afectada
            FROM registro_acciones
            {where}
            ORDER BY id DESC
            LIMIT ?
            """,
            [*params, limit],
        ).fetchall()
        return [dict(r) for r in filas]
    finally:
        conn.close()


# Inicialización automática al importar el módulo
crear_tabla_auditoria()
