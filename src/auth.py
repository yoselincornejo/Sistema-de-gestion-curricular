# auth.py — Lógica de autenticación del Sistema de Gestión Curricular
import sqlite3
from pathlib import Path

_RUTA_DB = Path(__file__).parent.parent / "data" / "sistema.db"

_USUARIOS_BUILTIN = [
    ("Administradora", "admin123",  "admin",     "Administradora"),
    ("Jefe de Carrera","jefe123",   "superuser", "Jefe de Carrera"),
    ("CCP",            "ccp123",    "superuser", "Comisión Curricular Permanente"),
    ("Dirección",      "dir123",    "superuser", "Dirección"),
]

ROLES = {
    "superuser": "Super usuario — acceso total, incluyendo configuración",
    "admin":     "Administrador — edición completa de programas y datos",
    "user":      "Usuario — visualización y edición limitada (solo sus propias asignaturas)",
}


def _conexion() -> sqlite3.Connection:
    conn = sqlite3.connect(str(_RUTA_DB))
    conn.execute("PRAGMA foreign_keys = ON")
    conn.row_factory = sqlite3.Row
    return conn


def _init_tabla() -> None:
    conn = _conexion()
    try:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS usuarios (
                username        TEXT PRIMARY KEY,
                password        TEXT NOT NULL,
                rol             TEXT NOT NULL DEFAULT 'user',
                nombre_completo TEXT NOT NULL DEFAULT ''
            )
        """)
        # Insertar builtin solo si la tabla está vacía
        n = conn.execute("SELECT COUNT(*) FROM usuarios").fetchone()[0]
        if n == 0:
            conn.executemany(
                "INSERT OR IGNORE INTO usuarios (username, password, rol, nombre_completo) VALUES (?,?,?,?)",
                _USUARIOS_BUILTIN,
            )
        conn.commit()
    finally:
        conn.close()


_init_tabla()


# ── API pública ───────────────────────────────────────────────────

def get_todos() -> dict:
    """Retorna dict {username: {password, rol, nombre_completo}}."""
    conn = _conexion()
    try:
        filas = conn.execute("SELECT username, password, rol, nombre_completo FROM usuarios").fetchall()
        return {r["username"]: {"password": r["password"], "rol": r["rol"],
                                "nombre_completo": r["nombre_completo"]} for r in filas}
    finally:
        conn.close()


def verificar_credenciales(usuario: str, password: str) -> str | None:
    conn = _conexion()
    try:
        fila = conn.execute("SELECT rol, password FROM usuarios WHERE username=?", (usuario,)).fetchone()
        if fila and password == fila["password"]:
            return fila["rol"]
        return None
    finally:
        conn.close()


def get_usuario_info(usuario: str) -> dict | None:
    conn = _conexion()
    try:
        fila = conn.execute(
            "SELECT rol, nombre_completo FROM usuarios WHERE username=?", (usuario,)
        ).fetchone()
        if fila:
            return {"rol": fila["rol"], "nombre_completo": fila["nombre_completo"]}
        return None
    finally:
        conn.close()


def crear_usuario(username: str, password: str, rol: str, nombre_completo: str) -> None:
    conn = _conexion()
    try:
        conn.execute(
            "INSERT INTO usuarios (username, password, rol, nombre_completo) VALUES (?,?,?,?)",
            (username, password, rol, nombre_completo or username),
        )
        conn.commit()
    finally:
        conn.close()


def actualizar_usuario(username: str, nombre_completo: str, rol: str, password: str | None = None) -> None:
    conn = _conexion()
    try:
        if password:
            conn.execute(
                "UPDATE usuarios SET nombre_completo=?, rol=?, password=? WHERE username=?",
                (nombre_completo, rol, password, username),
            )
        else:
            conn.execute(
                "UPDATE usuarios SET nombre_completo=?, rol=? WHERE username=?",
                (nombre_completo, rol, username),
            )
        conn.commit()
    finally:
        conn.close()


def eliminar_usuario(username: str) -> None:
    conn = _conexion()
    try:
        conn.execute("DELETE FROM usuarios WHERE username=?", (username,))
        conn.commit()
    finally:
        conn.close()


# Compatibilidad con código legacy que accede a _auth.USUARIOS directamente
@property
def _usuarios_compat():
    return get_todos()

USUARIOS: dict = {}  # stub — usar get_todos() en su lugar
