# auth.py — Lógica de autenticación del Sistema de Gestión Curricular
import json
from pathlib import Path

_USUARIOS_FILE = Path(__file__).parent.parent / "data" / "usuarios.json"

_USUARIOS_BUILTIN = {
    "Administradora": {
        "password": "admin123",
        "rol": "admin",
        "nombre_completo": "Administradora",
    },
    "Prof. X": {
        "password": "prof123",
        "rol": "user",
        "nombre_completo": "Prof. X",
    },
    "Jefe de Carrera": {
        "password": "jefe123",
        "rol": "superuser",
        "nombre_completo": "Jefe de Carrera",
    },
    "CCP": {
        "password": "ccp123",
        "rol": "superuser",
        "nombre_completo": "Comisión Curricular Permanente",
    },
    "Dirección": {
        "password": "dir123",
        "rol": "superuser",
        "nombre_completo": "Dirección",
    },
}

ROLES = {
    "superuser": "Super usuario — acceso total, incluyendo configuración",
    "admin":     "Administrador — edición completa de programas y datos",
    "user":      "Usuario — visualización y edición limitada (solo sus propias asignaturas)",
}


def _cargar_usuarios() -> dict:
    if _USUARIOS_FILE.exists():
        try:
            return json.loads(_USUARIOS_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return dict(_USUARIOS_BUILTIN)


USUARIOS: dict = _cargar_usuarios()


def guardar_usuarios() -> None:
    """Persiste USUARIOS en disco (data/usuarios.json)."""
    _USUARIOS_FILE.parent.mkdir(parents=True, exist_ok=True)
    _USUARIOS_FILE.write_text(
        json.dumps(USUARIOS, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def verificar_credenciales(usuario: str, password: str) -> str | None:
    datos = USUARIOS.get(usuario)
    if datos is None:
        return None
    if password == datos["password"]:
        return datos["rol"]
    return None


def get_usuario_info(usuario: str) -> dict | None:
    datos = USUARIOS.get(usuario)
    if datos is None:
        return None
    return {k: v for k, v in datos.items() if k != "password"}
