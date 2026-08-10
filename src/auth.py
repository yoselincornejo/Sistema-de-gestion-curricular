# auth.py — Lógica de autenticación del Sistema de Gestión Curricular
#
# SEGURIDAD (producción):
#   Las contraseñas aquí están en texto plano solo para desarrollo/pruebas.
#   Para producción, reemplazar cada valor "password" por un hash generado con:
#       import bcrypt
#       bcrypt.hashpw("contraseña".encode(), bcrypt.gensalt()).decode()
#   Y en verificar_credenciales(), usar:
#       bcrypt.checkpw(password.encode(), usuario["password"].encode())

USUARIOS = {
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

# Roles disponibles y sus permisos (referencia para la UI)
ROLES = {
    "superuser": "Super usuario — acceso total, incluyendo configuración",
    "admin":     "Administrador — edición completa de programas y datos",
    "user":      "Usuario — visualización y edición limitada",
}


def verificar_credenciales(usuario: str, password: str) -> str | None:
    """
    Verifica las credenciales del usuario.

    Retorna el 'rol' (str) si las credenciales son correctas,
    o None si el usuario no existe o la contraseña es incorrecta.

    En producción: reemplazar la comparación directa por bcrypt.checkpw().
    """
    datos = USUARIOS.get(usuario)
    if datos is None:
        return None
    # TODO (producción): bcrypt.checkpw(password.encode(), datos["password"].encode())
    if password == datos["password"]:
        return datos["rol"]
    return None


def get_usuario_info(usuario: str) -> dict | None:
    """Retorna el dict completo del usuario (sin exponer la contraseña)."""
    datos = USUARIOS.get(usuario)
    if datos is None:
        return None
    return {k: v for k, v in datos.items() if k != "password"}
