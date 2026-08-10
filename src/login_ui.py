import panel as pn
from auth import get_todos, verificar_credenciales

pn.extension()

CSS_LOGIN = """
:host(.login-card) {
    max-width: 420px;
    margin: 80px auto 0 auto;
    border-radius: 14px !important;
    box-shadow: 0 8px 32px rgba(0,0,0,0.18) !important;
    overflow: hidden;
}
"""


def crear_login(on_success=None):
    """
    Devuelve un pn.Column con el formulario de login.

    Parámetro opcional:
        on_success(rol, usuario): callback llamado cuando las credenciales
        son correctas. Se usará para conectar con la app principal.
    """
    # ── Widgets ──────────────────────────────────────────────────
    w_usuario = pn.widgets.Select(
        name="Usuario",
        options=list(get_todos().keys()),
        sizing_mode="stretch_width",
    )
    w_password = pn.widgets.PasswordInput(
        name="Contraseña",
        placeholder="Ingresa tu contraseña...",
        sizing_mode="stretch_width",
    )
    btn_ingresar = pn.widgets.Button(
        name="Ingresar",
        button_type="primary",
        sizing_mode="stretch_width",
        height=44,
    )
    msg_error = pn.pane.HTML(
        "",
        sizing_mode="stretch_width",
        styles={"min-height": "24px"},
    )

    # ── Lógica del botón ─────────────────────────────────────────
    def on_ingresar(event):
        rol = verificar_credenciales(w_usuario.value, w_password.value)
        if rol:
            msg_error.object = ""
            w_password.value = ""
            if on_success:
                on_success(rol, w_usuario.value)
        else:
            msg_error.object = (
                '<div style="background:#FEE2E2;border:1px solid #F87171;'
                'border-radius:6px;padding:8px 12px;color:#991B1B;font-size:13px">'
                '⚠️ Contraseña incorrecta. Inténtalo nuevamente.</div>'
            )
            w_password.value = ""

    btn_ingresar.on_click(on_ingresar)

    # ── Encabezado ───────────────────────────────────────────────
    header = pn.pane.HTML("""
        <div style="background:linear-gradient(135deg,#1e3a5f 0%,#2d6a9f 100%);
                    padding:28px 32px 20px 32px;margin:-16px -16px 20px -16px;
                    border-radius:14px 14px 0 0">
            <div style="display:flex;align-items:center;gap:14px">
                <div style="background:#fff;color:#1e3a5f;font-weight:800;
                            font-size:18px;width:46px;height:46px;border-radius:10px;
                            display:flex;align-items:center;justify-content:center">
                    ICM
                </div>
                <div>
                    <div style="color:#fff;font-size:18px;font-weight:700;
                                letter-spacing:.3px">Iniciar Sesión</div>
                    <div style="color:#93c5fd;font-size:12px;margin-top:2px">
                        Sistema de Gestión Curricular · UV
                    </div>
                </div>
            </div>
        </div>
    """, sizing_mode="stretch_width")

    # ── Tarjeta ──────────────────────────────────────────────────
    card = pn.Column(
        header,
        w_usuario,
        w_password,
        btn_ingresar,
        msg_error,
        sizing_mode="stretch_width",
        styles={
            "background": "#ffffff",
            "border": "1px solid #e2e8f0",
            "border-radius": "14px",
            "padding": "16px 24px 24px 24px",
            "box-shadow": "0 8px 32px rgba(0,0,0,0.15)",
            "max-width": "420px",
        },
    )

    # Contenedor centrado en la pantalla
    contenedor = pn.Column(
        pn.layout.VSpacer(),
        pn.Row(pn.layout.HSpacer(), card, pn.layout.HSpacer()),
        pn.layout.VSpacer(),
        sizing_mode="stretch_both",
        styles={"background": "#f1f5f9", "min-height": "100vh"},
    )

    return contenedor


# ── Preview standalone ────────────────────────────────────────────
if __name__ == "__main__":
    def demo_callback(rol, usuario):
        print(f"Login exitoso: {usuario} ({rol})")

    app = crear_login(on_success=demo_callback)
    app.show(port=5007, title="Login — SGC")
