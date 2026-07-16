"""Password gate simple. Una sola contraseña compartida, guardada en secrets."""
import hmac

import streamlit as st


def _password_correcta(intento: str) -> bool:
    esperado = st.secrets.get("APP_PASSWORD", "")
    if not esperado:
        return False
    return hmac.compare_digest(intento, esperado)


def exigir_login() -> None:
    """Bloquea el resto de la app hasta que se ingrese la contraseña correcta."""
    if st.session_state.get("autenticado"):
        return

    st.title("🔒 Tablero Comercial CEA")
    st.caption("Acceso restringido al equipo directivo de CEA Universidad.")

    with st.form("form_login"):
        intento = st.text_input("Contraseña", type="password")
        enviado = st.form_submit_button("Entrar", use_container_width=True)

    if enviado:
        if _password_correcta(intento):
            st.session_state["autenticado"] = True
            st.rerun()
        else:
            st.error("Contraseña incorrecta.")

    st.stop()


def cerrar_sesion() -> None:
    st.session_state["autenticado"] = False
    st.rerun()
