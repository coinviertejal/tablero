from __future__ import annotations

from datetime import datetime
from pathlib import Path

import streamlit as st

from data import MUNICIPIOS_JALISCO
from db import client_with_token, configured, public_client, upload_files, valid_official_email

st.set_page_config(page_title="COINVIERTE | Gestión Institucional", page_icon="🏛️", layout="wide")

st.markdown("""
<style>
:root { --azul:#143b5d; --dorado:#c99a45; --fondo:#f4f6f8; }
.stApp { background: var(--fondo); }
.hero { background:linear-gradient(120deg,#143b5d,#1e5c79); padding:2.2rem; border-radius:18px; color:white; margin-bottom:1.2rem; }
.hero h1 { margin:0; font-size:2.35rem; }
.hero p { opacity:.9; margin:.45rem 0 0; }
.card { background:white; border:1px solid #e3e7eb; border-top:5px solid #c99a45; border-radius:14px; padding:1.2rem; min-height:155px; }
.muted { color:#5d6872; }
div[data-testid="stForm"] { background:white; padding:1.3rem; border-radius:14px; border:1px solid #e3e7eb; }
</style>
""", unsafe_allow_html=True)


def logo_header():
    logo = Path("assets/logo_coinvierte.png")
    if logo.exists():
        st.image(str(logo), width=260)
    st.markdown("""
    <div class="hero"><h1>COINVIERTE</h1>
    <p>Plataforma institucional para la gestión, documentación y seguimiento de proyectos.</p></div>
    """, unsafe_allow_html=True)


def login():
    logo_header()
    st.subheader("Acceso institucional")
    if not configured():
        st.warning("Modo demostración: falta configurar Supabase. Puedes ingresar con cualquier correo @jalisco.gob.mx.")
    with st.form("login"):
        email = st.text_input("Correo institucional", placeholder="nombre@jalisco.gob.mx")
        password = st.text_input("Contraseña", type="password")
        submitted = st.form_submit_button("Ingresar", type="primary", use_container_width=True)
    if submitted:
        if not valid_official_email(email):
            st.error("El acceso está limitado a cuentas @jalisco.gob.mx.")
        elif not configured():
            st.session_state.user = {"email": email.lower(), "id": "demo"}
            st.rerun()
        else:
            try:
                auth = public_client().auth.sign_in_with_password({"email": email, "password": password})
                if not auth.user or not valid_official_email(auth.user.email or ""):
                    public_client().auth.sign_out()
                    st.error("La cuenta no pertenece al dominio autorizado.")
                    return
                st.session_state.user = {"email": auth.user.email, "id": str(auth.user.id)}
                st.session_state.access_token = auth.session.access_token
                st.rerun()
            except Exception:
                st.error("No fue posible iniciar sesión. Verifica el correo y la contraseña.")


def landing():
    logo_header()
    st.write(f"Sesión: **{st.session_state.user['email']}**")
    cols = st.columns(3)
    sections = [
        ("📁", "Programas / Proyectos", "Alta, consulta, edición y seguimiento de expedientes."),
        ("🏛️", "Junta de Gobierno", "Actas, acuerdos y documentación de las sesiones."),
        ("👥", "Comités", "Integración, sesiones, actas y dictaminación."),
    ]
    for col, (icon, title, text) in zip(cols, sections):
        with col:
            st.markdown(f'<div class="card"><h3>{icon} {title}</h3><p class="muted">{text}</p></div>', unsafe_allow_html=True)
            if st.button(f"Abrir {title}", key=title, use_container_width=True):
                st.session_state.page = title
                st.rerun()


def objective_fields(existing=None):
    existing = existing or [""]
    if "objective_count" not in st.session_state:
        st.session_state.objective_count = max(1, len(existing))
    values = []
    for index in range(st.session_state.objective_count):
        default = existing[index] if index < len(existing) else ""
        values.append(st.text_area(f"Objetivo específico {index + 1}", value=default, key=f"obj_{index}"))
    col1, col2 = st.columns([1, 4])
    if col1.form_submit_button("＋ Agregar objetivo"):
        st.session_state.objective_count += 1
        st.rerun()
    if st.session_state.objective_count > 1 and col2.form_submit_button("Quitar último"):
        st.session_state.objective_count -= 1
        st.rerun()
    return values


def project_form(direction: str, project=None):
    project = project or {}
    st.subheader("Editar proyecto" if project else "Dar de alta nuevo proyecto")
    with st.form("project_form", clear_on_submit=False):
        st.markdown("### Información General")
        name = st.text_input("Nombre del proyecto *", value=project.get("nombre", ""))
        applicant = st.text_input("Nombre del solicitante *", value=project.get("solicitante", ""))
        municipality = st.selectbox("Municipio de ejecución *", MUNICIPIOS_JALISCO,
                                    index=MUNICIPIOS_JALISCO.index(project["municipio"]) if project.get("municipio") in MUNICIPIOS_JALISCO else 0)
        c1, c2 = st.columns(2)
        year = c1.number_input("Año de inicio *", min_value=2000, max_value=2100,
                               value=int(project.get("anio_inicio", datetime.now().year)), step=1)
        amount = c2.number_input("Monto (MXN) *", min_value=0.0, value=float(project.get("monto", 0)), step=1000.0, format="%.2f")
        general = st.text_area("Objetivo general *", value=project.get("objetivo_general", ""), height=130)
        st.markdown("#### Objetivos específicos")
        objectives = objective_fields(project.get("objetivos_especificos"))

        st.markdown("### Gestión Documental")
        legal = st.file_uploader("Documentación jurídica", accept_multiple_files=True, key="legal")
        auxiliary = st.file_uploader("Documentación auxiliar", accept_multiple_files=True, key="aux")
        committee = st.file_uploader("Acta del Comité de Dictaminación", accept_multiple_files=False, key="committee")
        board = st.file_uploader("Acta de aprobación de Junta de Gobierno", accept_multiple_files=False, key="board")
        agreement = st.file_uploader("Convenio de colaboración", accept_multiple_files=False, key="agreement")

        st.markdown("### Evidencia fotográfica")
        photos = st.file_uploader("Fotografías (máximo 5 MB por archivo)", type=["jpg", "jpeg", "png", "webp"],
                                  accept_multiple_files=True, key="photos")

        st.markdown("### Monitoreo y Seguimiento")
        st.info("Bloque preparado. En el siguiente paso definiremos indicadores, metas, avances, fechas y responsables.")
        save = st.form_submit_button("Guardar proyecto", type="primary", use_container_width=True)

    if save:
        errors = []
        if not name.strip() or not applicant.strip() or not general.strip():
            errors.append("Completa todos los campos obligatorios.")
        clean_objectives = [o.strip() for o in objectives if o.strip()]
        if not clean_objectives:
            errors.append("Agrega al menos un objetivo específico.")
        oversized = [f.name for f in photos if f.size > 5 * 1024 * 1024]
        if oversized:
            errors.append("Estas fotografías exceden 5 MB: " + ", ".join(oversized))
        if errors:
            for error in errors:
                st.error(error)
            return
        payload = {"direccion": direction, "nombre": name.strip(), "solicitante": applicant.strip(),
                   "municipio": municipality, "anio_inicio": int(year), "monto": amount,
                   "objetivo_general": general.strip(), "objetivos_especificos": clean_objectives,
                   "creado_por": st.session_state.user["id"]}
        if not configured():
            st.success("Proyecto validado correctamente (modo demostración; aún no se guarda en base de datos).")
            st.json(payload)
            return
        try:
            client = client_with_token(st.session_state.access_token)
            if project:
                result = client.table("proyectos").update(payload).eq("id", project["id"]).execute()
            else:
                result = client.table("proyectos").insert(payload).execute()
            project_id = str(result.data[0]["id"])
            groups = {"juridica": legal, "auxiliar": auxiliary, "acta_comite": [committee] if committee else [],
                      "acta_junta": [board] if board else [], "convenio": [agreement] if agreement else [],
                      "fotografia": photos}
            for category, files in groups.items():
                upload_files(client, project_id, category, files)
            st.success("Proyecto guardado correctamente.")
        except Exception as exc:
            st.error(f"No fue posible guardar el proyecto: {exc}")


def programs():
    st.title("Programas / Proyectos")
    direction = st.radio("Dirección responsable", ["Dirección de Operaciones", "Dirección de Proyectos"], horizontal=True)
    action = st.radio("¿Qué deseas hacer?", ["Dar de alta nuevo proyecto", "Editar proyecto"], horizontal=True)
    if action == "Dar de alta nuevo proyecto":
        project_form(direction)
    elif not configured():
        st.info("La consulta y edición estarán disponibles al conectar Supabase.")
    else:
        try:
            client = client_with_token(st.session_state.access_token)
            rows = client.table("proyectos").select("*").eq("direccion", direction).order("updated_at", desc=True).execute().data
            if not rows:
                st.info("Todavía no hay proyectos registrados en esta dirección.")
            else:
                labels = {f"{p['nombre']} — {p['municipio']} ({p['anio_inicio']})": p for p in rows}
                selected = st.selectbox("Selecciona el proyecto", labels)
                project_form(direction, labels[selected])
        except Exception as exc:
            st.error(f"No fue posible consultar los proyectos: {exc}")


def placeholder(title: str):
    st.title(title)
    st.info("Módulo preparado para desarrollarse en la siguiente etapa.")


if "user" not in st.session_state:
    login()
else:
    with st.sidebar:
        st.markdown("## COINVIERTE")
        if st.button("Inicio", use_container_width=True):
            st.session_state.page = "Inicio"
            st.rerun()
        if st.button("Programas / Proyectos", use_container_width=True):
            st.session_state.page = "Programas / Proyectos"
            st.rerun()
        if st.button("Junta de Gobierno", use_container_width=True):
            st.session_state.page = "Junta de Gobierno"
            st.rerun()
        if st.button("Comités", use_container_width=True):
            st.session_state.page = "Comités"
            st.rerun()
        st.divider()
        if st.button("Cerrar sesión", use_container_width=True):
            st.session_state.clear()
            st.rerun()
    page = st.session_state.get("page", "Inicio")
    if page == "Inicio": landing()
    elif page == "Programas / Proyectos": programs()
    else: placeholder(page)

