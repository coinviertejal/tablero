from __future__ import annotations

from datetime import datetime
import base64
from pathlib import Path

import streamlit as st

from data import MUNICIPIOS_JALISCO
from db import client_with_token, configured, public_client, upload_files, valid_official_email

st.set_page_config(page_title="COINVIERTE | Gestión Institucional", page_icon="🏛️", layout="wide")

st.markdown("""
<style>
:root { --ink:#35434b; --gray:#858e93; --blue:#0798cf; --green:#009b4c; --teal:#16ad8f; --purple:#a990c7; --orange:#f68b08; --paper:#f6f8f9; }
.stApp { background:linear-gradient(180deg,#fbfcfc 0%,#f1f5f6 100%); color:var(--ink); }
.block-container { max-width:1480px; padding-top:2.1rem; padding-bottom:4rem; }
[data-testid="stSidebar"] { background:linear-gradient(180deg,#535f66 0%,#778187 100%); border-right:0; }
[data-testid="stSidebar"] * { color:white; }
[data-testid="stSidebar"] hr { border-color:rgba(255,255,255,.22); }
[data-testid="stSidebar"] .stButton button { background:rgba(255,255,255,.08); border:1px solid rgba(255,255,255,.22); color:white; text-align:left; padding:.68rem .85rem; border-radius:10px; }
[data-testid="stSidebar"] .stButton button:hover { background:var(--blue); border-color:#fff; }
.brand { display:flex; align-items:center; gap:13px; }
.brand-mark { width:44px; height:44px; flex:0 0 44px; }
.brand-name { font-size:1.32rem; font-weight:800; letter-spacing:.055em; line-height:1; }
.brand-sub { font-size:.67rem; letter-spacing:.08em; opacity:.72; margin-top:5px; text-transform:uppercase; }
.side-logo { background:#fff; border-radius:13px; padding:12px 10px; margin:.3rem 0 1.35rem; box-shadow:0 8px 20px rgba(0,0,0,.12); }
.side-logo img { display:block; width:100%; height:auto; }
.hero { position:relative; overflow:hidden; background:#fff; padding:2.25rem 2.55rem 1.65rem; border-radius:22px; margin-bottom:1.2rem; border:1px solid #e1e7e9; border-bottom:7px solid var(--orange); box-shadow:0 18px 42px rgba(53,67,75,.11); }
.hero:after { content:""; position:absolute; width:260px; height:260px; border:42px solid rgba(7,152,207,.06); border-radius:50%; right:-120px; top:-125px; box-shadow:0 0 0 38px rgba(0,155,76,.035); }
.hero-logo { display:block; width:min(760px,78%); max-height:145px; object-fit:contain; object-position:left center; position:relative; z-index:1; }
.hero-copy { max-width:820px; font-size:1.02rem; color:#667279; margin:1.25rem 0 0; line-height:1.55; position:relative; z-index:1; }
.welcome { background:white; padding:.8rem 1rem; border-radius:11px; border:1px solid #e5ebee; margin:.4rem 0 1.2rem; color:#52616d; font-size:.9rem; }
.card { background:rgba(255,255,255,.97); border:1px solid #dfe7e9; border-top:5px solid var(--accent,var(--blue)); border-radius:17px; padding:1.45rem; min-height:175px; box-shadow:0 8px 24px rgba(53,67,75,.06); transition:.2s ease; }
.card:hover { transform:translateY(-3px); box-shadow:0 13px 30px rgba(53,67,75,.11); border-color:var(--accent,var(--blue)); }
.card-blue { --accent:var(--blue); } .card-green { --accent:var(--green); } .card-purple { --accent:var(--purple); }
.card-icon { display:inline-flex; align-items:center; justify-content:center; width:46px; height:46px; border-radius:13px; background:color-mix(in srgb,var(--accent,var(--blue)) 13%,white); color:var(--accent,var(--blue)); font-size:1.3rem; font-weight:800; margin-bottom:.65rem; }
.card h3 { margin:.15rem 0 .65rem; color:var(--ink); font-size:1.25rem; }
.muted { color:#647580; line-height:1.5; }
div[data-testid="stForm"] { background:white; padding:1.55rem; border-radius:18px; border:1px solid #dfe7e9; box-shadow:0 8px 24px rgba(20,55,70,.045); }
div[data-testid="stForm"] h3 { color:var(--gray); border-left:5px solid var(--orange); border-bottom:1px solid #e4ebed; padding:.15rem 0 .65rem .75rem; margin-top:1.2rem; }
.stButton button, .stFormSubmitButton button { border-radius:10px; }
.stFormSubmitButton button[kind="primary"] { background:linear-gradient(90deg,var(--green),var(--teal)); border:0; }
.stFormSubmitButton button[kind="primary"]:hover { background:linear-gradient(90deg,var(--blue),var(--teal)); }
div[data-baseweb="radio"] div[aria-checked="true"] { color:var(--orange); }
div[data-baseweb="input"]:focus-within, div[data-baseweb="select"]:focus-within, div[data-baseweb="textarea"]:focus-within { border-color:var(--blue); box-shadow:0 0 0 1px var(--blue); }
[data-testid="stFileUploaderDropzone"] { background:#f5faf9; border-color:var(--teal); }
h1,h2,h3 { letter-spacing:-.018em; color:var(--ink); }
</style>
""", unsafe_allow_html=True)


def logo_data_uri():
    logo = Path("assets/logo_coinvierte.jpeg")
    if not logo.exists():
        return ""
    return "data:image/jpeg;base64," + base64.b64encode(logo.read_bytes()).decode()


def brand_html(sidebar=False):
    src = logo_data_uri()
    if src:
        css_class = "side-logo" if sidebar else ""
        return f'<div class="{css_class}"><img src="{src}" alt="COINVIERTE"></div>' if sidebar else f'<img class="hero-logo" src="{src}" alt="COINVIERTE">'
    return '<div class="brand"><div><div class="brand-name">COINVIERTE</div><div class="brand-sub">Agencia de Coinversión para el Desarrollo Sostenible de Jalisco</div></div></div>'


def logo_header():
    identity = brand_html()
    st.markdown(f'''<div class="hero">{identity}
    <p class="hero-copy">Plataforma institucional para la gestión integral, documentación y seguimiento de programas y proyectos.</p></div>''', unsafe_allow_html=True)


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
    st.markdown(f'<div class="welcome">Sesión institucional activa · <b>{st.session_state.user["email"]}</b></div>', unsafe_allow_html=True)
    cols = st.columns(3)
    sections = [
        ("01", "Programas / Proyectos", "Alta, consulta, edición y seguimiento de expedientes."),
        ("02", "Junta de Gobierno", "Actas, acuerdos y documentación de las sesiones."),
        ("03", "Comités", "Integración, sesiones, actas y dictaminación."),
    ]
    card_styles = ["card-blue", "card-green", "card-purple"]
    for col, (icon, title, text), card_style in zip(cols, sections, card_styles):
        with col:
            st.markdown(f'<div class="card {card_style}"><div class="card-icon">{icon}</div><h3>{title}</h3><p class="muted">{text}</p></div>', unsafe_allow_html=True)
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
        st.markdown(brand_html(sidebar=True), unsafe_allow_html=True)
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
