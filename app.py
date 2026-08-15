from __future__ import annotations

from datetime import datetime
import base64
import html
from pathlib import Path
import uuid

import pandas as pd
import streamlit as st

from data import MUNICIPIOS_JALISCO
from db import access_profile, client_with_token, configured, download_project_images, public_client, register_access, upload_files, valid_official_email
from exports import build_docx, build_pdf

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
.choice-card { background:#fff; border:1px solid #dfe7e9; border-top:7px solid var(--accent,var(--blue)); border-radius:20px; padding:2rem 1.8rem 1.65rem; min-height:205px; box-shadow:0 12px 32px rgba(53,67,75,.08); text-align:center; margin-top:.7rem; }
.choice-card .choice-icon { width:68px; height:68px; border-radius:20px; display:flex; align-items:center; justify-content:center; margin:0 auto 1rem; background:color-mix(in srgb,var(--accent,var(--blue)) 14%,white); color:var(--accent,var(--blue)); font-size:1.6rem; font-weight:800; }
.choice-card h3 { font-size:1.45rem; margin:.25rem 0 .7rem; }
.choice-card p { color:#6d7980; margin:0; line-height:1.45; }
.choice-operations { --accent:var(--blue); } .choice-projects { --accent:var(--green); }
.choice-new { --accent:var(--orange); } .choice-edit { --accent:var(--purple); } .choice-view { --accent:var(--teal); }
.choice-title { text-align:center; margin:.7rem 0 .25rem; }
.choice-subtitle { text-align:center; color:#69767d; margin-bottom:1.2rem; }
.metric-grid { display:grid; grid-template-columns:repeat(4,1fr); gap:12px; margin:1rem 0 1.25rem; }
.metric-box { background:#fff; border:1px solid #dfe7e9; border-radius:14px; padding:1rem; border-top:4px solid var(--metric,var(--blue)); }
.metric-box .metric-label { color:#748087; font-size:.78rem; font-weight:700; text-transform:uppercase; letter-spacing:.03em; }
.metric-box .metric-value { color:var(--ink); font-size:1.55rem; font-weight:800; margin-top:.25rem; }
.metric-blue{--metric:var(--blue)} .metric-green{--metric:var(--green)} .metric-orange{--metric:var(--orange)} .metric-purple{--metric:var(--purple)}
.goal-heading { padding:.75rem 1rem; border-radius:10px; margin:1rem 0 .8rem; font-weight:800; border-left:8px solid var(--status-color); background:color-mix(in srgb,var(--status-color) 10%,white); }
.status-red{--status-color:#d9534f}.status-yellow{--status-color:#f0ad00}.status-green{--status-color:#009b4c}.status-gray{--status-color:#858e93}
@media(max-width:900px){.metric-grid{grid-template-columns:repeat(2,1fr)}}
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
    login_tab, activation_tab = st.tabs(["Ingresar", "Activar acceso con código"])
    with login_tab:
        with st.form("login"):
            email = st.text_input("Correo institucional", placeholder="nombre@jalisco.gob.mx")
            password = st.text_input("Contraseña", type="password")
            submitted = st.form_submit_button("Ingresar", type="primary", use_container_width=True)
    with activation_tab:
        st.caption("Utiliza el código temporal entregado por el administrador. El código sólo puede usarse una vez.")
        with st.form("activate_access"):
            activation_email = st.text_input("Correo autorizado", placeholder="nombre@jalisco.gob.mx", key="activation_email")
            code = st.text_input("Código temporal", max_chars=8)
            new_password = st.text_input("Crea una contraseña", type="password", key="new_password")
            confirm_password = st.text_input("Confirma la contraseña", type="password")
            activate = st.form_submit_button("Activar mi acceso", type="primary", use_container_width=True)
    if activate:
        if not configured():
            st.error("Primero debes conectar Supabase.")
        elif not valid_official_email(activation_email):
            st.error("El correo debe pertenecer a @jalisco.gob.mx.")
        elif len(new_password) < 8:
            st.error("La contraseña debe tener al menos 8 caracteres.")
        elif new_password != confirm_password:
            st.error("Las contraseñas no coinciden.")
        else:
            try:
                auth = public_client().auth.sign_up({"email": activation_email.lower().strip(), "password": new_password})
                redeemed = public_client().rpc("canjear_codigo_acceso", {"p_email": activation_email.lower().strip(),
                                                                          "p_codigo": code.strip()}).execute().data
                if not redeemed:
                    st.error("El código es incorrecto, ya fue utilizado o está vencido.")
                else:
                    if auth.session and auth.user:
                        st.success("Acceso activado correctamente. Ya puedes ingresar.")
                    else:
                        st.success("Acceso activado. Revisa tu correo si Supabase solicita confirmar la cuenta.")
            except Exception as exc:
                st.error(f"No fue posible activar el acceso: {exc}")
    if submitted:
        if not valid_official_email(email):
            st.error("El acceso está limitado a cuentas @jalisco.gob.mx.")
        elif not configured():
            st.session_state.user = {"email": email.lower(), "id": "demo"}
            st.rerun()
        else:
            try:
                auth = public_client().auth.sign_in_with_password({"email": email, "password": password})
            except Exception as exc:
                message = str(exc)
                if "Invalid login credentials" in message:
                    st.error("Supabase rechazó el correo o la contraseña: Invalid login credentials.")
                elif "Email not confirmed" in message:
                    st.error("La cuenta existe, pero el correo todavía no está confirmado en Supabase.")
                else:
                    st.error(f"Supabase no pudo autenticar la cuenta: {message}")
                return
            if not auth.user or not valid_official_email(auth.user.email or ""):
                public_client().auth.sign_out()
                st.error("La cuenta no pertenece al dominio autorizado.")
                return
            st.session_state.access_token = auth.session.access_token
            st.session_state.refresh_token = auth.session.refresh_token
            try:
                user_client = client_with_token(auth.session.access_token, auth.session.refresh_token)
                profile = access_profile(user_client, auth.user.email)
            except Exception as exc:
                st.error(f"La contraseña fue aceptada, pero falló la consulta de autorización: {exc}")
                return
            if not profile or not profile.get("activo"):
                public_client().auth.sign_out()
                st.session_state.pop("access_token", None)
                st.session_state.pop("refresh_token", None)
                st.error("La contraseña fue aceptada, pero tu acceso no está autorizado o está suspendido.")
                return
            st.session_state.user = {"email": auth.user.email, "id": str(auth.user.id),
                                     "nombre": profile.get("nombre") or auth.user.email,
                                     "rol": profile.get("rol", "usuario")}
            register_access(user_client)
            st.rerun()


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


def execution_goal_fields(existing=None):
    existing = existing or []
    if "execution_goal_count" not in st.session_state:
        st.session_state.execution_goal_count = max(1, len(existing))
    goals, evidence_groups = [], []
    statuses = ["Por iniciar", "En progreso", "Terminada"]
    for index in range(st.session_state.execution_goal_count):
        saved = existing[index] if index < len(existing) else {}
        saved_status = saved.get("estatus", "Por iniciar")
        display_status = st.session_state.get(f"goal_status_{index}", saved_status)
        status_class = {"Por iniciar": "status-red", "En progreso": "status-yellow", "Terminada": "status-green"}.get(display_status, "status-gray")
        st.markdown(f'<div class="goal-heading {status_class}">Meta de ejecución {index + 1} · {display_status}</div>', unsafe_allow_html=True)
        name = st.text_input("Nombre de la meta", value=saved.get("nombre", ""), key=f"goal_name_{index}")
        description = st.text_area("Descripción", value=saved.get("descripcion", ""), key=f"goal_description_{index}", height=90)
        g1, g2 = st.columns(2)
        status = g1.selectbox("Estatus", statuses, index=statuses.index(saved_status) if saved_status in statuses else 0,
                              key=f"goal_status_{index}")
        target_date = g2.text_input("Fecha objetivo", value=saved.get("fecha_objetivo", ""),
                                    placeholder="Ej. 30/09/2026", key=f"goal_date_{index}")
        evidence = st.file_uploader("Evidencia de la meta (fotografías o documentos)", accept_multiple_files=True,
                                    key=f"goal_evidence_{index}")
        goals.append({"nombre": name.strip(), "descripcion": description.strip(), "estatus": status,
                      "fecha_objetivo": target_date.strip(),
                      "evidencias_nombres": [file.name for file in evidence]})
        evidence_groups.append(evidence)
    c1, c2 = st.columns([1, 3])
    add_goal = c1.form_submit_button("＋ Agregar meta", use_container_width=True)
    remove_goal = c2.form_submit_button("Quitar última meta", use_container_width=True) if st.session_state.execution_goal_count > 1 else False
    return goals, evidence_groups, add_goal, remove_goal


RISK_COLUMNS = ["id", "riesgo", "categoria", "descripcion", "causa", "consecuencia",
                "probabilidad", "impacto", "mitigacion", "responsable", "fecha_compromiso",
                "estatus", "observaciones", "puntaje", "nivel", "eliminar"]


def risk_level(score: int) -> str:
    if score <= 4: return "Bajo"
    if score <= 9: return "Moderado"
    if score <= 16: return "Alto"
    return "Crítico"


def normalize_risks(records) -> list[dict]:
    clean = []
    for raw in records or []:
        item = dict(raw)
        if item.get("eliminar") is True:
            continue
        name = str(item.get("riesgo") or "").strip()
        if not name:
            continue
        try:
            probability = max(1, min(5, int(float(item.get("probabilidad") or 1))))
            impact = max(1, min(5, int(float(item.get("impacto") or 1))))
        except (TypeError, ValueError):
            probability, impact = 1, 1
        score = probability * impact
        clean.append({"id": str(item.get("id") or uuid.uuid4()), "riesgo": name,
            "categoria": str(item.get("categoria") or "Otro").strip(),
            "descripcion": str(item.get("descripcion") or "").strip(),
            "causa": str(item.get("causa") or "").strip(),
            "consecuencia": str(item.get("consecuencia") or "").strip(),
            "probabilidad": probability, "impacto": impact,
            "mitigacion": str(item.get("mitigacion") or "").strip(),
            "responsable": str(item.get("responsable") or "").strip(),
            "fecha_compromiso": str(item.get("fecha_compromiso") or "").strip(),
            "estatus": str(item.get("estatus") or "Por iniciar").strip(),
            "observaciones": str(item.get("observaciones") or "").strip(),
            "puntaje": score, "nivel": risk_level(score), "eliminar": False})
    return clean


def risks_from_excel(uploaded_file) -> list[dict]:
    frame = pd.read_excel(uploaded_file, sheet_name="Matriz de riesgos", header=3)
    aliases = {"ID (no modificar)": "id", "Riesgo": "riesgo", "Categoría": "categoria",
        "Descripción": "descripcion", "Causa": "causa", "Consecuencia": "consecuencia",
        "Probabilidad (1-5)": "probabilidad", "Impacto (1-5)": "impacto",
        "Mitigación": "mitigacion", "Responsable": "responsable",
        "Fecha compromiso": "fecha_compromiso", "Estatus": "estatus", "Observaciones": "observaciones"}
    missing = [column for column in aliases if column not in frame.columns]
    if missing:
        raise ValueError("Faltan columnas de la plantilla: " + ", ".join(missing))
    frame = frame.rename(columns=aliases)[list(aliases.values())].where(pd.notna(frame), "")
    return normalize_risks(frame.to_dict("records"))


def risk_summary(risks: list[dict]) -> str:
    if not risks:
        return "Todavía no se han registrado riesgos."
    counts = {level: sum(1 for risk in risks if risk["nivel"] == level)
              for level in ["Bajo", "Moderado", "Alto", "Crítico"]}
    priority = sorted(risks, key=lambda risk: risk["puntaje"], reverse=True)[:3]
    main = ", ".join(f'{risk["riesgo"]} ({risk["puntaje"]}, {risk["nivel"]})' for risk in priority)
    materialized = sum(1 for risk in risks if risk["estatus"] == "Materializado")
    closed = sum(1 for risk in risks if risk["estatus"] == "Mitigado / cerrado")
    return (f'Riesgos: {len(risks)} · Críticos: {counts["Crítico"]} · Altos: {counts["Alto"]} · '
            f'Moderados: {counts["Moderado"]} · Bajos: {counts["Bajo"]}. Materializados: {materialized}; '
            f'mitigados/cerrados: {closed}. Principales: {main}.')


def project_form(direction: str, project=None):
    project = project or {}
    risk_context = str(project.get("id") or "new")
    if st.session_state.get("risk_context") != risk_context:
        st.session_state.risk_context = risk_context
        saved_risks = ((project.get("avance_proyecto") or {}).get("matriz_riesgos") or [])
        st.session_state.risk_rows = normalize_risks(saved_risks)
    st.subheader("Editar proyecto" if project else "Dar de alta nuevo proyecto")
    with st.form("project_form", clear_on_submit=False):
        is_projects = direction == "Dirección de Proyectos"
        if is_projects:
            general_tab, advance_tab, risks_tab = st.tabs(["Ficha general", "Avance", "Matriz de riesgos"])
        else:
            general_tab, advance_tab, risks_tab = st.container(), None, None

        with general_tab:
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
            photos = st.file_uploader("Fotografías generales (máximo 5 MB por archivo)", type=["jpg", "jpeg", "png", "webp"],
                                      accept_multiple_files=True, key="photos")

            if not is_projects:
                st.markdown("### Monitoreo y Seguimiento")
                previous_monitoring = project.get("monitoreo", {}) or {}
                m1, m2 = st.columns(2)
                statuses = ["Sin iniciar", "En planeación", "En ejecución", "Suspendido", "Concluido"]
                current_status = previous_monitoring.get("estatus", "Sin iniciar")
                status = m1.selectbox("Estatus del proyecto", statuses,
                                      index=statuses.index(current_status) if current_status in statuses else 0)
                responsible = m2.text_input("Responsable del seguimiento", value=previous_monitoring.get("responsable", ""))
                m3, m4 = st.columns(2)
                period = m3.text_input("Periodo de seguimiento", value=previous_monitoring.get("periodo", ""))
                progress = m4.slider("Porcentaje de avance", 0, 100, int(previous_monitoring.get("avance", 0)), 5)
                monitoring_progress = st.text_area("Principales avances", value=previous_monitoring.get("avances", ""), height=100)
                pending = st.text_area("Pendientes o riesgos", value=previous_monitoring.get("pendientes", ""), height=90)
                next_actions = st.text_area("Próximas acciones", value=previous_monitoring.get("proximas_acciones", ""), height=90)
                observations = st.text_area("Observaciones de seguimiento", value=previous_monitoring.get("observaciones", ""), height=90)

        execution_goals, goal_evidence_groups, add_goal, remove_goal = [], [], False, False
        budget_dispersed = 0.0
        if is_projects:
            with advance_tab:
                saved_advance = project.get("avance_proyecto", {}) or {}
                st.markdown("### Ejecución financiera")
                budget_dispersed = st.number_input("Presupuesto dispersado (MXN)", min_value=0.0,
                                                   value=float(saved_advance.get("presupuesto_dispersado", 0)), step=1000.0, format="%.2f")
                budget_pct = min(100.0, (budget_dispersed / amount * 100)) if amount else 0.0
                saved_goals = saved_advance.get("metas", []) or []
                st.markdown("### Metas de ejecución")
                execution_goals, goal_evidence_groups, add_goal, remove_goal = execution_goal_fields(saved_goals)
                completed = sum(1 for goal in execution_goals if goal["estatus"] == "Terminada")
                in_progress = sum(1 for goal in execution_goals if goal["estatus"] == "En progreso")
                physical_pct = ((completed * 100 + in_progress * 50) / len(execution_goals)) if execution_goals else 0
                traffic = "Verde" if physical_pct >= 80 else "Amarillo" if physical_pct >= 40 else "Rojo"
                traffic_class = "metric-green" if traffic == "Verde" else "metric-orange" if traffic == "Amarillo" else "metric-blue"
                st.markdown(f'''<div class="metric-grid">
                  <div class="metric-box metric-blue"><div class="metric-label">Avance financiero</div><div class="metric-value">{budget_pct:.1f}%</div></div>
                  <div class="metric-box metric-green"><div class="metric-label">Metas terminadas</div><div class="metric-value">{completed}/{len(execution_goals)}</div></div>
                  <div class="metric-box metric-purple"><div class="metric-label">Avance físico</div><div class="metric-value">{physical_pct:.1f}%</div></div>
                  <div class="metric-box {traffic_class}"><div class="metric-label">Semáforo general</div><div class="metric-value">{traffic}</div></div>
                </div>''', unsafe_allow_html=True)
                refresh_metrics = st.form_submit_button("Actualizar indicadores", use_container_width=True)

            with risks_tab:
                st.markdown("### Matriz de riesgos")
                st.caption("Edita la tabla. Para borrar un riesgo, marca Eliminar o usa el control de eliminación de filas.")
                template_path = Path("plantilla_matriz_riesgos_coinvierte.xlsx")
                if template_path.exists():
                    template_b64 = base64.b64encode(template_path.read_bytes()).decode()
                    st.markdown(f'<a download="plantilla_matriz_riesgos_coinvierte.xlsx" href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{template_b64}">⬇️ Descargar plantilla Excel</a>', unsafe_allow_html=True)
                risk_file = st.file_uploader("Cargar plantilla llena (.xlsx)", type=["xlsx"], key=f"risk_excel_{risk_context}")
                import_risks = st.form_submit_button("Ingestar riesgos desde Excel", use_container_width=True)
                risk_frame = pd.DataFrame(st.session_state.get("risk_rows", []), columns=RISK_COLUMNS)
                edited_risks = st.data_editor(risk_frame, key=f"risk_editor_{risk_context}", num_rows="dynamic",
                    use_container_width=True, hide_index=True, disabled=["id", "puntaje", "nivel"],
                    column_config={"id": None,
                        "riesgo": st.column_config.TextColumn("Riesgo", required=True, width="medium"),
                        "categoria": st.column_config.SelectboxColumn("Categoría", options=["Financiero", "Operativo", "Jurídico / normativo", "Técnico", "Ambiental", "Social", "Reputacional", "Cronograma", "Otro"], width="medium"),
                        "descripcion": st.column_config.TextColumn("Descripción", width="large"),
                        "causa": st.column_config.TextColumn("Causa", width="large"),
                        "consecuencia": st.column_config.TextColumn("Consecuencia", width="large"),
                        "probabilidad": st.column_config.NumberColumn("Probabilidad 1–5", min_value=1, max_value=5, step=1, required=True),
                        "impacto": st.column_config.NumberColumn("Impacto 1–5", min_value=1, max_value=5, step=1, required=True),
                        "mitigacion": st.column_config.TextColumn("Mitigación", width="large"),
                        "responsable": st.column_config.TextColumn("Responsable", width="medium"),
                        "fecha_compromiso": st.column_config.TextColumn("Fecha compromiso", help="Formato sugerido: AAAA-MM-DD"),
                        "estatus": st.column_config.SelectboxColumn("Estatus", options=["Por iniciar", "En seguimiento", "Materializado", "Mitigado / cerrado"], width="medium"),
                        "observaciones": st.column_config.TextColumn("Observaciones", width="large"),
                        "puntaje": st.column_config.NumberColumn("Puntaje", help="Probabilidad × Impacto"),
                        "nivel": st.column_config.TextColumn("Nivel"),
                        "eliminar": st.column_config.CheckboxColumn("Eliminar", help="Marca para borrar al guardar.")})
                recalculate_risks = st.form_submit_button("Actualizar puntajes y síntesis", use_container_width=True)
                current_risks = normalize_risks(edited_risks.to_dict("records"))
                st.info(risk_summary(current_risks))
        else:
            budget_pct, physical_pct, traffic, completed = 0, float(progress), current_status, 0
            refresh_metrics, import_risks, recalculate_risks, risk_file, edited_risks = False, False, False, None, None

        st.markdown("### Ficha del proyecto")
        st.caption("La ficha incluye información general, monitoreo y evidencia fotográfica. Gestión Documental no se muestra.")
        b1, b2 = st.columns(2)
        preview = b1.form_submit_button("Previsualizar ficha", use_container_width=True)
        save = b2.form_submit_button("Guardar proyecto", type="primary", use_container_width=True)

    if add_goal:
        st.session_state.execution_goal_count += 1
        st.rerun()
    if remove_goal:
        st.session_state.execution_goal_count -= 1
        st.rerun()
    if refresh_metrics:
        st.rerun()
    if import_risks:
        if risk_file is None:
            st.error("Selecciona primero el archivo Excel lleno.")
        else:
            try:
                st.session_state.risk_rows = risks_from_excel(risk_file)
                st.success(f"Se importaron {len(st.session_state.risk_rows)} riesgos.")
                st.rerun()
            except Exception as exc:
                st.error(f"No fue posible importar la plantilla: {exc}")
    if recalculate_risks:
        st.session_state.risk_rows = normalize_risks(edited_risks.to_dict("records"))
        st.rerun()

    if save or preview:
        errors = []
        if not name.strip() or not applicant.strip() or not general.strip():
            errors.append("Completa todos los campos obligatorios.")
        clean_objectives = [o.strip() for o in objectives if o.strip()]
        if not clean_objectives:
            errors.append("Agrega al menos un objetivo específico.")
        oversized = [f.name for f in photos if f.size > 5 * 1024 * 1024]
        oversized += [f.name for files in goal_evidence_groups for f in files if f.size > 5 * 1024 * 1024]
        if oversized:
            errors.append("Estas fotografías exceden 5 MB: " + ", ".join(oversized))
        if errors:
            for error in errors:
                st.error(error)
            return
        if is_projects:
            risk_rows = normalize_risks(edited_risks.to_dict("records"))
            st.session_state.risk_rows = risk_rows
            advance_data = {"presupuesto_dispersado": budget_dispersed, "porcentaje_financiero": round(budget_pct, 2),
                            "porcentaje_fisico": round(physical_pct, 2), "semaforo": traffic,
                            "metas_terminadas": completed, "metas": execution_goals,
                            "matriz_riesgos": risk_rows, "sintesis_riesgos": risk_summary(risk_rows)}
            monitoring_data = {"estatus": traffic, "responsable": "", "periodo": "", "avance": round(physical_pct),
                               "avances": f"{completed} de {len(execution_goals)} metas terminadas. Avance financiero: {budget_pct:.1f}%.",
                               "pendientes": "", "proximas_acciones": "", "observaciones": ""}
        else:
            advance_data = {}
            monitoring_data = {"estatus": status, "responsable": responsible.strip(), "periodo": period.strip(),
                               "avance": int(progress), "avances": monitoring_progress.strip(),
                               "pendientes": pending.strip(), "proximas_acciones": next_actions.strip(),
                               "observaciones": observations.strip()}
        payload = {"direccion": direction, "nombre": name.strip(), "solicitante": applicant.strip(),
                   "municipio": municipality, "anio_inicio": int(year), "monto": amount,
                   "objetivo_general": general.strip(), "objetivos_especificos": clean_objectives,
                   "monitoreo": monitoring_data, "avance_proyecto": advance_data,
                   "creado_por": st.session_state.user["id"]}
        if preview:
            photo_bytes = [photo.getvalue() for photo in photos]
            photo_bytes += [file.getvalue() for files in goal_evidence_groups for file in files
                            if (file.type or "").startswith("image/")]
            st.session_state.ficha_data = payload
            st.session_state.ficha_photos = photo_bytes
            st.session_state.ficha_pdf = build_pdf(payload, photo_bytes, "assets/logo_coinvierte.jpeg")
            st.session_state.ficha_docx = build_docx(payload, photo_bytes, "assets/logo_coinvierte.jpeg")
            st.rerun()
        if not configured():
            st.success("Proyecto validado correctamente (modo demostración; aún no se guarda en base de datos).")
            st.json(payload)
            return
        try:
            client = client_with_token(st.session_state.access_token, st.session_state.refresh_token)
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
            for files in goal_evidence_groups:
                upload_files(client, project_id, "evidencia_meta", files)
            st.success("Proyecto guardado correctamente.")
        except Exception as exc:
            st.error(f"No fue posible guardar el proyecto: {exc}")

    if st.session_state.get("ficha_data"):
        render_project_preview(st.session_state.ficha_data, st.session_state.get("ficha_photos", []))


def render_project_preview(data: dict, photos: list[bytes]):
    st.markdown("---")
    st.markdown("## Previsualización de la ficha")
    st.caption("Gestión Documental se excluye intencionalmente de esta ficha.")
    st.markdown(f'''<div class="card card-blue">
      <div class="card-icon">FP</div><h3>{data.get("nombre") or "Proyecto sin nombre"}</h3>
      <p class="muted"><b>{data.get("direccion", "")}</b><br>{data.get("solicitante", "")} · {data.get("municipio", "")} · {data.get("anio_inicio", "")}</p>
      <p><b>Monto:</b> ${float(data.get("monto", 0)):,.2f} MXN</p></div>''', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("### Información General")
        st.markdown(f"**Objetivo general**\n\n{data.get('objetivo_general') or 'Sin información'}")
        st.markdown("**Objetivos específicos**")
        for objective in data.get("objetivos_especificos", []):
            st.markdown(f"- {objective}")
    with c2:
        advance = data.get("avance_proyecto", {}) or {}
        if advance:
            st.markdown("### Avance del Proyecto")
            st.progress(int(advance.get("porcentaje_fisico", 0)), text=f"Avance físico: {advance.get('porcentaje_fisico', 0)}%")
            st.markdown(f"**Presupuesto dispersado:** ${float(advance.get('presupuesto_dispersado',0)):,.2f} MXN  \n"
                        f"**Avance financiero:** {advance.get('porcentaje_financiero',0)}%  \n"
                        f"**Semáforo:** {advance.get('semaforo','Sin información')}")
            for index, goal in enumerate(advance.get("metas", []), 1):
                color = {"Por iniciar":"🔴","En progreso":"🟡","Terminada":"🟢"}.get(goal.get("estatus"),"⚪")
                st.markdown(f"**{color} Meta {index}: {goal.get('nombre') or 'Sin nombre'}**  \n"
                            f"{goal.get('descripcion') or 'Sin descripción'}  \n"
                            f"Fecha objetivo: {goal.get('fecha_objetivo') or 'Sin fecha'}")
        else:
            st.markdown("### Monitoreo y Seguimiento")
            monitoring = data.get("monitoreo", {})
            st.progress(int(monitoring.get("avance", 0)), text=f"Avance: {monitoring.get('avance', 0)}%")
            st.markdown(f"**Estatus:** {monitoring.get('estatus', 'Sin información')}  \n"
                        f"**Responsable:** {monitoring.get('responsable') or 'Sin información'}  \n"
                        f"**Periodo:** {monitoring.get('periodo') or 'Sin información'}")
            for label, key in [("Principales avances","avances"),("Pendientes o riesgos","pendientes"),
                               ("Próximas acciones","proximas_acciones"),("Observaciones","observaciones")]:
                if monitoring.get(key):
                    st.markdown(f"**{label}**\n\n{monitoring[key]}")
    st.markdown("### Evidencia Fotográfica")
    if photos:
        columns = st.columns(2)
        for index, photo in enumerate(photos):
            columns[index % 2].image(photo, caption=f"Fotografía {index + 1}", use_container_width=True)
    else:
        st.info("No se cargó evidencia fotográfica para esta ficha.")
    name = "_".join((data.get("nombre") or "proyecto").lower().split())[:60]
    d1, d2, d3 = st.columns([1,1,2])
    d1.download_button("Descargar PDF", st.session_state.ficha_pdf, f"ficha_{name}.pdf", "application/pdf", use_container_width=True)
    d2.download_button("Descargar Word", st.session_state.ficha_docx, f"ficha_{name}.docx",
                       "application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
    if d3.button("Cerrar previsualización", use_container_width=True):
        for key in ["ficha_data", "ficha_photos", "ficha_pdf", "ficha_docx"]:
            st.session_state.pop(key, None)
        st.rerun()


def project_is_active(project: dict) -> bool:
    if project.get("direccion") == "Dirección de Proyectos":
        goals = (project.get("avance_proyecto") or {}).get("metas", [])
        return not goals or any(goal.get("estatus") != "Terminada" for goal in goals)
    return (project.get("monitoreo") or {}).get("estatus") != "Concluido"


def render_readonly_project(project: dict, photos: list[bytes]):
    name = html.escape(project.get("nombre") or "Proyecto sin nombre")
    st.markdown(f'''<div class="card card-blue"><div class="card-icon">VP</div><h3>{name}</h3>
      <p class="muted"><b>{html.escape(project.get("solicitante") or "")}</b><br>
      {html.escape(project.get("municipio") or "")} · {project.get("anio_inicio", "")}</p>
      <p><b>Monto total:</b> ${float(project.get("monto",0)):,.2f} MXN</p></div>''', unsafe_allow_html=True)

    is_projects = project.get("direccion") == "Dirección de Proyectos"
    tab_labels = ["Ficha general", "Avance", "Matriz de riesgos"] if is_projects else ["Ficha general", "Monitoreo y seguimiento"]
    tabs = st.tabs(tab_labels)
    general_tab, progress_tab = tabs[0], tabs[1]
    risks_tab = tabs[2] if is_projects else None
    with general_tab:
        st.markdown("### Información General")
        c1, c2, c3 = st.columns(3)
        c1.metric("Municipio", project.get("municipio") or "Sin información")
        c2.metric("Año de inicio", project.get("anio_inicio") or "Sin información")
        c3.metric("Monto", f'${float(project.get("monto",0)):,.2f}')
        st.markdown("#### Objetivo general")
        st.write(project.get("objetivo_general") or "Sin información")
        st.markdown("#### Objetivos específicos")
        for objective in project.get("objetivos_especificos", []):
            st.markdown(f"- {objective}")
        st.markdown("#### Evidencia fotográfica")
        if photos:
            columns = st.columns(2)
            for index, photo in enumerate(photos):
                columns[index % 2].image(photo, caption=f"Evidencia {index + 1}", use_container_width=True)
        else:
            st.info("Este proyecto todavía no tiene fotografías disponibles.")

    with progress_tab:
        if is_projects:
            advance = project.get("avance_proyecto", {}) or {}
            total_goals = len(advance.get("metas", []))
            st.markdown(f'''<div class="metric-grid">
              <div class="metric-box metric-blue"><div class="metric-label">Avance financiero</div><div class="metric-value">{advance.get("porcentaje_financiero",0)}%</div></div>
              <div class="metric-box metric-green"><div class="metric-label">Metas terminadas</div><div class="metric-value">{advance.get("metas_terminadas",0)}/{total_goals}</div></div>
              <div class="metric-box metric-purple"><div class="metric-label">Avance físico</div><div class="metric-value">{advance.get("porcentaje_fisico",0)}%</div></div>
              <div class="metric-box metric-orange"><div class="metric-label">Semáforo</div><div class="metric-value">{advance.get("semaforo","Sin datos")}</div></div>
            </div>''', unsafe_allow_html=True)
            st.markdown(f"**Presupuesto dispersado:** ${float(advance.get('presupuesto_dispersado',0)):,.2f} MXN")
            if not advance.get("metas"):
                st.info("Todavía no se han registrado metas de ejecución.")
            for index, goal in enumerate(advance.get("metas", []), 1):
                icon = {"Por iniciar":"🔴", "En progreso":"🟡", "Terminada":"🟢"}.get(goal.get("estatus"), "⚪")
                with st.expander(f"{icon} Meta {index}: {goal.get('nombre') or 'Sin nombre'} · {goal.get('estatus','')}", expanded=True):
                    st.write(goal.get("descripcion") or "Sin descripción")
                    st.caption(f"Fecha objetivo: {goal.get('fecha_objetivo') or 'Sin fecha'}")
                    evidence_names = goal.get("evidencias_nombres", [])
                    st.write("**Evidencias:** " + (", ".join(evidence_names) if evidence_names else "Sin evidencia"))
        else:
            monitoring = project.get("monitoreo", {}) or {}
            st.progress(int(monitoring.get("avance", 0)), text=f"Avance: {monitoring.get('avance',0)}%")
            for label, key in [("Estatus","estatus"),("Responsable","responsable"),("Periodo","periodo"),
                               ("Principales avances","avances"),("Pendientes o riesgos","pendientes"),
                               ("Próximas acciones","proximas_acciones"),("Observaciones","observaciones")]:
                st.markdown(f"**{label}:** {monitoring.get(key) or 'Sin información'}")

    if risks_tab:
        with risks_tab:
            risks = normalize_risks((project.get("avance_proyecto") or {}).get("matriz_riesgos", []))
            st.markdown("### Síntesis de riesgos")
            st.info(risk_summary(risks))
            if risks:
                display = pd.DataFrame(risks).drop(columns=["id", "eliminar"], errors="ignore")
                st.dataframe(display, use_container_width=True, hide_index=True)
            else:
                st.info("Todavía no se han registrado riesgos.")

    pdf = build_pdf(project, photos, "assets/logo_coinvierte.jpeg")
    docx = build_docx(project, photos, "assets/logo_coinvierte.jpeg")
    safe_name = "_".join((project.get("nombre") or "proyecto").lower().split())[:60]
    d1, d2 = st.columns(2)
    d1.download_button("Descargar ficha en PDF", pdf, f"ficha_{safe_name}.pdf", "application/pdf", use_container_width=True)
    d2.download_button("Descargar ficha en Word", docx, f"ficha_{safe_name}.docx",
                       "application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)


def view_active_projects(direction: str):
    if not configured():
        st.info("La visualización de proyectos estará disponible al conectar Supabase.")
        return
    client = client_with_token(st.session_state.access_token, st.session_state.refresh_token)
    rows = client.table("proyectos").select("*").eq("direccion", direction).order("updated_at", desc=True).execute().data
    active = [project for project in rows or [] if project_is_active(project)]
    selected_id = st.session_state.get("view_project_id")
    if selected_id:
        project = next((item for item in active if str(item["id"]) == str(selected_id)), None)
        if not project:
            st.session_state.pop("view_project_id", None)
            st.rerun()
        if st.button("← Volver a proyectos activos"):
            st.session_state.pop("view_project_id", None)
            st.session_state.pop("view_project_photos", None)
            st.rerun()
        if "view_project_photos" not in st.session_state:
            st.session_state.view_project_photos = download_project_images(client, str(project["id"]))
        render_readonly_project(project, st.session_state.view_project_photos)
        return

    st.markdown("## Proyectos activos")
    st.caption(f"{direction} · Selecciona un proyecto para consultar su información")
    if not active:
        st.info("No hay proyectos activos registrados en esta dirección.")
        return
    labels = {f"{p['nombre']} — {p['municipio']} ({p['anio_inicio']})": p for p in active}
    selected_label = st.selectbox("Proyecto", list(labels.keys()))
    selected = labels[selected_label]
    advance = selected.get("avance_proyecto", {}) or {}
    st.markdown(f'''<div class="card card-green"><div class="card-icon">{len(active)}</div>
      <h3>{html.escape(selected.get("nombre") or "")}</h3><p class="muted">{html.escape(selected.get("solicitante") or "")} · 
      {html.escape(selected.get("municipio") or "")}</p><p><b>Avance:</b> {advance.get("porcentaje_fisico", (selected.get("monitoreo") or {}).get("avance",0))}%</p></div>''', unsafe_allow_html=True)
    if st.button("Ver información del proyecto", type="primary", use_container_width=True):
        st.session_state.view_project_id = str(selected["id"])
        st.session_state.pop("view_project_photos", None)
        st.rerun()


def user_management():
    st.title("Gestión de usuarios")
    if st.session_state.user.get("rol") != "administrador":
        st.error("No tienes permisos para acceder a este módulo.")
        return
    client = client_with_token(st.session_state.access_token, st.session_state.refresh_token)
    create_tab, users_tab = st.tabs(["Generar código temporal", "Usuarios autorizados"])
    with create_tab:
        st.markdown("### Autorizar a una persona")
        with st.form("create_access_code"):
            name = st.text_input("Nombre de la persona")
            email = st.text_input("Correo institucional", placeholder="nombre@jalisco.gob.mx")
            hours = st.selectbox("Vigencia del código", [24, 48, 72, 168],
                                 format_func=lambda value: "7 días" if value == 168 else f"{value} horas")
            create_code = st.form_submit_button("Generar código de acceso", type="primary", use_container_width=True)
        if create_code:
            if not valid_official_email(email):
                st.error("Sólo se pueden autorizar correos @jalisco.gob.mx.")
            else:
                try:
                    result = client.rpc("crear_codigo_acceso", {"p_email": email.lower().strip(),
                                                                 "p_nombre": name.strip(), "p_horas": hours}).execute().data
                    record = result[0] if isinstance(result, list) else result
                    st.session_state.generated_code = record
                    st.session_state.generated_email = email.lower().strip()
                except Exception as exc:
                    st.error(f"No fue posible generar el código: {exc}")
        if st.session_state.get("generated_code"):
            record = st.session_state.generated_code
            st.success(f"Código generado para {st.session_state.generated_email}")
            st.code(record.get("codigo", ""), language=None)
            st.caption(f"Vence: {record.get('vence', '')}. Compártelo únicamente con la persona autorizada.")

    with users_tab:
        rows = client.table("usuarios_autorizados").select("id,email,nombre,rol,activo,ultimo_acceso,created_at").order("created_at", desc=True).execute().data
        if not rows:
            st.info("Todavía no hay usuarios registrados.")
        else:
            st.dataframe([{ "Nombre": row.get("nombre"), "Correo": row.get("email"), "Rol": row.get("rol"),
                            "Estado": "Activo" if row.get("activo") else "Suspendido / pendiente",
                            "Último acceso": row.get("ultimo_acceso") or "Sin acceso"} for row in rows],
                         use_container_width=True, hide_index=True)
            manageable = [row for row in rows if row.get("rol") != "administrador"]
            if manageable:
                labels = {f"{row.get('nombre') or 'Sin nombre'} — {row['email']}": row for row in manageable}
                selected_label = st.selectbox("Administrar usuario", list(labels.keys()))
                selected = labels[selected_label]
                c1, c2 = st.columns(2)
                if selected.get("activo"):
                    if c1.button("Suspender acceso", use_container_width=True):
                        client.table("usuarios_autorizados").update({"activo": False}).eq("id", selected["id"]).execute()
                        st.success("Acceso suspendido.")
                        st.rerun()
                else:
                    if c1.button("Reactivar acceso", use_container_width=True):
                        client.table("usuarios_autorizados").update({"activo": True}).eq("id", selected["id"]).execute()
                        st.success("Acceso reactivado.")
                        st.rerun()
                if c2.button("Generar nuevo código", use_container_width=True):
                    result = client.rpc("crear_codigo_acceso", {"p_email": selected["email"],
                                                                 "p_nombre": selected.get("nombre") or "", "p_horas": 24}).execute().data
                    st.session_state.generated_code = result[0] if isinstance(result, list) else result
                    st.session_state.generated_email = selected["email"]
                    st.success("Nuevo código generado. Consúltalo en la primera pestaña.")


def programs():
    direction = st.session_state.get("program_direction")
    action = st.session_state.get("program_action")

    if not direction:
        st.markdown('<h1 class="choice-title">Programas / Proyectos</h1>', unsafe_allow_html=True)
        st.markdown('<p class="choice-subtitle">Selecciona la dirección responsable para continuar</p>', unsafe_allow_html=True)
        c1, c2 = st.columns(2, gap="large")
        with c1:
            st.markdown('''<div class="choice-card choice-operations"><div class="choice-icon">DO</div>
                <h3>Dirección de Operaciones</h3><p>Gestión de los programas y proyectos correspondientes a esta dirección.</p></div>''', unsafe_allow_html=True)
            if st.button("Ingresar a Operaciones", key="choose_operations", use_container_width=True, type="primary"):
                st.session_state.program_direction = "Dirección de Operaciones"
                st.session_state.pop("program_action", None)
                st.rerun()
        with c2:
            st.markdown('''<div class="choice-card choice-projects"><div class="choice-icon">DP</div>
                <h3>Dirección de Proyectos</h3><p>Gestión de los programas y proyectos correspondientes a esta dirección.</p></div>''', unsafe_allow_html=True)
            if st.button("Ingresar a Proyectos", key="choose_projects", use_container_width=True, type="primary"):
                st.session_state.program_direction = "Dirección de Proyectos"
                st.session_state.pop("program_action", None)
                st.rerun()
        return

    if not action:
        top1, top2 = st.columns([1, 5])
        if top1.button("← Direcciones", use_container_width=True):
            st.session_state.pop("program_direction", None)
            st.rerun()
        top2.markdown(f"### {direction}")
        st.markdown('<p class="choice-subtitle">Selecciona la acción que deseas realizar</p>', unsafe_allow_html=True)
        c1, c2, c3 = st.columns(3, gap="large")
        with c1:
            st.markdown('''<div class="choice-card choice-new"><div class="choice-icon">＋</div>
                <h3>Dar de alta nuevo proyecto</h3><p>Crear un expediente e incorporar su información general, documentos y seguimiento.</p></div>''', unsafe_allow_html=True)
            if st.button("Crear nuevo proyecto", key="choose_new", use_container_width=True, type="primary"):
                st.session_state.program_action = "new"
                st.session_state.objective_count = 1
                st.rerun()
        with c2:
            st.markdown('''<div class="choice-card choice-edit"><div class="choice-icon">✎</div>
                <h3>Editar proyecto</h3><p>Consultar un expediente existente para actualizar su información y seguimiento.</p></div>''', unsafe_allow_html=True)
            if st.button("Consultar y editar", key="choose_edit", use_container_width=True, type="primary"):
                st.session_state.program_action = "edit"
                st.rerun()
        with c3:
            st.markdown('''<div class="choice-card choice-view"><div class="choice-icon">◉</div>
                <h3>Visualizar proyectos</h3><p>Consultar los proyectos activos y acceder a su ficha general y avance.</p></div>''', unsafe_allow_html=True)
            if st.button("Ver proyectos activos", key="choose_view", use_container_width=True, type="primary"):
                st.session_state.program_action = "view"
                st.session_state.pop("view_project_id", None)
                st.rerun()
        return

    nav1, nav2 = st.columns([1, 5])
    if nav1.button("← Acciones", use_container_width=True):
        st.session_state.pop("program_action", None)
        for key in ["ficha_data", "ficha_photos", "ficha_pdf", "ficha_docx"]:
            st.session_state.pop(key, None)
        st.session_state.pop("view_project_id", None)
        st.session_state.pop("view_project_photos", None)
        st.rerun()
    nav2.markdown(f"### {direction}")

    if action == "new":
        project_form(direction)
    elif action == "view":
        try:
            view_active_projects(direction)
        except Exception as exc:
            st.error(f"No fue posible consultar los proyectos activos: {exc}")
    elif not configured():
        st.info("La consulta y edición estarán disponibles al conectar Supabase.")
    else:
        try:
            client = client_with_token(st.session_state.access_token, st.session_state.refresh_token)
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
            st.session_state.pop("program_direction", None)
            st.session_state.pop("program_action", None)
            st.rerun()
        if st.button("Junta de Gobierno", use_container_width=True):
            st.session_state.page = "Junta de Gobierno"
            st.rerun()
        if st.button("Comités", use_container_width=True):
            st.session_state.page = "Comités"
            st.rerun()
        if st.session_state.user.get("rol") == "administrador":
            if st.button("Gestión de usuarios", use_container_width=True):
                st.session_state.page = "Gestión de usuarios"
                st.rerun()
        st.divider()
        if st.button("Cerrar sesión", use_container_width=True):
            st.session_state.clear()
            st.rerun()
    page = st.session_state.get("page", "Inicio")
    if page == "Inicio": landing()
    elif page == "Programas / Proyectos": programs()
    elif page == "Gestión de usuarios": user_management()
    else: placeholder(page)
