from __future__ import annotations

import re
import time
import uuid
from pathlib import Path

import streamlit as st
from supabase import Client, create_client

ALLOWED_DOMAIN = "@jalisco.gob.mx"
BUCKET = "expedientes"


def configured() -> bool:
    return bool(st.secrets.get("SUPABASE_URL") and public_key())


def public_key() -> str:
    return st.secrets.get("SUPABASE_PUBLISHABLE_KEY") or st.secrets.get("SUPABASE_ANON_KEY") or ""


@st.cache_resource
def public_client() -> Client:
    return create_client(st.secrets["SUPABASE_URL"], public_key())


@st.cache_resource(ttl=1800, show_spinner=False)
def client_with_token(access_token: str, refresh_token: str) -> Client:
    """Crea una sola vez el cliente autenticado por sesión/token.

    Streamlit vuelve a ejecutar el script con cada interacción. Sin la caché,
    ``set_session`` revalidaba los tokens contra Supabase en cada clic y una
    demora transitoria podía derribar toda la vista de Junta de Gobierno.
    """
    last_error = None
    for attempt in range(3):
        try:
            client = create_client(st.secrets["SUPABASE_URL"], public_key())
            client.auth.set_session(access_token, refresh_token)
            return client
        except Exception as exc:
            last_error = exc
            if exc.__class__.__name__ not in {"ReadTimeout", "ConnectTimeout", "ConnectError"}:
                raise
            if attempt < 2:
                time.sleep(0.6 * (attempt + 1))
    raise RuntimeError(
        "Supabase tardó demasiado en validar la sesión. "
        "La sesión no se cerró; espera unos segundos y vuelve a intentar."
    ) from last_error


def access_profile(client: Client, email: str) -> dict | None:
    rows = client.table("usuarios_autorizados").select("id,email,nombre,rol,activo,direccion,modulos,direcciones_proyectos").eq("email", email.lower()).execute().data
    return rows[0] if rows else None


def register_access(client: Client):
    try:
        client.rpc("registrar_acceso").execute()
    except Exception:
        pass


def valid_official_email(email: str) -> bool:
    return email.strip().lower().endswith(ALLOWED_DOMAIN)


def safe_name(name: str) -> str:
    clean = re.sub(r"[^A-Za-z0-9._-]+", "_", Path(name).name)
    return clean[:120] or "archivo"


def upload_files(client: Client, project_id: str, category: str, files: list) -> list[dict]:
    records = []
    for item in files:
        path = f"{project_id}/{category}/{uuid.uuid4().hex}_{safe_name(item.name)}"
        client.storage.from_(BUCKET).upload(
            path,
            item.getvalue(),
            {"content-type": item.type or "application/octet-stream"},
        )
        records.append({
            "proyecto_id": project_id,
            "categoria": category,
            "nombre_archivo": item.name,
            "ruta_storage": path,
            "mime_type": item.type,
            "tamano_bytes": item.size,
        })
    if records:
        client.table("documentos").insert(records).execute()
    return records


def download_project_images(client: Client, project_id: str) -> list[bytes]:
    """Descarga fotografías generales y evidencias de metas que sean imágenes."""
    rows = (client.table("documentos").select("ruta_storage,mime_type,categoria")
            .eq("proyecto_id", project_id).in_("categoria", ["fotografia", "evidencia_meta"]).execute().data)
    images = []
    for row in rows or []:
        if (row.get("mime_type") or "").startswith("image/"):
            try:
                images.append(client.storage.from_(BUCKET).download(row["ruta_storage"]))
            except Exception:
                continue
    return images
