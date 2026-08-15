from __future__ import annotations

import re
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


def client_with_token(access_token: str, refresh_token: str) -> Client:
    """Crea un cliente completo que actúa con la sesión del usuario."""
    client = create_client(st.secrets["SUPABASE_URL"], public_key())
    client.auth.set_session(access_token, refresh_token)
    return client


def access_profile(client: Client, email: str) -> dict | None:
    rows = client.table("usuarios_autorizados").select("id,email,nombre,rol,activo").eq("email", email.lower()).execute().data
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
