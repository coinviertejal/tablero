-- COINVIERTE · Vínculos Google Drive para Oficios Dirección General
-- Idempotente: puede ejecutarse aunque ya existan columnas previas.

alter table public.oficios_direccion_general
    add column if not exists drive_url text,
    add column if not exists drive_nombre_archivo text,
    add column if not exists drive_vinculado_at timestamptz;

create table if not exists public.ingestas_vinculos_drive_oficios_dg (
    id uuid primary key default gen_random_uuid(),
    nombre_archivo text,
    registros_detectados integer not null default 0,
    registros_vinculados integer not null default 0,
    registros_ambiguos integer not null default 0,
    registros_sin_coincidencia integer not null default 0,
    importado_por uuid references auth.users(id),
    autor_nombre text,
    detalle jsonb,
    created_at timestamptz not null default now()
);

alter table public.ingestas_vinculos_drive_oficios_dg enable row level security;

drop policy if exists "drive_links_select_auth" on public.ingestas_vinculos_drive_oficios_dg;
create policy "drive_links_select_auth"
on public.ingestas_vinculos_drive_oficios_dg
for select to authenticated
using (true);

drop policy if exists "drive_links_insert_auth" on public.ingestas_vinculos_drive_oficios_dg;
create policy "drive_links_insert_auth"
on public.ingestas_vinculos_drive_oficios_dg
for insert to authenticated
with check (auth.uid() = importado_por);
