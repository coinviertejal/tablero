-- COINVIERTE · Vinculación Google Drive ↔ Oficios Dirección General
-- NO duplica PDFs: guarda metadatos y vínculos a Google Drive.

alter table public.oficios_direccion_general
    add column if not exists drive_file_id text,
    add column if not exists drive_url text,
    add column if not exists drive_nombre_archivo text,
    add column if not exists drive_mime_type text,
    add column if not exists drive_size_bytes bigint,
    add column if not exists drive_modified_at timestamptz,
    add column if not exists drive_sincronizado_at timestamptz;

create unique index if not exists uq_oficios_dg_drive_file_id
    on public.oficios_direccion_general(drive_file_id)
    where drive_file_id is not null;

create table if not exists public.sincronizaciones_drive_oficios_dg (
    id uuid primary key default gen_random_uuid(),
    folder_id text not null,
    estado text not null default 'Procesando',
    archivos_detectados integer not null default 0,
    archivos_vinculados integer not null default 0,
    archivos_ambiguos integer not null default 0,
    archivos_sin_coincidencia integer not null default 0,
    detalle_error text,
    sincronizado_por uuid references auth.users(id),
    autor_nombre text,
    created_at timestamptz not null default now(),
    completed_at timestamptz
);

create table if not exists public.inventario_drive_oficios_dg (
    id uuid primary key default gen_random_uuid(),
    drive_file_id text not null unique,
    drive_nombre_archivo text not null,
    drive_url text,
    drive_mime_type text,
    drive_size_bytes bigint,
    drive_modified_at timestamptz,
    folio_detectado text,
    numero_detectado text,
    mes_detectado integer,
    anio_detectado integer,
    estado_match text not null default 'Sin coincidencia',
    criterio_match text,
    oficio_id uuid references public.oficios_direccion_general(id) on delete set null,
    ultima_sincronizacion_id uuid references public.sincronizaciones_drive_oficios_dg(id) on delete set null,
    created_at timestamptz not null default now(),
    updated_at timestamptz not null default now()
);

create index if not exists idx_inventario_drive_oficios_estado
    on public.inventario_drive_oficios_dg(estado_match);

alter table public.sincronizaciones_drive_oficios_dg enable row level security;
alter table public.inventario_drive_oficios_dg enable row level security;

drop policy if exists "sync_drive_oficios_select_auth" on public.sincronizaciones_drive_oficios_dg;
create policy "sync_drive_oficios_select_auth"
on public.sincronizaciones_drive_oficios_dg
for select to authenticated using (true);

drop policy if exists "sync_drive_oficios_insert_auth" on public.sincronizaciones_drive_oficios_dg;
create policy "sync_drive_oficios_insert_auth"
on public.sincronizaciones_drive_oficios_dg
for insert to authenticated
with check (auth.uid() = sincronizado_por);

drop policy if exists "sync_drive_oficios_update_auth" on public.sincronizaciones_drive_oficios_dg;
create policy "sync_drive_oficios_update_auth"
on public.sincronizaciones_drive_oficios_dg
for update to authenticated
using (true) with check (true);

drop policy if exists "inventory_drive_oficios_select_auth" on public.inventario_drive_oficios_dg;
create policy "inventory_drive_oficios_select_auth"
on public.inventario_drive_oficios_dg
for select to authenticated using (true);

drop policy if exists "inventory_drive_oficios_insert_auth" on public.inventario_drive_oficios_dg;
create policy "inventory_drive_oficios_insert_auth"
on public.inventario_drive_oficios_dg
for insert to authenticated with check (true);

drop policy if exists "inventory_drive_oficios_update_auth" on public.inventario_drive_oficios_dg;
create policy "inventory_drive_oficios_update_auth"
on public.inventario_drive_oficios_dg
for update to authenticated
using (true) with check (true);

-- La política UPDATE existente de oficios_direccion_general debe permitir
-- actualizar los campos drive_* a usuarios autenticados. Si ya corriste las
-- migraciones anteriores, esta política normalmente ya existe.
