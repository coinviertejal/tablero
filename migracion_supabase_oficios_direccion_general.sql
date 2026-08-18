-- COINVIERTE · Módulo Oficios Dirección General
-- Ejecutar en Supabase SQL Editor antes de publicar el nuevo app.py.

create extension if not exists pgcrypto;

create table if not exists public.oficios_direccion_general (
    id uuid primary key default gen_random_uuid(),
    anio integer not null check (anio between 2024 and 2030),
    mes integer not null check (mes between 1 and 12),
    numero_oficio text,
    fecha_oficio date not null,
    asunto text not null,
    destinatario text,
    notas text,
    nombre_archivo text not null,
    ruta_storage text not null,
    mime_type text,
    tamano_bytes bigint,
    subido_por uuid references auth.users(id),
    autor_nombre text,
    created_at timestamptz not null default now(),
    updated_at timestamptz not null default now()
);

create index if not exists idx_oficios_dg_anio_mes
    on public.oficios_direccion_general(anio, mes);

create index if not exists idx_oficios_dg_fecha
    on public.oficios_direccion_general(fecha_oficio desc);

create index if not exists idx_oficios_dg_numero
    on public.oficios_direccion_general(numero_oficio);

alter table public.oficios_direccion_general enable row level security;

drop policy if exists "oficios_dg_select_auth" on public.oficios_direccion_general;
create policy "oficios_dg_select_auth"
on public.oficios_direccion_general for select
to authenticated
using (true);

drop policy if exists "oficios_dg_insert_auth" on public.oficios_direccion_general;
create policy "oficios_dg_insert_auth"
on public.oficios_direccion_general for insert
to authenticated
with check (auth.uid() = subido_por);

drop policy if exists "oficios_dg_update_auth" on public.oficios_direccion_general;
create policy "oficios_dg_update_auth"
on public.oficios_direccion_general for update
to authenticated
using (true)
with check (true);

drop policy if exists "oficios_dg_delete_auth" on public.oficios_direccion_general;
create policy "oficios_dg_delete_auth"
on public.oficios_direccion_general for delete
to authenticated
using (true);
