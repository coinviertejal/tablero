-- COINVIERTE · Ingesta del archivo de control de Oficios Dirección General
-- Etapa 1: únicamente la hoja ENVIADOS DG.
-- Ejecutar en Supabase SQL Editor antes de publicar el app.py actualizado.

create extension if not exists pgcrypto;

-- 1) Bitácora de ingestiones.
create table if not exists public.ingestas_oficios_dg (
    id uuid primary key default gen_random_uuid(),
    nombre_archivo text not null,
    hash_archivo text,
    hoja text not null default 'ENVIADOS DG',
    registros_detectados integer not null default 0,
    registros_nuevos integer not null default 0,
    registros_actualizados integer not null default 0,
    registros_omitidos integer not null default 0,
    estado text not null default 'Procesando',
    detalle_error text,
    ingestado_por uuid references auth.users(id),
    autor_nombre text,
    created_at timestamptz not null default now(),
    completed_at timestamptz
);

alter table public.ingestas_oficios_dg enable row level security;

drop policy if exists "ingestas_oficios_dg_select_auth" on public.ingestas_oficios_dg;
create policy "ingestas_oficios_dg_select_auth"
on public.ingestas_oficios_dg for select
to authenticated
using (true);

drop policy if exists "ingestas_oficios_dg_insert_auth" on public.ingestas_oficios_dg;
create policy "ingestas_oficios_dg_insert_auth"
on public.ingestas_oficios_dg for insert
to authenticated
with check (auth.uid() = ingestado_por);

drop policy if exists "ingestas_oficios_dg_update_auth" on public.ingestas_oficios_dg;
create policy "ingestas_oficios_dg_update_auth"
on public.ingestas_oficios_dg for update
to authenticated
using (true)
with check (true);

-- 2) Ampliar la tabla existente para soportar registros provenientes del Excel
--    aun cuando todavía no se haya cargado el oficio firmado.
alter table public.oficios_direccion_general
    alter column fecha_oficio drop not null,
    alter column asunto drop not null,
    alter column nombre_archivo drop not null,
    alter column ruta_storage drop not null;

alter table public.oficios_direccion_general
    add column if not exists folio_control text,
    add column if not exists cargo text,
    add column if not exists dependencia text,
    add column if not exists fecha_control date,
    add column if not exists firma text,
    add column if not exists solicitado_por text,
    add column if not exists status_control text,
    add column if not exists archivo_fisico boolean,
    add column if not exists archivo_digital boolean,
    add column if not exists origen text not null default 'manual',
    add column if not exists hoja_origen text,
    add column if not exists fila_origen integer,
    add column if not exists clave_control text,
    add column if not exists ingesta_id uuid references public.ingestas_oficios_dg(id) on delete set null,
    add column if not exists registrado_por uuid references auth.users(id),
    add column if not exists registrado_por_nombre text;

create unique index if not exists uq_oficios_dg_clave_control
    on public.oficios_direccion_general(clave_control)
    where clave_control is not null;

create index if not exists idx_oficios_dg_ingesta
    on public.oficios_direccion_general(ingesta_id);

create index if not exists idx_oficios_dg_origen
    on public.oficios_direccion_general(origen);

-- Las políticas existentes de oficios_direccion_general siguen vigentes.
-- Los documentos firmados continúan guardándose en el bucket existente "expedientes".
