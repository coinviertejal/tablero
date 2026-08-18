-- COINVIERTE · Dirección de Proyectos
-- Transferencias, comprobación documental y avance financiero automático
-- Ejecutar una sola vez en Supabase SQL Editor.

create extension if not exists pgcrypto;

create table if not exists public.transferencias_proyecto (
    id uuid primary key default gen_random_uuid(),
    proyecto_id uuid not null references public.proyectos(id) on delete cascade,
    fecha_transferencia date not null,
    importe numeric(16,2) not null check (importe > 0),
    concepto text not null,
    beneficiario text,
    referencia text,
    creado_por uuid references auth.users(id),
    autor_nombre text,
    created_at timestamptz not null default now(),
    updated_at timestamptz not null default now()
);

create index if not exists idx_transferencias_proyecto_proyecto
    on public.transferencias_proyecto(proyecto_id);
create index if not exists idx_transferencias_proyecto_fecha
    on public.transferencias_proyecto(proyecto_id, fecha_transferencia);

create table if not exists public.documentos_transferencia_proyecto (
    id uuid primary key default gen_random_uuid(),
    transferencia_id uuid not null references public.transferencias_proyecto(id) on delete cascade,
    tipo_documento text not null check (
        tipo_documento in ('Comprobante de transferencia', 'Factura / CFDI', 'Documento soporte')
    ),
    nombre_visible text,
    nombre_archivo text not null,
    ruta_storage text not null,
    mime_type text,
    tamano_bytes bigint,
    subido_por uuid references auth.users(id),
    autor_nombre text,
    created_at timestamptz not null default now()
);

create index if not exists idx_documentos_transferencia_transferencia
    on public.documentos_transferencia_proyecto(transferencia_id);

alter table public.transferencias_proyecto enable row level security;
alter table public.documentos_transferencia_proyecto enable row level security;

drop policy if exists "transferencias_select_auth" on public.transferencias_proyecto;
create policy "transferencias_select_auth" on public.transferencias_proyecto
for select to authenticated using (true);

drop policy if exists "transferencias_insert_auth" on public.transferencias_proyecto;
create policy "transferencias_insert_auth" on public.transferencias_proyecto
for insert to authenticated with check (auth.uid() = creado_por);

drop policy if exists "transferencias_update_auth" on public.transferencias_proyecto;
create policy "transferencias_update_auth" on public.transferencias_proyecto
for update to authenticated using (true) with check (true);

drop policy if exists "transferencias_delete_auth" on public.transferencias_proyecto;
create policy "transferencias_delete_auth" on public.transferencias_proyecto
for delete to authenticated using (true);

drop policy if exists "documentos_transferencia_select_auth" on public.documentos_transferencia_proyecto;
create policy "documentos_transferencia_select_auth" on public.documentos_transferencia_proyecto
for select to authenticated using (true);

drop policy if exists "documentos_transferencia_insert_auth" on public.documentos_transferencia_proyecto;
create policy "documentos_transferencia_insert_auth" on public.documentos_transferencia_proyecto
for insert to authenticated with check (auth.uid() = subido_por);

drop policy if exists "documentos_transferencia_delete_auth" on public.documentos_transferencia_proyecto;
create policy "documentos_transferencia_delete_auth" on public.documentos_transferencia_proyecto
for delete to authenticated using (true);

-- Se reutiliza el bucket existente "expedientes".
-- Rutas: proyectos/{proyecto_id}/transferencias/{transferencia_id}/...
