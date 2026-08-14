-- Ejecutar una sola vez en Supabase > SQL Editor.
create extension if not exists pgcrypto;

create table if not exists public.proyectos (
  id uuid primary key default gen_random_uuid(),
  direccion text not null check (direccion in ('Dirección de Operaciones', 'Dirección de Proyectos')),
  nombre text not null,
  solicitante text not null,
  municipio text not null,
  anio_inicio integer not null check (anio_inicio between 2000 and 2100),
  monto numeric(16,2) not null check (monto >= 0),
  objetivo_general text not null,
  objetivos_especificos jsonb not null default '[]'::jsonb,
  creado_por uuid not null references auth.users(id),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.documentos (
  id uuid primary key default gen_random_uuid(),
  proyecto_id uuid not null references public.proyectos(id) on delete cascade,
  categoria text not null check (categoria in ('juridica','auxiliar','acta_comite','acta_junta','convenio','fotografia')),
  nombre_archivo text not null,
  ruta_storage text not null unique,
  mime_type text,
  tamano_bytes bigint not null,
  created_at timestamptz not null default now()
);

alter table public.proyectos enable row level security;
alter table public.documentos enable row level security;

create policy "usuarios oficiales consultan proyectos" on public.proyectos for select to authenticated
using ((auth.jwt()->>'email') like '%@jalisco.gob.mx');
create policy "usuarios oficiales crean proyectos" on public.proyectos for insert to authenticated
with check ((auth.jwt()->>'email') like '%@jalisco.gob.mx' and creado_por = auth.uid());
create policy "usuarios oficiales actualizan proyectos" on public.proyectos for update to authenticated
using ((auth.jwt()->>'email') like '%@jalisco.gob.mx') with check ((auth.jwt()->>'email') like '%@jalisco.gob.mx');

create policy "usuarios oficiales consultan documentos" on public.documentos for select to authenticated
using ((auth.jwt()->>'email') like '%@jalisco.gob.mx');
create policy "usuarios oficiales registran documentos" on public.documentos for insert to authenticated
with check ((auth.jwt()->>'email') like '%@jalisco.gob.mx');

insert into storage.buckets (id, name, public, file_size_limit)
values ('expedientes', 'expedientes', false, 52428800)
on conflict (id) do nothing;

create policy "usuarios oficiales suben expedientes" on storage.objects for insert to authenticated
with check (bucket_id = 'expedientes' and (auth.jwt()->>'email') like '%@jalisco.gob.mx');
create policy "usuarios oficiales leen expedientes" on storage.objects for select to authenticated
using (bucket_id = 'expedientes' and (auth.jwt()->>'email') like '%@jalisco.gob.mx');

