-- Documentos de seguimiento vinculados a acuerdos de Comités.
-- No modifica ni elimina acuerdos o documentos existentes.

create table if not exists public.archivos_acuerdo_comite (
  id uuid primary key default gen_random_uuid(),
  acuerdo_id uuid not null references public.acuerdos_comite(id) on delete cascade,
  nombre_visible text not null,
  nombre_archivo text not null,
  ruta_storage text not null unique,
  mime_type text,
  tamano_bytes bigint not null,
  subido_por uuid not null references auth.users(id),
  autor_nombre text not null,
  created_at timestamptz not null default now()
);

alter table public.archivos_acuerdo_comite enable row level security;

drop policy if exists "usuarios consultan archivos acuerdos comite" on public.archivos_acuerdo_comite;
drop policy if exists "usuarios crean archivos acuerdos comite" on public.archivos_acuerdo_comite;
drop policy if exists "administrador maestro elimina archivos acuerdos comite" on public.archivos_acuerdo_comite;

create policy "usuarios consultan archivos acuerdos comite"
on public.archivos_acuerdo_comite for select to authenticated
using (public.tiene_modulo('Comités'));

create policy "usuarios crean archivos acuerdos comite"
on public.archivos_acuerdo_comite for insert to authenticated
with check (public.tiene_modulo('Comités') and subido_por = auth.uid());

create policy "administrador maestro elimina archivos acuerdos comite"
on public.archivos_acuerdo_comite for delete to authenticated
using (public.es_administrador_maestro());

create index if not exists archivos_acuerdo_comite_idx
on public.archivos_acuerdo_comite (acuerdo_id, created_at);

notify pgrst, 'reload schema';
