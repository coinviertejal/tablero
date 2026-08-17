-- Documentos visibles de sesiones y acuerdos.
alter table public.archivos_acuerdo add column if not exists nombre_visible text;
alter table public.archivos_acuerdo add column if not exists autor_nombre text;

create table if not exists public.documentos_sesion_junta (
  id uuid primary key default gen_random_uuid(),
  sesion_id uuid not null references public.sesiones_junta(id) on delete cascade,
  tipo_documento text not null,
  nombre_visible text not null,
  nombre_archivo text not null,
  ruta_storage text not null unique,
  mime_type text,
  tamano_bytes bigint not null,
  subido_por uuid not null references auth.users(id),
  autor_nombre text not null,
  created_at timestamptz not null default now()
);

alter table public.documentos_sesion_junta enable row level security;
drop policy if exists "usuarios consultan documentos de sesión" on public.documentos_sesion_junta;
drop policy if exists "usuarios crean documentos de sesión" on public.documentos_sesion_junta;
create policy "usuarios consultan documentos de sesión" on public.documentos_sesion_junta
for select to authenticated using (public.tiene_modulo('Junta de Gobierno'));
create policy "usuarios crean documentos de sesión" on public.documentos_sesion_junta
for insert to authenticated with check (public.tiene_modulo('Junta de Gobierno') and subido_por=auth.uid());

notify pgrst, 'reload schema';
