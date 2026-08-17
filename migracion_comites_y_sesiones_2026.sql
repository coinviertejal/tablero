-- Comités institucionales y sesiones 2025-2030.
-- No modifica ni elimina sesiones de Junta de Gobierno.

create table if not exists public.sesiones_comite (
  id uuid primary key default gen_random_uuid(),
  comite text not null check (comite in (
    'Comité de Ética', 'Comité de Igualdad de Género',
    'Comité de Archivo', 'Comité de Control Interno'
  )),
  anio integer not null check (anio between 2025 and 2030),
  tipo text not null check (tipo in ('Ordinaria','Extraordinaria')),
  nombre text not null,
  fecha_sesion date,
  creado_por uuid not null references auth.users(id),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (comite, anio, tipo, nombre)
);

alter table public.sesiones_comite enable row level security;

drop policy if exists "usuarios consultan sesiones comite" on public.sesiones_comite;
drop policy if exists "usuarios crean sesiones comite" on public.sesiones_comite;
drop policy if exists "usuarios actualizan sesiones comite" on public.sesiones_comite;

create policy "usuarios consultan sesiones comite"
on public.sesiones_comite for select to authenticated
using (public.tiene_modulo('Comités'));

create policy "usuarios crean sesiones comite"
on public.sesiones_comite for insert to authenticated
with check (public.tiene_modulo('Comités') and creado_por = auth.uid());

create policy "usuarios actualizan sesiones comite"
on public.sesiones_comite for update to authenticated
using (public.tiene_modulo('Comités'))
with check (public.tiene_modulo('Comités'));

create index if not exists sesiones_comite_consulta_idx
on public.sesiones_comite (comite, anio, fecha_sesion, created_at);

notify pgrst, 'reload schema';
