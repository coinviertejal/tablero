-- Expediente, acuerdos e informes de las sesiones de Comités.
-- Conserva intactas las sesiones ya creadas en public.sesiones_comite.

create table if not exists public.documentos_sesion_comite (
  id uuid primary key default gen_random_uuid(),
  sesion_id uuid not null references public.sesiones_comite(id) on delete cascade,
  tipo_documento text not null check (tipo_documento in ('Convocatoria / orden del día','Acta de la sesión','Otro')),
  nombre_visible text not null,
  nombre_archivo text not null,
  ruta_storage text not null unique,
  mime_type text,
  tamano_bytes bigint,
  subido_por uuid not null references auth.users(id),
  autor_nombre text,
  created_at timestamptz not null default now()
);

create table if not exists public.acuerdos_comite (
  id uuid primary key default gen_random_uuid(),
  sesion_id uuid not null references public.sesiones_comite(id) on delete cascade,
  numero text not null,
  tipo_registro text not null default 'Acuerdo' check (tipo_registro in ('Acuerdo','Informe')),
  titulo text not null,
  texto text,
  areas text[] not null default '{}',
  estatus text not null default 'Por iniciar' check (estatus in ('Por iniciar','En proceso','Terminada')),
  resultado text not null default 'Pendiente' check (resultado in ('Pendiente','Aprobado','Rechazado')),
  fecha_compromiso date,
  comentario_seguimiento text,
  actualizado_por uuid references auth.users(id),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (sesion_id, numero)
);

alter table public.documentos_sesion_comite enable row level security;
alter table public.acuerdos_comite enable row level security;

drop policy if exists "usuarios consultan documentos comite" on public.documentos_sesion_comite;
drop policy if exists "usuarios crean documentos comite" on public.documentos_sesion_comite;
drop policy if exists "usuarios eliminan documentos comite" on public.documentos_sesion_comite;
drop policy if exists "usuarios consultan acuerdos comite" on public.acuerdos_comite;
drop policy if exists "usuarios crean acuerdos comite" on public.acuerdos_comite;
drop policy if exists "usuarios actualizan acuerdos comite" on public.acuerdos_comite;

create policy "usuarios consultan documentos comite"
on public.documentos_sesion_comite for select to authenticated
using (public.tiene_modulo('Comités'));

create policy "usuarios crean documentos comite"
on public.documentos_sesion_comite for insert to authenticated
with check (public.tiene_modulo('Comités') and subido_por = auth.uid());

create policy "usuarios eliminan documentos comite"
on public.documentos_sesion_comite for delete to authenticated
using (public.tiene_modulo('Comités'));

create policy "usuarios consultan acuerdos comite"
on public.acuerdos_comite for select to authenticated
using (public.tiene_modulo('Comités'));

create policy "usuarios crean acuerdos comite"
on public.acuerdos_comite for insert to authenticated
with check (public.tiene_modulo('Comités'));

create policy "usuarios actualizan acuerdos comite"
on public.acuerdos_comite for update to authenticated
using (public.tiene_modulo('Comités'))
with check (public.tiene_modulo('Comités'));

create index if not exists documentos_sesion_comite_idx
on public.documentos_sesion_comite (sesion_id, created_at);

create index if not exists acuerdos_comite_sesion_idx
on public.acuerdos_comite (sesion_id, numero);

notify pgrst, 'reload schema';
