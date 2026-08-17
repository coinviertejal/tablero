-- Ejecutar una sola vez en Supabase > SQL Editor.
create extension if not exists pgcrypto;

create table if not exists public.usuarios_autorizados (
  id uuid primary key default gen_random_uuid(),
  user_id uuid unique references auth.users(id) on delete set null,
  email text not null unique check (lower(email) like '%@jalisco.gob.mx'),
  nombre text,
  rol text not null default 'usuario' check (rol in ('administrador','usuario')),
  direccion text check (direccion in ('Dirección General','Dirección Jurídica','Dirección de Operaciones','Dirección de Planeación')),
  modulos jsonb not null default '["Programas / Proyectos","Junta de Gobierno","Comités"]'::jsonb,
  direcciones_proyectos jsonb not null default '["Dirección de Operaciones","Dirección de Proyectos"]'::jsonb,
  activo boolean not null default false,
  ultimo_acceso timestamptz,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);
alter table public.usuarios_autorizados add column if not exists direccion text;
alter table public.usuarios_autorizados add column if not exists modulos jsonb not null default '["Programas / Proyectos","Junta de Gobierno","Comités"]'::jsonb;
alter table public.usuarios_autorizados add column if not exists direcciones_proyectos jsonb not null default '["Dirección de Operaciones","Dirección de Proyectos"]'::jsonb;

create table if not exists public.codigos_acceso (
  id uuid primary key default gen_random_uuid(),
  email text not null,
  codigo_hash text not null,
  expires_at timestamptz not null,
  usado_at timestamptz,
  creado_por uuid references auth.users(id),
  created_at timestamptz not null default now()
);

create or replace function public.es_administrador()
returns boolean language sql stable security definer set search_path=public as $$
  select exists(select 1 from public.usuarios_autorizados
                where user_id=auth.uid() and activo=true and rol='administrador');
$$;

create or replace function public.esta_autorizado()
returns boolean language sql stable security definer set search_path=public as $$
  select exists(select 1 from public.usuarios_autorizados where user_id=auth.uid() and activo=true);
$$;

create or replace function public.tiene_modulo(p_modulo text)
returns boolean language sql stable security definer set search_path=public as $$
  select exists(select 1 from public.usuarios_autorizados
    where user_id=auth.uid() and activo=true and (rol='administrador' or modulos ? p_modulo));
$$;

create or replace function public.tiene_direccion_proyecto(p_direccion text)
returns boolean language sql stable security definer set search_path=public as $$
  select exists(select 1 from public.usuarios_autorizados
    where user_id=auth.uid() and activo=true and (rol='administrador' or direcciones_proyectos ? p_direccion));
$$;

create or replace function public.vincular_usuario_autorizado()
returns trigger language plpgsql security definer set search_path=public as $$
begin
  update public.usuarios_autorizados set user_id=new.id, updated_at=now()
  where lower(email)=lower(new.email) and user_id is null;
  return new;
end; $$;

drop trigger if exists vincular_usuario_autorizado_trigger on auth.users;
create trigger vincular_usuario_autorizado_trigger after insert or update of email on auth.users
for each row execute function public.vincular_usuario_autorizado();

drop function if exists public.crear_codigo_acceso(text,text,integer);
create or replace function public.crear_codigo_acceso(
  p_email text, p_nombre text default null, p_horas integer default 24,
  p_direccion text default null, p_modulos jsonb default '[]'::jsonb,
  p_direcciones_proyectos jsonb default '[]'::jsonb)
returns table(codigo text, vence timestamptz) language plpgsql security definer set search_path=public as $$
declare v_codigo text; v_vence timestamptz;
begin
  if not public.es_administrador() then raise exception 'Acceso no autorizado'; end if;
  p_email := lower(trim(p_email));
  if p_email not like '%@jalisco.gob.mx' then raise exception 'Correo institucional no válido'; end if;
  v_codigo := upper(substr(encode(gen_random_bytes(6),'hex'),1,8));
  v_vence := now() + make_interval(hours => greatest(1,least(p_horas,168)));
  insert into public.usuarios_autorizados(email,nombre,rol,direccion,modulos,direcciones_proyectos,activo)
  values(p_email,nullif(trim(p_nombre),''),'usuario',p_direccion,coalesce(p_modulos,'[]'::jsonb),coalesce(p_direcciones_proyectos,'[]'::jsonb),false)
  on conflict(email) do update set nombre=coalesce(excluded.nombre,usuarios_autorizados.nombre),
    direccion=excluded.direccion,modulos=excluded.modulos,direcciones_proyectos=excluded.direcciones_proyectos,updated_at=now();
  update public.codigos_acceso set usado_at=now() where lower(email)=p_email and usado_at is null;
  insert into public.codigos_acceso(email,codigo_hash,expires_at,creado_por)
  values(p_email,encode(digest(v_codigo,'sha256'),'hex'),v_vence,auth.uid());
  return query select v_codigo,v_vence;
end; $$;

create or replace function public.remover_usuario_autorizado(p_usuario_id uuid)
returns boolean language plpgsql security definer set search_path=public as $$
declare v_email text; v_rol text;
begin
  if not public.es_administrador() then raise exception 'Acceso no autorizado'; end if;
  select email,rol into v_email,v_rol from public.usuarios_autorizados where id=p_usuario_id;
  if v_email is null then return false; end if;
  if v_rol='administrador' then raise exception 'No se puede remover al administrador'; end if;
  update public.codigos_acceso set usado_at=coalesce(usado_at,now()) where lower(email)=lower(v_email);
  delete from public.usuarios_autorizados where id=p_usuario_id;
  return true;
end; $$;

create or replace function public.canjear_codigo_acceso(p_email text, p_codigo text)
returns boolean language plpgsql security definer set search_path=public as $$
declare v_id uuid;
begin
  select id into v_id from public.codigos_acceso
  where lower(email)=lower(trim(p_email)) and codigo_hash=encode(digest(upper(trim(p_codigo)),'sha256'),'hex')
    and usado_at is null and expires_at>now() order by created_at desc limit 1 for update;
  if v_id is null then return false; end if;
  update public.codigos_acceso set usado_at=now() where id=v_id;
  update public.usuarios_autorizados set activo=true,updated_at=now() where lower(email)=lower(trim(p_email));
  return true;
end; $$;

create or replace function public.registrar_acceso()
returns void language sql security definer set search_path=public as $$
  update public.usuarios_autorizados set ultimo_acceso=now(),updated_at=now() where user_id=auth.uid() and activo=true;
$$;

grant execute on function public.canjear_codigo_acceso(text,text) to anon, authenticated;
grant execute on function public.crear_codigo_acceso(text,text,integer,text,jsonb,jsonb) to authenticated;
grant execute on function public.remover_usuario_autorizado(uuid) to authenticated;
grant execute on function public.registrar_acceso() to authenticated;

alter table public.usuarios_autorizados enable row level security;
alter table public.codigos_acceso enable row level security;
drop policy if exists "usuarios consultan su acceso" on public.usuarios_autorizados;
drop policy if exists "administrador gestiona usuarios" on public.usuarios_autorizados;
create policy "usuarios consultan su acceso" on public.usuarios_autorizados for select to authenticated
using (user_id=auth.uid() or public.es_administrador());
create policy "administrador gestiona usuarios" on public.usuarios_autorizados for update to authenticated
using (public.es_administrador()) with check (public.es_administrador());

-- Administrador inicial. Cambiar el correo aquí si fuera necesario.
insert into public.usuarios_autorizados(email,nombre,rol,activo)
values('yani.limberopulos@jalisco.gob.mx','Yani Limberopulos','administrador',true)
on conflict(email) do update set rol='administrador',activo=true,updated_at=now();
update public.usuarios_autorizados au set user_id=u.id from auth.users u
where lower(au.email)=lower(u.email) and au.user_id is null;

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
  monitoreo jsonb not null default '{}'::jsonb,
  avance_proyecto jsonb not null default '{}'::jsonb,
  creado_por uuid not null references auth.users(id),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

-- Compatible con bases creadas con una versión anterior del esquema.
alter table public.proyectos add column if not exists monitoreo jsonb not null default '{}'::jsonb;
alter table public.proyectos add column if not exists avance_proyecto jsonb not null default '{}'::jsonb;

create table if not exists public.documentos (
  id uuid primary key default gen_random_uuid(),
  proyecto_id uuid not null references public.proyectos(id) on delete cascade,
  categoria text not null check (categoria in ('juridica','auxiliar','acta_comite','acta_junta','convenio','fotografia','evidencia_meta')),
  nombre_archivo text not null,
  ruta_storage text not null unique,
  mime_type text,
  tamano_bytes bigint not null,
  created_at timestamptz not null default now()
);

-- Amplía las categorías permitidas en instalaciones existentes.
alter table public.documentos drop constraint if exists documentos_categoria_check;
alter table public.documentos add constraint documentos_categoria_check
check (categoria in ('juridica','auxiliar','acta_comite','acta_junta','convenio','fotografia','evidencia_meta'));

alter table public.proyectos enable row level security;
alter table public.documentos enable row level security;

drop policy if exists "usuarios oficiales consultan proyectos" on public.proyectos;
drop policy if exists "usuarios oficiales crean proyectos" on public.proyectos;
drop policy if exists "usuarios oficiales actualizan proyectos" on public.proyectos;
drop policy if exists "usuarios oficiales consultan documentos" on public.documentos;
drop policy if exists "usuarios oficiales registran documentos" on public.documentos;
create policy "usuarios oficiales consultan proyectos" on public.proyectos for select to authenticated
using (public.tiene_modulo('Programas / Proyectos') and public.tiene_direccion_proyecto(direccion));
create policy "usuarios oficiales crean proyectos" on public.proyectos for insert to authenticated
with check (public.tiene_modulo('Programas / Proyectos') and public.tiene_direccion_proyecto(direccion) and creado_por = auth.uid());
create policy "usuarios oficiales actualizan proyectos" on public.proyectos for update to authenticated
using (public.tiene_modulo('Programas / Proyectos') and public.tiene_direccion_proyecto(direccion))
with check (public.tiene_modulo('Programas / Proyectos') and public.tiene_direccion_proyecto(direccion));

create policy "usuarios oficiales consultan documentos" on public.documentos for select to authenticated
using (exists(select 1 from public.proyectos p where p.id=proyecto_id));
create policy "usuarios oficiales registran documentos" on public.documentos for insert to authenticated
with check (exists(select 1 from public.proyectos p where p.id=proyecto_id));

insert into storage.buckets (id, name, public, file_size_limit)
values ('expedientes', 'expedientes', false, 52428800)
on conflict (id) do nothing;

drop policy if exists "usuarios oficiales suben expedientes" on storage.objects;
drop policy if exists "usuarios oficiales leen expedientes" on storage.objects;
create policy "usuarios oficiales suben expedientes" on storage.objects for insert to authenticated
with check (bucket_id = 'expedientes' and public.esta_autorizado());
create policy "usuarios oficiales leen expedientes" on storage.objects for select to authenticated
using (bucket_id = 'expedientes' and public.esta_autorizado());

-- Junta de Gobierno: catálogo de sesiones y acuerdos consultables.
create table if not exists public.sesiones_junta (
  id uuid primary key default gen_random_uuid(),
  anio integer not null check (anio between 2025 and 2030),
  tipo text not null check (tipo in ('Ordinaria','Extraordinaria')),
  nombre text not null,
  fecha_sesion date,
  creado_por uuid not null references auth.users(id),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (anio, tipo, nombre)
);
alter table public.sesiones_junta add column if not exists fecha_sesion date;

-- Sólo ejercicios oficiales 2025-2030.
delete from public.sesiones_junta where anio = 0;
alter table public.sesiones_junta drop constraint if exists sesiones_junta_anio_check;
alter table public.sesiones_junta add constraint sesiones_junta_anio_check
check (anio between 2025 and 2030);

create table if not exists public.acuerdos_junta (
  id uuid primary key default gen_random_uuid(),
  sesion_id uuid not null references public.sesiones_junta(id) on delete cascade,
  numero text,
  titulo text not null,
  texto text not null default '',
  tipo_registro text not null default 'Acuerdo' check (tipo_registro in ('Acuerdo','Informe')),
  areas jsonb not null default '[]'::jsonb,
  estatus text not null default 'Por iniciar' check (estatus in ('Por iniciar','En proceso','Terminada')),
  fecha_compromiso date,
  fecha_cierre date,
  cumplimiento text check (cumplimiento in ('En tiempo','Extemporáneo','Sin fecha compromiso')),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

alter table public.acuerdos_junta add column if not exists tipo_registro text not null default 'Acuerdo';
alter table public.acuerdos_junta add column if not exists areas jsonb not null default '[]'::jsonb;
alter table public.acuerdos_junta add column if not exists estatus text not null default 'Por iniciar';
alter table public.acuerdos_junta add column if not exists fecha_compromiso date;
alter table public.acuerdos_junta add column if not exists fecha_cierre date;
alter table public.acuerdos_junta add column if not exists cumplimiento text;

create table if not exists public.comentarios_acuerdo (
  id uuid primary key default gen_random_uuid(), acuerdo_id uuid not null references public.acuerdos_junta(id) on delete cascade,
  autor_id uuid not null references auth.users(id), autor_nombre text not null, comentario text not null,
  created_at timestamptz not null default now()
);
create table if not exists public.archivos_acuerdo (
  id uuid primary key default gen_random_uuid(), acuerdo_id uuid not null references public.acuerdos_junta(id) on delete cascade,
  nombre_archivo text not null, ruta_storage text not null unique, mime_type text, tamano_bytes bigint not null,
  subido_por uuid not null references auth.users(id), created_at timestamptz not null default now()
);
create table if not exists public.historial_acuerdo (
  id uuid primary key default gen_random_uuid(), acuerdo_id uuid not null references public.acuerdos_junta(id) on delete cascade,
  autor_id uuid not null references auth.users(id), autor_nombre text not null, descripcion text not null,
  created_at timestamptz not null default now()
);

alter table public.sesiones_junta enable row level security;
alter table public.acuerdos_junta enable row level security;
alter table public.comentarios_acuerdo enable row level security;
alter table public.archivos_acuerdo enable row level security;
alter table public.historial_acuerdo enable row level security;
drop policy if exists "usuarios consultan sesiones junta" on public.sesiones_junta;
drop policy if exists "usuarios crean sesiones junta" on public.sesiones_junta;
drop policy if exists "usuarios actualizan sesiones junta" on public.sesiones_junta;
drop policy if exists "usuarios consultan acuerdos junta" on public.acuerdos_junta;
drop policy if exists "usuarios crean acuerdos junta" on public.acuerdos_junta;
drop policy if exists "usuarios actualizan acuerdos junta" on public.acuerdos_junta;
create policy "usuarios consultan sesiones junta" on public.sesiones_junta for select to authenticated
using (public.tiene_modulo('Junta de Gobierno'));
create policy "usuarios crean sesiones junta" on public.sesiones_junta for insert to authenticated
with check (public.tiene_modulo('Junta de Gobierno') and creado_por=auth.uid());
create policy "usuarios actualizan sesiones junta" on public.sesiones_junta for update to authenticated
using (public.tiene_modulo('Junta de Gobierno')) with check (public.tiene_modulo('Junta de Gobierno'));
create policy "usuarios consultan acuerdos junta" on public.acuerdos_junta for select to authenticated
using (public.tiene_modulo('Junta de Gobierno'));
create policy "usuarios crean acuerdos junta" on public.acuerdos_junta for insert to authenticated
with check (public.tiene_modulo('Junta de Gobierno'));
create policy "usuarios actualizan acuerdos junta" on public.acuerdos_junta for update to authenticated
using (public.tiene_modulo('Junta de Gobierno')) with check (public.tiene_modulo('Junta de Gobierno'));
drop policy if exists "usuarios consultan comentarios junta" on public.comentarios_acuerdo;
drop policy if exists "usuarios crean comentarios junta" on public.comentarios_acuerdo;
create policy "usuarios consultan comentarios junta" on public.comentarios_acuerdo for select to authenticated using (public.tiene_modulo('Junta de Gobierno'));
create policy "usuarios crean comentarios junta" on public.comentarios_acuerdo for insert to authenticated with check (public.tiene_modulo('Junta de Gobierno') and autor_id=auth.uid());
drop policy if exists "usuarios consultan archivos junta" on public.archivos_acuerdo;
drop policy if exists "usuarios crean archivos junta" on public.archivos_acuerdo;
create policy "usuarios consultan archivos junta" on public.archivos_acuerdo for select to authenticated using (public.tiene_modulo('Junta de Gobierno'));
create policy "usuarios crean archivos junta" on public.archivos_acuerdo for insert to authenticated with check (public.tiene_modulo('Junta de Gobierno') and subido_por=auth.uid());
drop policy if exists "usuarios consultan historial junta" on public.historial_acuerdo;
drop policy if exists "usuarios crean historial junta" on public.historial_acuerdo;
create policy "usuarios consultan historial junta" on public.historial_acuerdo for select to authenticated using (public.tiene_modulo('Junta de Gobierno'));
create policy "usuarios crean historial junta" on public.historial_acuerdo for insert to authenticated with check (public.tiene_modulo('Junta de Gobierno') and autor_id=auth.uid());

create index if not exists acuerdos_junta_busqueda_idx
on public.acuerdos_junta using gin (to_tsvector('spanish', coalesce(numero,'') || ' ' || titulo || ' ' || texto));

-- Sesiones iniciales solicitadas para 2025 y 2026.
insert into public.sesiones_junta (anio,tipo,nombre,creado_por)
select y.anio, t.tipo, n.nombre, au.user_id
from (values (2025),(2026)) as y(anio)
cross join (values ('Ordinaria'),('Extraordinaria')) as t(tipo)
cross join (values ('Primera (1era)'),('Segunda (2da)'),('Tercera (3ra)')) as n(nombre)
cross join lateral (
  select user_id from public.usuarios_autorizados
  where rol='administrador' and activo=true and user_id is not null limit 1
) au
on conflict (anio,tipo,nombre) do nothing;
