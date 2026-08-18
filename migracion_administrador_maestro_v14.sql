-- Autoridad destructiva exclusiva para el Administrador Maestro.
-- No elimina datos al ejecutarse; únicamente crea las reglas de autorización.

create or replace function public.es_administrador_maestro()
returns boolean
language sql
stable
security definer
set search_path = public
as $$
  select lower(coalesce(auth.jwt() ->> 'email', '')) = 'yani.limberopulos@jalisco.gob.mx'
     and exists (
       select 1
       from public.usuarios_autorizados u
       where u.user_id = auth.uid()
         and lower(u.email) = 'yani.limberopulos@jalisco.gob.mx'
         and u.activo = true
         and u.rol = 'administrador'
     );
$$;

revoke all on function public.es_administrador_maestro() from public;
grant execute on function public.es_administrador_maestro() to authenticated;

-- Sólo el Administrador Maestro puede borrar las entidades principales.
drop policy if exists "administrador maestro elimina proyectos" on public.proyectos;
create policy "administrador maestro elimina proyectos"
on public.proyectos for delete to authenticated
using (public.es_administrador_maestro());

drop policy if exists "administrador maestro elimina sesiones junta" on public.sesiones_junta;
create policy "administrador maestro elimina sesiones junta"
on public.sesiones_junta for delete to authenticated
using (public.es_administrador_maestro());

drop policy if exists "administrador maestro elimina sesiones comite" on public.sesiones_comite;
create policy "administrador maestro elimina sesiones comite"
on public.sesiones_comite for delete to authenticated
using (public.es_administrador_maestro());

-- Revoca políticas anteriores que permitían borrar documentos a cualquier
-- usuario del módulo y concentra también esa facultad en el Administrador Maestro.
drop policy if exists "usuarios eliminan documentos de sesión" on public.documentos_sesion_junta;
drop policy if exists "usuarios eliminan archivos junta" on public.archivos_acuerdo;
drop policy if exists "usuarios eliminan documentos comite" on public.documentos_sesion_comite;

drop policy if exists "administrador maestro elimina documentos junta" on public.documentos_sesion_junta;
create policy "administrador maestro elimina documentos junta"
on public.documentos_sesion_junta for delete to authenticated
using (public.es_administrador_maestro());

drop policy if exists "administrador maestro elimina archivos acuerdo" on public.archivos_acuerdo;
create policy "administrador maestro elimina archivos acuerdo"
on public.archivos_acuerdo for delete to authenticated
using (public.es_administrador_maestro());

drop policy if exists "administrador maestro elimina documentos comite" on public.documentos_sesion_comite;
create policy "administrador maestro elimina documentos comite"
on public.documentos_sesion_comite for delete to authenticated
using (public.es_administrador_maestro());

-- Permite retirar del bucket los archivos vinculados antes de borrar el registro.
drop policy if exists "administrador maestro elimina expedientes" on storage.objects;
drop policy if exists "usuarios oficiales eliminan expedientes" on storage.objects;
create policy "administrador maestro elimina expedientes"
on storage.objects for delete to authenticated
using (bucket_id = 'expedientes' and public.es_administrador_maestro());

notify pgrst, 'reload schema';
