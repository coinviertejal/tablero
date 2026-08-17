-- Permite borrar documentos de Junta de Gobierno sin alterar sesiones ni acuerdos.
-- Ejecutar una sola vez en Supabase SQL Editor.

alter table public.documentos_sesion_junta enable row level security;
alter table public.archivos_acuerdo enable row level security;

drop policy if exists "usuarios eliminan documentos de sesión" on public.documentos_sesion_junta;
create policy "usuarios eliminan documentos de sesión"
on public.documentos_sesion_junta
for delete to authenticated
using (public.tiene_modulo('Junta de Gobierno'));

drop policy if exists "usuarios eliminan archivos junta" on public.archivos_acuerdo;
create policy "usuarios eliminan archivos junta"
on public.archivos_acuerdo
for delete to authenticated
using (public.tiene_modulo('Junta de Gobierno'));

drop policy if exists "usuarios oficiales eliminan expedientes" on storage.objects;
create policy "usuarios oficiales eliminan expedientes"
on storage.objects
for delete to authenticated
using (
  bucket_id = 'expedientes'
  and public.tiene_modulo('Junta de Gobierno')
);

notify pgrst, 'reload schema';
