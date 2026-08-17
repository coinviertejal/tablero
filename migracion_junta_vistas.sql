-- Migración puntual: metadatos de sesión y Órgano Interno de Control.
-- No elimina sesiones, acuerdos, usuarios, comentarios ni archivos.

alter table public.sesiones_junta add column if not exists videograbacion_url text;
alter table public.sesiones_junta add column if not exists acta_firmada_nombre text;
alter table public.sesiones_junta add column if not exists acta_firmada_ruta text;

alter table public.usuarios_autorizados drop constraint if exists usuarios_autorizados_direccion_check;
alter table public.usuarios_autorizados add constraint usuarios_autorizados_direccion_check
check (
  direccion is null or direccion in (
    'Dirección General',
    'Dirección Jurídica',
    'Dirección de Operaciones',
    'Dirección de Planeación',
    'Órgano Interno de Control'
  )
);

-- Retira únicamente puntos protocolarios que hubieran quedado guardados por
-- ingestas anteriores. Los documentos y demás acuerdos se conservan.
delete from public.acuerdos_junta
where titulo ~* '(clausura([[:space:]]+de)?[[:space:]]+la[[:space:]]+sesión|asuntos[[:space:]]+varios)';

notify pgrst, 'reload schema';

select column_name, data_type
from information_schema.columns
where table_schema = 'public'
  and table_name = 'sesiones_junta'
  and column_name in ('videograbacion_url','acta_firmada_nombre','acta_firmada_ruta')
order by column_name;
