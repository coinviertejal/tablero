-- Resultado de acuerdos para la analítica de Junta de Gobierno.
-- Los acuerdos existentes permanecen pendientes hasta su clasificación manual.

alter table public.acuerdos_junta
add column if not exists resultado text not null default 'Pendiente';

alter table public.acuerdos_junta
drop constraint if exists acuerdos_junta_resultado_check;

alter table public.acuerdos_junta
add constraint acuerdos_junta_resultado_check
check (resultado in ('Pendiente','Aprobado','Rechazado'));

create index if not exists acuerdos_junta_resultado_idx
on public.acuerdos_junta (resultado, estatus);

notify pgrst, 'reload schema';
