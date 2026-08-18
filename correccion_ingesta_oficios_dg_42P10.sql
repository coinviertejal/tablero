-- Corrección COINVIERTE · Ingesta Oficios Dirección General
-- Soluciona error 42P10: no unique or exclusion constraint matching ON CONFLICT.

-- El índice parcial anterior no puede ser inferido por ON CONFLICT (clave_control).
drop index if exists public.uq_oficios_dg_clave_control;

-- Un índice UNIQUE normal permite múltiples NULL en PostgreSQL,
-- por lo que no necesitamos la cláusula WHERE.
create unique index if not exists uq_oficios_dg_clave_control
    on public.oficios_direccion_general (clave_control);

-- Verificación opcional: debe devolver una fila con indisunique = true e indpred = null.
select
    indexrelid::regclass as indice,
    indisunique,
    pg_get_expr(indpred, indrelid) as predicado
from pg_index
where indexrelid = 'public.uq_oficios_dg_clave_control'::regclass;
