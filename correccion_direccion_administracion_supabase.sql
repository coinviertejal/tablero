
-- COINVIERTE · Agregar Dirección de Administración al catálogo permitido
-- de usuarios_autorizados.
-- Idempotente y seguro para volver a ejecutar.

DO $$
DECLARE
    r record;
BEGIN
    -- Elimina únicamente CHECK constraints de usuarios_autorizados
    -- cuyo texto haga referencia a la columna direccion.
    FOR r IN
        SELECT c.conname
        FROM pg_constraint c
        JOIN pg_class t ON t.oid = c.conrelid
        JOIN pg_namespace n ON n.oid = t.relnamespace
        WHERE n.nspname = 'public'
          AND t.relname = 'usuarios_autorizados'
          AND c.contype = 'c'
          AND pg_get_constraintdef(c.oid) ILIKE '%direccion%'
    LOOP
        EXECUTE format(
            'ALTER TABLE public.usuarios_autorizados DROP CONSTRAINT IF EXISTS %I',
            r.conname
        );
    END LOOP;
END $$;

ALTER TABLE public.usuarios_autorizados
DROP CONSTRAINT IF EXISTS usuarios_autorizados_direccion_check;

ALTER TABLE public.usuarios_autorizados
ADD CONSTRAINT usuarios_autorizados_direccion_check
CHECK (
    direccion IS NULL
    OR direccion IN (
        'Dirección General',
        'Dirección de Administración',
        'Dirección Jurídica',
        'Dirección de Operaciones',
        'Dirección de Planeación',
        'Órgano Interno de Control'
    )
);

-- Verificación:
SELECT
    conname,
    pg_get_constraintdef(oid) AS definicion
FROM pg_constraint
WHERE conrelid = 'public.usuarios_autorizados'::regclass
  AND contype = 'c'
ORDER BY conname;
