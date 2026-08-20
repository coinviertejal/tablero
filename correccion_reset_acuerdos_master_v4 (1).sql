
-- ================================================================
-- COINVIERTE · Corrección definitiva v4
-- Reinicio maestro de seguimiento
-- "areas" es JSONB y NOT NULL
-- ================================================================

CREATE OR REPLACE FUNCTION public._es_admin_maestro_coinvierte()
RETURNS boolean
LANGUAGE sql
STABLE
SECURITY DEFINER
SET search_path = public
AS $$
    SELECT lower(coalesce(auth.jwt() ->> 'email', '')) = 'yani.limberopulos@jalisco.gob.mx';
$$;

REVOKE ALL ON FUNCTION public._es_admin_maestro_coinvierte() FROM PUBLIC;
GRANT EXECUTE ON FUNCTION public._es_admin_maestro_coinvierte() TO authenticated;


CREATE OR REPLACE FUNCTION public.reset_acuerdo_junta_master(p_acuerdo_id uuid)
RETURNS void
LANGUAGE plpgsql
SECURITY DEFINER
SET search_path = public
AS $$
BEGIN
    IF NOT public._es_admin_maestro_coinvierte() THEN
        RAISE EXCEPTION 'Acción reservada al administrador maestro';
    END IF;

    DELETE FROM public.comentarios_acuerdo
    WHERE acuerdo_id = p_acuerdo_id;

    DELETE FROM public.historial_acuerdo
    WHERE acuerdo_id = p_acuerdo_id;

    DELETE FROM public.notificaciones_acuerdos
    WHERE acuerdo_id = p_acuerdo_id
      AND origen = 'Junta de Gobierno';

    UPDATE public.acuerdos_junta
    SET
        areas = '[]'::jsonb,
        estatus = 'Por iniciar',
        resultado = 'Pendiente',
        fecha_compromiso = NULL,
        fecha_cierre = NULL,
        cumplimiento = NULL,
        responsables_notificacion = '[]'::jsonb,
        responsable_usuario_id = NULL,
        responsable_nombre = NULL,
        responsable_email = NULL,
        notificar_email = false,
        updated_at = now()
    WHERE id = p_acuerdo_id;
END;
$$;

REVOKE ALL ON FUNCTION public.reset_acuerdo_junta_master(uuid) FROM PUBLIC;
GRANT EXECUTE ON FUNCTION public.reset_acuerdo_junta_master(uuid) TO authenticated;


CREATE OR REPLACE FUNCTION public.reset_acuerdo_comite_master(p_acuerdo_id uuid)
RETURNS void
LANGUAGE plpgsql
SECURITY DEFINER
SET search_path = public
AS $$
BEGIN
    IF NOT public._es_admin_maestro_coinvierte() THEN
        RAISE EXCEPTION 'Acción reservada al administrador maestro';
    END IF;

    DELETE FROM public.notificaciones_acuerdos
    WHERE acuerdo_id = p_acuerdo_id
      AND origen = 'Comité';

    UPDATE public.acuerdos_comite
    SET
        areas = '[]'::jsonb,
        estatus = 'Por iniciar',
        resultado = 'Pendiente',
        fecha_compromiso = NULL,
        comentario_seguimiento = NULL,
        responsables_notificacion = '[]'::jsonb,
        responsable_usuario_id = NULL,
        responsable_nombre = NULL,
        responsable_email = NULL,
        notificar_email = false
    WHERE id = p_acuerdo_id;
END;
$$;

REVOKE ALL ON FUNCTION public.reset_acuerdo_comite_master(uuid) FROM PUBLIC;
GRANT EXECUTE ON FUNCTION public.reset_acuerdo_comite_master(uuid) TO authenticated;

-- Verificación opcional
SELECT
    table_name,
    column_name,
    data_type,
    udt_name,
    is_nullable
FROM information_schema.columns
WHERE table_schema = 'public'
  AND table_name IN ('acuerdos_junta', 'acuerdos_comite')
  AND column_name IN ('areas', 'responsables_notificacion');
