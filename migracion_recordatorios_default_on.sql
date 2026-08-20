
-- COINVIERTE · Recordatorio por correo activado por defecto
-- Aplica a Junta de Gobierno y Comités.

ALTER TABLE public.acuerdos_junta
ALTER COLUMN notificar_email SET DEFAULT true;

ALTER TABLE public.acuerdos_comite
ALTER COLUMN notificar_email SET DEFAULT true;

-- Los registros con NULL se normalizan a true.
UPDATE public.acuerdos_junta
SET notificar_email = true
WHERE notificar_email IS NULL;

UPDATE public.acuerdos_comite
SET notificar_email = true
WHERE notificar_email IS NULL;

-- Ajusta los reinicios maestros para volver a dejar el recordatorio activado.
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
        notificar_email = true,
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
        areas = ARRAY[]::text[],
        estatus = 'Por iniciar',
        resultado = 'Pendiente',
        fecha_compromiso = NULL,
        comentario_seguimiento = NULL,
        responsables_notificacion = '[]'::jsonb,
        responsable_usuario_id = NULL,
        responsable_nombre = NULL,
        responsable_email = NULL,
        notificar_email = true
    WHERE id = p_acuerdo_id;
END;
$$;

REVOKE ALL ON FUNCTION public.reset_acuerdo_comite_master(uuid) FROM PUBLIC;
GRANT EXECUTE ON FUNCTION public.reset_acuerdo_comite_master(uuid) TO authenticated;
