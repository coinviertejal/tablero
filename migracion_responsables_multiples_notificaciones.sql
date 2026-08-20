
-- ================================================================
-- COINVIERTE · Responsables múltiples + notificaciones individuales
-- Junta de Gobierno y Comités
-- ================================================================

ALTER TABLE public.acuerdos_junta
ADD COLUMN IF NOT EXISTS responsables_notificacion jsonb NOT NULL DEFAULT '[]'::jsonb;

ALTER TABLE public.acuerdos_comite
ADD COLUMN IF NOT EXISTS responsables_notificacion jsonb NOT NULL DEFAULT '[]'::jsonb;

-- Migra responsable único anterior al arreglo nuevo, si existe.
UPDATE public.acuerdos_junta
SET responsables_notificacion = jsonb_build_array(
    jsonb_build_object(
        'id', responsable_usuario_id,
        'nombre', responsable_nombre,
        'email', responsable_email
    )
)
WHERE responsable_email IS NOT NULL
  AND btrim(responsable_email) <> ''
  AND responsables_notificacion = '[]'::jsonb;

UPDATE public.acuerdos_comite
SET responsables_notificacion = jsonb_build_array(
    jsonb_build_object(
        'id', responsable_usuario_id,
        'nombre', responsable_nombre,
        'email', responsable_email
    )
)
WHERE responsable_email IS NOT NULL
  AND btrim(responsable_email) <> ''
  AND responsables_notificacion = '[]'::jsonb;

-- Ahora una misma tarea puede generar un recordatorio por cada destinatario.
DROP INDEX IF EXISTS public.uq_notificacion_acuerdo_3dias;

CREATE UNIQUE INDEX IF NOT EXISTS uq_notificacion_acuerdo_3dias_email
ON public.notificaciones_acuerdos(
    origen,
    acuerdo_id,
    fecha_compromiso,
    dias_anticipacion,
    destinatario_email
);

CREATE OR REPLACE FUNCTION public.preparar_recordatorios_acuerdos()
RETURNS jsonb
LANGUAGE plpgsql
SECURITY DEFINER
SET search_path = public
AS $$
DECLARE
    v_junta integer := 0;
    v_comite integer := 0;
BEGIN
    INSERT INTO public.notificaciones_acuerdos (
        origen,
        acuerdo_id,
        numero_acuerdo,
        titulo,
        responsable_usuario_id,
        responsable_nombre,
        destinatario_email,
        fecha_compromiso,
        dias_anticipacion,
        asunto,
        cuerpo
    )
    SELECT
        'Junta de Gobierno',
        a.id,
        a.numero,
        COALESCE(NULLIF(a.titulo,''), NULLIF(a.texto,''), 'Acuerdo'),
        CASE
            WHEN nullif(r->>'id','') IS NULL THEN NULL
            ELSE (r->>'id')::uuid
        END,
        COALESCE(NULLIF(r->>'nombre',''), r->>'email'),
        r->>'email',
        a.fecha_compromiso::date,
        3,
        'COINVIERTE · Acuerdo próximo a vencer',
        'Hola ' || COALESCE(NULLIF(r->>'nombre',''), 'responsable') || E',\n\n' ||
        'Tienes un acuerdo de Junta de Gobierno próximo a vencer.' || E'\n\n' ||
        'Acuerdo: ' || COALESCE(a.numero, 'Sin número') || E'\n' ||
        'Tarea / asunto: ' ||
            COALESCE(NULLIF(a.titulo,''), NULLIF(a.texto,''), 'Sin descripción') || E'\n' ||
        'Fecha límite: ' || to_char(a.fecha_compromiso::date, 'DD/MM/YYYY') || E'\n' ||
        'Faltan: 3 días' || E'\n\n' ||
        'Consulta el seguimiento en la plataforma institucional COINVIERTE.'
    FROM public.acuerdos_junta a
    CROSS JOIN LATERAL jsonb_array_elements(
        CASE
            WHEN jsonb_array_length(a.responsables_notificacion) > 0
                THEN a.responsables_notificacion
            WHEN a.responsable_email IS NOT NULL AND btrim(a.responsable_email) <> ''
                THEN jsonb_build_array(jsonb_build_object(
                    'id', a.responsable_usuario_id,
                    'nombre', a.responsable_nombre,
                    'email', a.responsable_email
                ))
            ELSE '[]'::jsonb
        END
    ) AS r
    WHERE a.notificar_email IS TRUE
      AND NULLIF(btrim(r->>'email'),'') IS NOT NULL
      AND a.fecha_compromiso IS NOT NULL
      AND a.fecha_compromiso::date = current_date + 3
      AND COALESCE(a.estatus, '') <> 'Terminada'
    ON CONFLICT (
        origen, acuerdo_id, fecha_compromiso, dias_anticipacion, destinatario_email
    ) DO NOTHING;

    GET DIAGNOSTICS v_junta = ROW_COUNT;

    INSERT INTO public.notificaciones_acuerdos (
        origen,
        acuerdo_id,
        numero_acuerdo,
        titulo,
        responsable_usuario_id,
        responsable_nombre,
        destinatario_email,
        fecha_compromiso,
        dias_anticipacion,
        asunto,
        cuerpo
    )
    SELECT
        'Comité',
        a.id,
        a.numero,
        COALESCE(NULLIF(a.titulo,''), NULLIF(a.texto,''), 'Acuerdo'),
        CASE
            WHEN nullif(r->>'id','') IS NULL THEN NULL
            ELSE (r->>'id')::uuid
        END,
        COALESCE(NULLIF(r->>'nombre',''), r->>'email'),
        r->>'email',
        a.fecha_compromiso::date,
        3,
        'COINVIERTE · Acuerdo próximo a vencer',
        'Hola ' || COALESCE(NULLIF(r->>'nombre',''), 'responsable') || E',\n\n' ||
        'Tienes un acuerdo de Comité próximo a vencer.' || E'\n\n' ||
        'Acuerdo: ' || COALESCE(a.numero, 'Sin número') || E'\n' ||
        'Tarea / asunto: ' ||
            COALESCE(NULLIF(a.titulo,''), NULLIF(a.texto,''), 'Sin descripción') || E'\n' ||
        'Fecha límite: ' || to_char(a.fecha_compromiso::date, 'DD/MM/YYYY') || E'\n' ||
        'Faltan: 3 días' || E'\n\n' ||
        'Consulta el seguimiento en la plataforma institucional COINVIERTE.'
    FROM public.acuerdos_comite a
    CROSS JOIN LATERAL jsonb_array_elements(
        CASE
            WHEN jsonb_array_length(a.responsables_notificacion) > 0
                THEN a.responsables_notificacion
            WHEN a.responsable_email IS NOT NULL AND btrim(a.responsable_email) <> ''
                THEN jsonb_build_array(jsonb_build_object(
                    'id', a.responsable_usuario_id,
                    'nombre', a.responsable_nombre,
                    'email', a.responsable_email
                ))
            ELSE '[]'::jsonb
        END
    ) AS r
    WHERE a.notificar_email IS TRUE
      AND NULLIF(btrim(r->>'email'),'') IS NOT NULL
      AND a.fecha_compromiso IS NOT NULL
      AND a.fecha_compromiso::date = current_date + 3
      AND COALESCE(a.estatus, '') <> 'Terminada'
    ON CONFLICT (
        origen, acuerdo_id, fecha_compromiso, dias_anticipacion, destinatario_email
    ) DO NOTHING;

    GET DIAGNOSTICS v_comite = ROW_COUNT;

    RETURN jsonb_build_object(
        'junta_preparados', v_junta,
        'comite_preparados', v_comite,
        'total_preparados', v_junta + v_comite
    );
END;
$$;

REVOKE ALL ON FUNCTION public.preparar_recordatorios_acuerdos() FROM PUBLIC;
GRANT EXECUTE ON FUNCTION public.preparar_recordatorios_acuerdos() TO service_role;

SELECT
    'acuerdos_junta' AS tabla,
    count(*) AS total,
    count(*) FILTER (
        WHERE jsonb_array_length(responsables_notificacion) > 0
    ) AS con_responsables
FROM public.acuerdos_junta
UNION ALL
SELECT
    'acuerdos_comite',
    count(*),
    count(*) FILTER (
        WHERE jsonb_array_length(responsables_notificacion) > 0
    )
FROM public.acuerdos_comite;
