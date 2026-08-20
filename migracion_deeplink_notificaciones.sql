
-- ================================================================
-- COINVIERTE · Rutas de navegación y vínculos directos en avisos
-- ================================================================

ALTER TABLE public.notificaciones_acuerdos
    ADD COLUMN IF NOT EXISTS sesion_id uuid,
    ADD COLUMN IF NOT EXISTS anio integer,
    ADD COLUMN IF NOT EXISTS comite_nombre text,
    ADD COLUMN IF NOT EXISTS sesion_nombre text,
    ADD COLUMN IF NOT EXISTS sesion_tipo text,
    ADD COLUMN IF NOT EXISTS ruta_navegacion text;

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
        origen, acuerdo_id, numero_acuerdo, titulo,
        responsable_usuario_id, responsable_nombre, destinatario_email,
        fecha_compromiso, dias_anticipacion, asunto, cuerpo,
        sesion_id, anio, sesion_nombre, sesion_tipo, ruta_navegacion
    )
    SELECT
        'Junta de Gobierno',
        a.id,
        a.numero,
        COALESCE(NULLIF(a.titulo,''), NULLIF(a.texto,''), 'Acuerdo'),
        CASE WHEN nullif(r->>'id','') IS NULL THEN NULL ELSE (r->>'id')::uuid END,
        COALESCE(NULLIF(r->>'nombre',''), r->>'email'),
        r->>'email',
        a.fecha_compromiso::date,
        3,
        'COINVIERTE · Acuerdo próximo a vencer',
        'Hola ' || COALESCE(NULLIF(r->>'nombre',''), 'responsable') || E',\n\n' ||
        'Tienes un acuerdo de Junta de Gobierno próximo a vencer.' || E'\n\n' ||
        'Acuerdo: ' || COALESCE(a.numero, 'Sin número') || E'\n' ||
        'Tarea / asunto: ' || COALESCE(NULLIF(a.titulo,''), NULLIF(a.texto,''), 'Sin descripción') || E'\n' ||
        'Fecha límite: ' || to_char(a.fecha_compromiso::date, 'DD/MM/YYYY') || E'\n' ||
        'Faltan: 3 días' || E'\n\n' ||
        'Ubicación: Junta de Gobierno → ' || s.anio::text || ' → ' ||
        COALESCE(s.nombre, s.tipo, 'Sesión') || ' → ' || COALESCE(a.numero, 'Acuerdo') || E'\n\n' ||
        'Consulta el seguimiento en la plataforma institucional COINVIERTE.',
        a.sesion_id,
        s.anio,
        s.nombre,
        s.tipo,
        'Junta de Gobierno → ' || s.anio::text || ' → ' ||
        COALESCE(s.nombre, s.tipo, 'Sesión') || ' → ' || COALESCE(a.numero, 'Acuerdo')
    FROM public.acuerdos_junta a
    JOIN public.sesiones_junta s ON s.id = a.sesion_id
    CROSS JOIN LATERAL jsonb_array_elements(
        CASE
            WHEN jsonb_array_length(a.responsables_notificacion) > 0 THEN a.responsables_notificacion
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
    ) DO UPDATE SET
        sesion_id = EXCLUDED.sesion_id,
        anio = EXCLUDED.anio,
        sesion_nombre = EXCLUDED.sesion_nombre,
        sesion_tipo = EXCLUDED.sesion_tipo,
        ruta_navegacion = EXCLUDED.ruta_navegacion,
        cuerpo = EXCLUDED.cuerpo,
        updated_at = now();

    GET DIAGNOSTICS v_junta = ROW_COUNT;

    INSERT INTO public.notificaciones_acuerdos (
        origen, acuerdo_id, numero_acuerdo, titulo,
        responsable_usuario_id, responsable_nombre, destinatario_email,
        fecha_compromiso, dias_anticipacion, asunto, cuerpo,
        sesion_id, anio, comite_nombre, sesion_nombre, sesion_tipo, ruta_navegacion
    )
    SELECT
        'Comité',
        a.id,
        a.numero,
        COALESCE(NULLIF(a.titulo,''), NULLIF(a.texto,''), 'Acuerdo'),
        CASE WHEN nullif(r->>'id','') IS NULL THEN NULL ELSE (r->>'id')::uuid END,
        COALESCE(NULLIF(r->>'nombre',''), r->>'email'),
        r->>'email',
        a.fecha_compromiso::date,
        3,
        'COINVIERTE · Acuerdo próximo a vencer',
        'Hola ' || COALESCE(NULLIF(r->>'nombre',''), 'responsable') || E',\n\n' ||
        'Tienes un acuerdo de Comité próximo a vencer.' || E'\n\n' ||
        'Acuerdo: ' || COALESCE(a.numero, 'Sin número') || E'\n' ||
        'Tarea / asunto: ' || COALESCE(NULLIF(a.titulo,''), NULLIF(a.texto,''), 'Sin descripción') || E'\n' ||
        'Fecha límite: ' || to_char(a.fecha_compromiso::date, 'DD/MM/YYYY') || E'\n' ||
        'Faltan: 3 días' || E'\n\n' ||
        'Ubicación: Comités → ' || COALESCE(s.comite, 'Comité') || ' → ' || s.anio::text ||
        ' → ' || COALESCE(s.nombre, s.tipo, 'Sesión') || ' → ' || COALESCE(a.numero, 'Acuerdo') || E'\n\n' ||
        'Consulta el seguimiento en la plataforma institucional COINVIERTE.',
        a.sesion_id,
        s.anio,
        s.comite,
        s.nombre,
        s.tipo,
        'Comités → ' || COALESCE(s.comite, 'Comité') || ' → ' || s.anio::text ||
        ' → ' || COALESCE(s.nombre, s.tipo, 'Sesión') || ' → ' || COALESCE(a.numero, 'Acuerdo')
    FROM public.acuerdos_comite a
    JOIN public.sesiones_comite s ON s.id = a.sesion_id
    CROSS JOIN LATERAL jsonb_array_elements(
        CASE
            WHEN jsonb_array_length(a.responsables_notificacion) > 0 THEN a.responsables_notificacion
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
    ) DO UPDATE SET
        sesion_id = EXCLUDED.sesion_id,
        anio = EXCLUDED.anio,
        comite_nombre = EXCLUDED.comite_nombre,
        sesion_nombre = EXCLUDED.sesion_nombre,
        sesion_tipo = EXCLUDED.sesion_tipo,
        ruta_navegacion = EXCLUDED.ruta_navegacion,
        cuerpo = EXCLUDED.cuerpo,
        updated_at = now();

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

-- Enriquecer notificaciones ya existentes con el contexto de su sesión.
UPDATE public.notificaciones_acuerdos n
SET
    sesion_id = a.sesion_id,
    anio = s.anio,
    sesion_nombre = s.nombre,
    sesion_tipo = s.tipo,
    ruta_navegacion = 'Junta de Gobierno → ' || s.anio::text || ' → ' ||
        COALESCE(s.nombre, s.tipo, 'Sesión') || ' → ' || COALESCE(a.numero, 'Acuerdo'),
    updated_at = now()
FROM public.acuerdos_junta a
JOIN public.sesiones_junta s ON s.id = a.sesion_id
WHERE n.origen = 'Junta de Gobierno'
  AND n.acuerdo_id = a.id;

UPDATE public.notificaciones_acuerdos n
SET
    sesion_id = a.sesion_id,
    anio = s.anio,
    comite_nombre = s.comite,
    sesion_nombre = s.nombre,
    sesion_tipo = s.tipo,
    ruta_navegacion = 'Comités → ' || COALESCE(s.comite, 'Comité') || ' → ' || s.anio::text ||
        ' → ' || COALESCE(s.nombre, s.tipo, 'Sesión') || ' → ' || COALESCE(a.numero, 'Acuerdo'),
    updated_at = now()
FROM public.acuerdos_comite a
JOIN public.sesiones_comite s ON s.id = a.sesion_id
WHERE n.origen = 'Comité'
  AND n.acuerdo_id = a.id;
