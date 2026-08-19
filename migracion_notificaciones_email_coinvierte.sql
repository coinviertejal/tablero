
-- ============================================================
-- COINVIERTE · Recordatorios de acuerdos por correo
-- Junta de Gobierno + Comités
-- ============================================================

-- 1) Campos de responsable individual y preferencia de correo
ALTER TABLE public.acuerdos_junta
    ADD COLUMN IF NOT EXISTS responsable_usuario_id uuid,
    ADD COLUMN IF NOT EXISTS responsable_nombre text,
    ADD COLUMN IF NOT EXISTS responsable_email text,
    ADD COLUMN IF NOT EXISTS notificar_email boolean NOT NULL DEFAULT true;

ALTER TABLE public.acuerdos_comite
    ADD COLUMN IF NOT EXISTS responsable_usuario_id uuid,
    ADD COLUMN IF NOT EXISTS responsable_nombre text,
    ADD COLUMN IF NOT EXISTS responsable_email text,
    ADD COLUMN IF NOT EXISTS notificar_email boolean NOT NULL DEFAULT true;

-- FK opcionales: si el usuario se elimina, conservamos nombre/correo histórico.
DO $$
BEGIN
    IF NOT EXISTS (
        SELECT 1 FROM pg_constraint
        WHERE conname = 'acuerdos_junta_responsable_usuario_fk'
    ) THEN
        ALTER TABLE public.acuerdos_junta
        ADD CONSTRAINT acuerdos_junta_responsable_usuario_fk
        FOREIGN KEY (responsable_usuario_id)
        REFERENCES public.usuarios_autorizados(id)
        ON DELETE SET NULL;
    END IF;

    IF NOT EXISTS (
        SELECT 1 FROM pg_constraint
        WHERE conname = 'acuerdos_comite_responsable_usuario_fk'
    ) THEN
        ALTER TABLE public.acuerdos_comite
        ADD CONSTRAINT acuerdos_comite_responsable_usuario_fk
        FOREIGN KEY (responsable_usuario_id)
        REFERENCES public.usuarios_autorizados(id)
        ON DELETE SET NULL;
    END IF;
END $$;

-- 2) Cola / bitácora de notificaciones
CREATE TABLE IF NOT EXISTS public.notificaciones_acuerdos (
    id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    origen text NOT NULL CHECK (origen IN ('Junta de Gobierno', 'Comité')),
    acuerdo_id uuid NOT NULL,
    numero_acuerdo text,
    titulo text,
    responsable_usuario_id uuid,
    responsable_nombre text,
    destinatario_email text NOT NULL,
    fecha_compromiso date NOT NULL,
    dias_anticipacion integer NOT NULL DEFAULT 3,
    asunto text NOT NULL,
    cuerpo text NOT NULL,
    estado text NOT NULL DEFAULT 'Pendiente'
        CHECK (estado IN ('Pendiente', 'Enviada', 'Error', 'Cancelada')),
    intentos integer NOT NULL DEFAULT 0,
    ultimo_error text,
    preparada_at timestamptz NOT NULL DEFAULT now(),
    enviada_at timestamptz,
    created_at timestamptz NOT NULL DEFAULT now(),
    updated_at timestamptz NOT NULL DEFAULT now()
);

CREATE UNIQUE INDEX IF NOT EXISTS uq_notificacion_acuerdo_3dias
ON public.notificaciones_acuerdos(origen, acuerdo_id, fecha_compromiso, dias_anticipacion);

CREATE INDEX IF NOT EXISTS idx_notificaciones_estado
ON public.notificaciones_acuerdos(estado, fecha_compromiso);

ALTER TABLE public.notificaciones_acuerdos ENABLE ROW LEVEL SECURITY;

-- Usuarios autenticados pueden consultar la bitácora desde la app si se agrega una vista posteriormente.
DROP POLICY IF EXISTS notificaciones_select_authenticated
ON public.notificaciones_acuerdos;

CREATE POLICY notificaciones_select_authenticated
ON public.notificaciones_acuerdos
FOR SELECT TO authenticated
USING (true);

-- 3) Función que prepara recordatorios cuando faltan exactamente 3 días.
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
        fecha_compromiso, dias_anticipacion, asunto, cuerpo
    )
    SELECT
        'Junta de Gobierno',
        a.id,
        a.numero,
        COALESCE(a.titulo, a.texto, 'Acuerdo'),
        a.responsable_usuario_id,
        a.responsable_nombre,
        a.responsable_email,
        a.fecha_compromiso::date,
        3,
        'COINVIERTE · Acuerdo próximo a vencer',
        'Hola ' || COALESCE(NULLIF(a.responsable_nombre,''), 'responsable') || E',\n\n' ||
        'Tienes un acuerdo de Junta de Gobierno próximo a vencer.' || E'\n\n' ||
        'Acuerdo: ' || COALESCE(a.numero, 'Sin número') || E'\n' ||
        'Tarea / asunto: ' || COALESCE(NULLIF(a.titulo,''), NULLIF(a.texto,''), 'Sin descripción') || E'\n' ||
        'Fecha límite: ' || to_char(a.fecha_compromiso::date, 'DD/MM/YYYY') || E'\n' ||
        'Faltan: 3 días' || E'\n\n' ||
        'Consulta el seguimiento en la plataforma institucional COINVIERTE.'
    FROM public.acuerdos_junta a
    WHERE a.notificar_email IS TRUE
      AND a.responsable_email IS NOT NULL
      AND btrim(a.responsable_email) <> ''
      AND a.fecha_compromiso IS NOT NULL
      AND a.fecha_compromiso::date = current_date + 3
      AND COALESCE(a.estatus, '') <> 'Terminada'
    ON CONFLICT (origen, acuerdo_id, fecha_compromiso, dias_anticipacion)
    DO NOTHING;

    GET DIAGNOSTICS v_junta = ROW_COUNT;

    INSERT INTO public.notificaciones_acuerdos (
        origen, acuerdo_id, numero_acuerdo, titulo,
        responsable_usuario_id, responsable_nombre, destinatario_email,
        fecha_compromiso, dias_anticipacion, asunto, cuerpo
    )
    SELECT
        'Comité',
        a.id,
        a.numero,
        COALESCE(a.titulo, a.texto, 'Acuerdo'),
        a.responsable_usuario_id,
        a.responsable_nombre,
        a.responsable_email,
        a.fecha_compromiso::date,
        3,
        'COINVIERTE · Acuerdo próximo a vencer',
        'Hola ' || COALESCE(NULLIF(a.responsable_nombre,''), 'responsable') || E',\n\n' ||
        'Tienes un acuerdo de Comité próximo a vencer.' || E'\n\n' ||
        'Acuerdo: ' || COALESCE(a.numero, 'Sin número') || E'\n' ||
        'Tarea / asunto: ' || COALESCE(NULLIF(a.titulo,''), NULLIF(a.texto,''), 'Sin descripción') || E'\n' ||
        'Fecha límite: ' || to_char(a.fecha_compromiso::date, 'DD/MM/YYYY') || E'\n' ||
        'Faltan: 3 días' || E'\n\n' ||
        'Consulta el seguimiento en la plataforma institucional COINVIERTE.'
    FROM public.acuerdos_comite a
    WHERE a.notificar_email IS TRUE
      AND a.responsable_email IS NOT NULL
      AND btrim(a.responsable_email) <> ''
      AND a.fecha_compromiso IS NOT NULL
      AND a.fecha_compromiso::date = current_date + 3
      AND COALESCE(a.estatus, '') <> 'Terminada'
    ON CONFLICT (origen, acuerdo_id, fecha_compromiso, dias_anticipacion)
    DO NOTHING;

    GET DIAGNOSTICS v_comite = ROW_COUNT;

    RETURN jsonb_build_object(
        'junta_preparados', v_junta,
        'comite_preparados', v_comite,
        'total_preparados', v_junta + v_comite
    );
END;
$$;

-- Service role podrá ejecutar las funciones desde Apps Script.
REVOKE ALL ON FUNCTION public.preparar_recordatorios_acuerdos() FROM PUBLIC;
GRANT EXECUTE ON FUNCTION public.preparar_recordatorios_acuerdos() TO service_role;

-- 4) Función para marcar una notificación como enviada
CREATE OR REPLACE FUNCTION public.marcar_recordatorio_enviado(p_id uuid)
RETURNS void
LANGUAGE plpgsql
SECURITY DEFINER
SET search_path = public
AS $$
BEGIN
    UPDATE public.notificaciones_acuerdos
    SET estado = 'Enviada',
        enviada_at = now(),
        intentos = intentos + 1,
        ultimo_error = NULL,
        updated_at = now()
    WHERE id = p_id;
END;
$$;

REVOKE ALL ON FUNCTION public.marcar_recordatorio_enviado(uuid) FROM PUBLIC;
GRANT EXECUTE ON FUNCTION public.marcar_recordatorio_enviado(uuid) TO service_role;

-- 5) Función para registrar error de envío
CREATE OR REPLACE FUNCTION public.marcar_recordatorio_error(
    p_id uuid,
    p_error text
)
RETURNS void
LANGUAGE plpgsql
SECURITY DEFINER
SET search_path = public
AS $$
BEGIN
    UPDATE public.notificaciones_acuerdos
    SET estado = 'Error',
        intentos = intentos + 1,
        ultimo_error = left(COALESCE(p_error, 'Error desconocido'), 1000),
        updated_at = now()
    WHERE id = p_id;
END;
$$;

REVOKE ALL ON FUNCTION public.marcar_recordatorio_error(uuid, text) FROM PUBLIC;
GRANT EXECUTE ON FUNCTION public.marcar_recordatorio_error(uuid, text) TO service_role;

-- Verificación rápida
SELECT
    'acuerdos_junta' AS tabla,
    count(*) FILTER (WHERE notificar_email) AS avisos_activos
FROM public.acuerdos_junta
UNION ALL
SELECT
    'acuerdos_comite',
    count(*) FILTER (WHERE notificar_email)
FROM public.acuerdos_comite;
