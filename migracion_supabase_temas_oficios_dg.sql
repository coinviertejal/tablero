-- COINVIERTE · Clasificación temática de Oficios Dirección General
-- Ejecutar una sola vez en Supabase SQL Editor.

alter table public.oficios_direccion_general
    add column if not exists tema text,
    add column if not exists subtema text,
    add column if not exists clasificacion_confianza numeric(4,3),
    add column if not exists clasificacion_fuente text,
    add column if not exists clasificacion_manual boolean not null default false,
    add column if not exists clasificado_at timestamptz;

create index if not exists idx_oficios_dg_tema
    on public.oficios_direccion_general (tema);

create index if not exists idx_oficios_dg_anio_tema
    on public.oficios_direccion_general (anio, tema);

comment on column public.oficios_direccion_general.tema
    is 'Categoría temática principal del oficio.';
comment on column public.oficios_direccion_general.subtema
    is 'Subcategoría temática del oficio.';
comment on column public.oficios_direccion_general.clasificacion_confianza
    is 'Confianza heurística de la clasificación automática, entre 0 y 1.';
comment on column public.oficios_direccion_general.clasificacion_fuente
    is 'Método utilizado para clasificar el oficio; por ejemplo reglas_v1 o manual.';
comment on column public.oficios_direccion_general.clasificacion_manual
    is 'TRUE cuando una persona sustituyó manualmente la clasificación automática.';
comment on column public.oficios_direccion_general.clasificado_at
    is 'Fecha de la última clasificación temática.';
