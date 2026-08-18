-- COINVIERTE · Corrección RLS para ingesta ENVIADOS DG
-- Permite crear registros desde Excel usando registrado_por,
-- conservando el esquema actual para cargas manuales con subido_por.

alter table public.oficios_direccion_general enable row level security;

drop policy if exists "oficios_dg_insert_auth" on public.oficios_direccion_general;
create policy "oficios_dg_insert_auth"
on public.oficios_direccion_general
for insert
to authenticated
with check (
    auth.uid() = subido_por
    or auth.uid() = registrado_por
);

-- Aseguramos que un UPSERT pueda actualizar registros existentes.
drop policy if exists "oficios_dg_update_auth" on public.oficios_direccion_general;
create policy "oficios_dg_update_auth"
on public.oficios_direccion_general
for update
to authenticated
using (true)
with check (
    subido_por is null
    or auth.uid() = subido_por
    or registrado_por is null
    or auth.uid() = registrado_por
);

-- Consulta disponible para usuarios autenticados del aplicativo.
drop policy if exists "oficios_dg_select_auth" on public.oficios_direccion_general;
create policy "oficios_dg_select_auth"
on public.oficios_direccion_general
for select
to authenticated
using (true);
