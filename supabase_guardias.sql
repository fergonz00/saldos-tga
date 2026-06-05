-- =====================================================================
-- saldos_guardias  ·  overrides del cronograma de Guardias Ventas
-- Proyecto Supabase: wjfglsafgaltusmbnccl (el mismo de saldos/portal)
-- Correr una vez en: Supabase → SQL Editor.
-- =====================================================================
-- La grilla BASE (jun–jul 2026, sacada del Excel) vive hardcodeada en
-- index.html. Esta tabla guarda SOLO las modificaciones (overrides):
-- una fila por (fecha, sucursal, vendedor) que se editó desde la web.
-- Si no hay fila, vale lo que dice la base.

create table if not exists public.saldos_guardias (
  fecha       date        not null,
  sucursal    text        not null,
  vendedor    text        not null,
  estado      text        not null,   -- trabaja | franco | vacaciones | curso | falta
  updated_by  text,
  updated_at  timestamptz not null default now(),
  primary key (fecha, sucursal, vendedor)
);

-- Escritura/lectura con la anon key (igual que saldos_fondos_diario).
alter table public.saldos_guardias enable row level security;

drop policy if exists "guardias anon read"  on public.saldos_guardias;
drop policy if exists "guardias anon write" on public.saldos_guardias;

create policy "guardias anon read"
  on public.saldos_guardias for select
  using (true);

create policy "guardias anon write"
  on public.saldos_guardias for all
  using (true) with check (true);

-- Mantener updated_at fresco en cada upsert
create or replace function public.touch_saldos_guardias()
returns trigger language plpgsql as $$
begin
  new.updated_at = now();
  return new;
end $$;

drop trigger if exists trg_touch_saldos_guardias on public.saldos_guardias;
create trigger trg_touch_saldos_guardias
  before update on public.saldos_guardias
  for each row execute function public.touch_saldos_guardias();
