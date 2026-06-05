-- =====================================================================
-- saldos_novedades · acumulación histórica de novedades del personal
-- Proyecto Supabase: wjfglsafgaltusmbnccl
-- (ya creada vía Management API; queda acá como documentación)
-- =====================================================================
-- Guarda eventos del personal GENERAL (del Sheet INFORME) para reportes:
-- faltas, llegadas tarde, etc. Los de VENTAS salen directo de saldos_guardias.
-- dedup_key evita duplicar cuando el snapshot corre varias veces.

create table if not exists public.saldos_novedades (
  id         bigint generated always as identity primary key,
  fecha      date,
  area       text not null default 'general',  -- general | ventas
  tipo       text not null,                    -- falta | llegada_tarde | otro
  persona    text not null,
  detalle    text,
  origen     text not null default 'sheet',     -- sheet | guardias
  dedup_key  text unique,
  created_at timestamptz not null default now()
);

alter table public.saldos_novedades enable row level security;
drop policy if exists novedades_anon_read  on public.saldos_novedades;
drop policy if exists novedades_anon_write on public.saldos_novedades;
create policy novedades_anon_read  on public.saldos_novedades for select using (true);
create policy novedades_anon_write on public.saldos_novedades for all using (true) with check (true);
