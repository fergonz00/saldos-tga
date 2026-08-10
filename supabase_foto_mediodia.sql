-- =====================================================================
-- saldos-tga - LA FOTO DEL INFORME AL MEDIODIA (pedido de Fer, 10-ago-2026)
-- =====================================================================
-- El INFORME (Sheet) es una foto del dia, no un log, y la planilla se termina
-- de ordenar durante la maniana: la ausencia de ayer queda pegada aunque la
-- persona ya haya fichado. Hasta ahora la unica foto que existia era la que
-- sacaba el navegador al abrir la app, y se abre casi siempre entre las 9 y
-- las 11 -> quedaban registradas faltas que no eran (caso Antonella, 7-ago).
--
-- Esto saca la foto SOLA todos los dias habiles a las 12:30 de Argentina, sin
-- depender de que nadie abra nada, y reconcilia: lo que se habia guardado hoy
-- mas temprano y ya no figura en el INFORME se borra.
--
-- Espeja exactamente lo que hace snapshotNovedadesGenerales() en index.html
-- (misma dedup_key, misma fecha, mismo detalle) para que las dos fuentes no
-- se pisen ni dupliquen.
-- =====================================================================

create table if not exists saldos_foto_req (
  id           bigserial primary key,
  request_id   bigint not null,
  pedido_at    timestamptz not null default now(),
  procesado_at timestamptz,
  resultado    jsonb
);
alter table saldos_foto_req enable row level security;
revoke all on table saldos_foto_req from anon, authenticated;

-- Paso 1 (12:30): le pide el INFORME al Apps Script. pg_net es asincrono:
-- devuelve un id y la respuesta llega despues a net._http_response.
create or replace function saldos_foto_pedir()
returns bigint
language plpgsql
security definer
set search_path = public, net
as $fn$
declare rid bigint;
begin
  select net.http_get(
    url := 'https://script.google.com/macros/s/AKfycbyRTqqpQMjKDL82Z5Cjd9IJWPQnINF0LAEvji8FizfXMBO8Cz0IVbTSnQnNmH_rRxz9yg/exec?token=tga-saldos-K9Mx2P7vQ&tipo=informe',
    timeout_milliseconds := 25000
  ) into rid;
  insert into saldos_foto_req(request_id) values (rid);
  return rid;
end;
$fn$;

-- Paso 2 (12:33): parsea la respuesta y la guarda.
create or replace function saldos_foto_procesar()
returns jsonb
language plpgsql
security definer
set search_path = public, net
as $fn$
declare
  req      record;
  resp     record;
  hoy      date := (now() at time zone 'America/Argentina/Buenos_Aires')::date;
  personal jsonb;
  it       jsonb;
  seccion  text := '';
  ev_tipo  text;
  txt      text;
  fstr     text;
  mes3     text;
  f        date;
  k        text;
  claves   text[] := array[]::text[];
  nuevas   int := 0;
  borradas int := 0;
  res      jsonb;
begin
  select * into req from saldos_foto_req where procesado_at is null order by id desc limit 1;
  if not found then return jsonb_build_object('ok', false, 'motivo', 'sin pedido pendiente'); end if;

  select * into resp from net._http_response where id = req.request_id;
  if not found then return jsonb_build_object('ok', false, 'motivo', 'la respuesta todavia no llego'); end if;

  if resp.status_code is distinct from 200 then
    res := jsonb_build_object('ok', false, 'status', resp.status_code, 'error', resp.error_msg);
    update saldos_foto_req set procesado_at = now(), resultado = res where id = req.id;
    return res;
  end if;

  personal := (resp.content::jsonb) -> 'personal';
  if personal is null or jsonb_typeof(personal) <> 'array' then
    res := jsonb_build_object('ok', false, 'motivo', 'la respuesta no trae bloque personal');
    update saldos_foto_req set procesado_at = now(), resultado = res where id = req.id;
    return res;
  end if;

  for it in select * from jsonb_array_elements(personal) loop
    if it->>'tipo' = 'header' then
      seccion := '';
    elsif it->>'tipo' = 'subhead' then
      seccion := coalesce(it->>'texto', '');
    elsif it->>'tipo' = 'item' then
      txt := btrim(coalesce(it->>'texto', ''));
      continue when txt = '';

      -- misma clasificacion que clasificarTipoNovedad() en index.html
      ev_tipo := case
        when seccion ~* 'tarde'        then 'llegada_tarde'
        when seccion ~* 'ausent|falt'  then 'falta'
        else 'otro' end;

      -- misma fecha que parseFechaCorta(): en Ausentes la planilla escribe la
      -- fecha de REINCORPORACION; si no hay nada, es de hoy.
      fstr := btrim(coalesce(it->>'fecha', ''));
      f := null;
      if fstr ~ '^[0-9]{4}-[0-9]{1,2}-[0-9]{1,2}$' then
        f := to_date(fstr, 'YYYY-MM-DD');
      elsif fstr ~ '^[0-9]{1,2}/[0-9]{1,2}/[0-9]{2,4}$' then
        f := to_date(fstr, case when length(split_part(fstr, '/', 3)) = 2 then 'DD/MM/YY' else 'DD/MM/YYYY' end);
      elsif lower(fstr) ~ '^[0-9]{1,2}-[a-z]{3}' then
        mes3 := lower(substring(fstr from '^[0-9]{1,2}-([a-z]{3})'));
        if mes3 = any (array['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic']) then
          f := make_date(
                 extract(year from hoy)::int,
                 array_position(array['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'], mes3),
                 substring(fstr from '^([0-9]{1,2})-')::int);
        end if;
      end if;
      f := coalesce(f, hoy);

      k := f::text || '|general|' || ev_tipo || '|' || lower(txt);
      claves := claves || k;

      insert into saldos_novedades (fecha, area, tipo, persona, detalle, origen, dedup_key)
      values (f, 'general', ev_tipo, txt, nullif(fstr, ''), 'sheet', k)
      on conflict (dedup_key) do nothing;
      if found then nuevas := nuevas + 1; end if;
    end if;
  end loop;

  -- Reconciliacion: lo que se snapshoteo HOY mas temprano (navegador o este
  -- mismo cron) y ya no esta en la planilla era arrastre. Solo toca filas del
  -- Sheet creadas hoy: ni el historico viejo (incluidas las ausencias con
  -- fecha de reincorporacion futura) ni lo cargado a mano se tocan.
  if array_length(claves, 1) > 0 then
    delete from saldos_novedades
     where origen = 'sheet'
       and created_at >= (hoy + time '00:00') at time zone 'America/Argentina/Buenos_Aires'
       and not (dedup_key = any (claves));
    get diagnostics borradas = row_count;
  end if;

  res := jsonb_build_object('ok', true, 'fecha', hoy, 'leidas', array_length(claves, 1),
                            'nuevas', nuevas, 'borradas', borradas);
  update saldos_foto_req set procesado_at = now(), resultado = res where id = req.id;
  return res;
end;
$fn$;

-- Que nadie las pueda disparar desde la API publica: solo el cron (postgres).
revoke execute on function saldos_foto_pedir()    from public, anon, authenticated;
revoke execute on function saldos_foto_procesar() from public, anon, authenticated;

-- 15:30 / 15:33 UTC = 12:30 / 12:33 de Argentina, lunes a viernes.
select cron.schedule('saldos-foto-mediodia-pedir',    '30 15 * * 1-5', 'select public.saldos_foto_pedir()');
select cron.schedule('saldos-foto-mediodia-procesar', '33 15 * * 1-5', 'select public.saldos_foto_procesar()');
