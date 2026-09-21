// Edge Function: saldos-proxy
// Proxy PostgREST autenticado por firma tga_session para las tablas del panel de
// Saldos (saldos.titogonzalez.online), que hoy el navegador toca DIRECTO con la
// anon key pública (cualquiera con esa key podía leer/escribir/borrar pagos,
// tesorería, ventas y novedades). El cliente manda { session, path, method, body,
// prefer } con su sesión firmada del SSO; el proxy verifica el HMAC y reenvía a
// PostgREST con service_role SOLO para las tablas de la whitelist. Así se cierra
// el acceso anónimo sin reescribir cada llamada (solo el helper sbFetch de la app).
// Gemelo de consulta-0km/db-proxy, con otra whitelist. Mismo secreto de firma.

const SUPABASE_URL = Deno.env.get("SUPABASE_URL")!;
const SERVICE_KEY = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;

// Tablas que el proxy acepta. Cualquier otra ruta se rechaza.
const TABLAS_OK = new Set([
  "tesoreria_items",
  "adm_ventas",
  "saldos_fondos_diario",
  "saldos_novedades",
  "saldos_guardias",
  "saldos_guardias_base", // grilla del Excel de guardias (tambien la lee el CRM)
  "compras_vw",
]);

// QUIÉN puede usar cada tabla. La firma sola prueba que la sesión es de alguien
// del portal, no que ese alguien tenga Saldos: sin esto, cualquier usuario del
// SSO podía leer y escribir tesorería y pagos llamando al proxy directo, aunque
// la página no le mostrara la solapa. Espeja las listas de index.html
// (ADMIN_USERS, TAB_ACCESS) y se evalúa sobre el usuario FIRMANTE: si un dueño
// impersona a otro, manda el dueño. Al sumar a alguien a la página, sumarlo acá.
const ADMINS = new Set(["fngonzalez", "fgonzalez", "cgonzalez", "vreyna"]);
// Usuarios restringidos → tablas que tocan. marianom carga guardias y novedades
// (solapa Personal): lee y escribe esas dos, y lee la grilla base.
const ACCESO_RESTRINGIDO: Record<string, Set<string>> = {
  marianom: new Set(["saldos_novedades", "saldos_guardias", "saldos_guardias_base"]),
};
// Tablas que la web solo lee: se cargan desde afuera (el Excel de guardias).
const SOLO_LECTURA = new Set(["saldos_guardias_base"]);

function puede(usuario: string, tabla: string, method: string): boolean {
  if (SOLO_LECTURA.has(tabla) && method !== "GET") return false;
  if (ADMINS.has(usuario)) return true;
  return !!ACCESO_RESTRINGIDO[usuario]?.has(tabla);
}

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
  "Access-Control-Allow-Headers": "authorization, apikey, content-type, x-client-info",
};

function json(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { ...CORS_HEADERS, "Content-Type": "application/json" },
  });
}

let _secret: string | null = null;
async function getSecret(): Promise<string> {
  if (_secret) return _secret;
  const r = await fetch(`${SUPABASE_URL}/rest/v1/app_config?clave=eq.tga_session_secret&select=valor`, {
    headers: { apikey: SERVICE_KEY, Authorization: `Bearer ${SERVICE_KEY}` },
  });
  const rows = await r.json().catch(() => []);
  _secret = (Array.isArray(rows) && rows[0]?.valor) || "";
  return _secret;
}
function toHex(buf: ArrayBuffer): string {
  return [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, "0")).join("");
}
function timingSafeEqual(a: string, b: string): boolean {
  if (a.length !== b.length) return false;
  let d = 0;
  for (let i = 0; i < a.length; i++) d |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return d === 0;
}
async function sesionValida(sess: any): Promise<boolean> {
  if (!sess || typeof sess !== "object") return false;
  const usuario = String(sess.usuario || "").trim().toLowerCase();
  const exp = Number(sess.session_exp);
  const sig = String(sess.session_sig || "");
  if (!usuario || !exp || !sig) return false;
  if (exp < Math.floor(Date.now() / 1000)) return false;
  const secret = await getSecret();
  if (!secret) return false;
  const key = await crypto.subtle.importKey(
    "raw",
    new TextEncoder().encode(secret),
    { name: "HMAC", hash: "SHA-256" },
    false,
    ["sign"],
  );
  const mac = await crypto.subtle.sign("HMAC", key, new TextEncoder().encode(`${usuario}.${exp}`));
  return timingSafeEqual(toHex(mac), sig);
}

// Extrae el nombre de tabla del path PostgREST ("tesoreria_items?foo=bar" -> "tesoreria_items").
function tablaDe(path: string): string {
  const limpio = String(path || "").replace(/^\/+/, "").replace(/^rest\/v1\//, "");
  const m = limpio.match(/^([a-zA-Z0-9_]+)/);
  return m ? m[1] : "";
}

Deno.serve(async (req: Request) => {
  if (req.method === "OPTIONS") return new Response(null, { headers: CORS_HEADERS });
  if (req.method !== "POST") return json({ error: "Método no permitido" }, 405);

  let body: any;
  try {
    body = await req.json();
  } catch {
    return json({ error: "JSON inválido" }, 400);
  }

  if (!(await sesionValida(body?.session))) return json({ error: "No autorizado" }, 401);

  const path = String(body?.path || "");
  const tabla = tablaDe(path);
  if (!TABLAS_OK.has(tabla)) return json({ error: "Tabla no permitida" }, 403);

  const method = String(body?.method || "GET").toUpperCase();
  if (!["GET", "POST", "PATCH", "DELETE"].includes(method)) return json({ error: "Método inválido" }, 400);

  const usuario = String(body.session.usuario).trim().toLowerCase();
  if (!puede(usuario, tabla, method)) return json({ error: "Sin permiso sobre esta tabla" }, 403);

  // Reenvío a PostgREST con service_role.
  const headers: Record<string, string> = {
    apikey: SERVICE_KEY,
    Authorization: `Bearer ${SERVICE_KEY}`,
  };
  if (body?.prefer) headers["Prefer"] = String(body.prefer);
  const init: RequestInit = { method, headers };
  if (method === "POST" || method === "PATCH") {
    headers["Content-Type"] = "application/json";
    init.body = typeof body.body === "string" ? body.body : JSON.stringify(body.body ?? {});
  }

  try {
    const clean = path.replace(/^\/+/, "").replace(/^rest\/v1\//, "");
    const resp = await fetch(`${SUPABASE_URL}/rest/v1/${clean}`, init);
    const text = await resp.text();
    // 204/205/304 no pueden llevar cuerpo: pasar null o el constructor de Response
    // tira "Response with null body status cannot have body" (los DELETE y los
    // PATCH/POST con return=minimal devuelven 204). La escritura ya se aplicó.
    const sinBody = resp.status === 204 || resp.status === 205 || resp.status === 304;
    return new Response(sinBody ? null : text, {
      status: resp.status,
      headers: {
        ...CORS_HEADERS,
        "Content-Type": resp.headers.get("Content-Type") || "application/json",
        ...(resp.headers.get("Content-Range") ? { "Content-Range": resp.headers.get("Content-Range")! } : {}),
      },
    });
  } catch (e) {
    return json({ error: String(e instanceof Error ? e.message : e) }, 502);
  }
});
