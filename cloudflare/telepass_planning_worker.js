// Cloudflare Worker: planning + Grab & Go shared sync + simple shared login
const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET,POST,OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type,X-Admin-Key,Authorization,X-Session-Token",
  "Content-Type": "application/json; charset=utf-8",
};

const enc = new TextEncoder();
const dec = new TextDecoder();

function json(data, status = 200) {
  return new Response(JSON.stringify(data), { status, headers: CORS_HEADERS });
}

function normPdv(value) {
  const digits = String(value || "").match(/\d+/g);
  if (!digits) return "";
  return digits[0].padStart(5, "0");
}

function bytesToBase64Url(bytes) {
  let out = "";
  for (const b of bytes) out += String.fromCharCode(b);
  return btoa(out).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
}

function base64UrlToBytes(value) {
  let s = String(value || "").replace(/-/g, "+").replace(/_/g, "/");
  while (s.length % 4) s += "=";
  const bin = atob(s);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

async function hmac(secret, text) {
  const key = await crypto.subtle.importKey(
    "raw",
    enc.encode(String(secret || "")),
    { name: "HMAC", hash: "SHA-256" },
    false,
    ["sign"]
  );
  const sig = await crypto.subtle.sign("HMAC", key, enc.encode(String(text || "")));
  return bytesToBase64Url(new Uint8Array(sig));
}

async function signSession(username, env) {
  if (!env.ADMIN_KEY) throw new Error("ADMIN_KEY non configurata");
  const exp = Date.now() + 1000 * 60 * 60 * 24 * 30;
  const payload = bytesToBase64Url(enc.encode(JSON.stringify({ u: String(username || "utente"), exp })));
  const sig = await hmac(env.ADMIN_KEY, payload);
  return { token: payload + "." + sig, expires_at: new Date(exp).toISOString() };
}

function readBearerToken(request) {
  const auth = request.headers.get("Authorization") || "";
  if (auth.toLowerCase().startsWith("bearer ")) return auth.slice(7).trim();
  return request.headers.get("X-Session-Token") || "";
}

async function verifySessionToken(token, env) {
  if (!token || !env.ADMIN_KEY) return false;
  const parts = String(token).split(".");
  if (parts.length !== 2) return false;
  const [payload, sig] = parts;
  const expected = await hmac(env.ADMIN_KEY, payload);
  if (sig !== expected) return false;
  const body = JSON.parse(dec.decode(base64UrlToBytes(payload)));
  if (!body.exp || Date.now() > Number(body.exp)) return false;
  return body;
}

function readAdminKey(request) {
  return request.headers.get("X-Admin-Key") || "";
}

async function isAdmin(request, env) {
  const direct = readAdminKey(request);
  if (direct && env.ADMIN_KEY && direct === env.ADMIN_KEY) return true;
  const session = await verifySessionToken(readBearerToken(request), env).catch(() => false);
  return !!session;
}

async function login(request, env) {
  const configuredUser = String(env.LOGIN_USER || "").trim();
  const configuredPass = String(env.LOGIN_PASS || "");
  if (!configuredUser || !configuredPass) {
    return json({ ok: false, error: "LOGIN_USER o LOGIN_PASS non configurati su Cloudflare" }, 500);
  }
  const body = await request.json().catch(() => ({}));
  const username = String(body.username || "").trim();
  const password = String(body.password || "");
  if (username !== configuredUser || password !== configuredPass) {
    return json({ ok: false, error: "Utente o password non validi" }, 401);
  }
  const session = await signSession(username, env);
  return json({ ok: true, username, ...session });
}

async function sessionStatus(request, env) {
  const session = await verifySessionToken(readBearerToken(request), env).catch(() => false);
  if (!session) return json({ ok: false, logged: false }, 401);
  return json({ ok: true, logged: true, username: session.u, expires_at: new Date(Number(session.exp)).toISOString() });
}

async function ensureGrabVisiteTable(env) {
  await env.DB.prepare(`
    CREATE TABLE IF NOT EXISTS grab_visite (
      pdv TEXT PRIMARY KEY,
      month INTEGER,
      year INTEGER,
      saved_at TEXT,
      updated_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
  `).run();
}

async function listModifiche(env) {
  const rows = await env.DB.prepare(`
    SELECT id, action, pdv, agente, tipo, citta, indirizzo, latitudine, longitudine, note, attivo, created_at, updated_at
    FROM pv_modifiche
    WHERE attivo = 1
    ORDER BY updated_at ASC, id ASC
  `).all();
  return rows.results || [];
}

async function saveModifica(request, env) {
  if (!(await isAdmin(request, env))) {
    return json({ ok: false, error: "Login non valido o ADMIN_KEY non valida" }, 401);
  }

  const body = await request.json().catch(() => ({}));
  const pdv = normPdv(body.pdv);
  const action = String(body.action || "").trim().toUpperCase();

  if (!pdv) return json({ ok: false, error: "PV mancante" }, 400);
  if (!["AGGIUNGI", "ESCLUDI"].includes(action)) {
    return json({ ok: false, error: "Azione non valida. Usa AGGIUNGI o ESCLUDI" }, 400);
  }

  await env.DB.prepare(`UPDATE pv_modifiche SET attivo = 0, updated_at = CURRENT_TIMESTAMP WHERE pdv = ? AND attivo = 1`).bind(pdv).run();

  await env.DB.prepare(`
    INSERT INTO pv_modifiche
    (action, pdv, agente, tipo, citta, indirizzo, latitudine, longitudine, note, attivo, created_at, updated_at)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 1, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP)
  `).bind(
    action,
    pdv,
    String(body.agente || body.agent || ""),
    String(body.tipo || body.type || ""),
    String(body.citta || body.city || ""),
    String(body.indirizzo || body.address || ""),
    body.latitudine === "" || body.latitudine === undefined ? null : Number(body.latitudine),
    body.longitudine === "" || body.longitudine === undefined ? null : Number(body.longitudine),
    String(body.note || "")
  ).run();

  return json({ ok: true, pdv, action });
}

async function ripristinaPdv(request, env) {
  if (!(await isAdmin(request, env))) {
    return json({ ok: false, error: "Login non valido o ADMIN_KEY non valida" }, 401);
  }
  const body = await request.json().catch(() => ({}));
  const pdv = normPdv(body.pdv);
  if (!pdv) return json({ ok: false, error: "PV mancante" }, 400);
  await env.DB.prepare(`UPDATE pv_modifiche SET attivo = 0, updated_at = CURRENT_TIMESTAMP WHERE pdv = ? AND attivo = 1`).bind(pdv).run();
  return json({ ok: true, pdv, action: "RIPRISTINA" });
}

async function listGrabVisite(env) {
  await ensureGrabVisiteTable(env);
  const rows = await env.DB.prepare(`
    SELECT pdv, month, year, saved_at, updated_at
    FROM grab_visite
    ORDER BY updated_at DESC
  `).all();
  return rows.results || [];
}

async function saveGrabVisita(request, env) {
  if (!(await isAdmin(request, env))) {
    return json({ ok: false, error: "Login non valido o ADMIN_KEY non valida" }, 401);
  }
  await ensureGrabVisiteTable(env);
  const body = await request.json().catch(() => ({}));
  const pdv = normPdv(body.pdv);
  const month = Number(body.month || 0);
  const year = Number(body.year || new Date().getFullYear());
  if (!pdv) return json({ ok: false, error: "PV mancante" }, 400);
  if (!month) {
    await env.DB.prepare(`DELETE FROM grab_visite WHERE pdv = ?`).bind(pdv).run();
    return json({ ok: true, pdv, action: "DELETE" });
  }
  if (month < 1 || month > 12) return json({ ok: false, error: "Mese non valido" }, 400);
  await env.DB.prepare(`
    INSERT INTO grab_visite (pdv, month, year, saved_at, updated_at)
    VALUES (?, ?, ?, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP)
    ON CONFLICT(pdv) DO UPDATE SET month=excluded.month, year=excluded.year, saved_at=CURRENT_TIMESTAMP, updated_at=CURRENT_TIMESTAMP
  `).bind(pdv, month, year).run();
  return json({ ok: true, pdv, month, year });
}

export default {
  async fetch(request, env) {
    if (request.method === "OPTIONS") return new Response(null, { headers: CORS_HEADERS });

    const url = new URL(request.url);

    try {
      if (url.pathname === "/health") {
        return json({ ok: true, service: "telepass-planning-api" });
      }

      if (url.pathname === "/login" && request.method === "POST") {
        return login(request, env);
      }

      if (url.pathname === "/session" && request.method === "GET") {
        return sessionStatus(request, env);
      }

      if (url.pathname === "/modifiche" && request.method === "GET") {
        return json({ ok: true, modifiche: await listModifiche(env) });
      }

      if (url.pathname === "/modifica" && request.method === "POST") {
        return saveModifica(request, env);
      }

      if (url.pathname === "/ripristina" && request.method === "POST") {
        return ripristinaPdv(request, env);
      }

      if (url.pathname === "/grab-visite" && request.method === "GET") {
        return json({ ok: true, visite: await listGrabVisite(env) });
      }

      if (url.pathname === "/grab-visita" && request.method === "POST") {
        return saveGrabVisita(request, env);
      }

      return json({ ok: false, error: "Endpoint non trovato" }, 404);
    } catch (error) {
      return json({ ok: false, error: String(error && error.message ? error.message : error) }, 500);
    }
  },
};
