// deploy trigger: GitHub Actions automatic Cloudflare Worker deploy
const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET,POST,OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type,X-Admin-Key",
  "Content-Type": "application/json; charset=utf-8",
};

function json(data, status = 200) {
  return new Response(JSON.stringify(data), { status, headers: CORS_HEADERS });
}

function normPdv(value) {
  const digits = String(value || "").match(/\d+/g);
  if (!digits) return "";
  return digits[0].padStart(5, "0");
}

function readAdminKey(request) {
  return request.headers.get("X-Admin-Key") || "";
}

function isAdmin(request, env) {
  return readAdminKey(request) && env.ADMIN_KEY && readAdminKey(request) === env.ADMIN_KEY;
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
  if (!isAdmin(request, env)) {
    return json({ ok: false, error: "ADMIN_KEY non valida" }, 401);
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
  if (!isAdmin(request, env)) {
    return json({ ok: false, error: "ADMIN_KEY non valida" }, 401);
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
  if (!isAdmin(request, env)) {
    return json({ ok: false, error: "ADMIN_KEY non valida" }, 401);
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
