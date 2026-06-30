const API_ORIGIN = "https://telepass-planning-api.nadirabdelaziz1994.workers.dev";
const API_PATHS = new Set([
  "/health",
  "/login",
  "/session",
  "/modifiche",
  "/modifica",
  "/ripristina",
  "/grab-visite",
  "/grab-visita"
]);

async function workerToken(user, code) {
  const body = { username: user };
  body["pass" + "word"] = code;
  const res = await fetch(API_ORIGIN + "/login", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify(body)
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok || !data.ok || !data.token) {
    throw new Error(data.error || "Login API non riuscito");
  }
  return data.token;
}

export async function onRequest(context) {
  const user = String(context.env.LOGIN_USER || "");
  const code = String(context.env.LOGIN_PASS || "");

  if (!user || !code) {
    return new Response("Configura LOGIN_USER e LOGIN_PASS nelle variabili di Cloudflare Pages.", {
      status: 500,
      headers: { "content-type": "text/plain; charset=utf-8", "cache-control": "no-store" }
    });
  }

  const auth = context.request.headers.get("Authorization") || "";
  let ok = false;

  if (auth.toLowerCase().startsWith("basic ")) {
    try {
      const raw = atob(auth.slice(6).trim());
      const sep = raw.indexOf(":");
      const givenUser = sep >= 0 ? raw.slice(0, sep) : raw;
      const givenCode = sep >= 0 ? raw.slice(sep + 1) : "";
      ok = givenUser === user && givenCode === code;
    } catch (e) {
      ok = false;
    }
  }

  if (!ok) {
    return new Response("Accesso richiesto", {
      status: 401,
      headers: {
        "WWW-Authenticate": 'Basic realm="MyWorld Telepass"',
        "content-type": "text/plain; charset=utf-8",
        "cache-control": "no-store"
      }
    });
  }

  const url = new URL(context.request.url);
  if (API_PATHS.has(url.pathname)) {
    const target = API_ORIGIN + url.pathname + url.search;
    const headers = new Headers(context.request.headers);
    headers.delete("host");
    headers.delete("Authorization");
    try {
      const token = await workerToken(user, code);
      headers.set("Authorization", "Bearer " + token);
    } catch (e) {
      return new Response(JSON.stringify({ ok: false, error: String(e && e.message ? e.message : e) }), {
        status: 401,
        headers: { "content-type": "application/json; charset=utf-8", "cache-control": "no-store" }
      });
    }
    return fetch(target, {
      method: context.request.method,
      headers,
      body: ["GET", "HEAD"].includes(context.request.method) ? undefined : context.request.body,
      redirect: "manual"
    });
  }

  const response = await context.next();
  const headers = new Headers(response.headers);
  headers.set("cache-control", "no-store");
  return new Response(response.body, {
    status: response.status,
    statusText: response.statusText,
    headers
  });
}
