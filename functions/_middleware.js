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

  const response = await context.next();
  const headers = new Headers(response.headers);
  headers.set("cache-control", "no-store");
  return new Response(response.body, {
    status: response.status,
    statusText: response.statusText,
    headers
  });
}
