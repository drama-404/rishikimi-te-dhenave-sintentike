// Cross-device persistence for the set of "reviewed" respondent IDs.
//
// Backed by an Upstash Redis / Vercel KV store via its REST API. Provision a
// store in the Vercel dashboard (Storage -> Upstash/KV) and connect it to this
// project; that injects the env vars below automatically. Either naming works:
//   KV_REST_API_URL        / KV_REST_API_TOKEN          (Vercel KV / Upstash)
//   UPSTASH_REDIS_REST_URL / UPSTASH_REDIS_REST_TOKEN
//
// Endpoints:
//   GET    /api/reviewed        -> { reviewed: ["1","2",...] }
//   POST   /api/reviewed {id}   -> marks one id reviewed
//   DELETE /api/reviewed        -> clears all (used by the Rivendos button)

const KEY = "reviewed";

// Resolve the Upstash REST endpoint + write token from env, regardless of the
// prefix the Vercel/Upstash integration used (KV_REST_API_*, UPSTASH_REDIS_*,
// or a custom prefix like STORAGE_REST_URL / STORAGE_REST_TOKEN).
function resolveStore() {
  const e = process.env;
  let url = e.KV_REST_API_URL || e.UPSTASH_REDIS_REST_URL;
  let token = e.KV_REST_API_TOKEN || e.UPSTASH_REDIS_REST_TOKEN;
  if (!url) {
    const k = Object.keys(e).find((k) => /REST_URL$/.test(k) && /^https:\/\//.test(e[k] || ""));
    if (k) url = e[k];
  }
  if (!token) {
    // Prefer a full (non read-only) token so writes (SADD/DEL) work.
    const k = Object.keys(e).find((k) => /REST_TOKEN$/.test(k) && !/READ_?ONLY/i.test(k));
    if (k) token = e[k];
  }
  return { url, token };
}

async function redis(store, command) {
  const res = await fetch(store.url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${store.token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(command),
  });
  if (!res.ok) throw new Error(`redis ${res.status}: ${await res.text()}`);
  const data = await res.json();
  return data.result;
}

module.exports = async (req, res) => {
  const store = resolveStore();
  if (!store.url || !store.token) {
    res.status(500).json({ error: "Storage not configured (no *_REST_URL / *_REST_TOKEN env vars found)." });
    return;
  }
  try {
    if (req.method === "GET") {
      const ids = await redis(store, ["SMEMBERS", KEY]);
      res.status(200).json({ reviewed: Array.isArray(ids) ? ids : [] });
      return;
    }
    if (req.method === "POST") {
      let body = req.body;
      if (typeof body === "string") { try { body = JSON.parse(body); } catch (_) { body = {}; } }
      const id = body && body.id != null ? String(body.id) : "";
      if (!id) { res.status(400).json({ error: "missing id" }); return; }
      await redis(store, ["SADD", KEY, id]);
      res.status(200).json({ ok: true });
      return;
    }
    if (req.method === "DELETE") {
      await redis(store, ["DEL", KEY]);
      res.status(200).json({ ok: true });
      return;
    }
    res.status(405).json({ error: "method not allowed" });
  } catch (e) {
    res.status(500).json({ error: String((e && e.message) || e) });
  }
};
