/**
 * 🔑 TifoTV beIN Bypass Worker (Smart v2) — Cloudflare
 * يكتشف تلقائياً المشغّل الصحيح لكل قناة (daddy2/3/4/5)
 */
export default {
  async fetch(request) {
    const url = new URL(request.url);
    const id = url.searchParams.get("id");
    const proxyTarget = url.searchParams.get("proxy");

    const cors = {
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "GET, HEAD, OPTIONS",
      "Access-Control-Allow-Headers": "*",
      "Access-Control-Expose-Headers": "*",
    };
    if (request.method === "OPTIONS") return new Response(null, { headers: cors });

    try {
      if (proxyTarget) return await proxyHls(proxyTarget, cors);
      if (!id) return jsonRes({ name: "TifoTV Smart v2", usage: "?id=91" }, cors);

      // STEP 1: Get correct iframe from dlhd.pk
      let playerUrl = null;
      try {
        const r = await fetch(`https://dlhd.pk/stream/stream-${encodeURIComponent(id)}.php`, {
          headers: { "Referer": "https://dlhd.pk/", "User-Agent": "Mozilla/5.0" },
        });
        if (r.ok) {
          const h = await r.text();
          const m = h.match(/iframe\s+src=["']([^"']+)["']/i);
          if (m) playerUrl = m[1];
        }
      } catch (e) {}

      const candidates = playerUrl ? [playerUrl] : [
        `https://hamis.romponalis.st/premiumtv/daddy4.php?id=${id}`,
        `https://hamis.romponalis.st/premiumtv/daddy5.php?id=${id}`,
        `https://hamis.romponalis.st/premiumtv/daddy3.php?id=${id}`,
        `https://hamis.romponalis.st/premiumtv/daddy2.php?id=${id}`,
      ];

      let m3u8 = null, lastErr = null;
      for (const purl of candidates) {
        try {
          const r = await fetch(purl, {
            headers: {
              "Referer": "https://dlhd.pk/",
              "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            },
          });
          if (!r.ok) { lastErr = "player " + r.status; continue; }
          const html = await r.text();
          const m = html.match(/atob\(['"]([A-Za-z0-9+/=_-]+)['"]\)/);
          if (!m) { lastErr = "no m3u8 in page"; continue; }
          const cand = atob(m[1]);
          const test = await fetch(cand, { headers: { "User-Agent": "Mozilla/5.0" } });
          if (test.ok) {
            const t = await test.text();
            if (t.includes("#EXTM3U")) { m3u8 = cand; break; }
          }
          lastErr = "stream " + test.status;
        } catch (e) { lastErr = e.message; }
      }
      if (!m3u8) return jsonRes({ error: "no working stream", detail: lastErr }, cors, 502);

      if (url.searchParams.get("json") === "1") {
        return jsonRes({ id, m3u8, proxied: `${url.origin}/?proxy=${encodeURIComponent(m3u8)}` }, cors);
      }
      return await proxyHls(m3u8, cors);

    } catch (e) {
      return jsonRes({ error: String(e) }, cors, 500);
    }
  },
};

async function proxyHls(target, cors) {
  const upstream = await fetch(target, {
    headers: { "Referer": "https://hamis.romponalis.st/", "User-Agent": "Mozilla/5.0" },
  });
  const ct = (upstream.headers.get("content-type") || "").toLowerCase();
  const isPL = ct.includes("mpegurl") || target.includes(".m3u8") || ct.includes("text/plain");
  if (!isPL) {
    const h = new Headers(upstream.headers);
    for (const [k, v] of Object.entries(cors)) h.set(k, v);
    return new Response(upstream.body, { status: upstream.status, headers: h });
  }
  let text = await upstream.text();
  const base = target.replace(/[?#].*$/, "").replace(/\/[^\/]*$/, "/");
  text = text.split("\n").map(line => {
    const t = line.trim();
    if (!t) return line;
    if (t.startsWith("#")) return line.replace(/URI="([^"]+)"/g, (_, u) => `URI="?proxy=${encodeURIComponent(new URL(u, base).href)}"`);
    return `?proxy=${encodeURIComponent(new URL(t, base).href)}`;
  }).join("\n");
  return new Response(text, {
    status: upstream.status,
    headers: { ...cors, "Content-Type": "application/vnd.apple.mpegurl", "Cache-Control": "no-cache, max-age=2" },
  });
}

function jsonRes(o, cors, s = 200) {
  return new Response(JSON.stringify(o, null, 2), {
    status: s,
    headers: { ...cors, "Content-Type": "application/json" },
  });
}
