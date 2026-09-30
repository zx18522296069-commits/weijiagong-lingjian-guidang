const UPSTREAM = "https://production-control-api.zx18522296069.chatgpt.site";

const ALLOWED_PATHS = new Set([
  "/api/status",
  "/api/run/parts",
  "/api/run/split",
  "/api/results/parts",
  "/api/results/split",
]);

module.exports = async function handler(req, res) {
  try {
    const rawPath = Array.isArray(req.query.path) ? req.query.path[0] : req.query.path;
    const path = String(rawPath || "");

    if (!ALLOWED_PATHS.has(path)) {
      res.status(400).json({ status: "error", message: "Unsupported proxy path" });
      return;
    }

    const method = String(req.method || "GET").toUpperCase();
    if (!["GET", "POST", "OPTIONS"].includes(method)) {
      res.status(405).json({ status: "error", message: "Method not allowed" });
      return;
    }

    if (method === "OPTIONS") {
      res.status(204).end();
      return;
    }

    const headers = {
      "Content-Type": "application/json",
      "X-Control-Key": String(req.headers["x-control-key"] || ""),
    };

    const init = { method, headers };
    if (method !== "GET") {
      init.body = typeof req.body === "string"
        ? req.body
        : JSON.stringify(req.body || {});
    }

    const upstream = await fetch(UPSTREAM + path, init);
    const body = await upstream.text();
    const contentType = upstream.headers.get("content-type") || "application/json; charset=utf-8";

    res.status(upstream.status);
    res.setHeader("Content-Type", contentType);
    res.setHeader("Cache-Control", "no-store");
    res.send(body);
  } catch (error) {
    res.status(502).json({
      status: "error",
      message: "控制服务代理失败",
      detail: error instanceof Error ? error.message : String(error),
    });
  }
};
