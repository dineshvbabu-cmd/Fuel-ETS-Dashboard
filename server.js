const http = require("http");
const fs = require("fs");
const path = require("path");
const { URL } = require("url");

const PORT = Number(process.env.PORT || 3000);
const APP_DIR = path.join(__dirname, "compliance_dashboard");
const EUA_MARKET_URL = "https://tradingeconomics.com/commodity/carbon";
const MARKET_CACHE_MS = 30 * 60 * 1000;
const marketCache = {
  fetchedAt: 0,
  payload: null,
};

const contentTypes = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".txt": "text/plain; charset=utf-8",
};

function sendJson(res, statusCode, body) {
  const payload = JSON.stringify(body);
  res.writeHead(statusCode, {
    "Content-Type": "application/json; charset=utf-8",
    "Content-Length": Buffer.byteLength(payload),
  });
  res.end(payload);
}

function sendFile(res, filePath) {
  fs.readFile(filePath, (error, data) => {
    if (error) {
      res.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
      res.end("Not found");
      return;
    }
    const extension = path.extname(filePath).toLowerCase();
    res.writeHead(200, {
      "Content-Type": contentTypes[extension] || "application/octet-stream",
    });
    res.end(data);
  });
}

function stripHtml(html) {
  return html
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/gi, " ")
    .replace(/&#39;/g, "'")
    .replace(/&amp;/g, "&")
    .replace(/\s+/g, " ")
    .trim();
}

function parseSignedNumber(text) {
  return Number(String(text).replace(/[^\d.-]/g, ""));
}

async function fetchEuaMarketSnapshot() {
  if (marketCache.payload && Date.now() - marketCache.fetchedAt < MARKET_CACHE_MS) {
    return marketCache.payload;
  }

  const response = await fetch(EUA_MARKET_URL, {
    headers: {
      "User-Agent": "Mozilla/5.0 (compatible; Fuel-ETS-Dashboard/1.0; +https://railway.app)",
      Accept: "text/html,application/xhtml+xml",
    },
  });

  if (!response.ok) {
    throw new Error(`Source responded with ${response.status}`);
  }

  const html = await response.text();
  const text = stripHtml(html);

  const summaryMatch = text.match(
    /EU Carbon Permits .*? to ([\d.]+) EUR on ([A-Za-z]+ \d{1,2}, \d{4}), .*? ([\d.-]+)% from the previous day\. Over the past month, EU Carbon Permits(?:'s)? price has .*? ([\d.-]+)%, but it is still ([\d.-]+)% .*? than a year ago/i
  );
  const statsMatch = text.match(/Actual Previous Highest Lowest Dates Unit Frequency ([\d.]+) ([\d.]+) ([\d.]+) ([\d.]+) ([\d]{4} - [\d]{4}) EUR Daily/i);

  if (!summaryMatch && !statsMatch) {
    throw new Error("Could not parse EUA market page");
  }

  const price = summaryMatch ? Number(summaryMatch[1]) : Number(statsMatch[1]);
  const asOfDate = summaryMatch ? summaryMatch[2] : "Latest session";
  const dayChangePercent = summaryMatch ? parseSignedNumber(summaryMatch[3]) : 0;
  const monthChangePercent = summaryMatch ? parseSignedNumber(summaryMatch[4]) : 0;
  const yearChangePercent = summaryMatch ? parseSignedNumber(summaryMatch[5]) : 0;
  const previous = statsMatch ? Number(statsMatch[2]) : null;

  const payload = {
    commodity: "EU Carbon Permits",
    price,
    previous,
    dayChangePercent,
    monthChangePercent,
    yearChangePercent,
    asOfDate,
    sourceUrl: EUA_MARKET_URL,
    fetchedAt: new Date().toISOString(),
  };

  marketCache.payload = payload;
  marketCache.fetchedAt = Date.now();
  return payload;
}

const server = http.createServer(async (req, res) => {
  const parsed = new URL(req.url, `http://${req.headers.host}`);

  if (req.method === "GET" && parsed.pathname === "/api/health") {
    sendJson(res, 200, {
      ok: true,
      service: "eu-ets-fueleu-compliance-dashboard",
      status: "healthy",
    });
    return;
  }

  if (req.method === "GET" && parsed.pathname === "/api/market/eua") {
    try {
      const payload = await fetchEuaMarketSnapshot();
      sendJson(res, 200, payload);
    } catch (error) {
      sendJson(res, 502, {
        ok: false,
        message: error.message,
        sourceUrl: EUA_MARKET_URL,
      });
    }
    return;
  }

  if (req.method === "GET" && parsed.pathname === "/favicon.ico") {
    res.writeHead(204);
    res.end();
    return;
  }

  if (req.method !== "GET") {
    res.writeHead(405, { "Content-Type": "text/plain; charset=utf-8" });
    res.end("Method not allowed");
    return;
  }

  const requestedPath = parsed.pathname === "/" ? "/index.html" : parsed.pathname;
  const candidate = path.join(APP_DIR, requestedPath.replace(/^\/+/, ""));

  if (!candidate.startsWith(APP_DIR)) {
    res.writeHead(403, { "Content-Type": "text/plain; charset=utf-8" });
    res.end("Forbidden");
    return;
  }

  sendFile(res, candidate);
});

server.listen(PORT, () => {
  console.log(`EU ETS and FuelEU dashboard listening on port ${PORT}`);
});
