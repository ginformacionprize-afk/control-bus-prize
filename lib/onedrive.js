// Conexión con el Excel de OneDrive mediante Microsoft Graph (permiso de aplicación).
const crypto = require("crypto");

const U = v => String(v ?? "").trim().toUpperCase();
const norm = s => U(s).normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^A-Z0-9]/g, "");
const pad2 = n => String(n).padStart(2, "0");
const hoyLima = () => new Date(Date.now() - 5 * 3600 * 1000).toISOString().slice(0, 10);
const addDays = (iso, n) => { const d = new Date(iso + "T12:00:00Z"); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); };
const serial = iso => { const [y, m, d] = iso.split("-").map(Number); return (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000; };
const padDni = v => { const d = String(v ?? "").replace(/\D/g, ""); return d ? d.padStart(8, "0") : ""; };

function toISO(v) {
  if (v === "" || v == null) return "";
  if (typeof v === "number") return new Date(Math.round((v - 25569) * 86400000)).toISOString().slice(0, 10);
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/); if (m) return `${m[1]}-${pad2(m[2])}-${pad2(m[3])}`;
  m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);
  if (m) { const y = m[3].length === 2 ? "20" + m[3] : m[3]; return `${y}-${pad2(m[2])}-${pad2(m[1])}`; }
  return "";
}
function hora(v) {
  if (typeof v === "number" && v < 1) { const t = Math.round(v * 1440); return pad2(Math.floor(t / 60)) + ":" + pad2(t % 60); }
  return U(v);
}

// ---------- token de Microsoft ----------
let tok = null, tokExp = 0, tokP = null;
function graphToken() {
  if (tok && Date.now() < tokExp - 60000) return Promise.resolve(tok);
  // Una sola petición de token aunque lleguen muchas solicitudes a la vez
  return tokP ??= newToken().finally(() => { tokP = null; });
}
async function newToken() {
  const r = await fetch(`https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: process.env.CLIENT_ID, client_secret: process.env.CLIENT_SECRET,
      scope: "https://graph.microsoft.com/.default", grant_type: "client_credentials"
    })
  });
  const j = await r.json();
  if (!r.ok) throw new Error("No se pudo conectar con Microsoft: " + (j.error_description || r.status));
  tok = j.access_token; tokExp = Date.now() + j.expires_in * 1000;
  return tok;
}
function fileUrl(file) {
  const path = [process.env.BASE_PATH, file].filter(Boolean).join("/").split("/").map(encodeURIComponent).join("/");
  const drive = process.env.DRIVE_ID
    ? `drives/${encodeURIComponent(process.env.DRIVE_ID)}`
    : `users/${encodeURIComponent(process.env.DRIVE_USER)}/drive`;
  return `https://graph.microsoft.com/v1.0/${drive}/root:/${path}:`;
}
const sleep = ms => new Promise(r => setTimeout(r, ms));
async function graph(url, opt = {}) {
  const t = await graphToken();
  let r;
  // Si Microsoft limita (429) o falla un momento (503/504), reintenta respetando Retry-After
  for (let intento = 0; ; intento++) {
    r = await fetch(url, { ...opt, headers: { Authorization: "Bearer " + t, "Content-Type": "application/json", ...(opt.headers || {}) } });
    if (![429, 503, 504].includes(r.status) || intento >= 2) break;
    const ra = Number(r.headers.get("retry-after"));
    await sleep(Math.min(ra > 0 ? ra * 1000 : 500 * 2 ** intento, 4000) + Math.random() * 250);
  }
  if (r.status === 204) return null;
  const j = await r.json().catch(() => null);
  if (!r.ok) { const e = new Error("OneDrive " + r.status + ": " + ((j && j.error && j.error.message) || "")); e.status = r.status; throw e; }
  return j;
}

// ---------- bloqueo compartido entre servidores ----------
// Crea un archivo de bloqueo en la carpeta de OneDrive; si ya existe, otro servidor está escribiendo y se espera.
// Así dos celulares que sincronizan a la vez no leen y escriben la tabla al mismo tiempo.
const LOCK_STALE_MS = 60000, LOCK_WAIT_MS = 45000;
async function withLock(name, fn) {
  const url = fileUrl(name);
  const fin = Date.now() + LOCK_WAIT_MS;
  for (;;) {
    try {
      await graph(url + "/content?@microsoft.graph.conflictBehavior=fail", {
        method: "PUT", headers: { "Content-Type": "text/plain" }, body: new Date().toISOString()
      });
      break;
    } catch (err) {
      if (err.status !== 409) throw err;
      // Bloqueo abandonado (un servidor que se cayó): se libera después de 60 s
      const meta = await graph(url).catch(() => null);
      const edad = meta ? Date.now() - Date.parse(meta.lastModifiedDateTime) : 0;
      if (edad > LOCK_STALE_MS) { await graph(url, { method: "DELETE" }).catch(() => {}); continue; }
      if (Date.now() > fin) throw new Error("OneDrive ocupado escribiendo otros registros. Se reintentará al sincronizar.");
      await sleep(300 + Math.random() * 500);
    }
  }
  try { return await fn(); }
  finally { await graph(url, { method: "DELETE" }).catch(() => {}); }
}

// Lee una tabla de Excel: devuelve encabezados y filas (valores crudos)
async function readTable(file, table) {
  // $select=values: pide solo los valores (sin formatos, fórmulas ni tipos) -> respuesta mucho más liviana
  const j = await graph(fileUrl(file) + `/workbook/tables/${encodeURIComponent(table)}/range?$select=values`);
  const [head, ...rows] = j.values;
  return { head, rows };
}
const asObjects = ({ head, rows }) => rows
  .filter(r => r.some(v => v !== "" && v != null))
  .map(r => Object.fromEntries(head.map((k, i) => [norm(k), r[i]])));
const pick = (o, names) => { for (const n of names) if (o[n] !== undefined && o[n] !== "") return o[n]; return ""; };

// ---------- base maestra (con caché corta para no saturar OneDrive) ----------
// Permite terminar la relectura del Excel después de responder (en Vercel); en local no hace nada.
let waitUntil = () => {};
try { ({ waitUntil } = require("@vercel/functions")); } catch (e) {}

let cache = null, cacheAt = 0, baseP = null;
const CACHE_MS = 15000;          // datos frescos: se vuelven a leer cada 15 s
const SWR_MS = 2 * 60 * 1000;    // hasta 2 min: se responde al instante y se relee el Excel en segundo plano
const STALE_MS = 10 * 60 * 1000; // si OneDrive falla, se usan los últimos datos hasta 10 min
function base() {
  const edad = Date.now() - cacheAt;
  if (cache && edad < CACHE_MS) return Promise.resolve(cache);
  // Una sola lectura del Excel aunque lleguen muchas solicitudes a la vez
  baseP ??= readBase().finally(() => { baseP = null; });
  if (!cache) return baseP;
  if (edad < SWR_MS) {
    // Nadie espera a OneDrive: se responde con lo que hay y se actualiza por detrás
    const p = baseP.catch(() => {}); waitUntil(p);
    return Promise.resolve(cache);
  }
  return baseP.catch(err => { if (Date.now() - cacheAt < STALE_MS) return cache; throw err; });
}
async function readBase() {
  const file = process.env.BASE_FILE || "BASE_MAESTRA_ASIGNACIONES.xlsx";
  const [us, cap, asg] = await Promise.all([
    readTable(file, "USUARIOS"), readTable(file, "CAPACIDADES"), readTable(file, "ASIGNACIONES")
  ]);
  cache = {
    usuarios: asObjects(us).map(o => ({
      PLACA: U(pick(o, ["PLACA", "USUARIO"])), CLAVE: String(pick(o, ["CONTRASENA", "CLAVE"])).trim(),
      ROL: U(pick(o, ["ROL"])) || "CONDUCTOR", CONDUCTOR: U(pick(o, ["CONDUCTOR", "NOMBRE"])),
      INICIO: toISO(pick(o, ["FECHAINICIO", "INICIO"])), FIN: toISO(pick(o, ["FECHAFIN", "FIN"]))
    })).filter(x => x.PLACA),
    capacidades: asObjects(cap).map(o => ({
      PLACA: U(pick(o, ["PLACA"])), INICIO: toISO(pick(o, ["FECHAINICIO", "INICIO"])),
      FIN: toISO(pick(o, ["FECHAFIN", "FIN"])), CAPACIDAD: Number(pick(o, ["CAPACIDAD", "ASIENTOS"])) || 0
    })).filter(x => x.PLACA),
    asignaciones: asObjects(asg).map(o => ({
      FECHA: toISO(pick(o, ["FECHA"])), DNI: padDni(pick(o, ["DNI"])), APELLIDOS: U(pick(o, ["APELLIDOS", "APELLIDOSYNOMBRES"])),
      PLACA: U(pick(o, ["PLACA"])), TIPO: U(pick(o, ["TIPO"])), HORARIO: hora(pick(o, ["HORARIO", "HORA"])), RUTA: U(pick(o, ["RUTA"]))
    })).filter(x => x.DNI)
  };
  cacheAt = Date.now();
  return cache;
}
const vigente = (r, f) => (!r.INICIO || r.INICIO <= f) && (!r.FIN || r.FIN >= f);

// ---------- sesión firmada ----------
function sign(payload) {
  const b = Buffer.from(JSON.stringify(payload)).toString("base64url");
  const s = crypto.createHmac("sha256", process.env.APP_SECRET).update(b).digest("base64url");
  return b + "." + s;
}
function session(req) {
  const t = String(req.headers.authorization || "").replace(/^Bearer\s+/i, "");
  const [b, s] = t.split(".");
  if (!b || !s) return null;
  const e = crypto.createHmac("sha256", process.env.APP_SECRET).update(b).digest("base64url");
  if (e.length !== s.length || !crypto.timingSafeEqual(Buffer.from(e), Buffer.from(s))) return null;
  try { const p = JSON.parse(Buffer.from(b, "base64url").toString()); return p.exp > Date.now() ? p : null; } catch (e) { return null; }
}
const isAdmin = rol => /ADMIN|SUPERV/.test(rol || "");

function fail(res, err) {
  console.error(err);
  res.status(500).json({ error: "Error del servidor: " + (err && err.message ? err.message : err) });
}

module.exports = { U, norm, hoyLima, addDays, serial, graph, fileUrl, readTable, withLock, base, vigente, sign, session, isAdmin, fail };
