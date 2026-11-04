// GET  /api/login?placa=XXX  -> rol de la placa (para mostrarlo al escribir)
// POST /api/login {placa, clave} -> sesión
const { U, hoyLima, base, vigente, sign, fail } = require("../lib/onedrive");
const limite = require("../lib/limite");

module.exports = async (req, res) => {
  res.setHeader("Cache-Control", "no-store");
  try {
    const { usuarios } = await base();
    const hoy = hoyLima();
    if (req.method === "GET") {
      const placa = U(req.query.placa);
      const rows = usuarios.filter(x => x.PLACA === placa);
      const r = rows.find(x => vigente(x, hoy)) || rows[0];
      // El rol por placa no es sensible: el CDN de Vercel lo guarda 15 s para no saturar OneDrive
      res.setHeader("Cache-Control", "public, max-age=0, s-maxage=15, stale-while-revalidate=60");
      return res.json({ rol: r ? r.ROL : "" });
    }
    if (req.method !== "POST") return res.status(405).json({ error: "Método no permitido" });
    const placa = U(req.body && req.body.placa), clave = String((req.body && req.body.clave) || "").trim();
    const espera = limite.bloqueado(req, placa);
    if (espera) {
      res.setHeader("Retry-After", String(espera));
      return res.status(429).json({ error: `Demasiados intentos fallidos. Espera ${Math.ceil(espera / 60)} min e inténtalo de nuevo.` });
    }
    const rows = usuarios.filter(x => x.PLACA === placa && x.CLAVE === clave);
    if (!rows.length) { limite.falla(req, placa); return res.status(401).json({ error: "Placa o contraseña incorrectas." }); }
    limite.acierta(req, placa);
    const r = rows.find(x => vigente(x, hoy));
    if (!r) return res.status(403).json({ error: `La placa ${placa} no tiene conductor asignado para hoy. Revisa las fechas en USUARIOS.` });
    const user = { PLACA: r.PLACA, ROL: r.ROL, CONDUCTOR: r.CONDUCTOR };
    const token = sign({ p: r.PLACA, r: r.ROL, c: r.CONDUCTOR, exp: Date.now() + 24 * 3600 * 1000 });
    res.json({ token, user });
  } catch (err) { fail(res, err); }
};
