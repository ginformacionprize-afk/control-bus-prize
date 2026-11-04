// GET /api/estado -> hora de la última lectura del Excel en OneDrive (para mostrarla en el login).
// No devuelve datos de usuarios ni asignaciones.
const { base, fail } = require("../lib/onedrive");

module.exports = async (req, res) => {
  try {
    const b = await base();
    res.setHeader("Cache-Control", "public, max-age=0, s-maxage=15, stale-while-revalidate=60");
    res.json({ actualizado: b.leido });
  } catch (err) { res.setHeader("Cache-Control", "no-store"); fail(res, err); }
};
