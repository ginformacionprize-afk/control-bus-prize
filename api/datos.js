// GET /api/datos -> capacidades, asignaciones (de ayer a 7 días) y placas. Nunca envía contraseñas.
const { hoyLima, addDays, base, session, fail } = require("../lib/onedrive");

module.exports = async (req, res) => {
  res.setHeader("Cache-Control", "no-store");
  const ses = session(req);
  if (!ses) return res.status(401).json({ error: "Sesión vencida" });
  try {
    const b = await base();
    const hoy = hoyLima(), desde = addDays(hoy, -1), hasta = addDays(hoy, 7);
    res.json({
      maestra: b.usuarios.map(({ CLAVE, ...x }) => x),
      capacidades: b.capacidades,
      asignaciones: b.asignaciones.filter(a => a.FECHA >= desde && a.FECHA <= hasta),
      actualizado: b.leido
    });
  } catch (err) { fail(res, err); }
};
