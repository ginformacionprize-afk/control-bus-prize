// DIAGNÓSTICO TEMPORAL (solo rama desarrollo / Preview). En producción responde 404.
// Sirve para ubicar la carpeta nueva en OneDrive antes de migrar y comprobar que tiene los Excel y sus tablas.
//   /api/diagnostico?user=correo&buscar=TEXTO        busca archivos/carpetas en el OneDrive de ese usuario y muestra su ruta
//   /api/diagnostico?user=correo&path=ruta           lista el contenido de esa carpeta
//   /api/diagnostico?user=correo&path=ruta&tablas=archivo.xlsx   lista las tablas de ese Excel
// Solo lectura. NO se debe fusionar a main (se quita antes de migrar).
const { graph } = require("../lib/onedrive");

const seg = p => String(p || "").split("/").filter(Boolean).map(encodeURIComponent).join("/");

module.exports = async (req, res) => {
  res.setHeader("Cache-Control", "no-store");
  if (process.env.VERCEL_ENV === "production") return res.status(404).json({ error: "No disponible." });
  const user = String(req.query.user || process.env.DRIVE_USER || "");
  const drive = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(user)}/drive`;
  try {
    if (req.query.buscar) {
      const q = String(req.query.buscar).replace(/'/g, "''");
      const j = await graph(`${drive}/root/search(q='${encodeURIComponent(q)}')?$top=50`);
      return res.json({
        user,
        resultados: (j.value || []).map(x => ({
          nombre: x.name, tipo: x.folder ? "carpeta" : "archivo",
          carpeta: ((x.parentReference && x.parentReference.path) || "").replace(/^\/drive\/root:?/, "") || "/",
          modificado: x.lastModifiedDateTime
        }))
      });
    }
    const path = seg(req.query.path);
    const base = path ? `${drive}/root:/${path}:` : `${drive}/root`;
    const j = await graph(`${base}/children?$top=200`);
    const out = {
      user, path: String(req.query.path || "/"),
      contenido: (j.value || []).map(x => ({ nombre: x.name, tipo: x.folder ? "carpeta" : "archivo", tamano: x.size, modificado: x.lastModifiedDateTime }))
    };
    if (req.query.tablas) {
      const t = await graph(`${path ? `${drive}/root:/${path}/${encodeURIComponent(req.query.tablas)}:` : `${drive}/root:/${encodeURIComponent(req.query.tablas)}:`}/workbook/tables`);
      out.tablas = (t.value || []).map(x => x.name);
    }
    res.json(out);
  } catch (err) { res.status(500).json({ error: String(err && err.message || err) }); }
};
