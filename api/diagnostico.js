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
    // Prueba de punta a punta con el código real: login (rol), datos y envío de un registro de prueba (que luego se borra).
    // Se niega a escribir si la configuración apunta al OneDrive de producción.
    if (req.query.probarFlujo) {
      if (process.env.VERCEL_ENV !== "preview" || /^jcontreras@/i.test(process.env.DRIVE_USER || ""))
        return res.status(403).json({ error: "Prueba de escritura bloqueada: solo en Preview y fuera del OneDrive de producción." });
      const { sign, readTable, graph: g, fileUrl, norm } = require("../lib/onedrive");
      const llama = (handler, r) => new Promise(ok => {
        const resp = { code: 200, setHeader() {}, status(c) { this.code = c; return this; }, json(b) { ok({ code: this.code, b }); } };
        Promise.resolve(handler({ headers: {}, query: {}, body: {}, ...r }, resp)).catch(e => ok({ code: 500, b: { error: String(e && e.message || e) } }));
      });
      const out = { config: { DRIVE_USER: process.env.DRIVE_USER, BASE_PATH: process.env.BASE_PATH, BASE_FILE: process.env.BASE_FILE, REG_FILE: process.env.REG_FILE } };
      const placa = String(req.query.placa || "CFY-121");
      const rol = await llama(require("./login"), { method: "GET", query: { placa } });
      out.login_rol = { http: rol.code, rol: rol.b.rol, error: rol.b.error };
      const token = sign({ p: "PRUEBA-MIGRACION", r: "ADMINISTRADOR", c: "PRUEBA", exp: Date.now() + 600000 });
      const auth = { authorization: "Bearer " + token };
      const dat = await llama(require("./datos"), { method: "GET", headers: auth });
      out.datos = dat.code === 200
        ? { http: 200, usuarios: dat.b.maestra.length, capacidades: dat.b.capacidades.length, asignaciones_ayer_a_7d: dat.b.asignaciones.length, actualizado: dat.b.actualizado, expone_claves: dat.b.maestra.some(u => "CLAVE" in u) }
        : { http: dat.code, error: dat.b.error };
      const file = process.env.REG_FILE || "REGISTROS_EMBARQUE.xlsx";
      const antes = (await readTable(file, "REGISTROS")).rows.length;
      const id = "PRUEBA-MIGRACION-" + Date.now().toString(36);
      const hoy = new Date().toISOString().slice(0, 10);
      const env = await llama(require("./registros"), { method: "POST", headers: auth, body: { registros: [{
        ID: id, REEMPLAZA: [], LECTURA: "DIGITADO", FECHA: hoy, DNI: "00000000", APELLIDOS: "PRUEBA MIGRACION (borrar)",
        PLACA_PROGRAMADA: placa, TURNO_PROGRAMADO: "INGRESO", HORARIO_PROGRAMADO: "00:00", RUTA_PROGRAMADA: "PRUEBA",
        PLACA_EJECUTADA: placa, TURNO_EJECUTADO: "INGRESO", HORA_EJECUTADA: "00:00", CONDUCTOR: "PRUEBA", ESTADO: "PRUEBA", COINCIDE_PLACA: "SI" }] } });
      out.registros_envio = { http: env.code, resultado: env.b };
      const { head, rows } = await readTable(file, "REGISTROS");
      const idCol = head.findIndex(h => norm(h) === "IDREGISTRO");
      const i = rows.findIndex(r => String(r[idCol]) === id);
      out.registros_fila_visible_en_excel = i >= 0;
      if (i >= 0) { await g(fileUrl(file) + `/workbook/tables/REGISTROS/rows/itemAt(index=${i})`, { method: "DELETE" }); }
      const despues = (await readTable(file, "REGISTROS")).rows.length;
      out.registros_limpieza = { filas_antes: antes, filas_despues_de_borrar: despues, queda_igual: antes === despues };
      return res.json(out);
    }
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
