// POST /api/registros {registros:[...]} -> agrega filas a la tabla REGISTROS de REGISTROS_EMBARQUE.xlsx
// No duplica (usa ID REGISTRO) y reemplaza los "INTENTÓ SUBIR" que luego subieron.
const { U, norm, serial, graph, fileUrl, readTable, session, isAdmin, fail } = require("../lib/onedrive");

module.exports = async (req, res) => {
  res.setHeader("Cache-Control", "no-store");
  if (req.method !== "POST") return res.status(405).json({ error: "Método no permitido" });
  const ses = session(req);
  if (!ses) return res.status(401).json({ error: "Sesión vencida" });
  const regs = (req.body && Array.isArray(req.body.registros)) ? req.body.registros : [];
  if (!regs.length) return res.json({ ok: true, agregados: 0, reemplazados: 0 });
  if (!isAdmin(ses.r) && regs.some(r => U(r.PLACA_EJECUTADA) !== ses.p))
    return res.status(403).json({ error: "Solo puedes enviar registros de tu placa." });
  try {
    const file = process.env.REG_FILE || "REGISTROS_EMBARQUE.xlsx";
    const tbl = fileUrl(file) + "/workbook/tables/REGISTROS";
    const { head, rows } = await readTable(file, "REGISTROS");
    const col = name => head.findIndex(h => norm(h) === name);
    const idCol = col("IDREGISTRO");
    if (idCol < 0) throw new Error("Falta la columna ID REGISTRO en la tabla REGISTROS.");
    const ids = new Map();
    rows.forEach((r, i) => { if (r[idCol]) ids.set(String(r[idCol]), i); });

    // 1) Borrar los "intentó subir" reemplazados (de abajo hacia arriba para no mover índices)
    const borrar = [...new Set(regs.flatMap(r => r.REEMPLAZA || []))]
      .map(id => ids.get(String(id))).filter(i => i !== undefined).sort((a, b) => b - a);
    for (const i of borrar) await graph(tbl + `/rows/itemAt(index=${i})`, { method: "DELETE" });

    // 2) Agregar solo los que aún no existen
    const nuevos = regs.filter(r => r.ID && !ids.has(String(r.ID)));
    if (nuevos.length) {
      const map = {
        FECHA: r => serial(r.FECHA), DNI: r => String(r.DNI), APELLIDOS: r => r.APELLIDOS,
        PLACAPROGRAMADA: r => r.PLACA_PROGRAMADA, TURNOPROGRAMADO: r => r.TURNO_PROGRAMADO,
        HORARIOPROGRAMADO: r => r.HORARIO_PROGRAMADO, RUTAPROGRAMADA: r => r.RUTA_PROGRAMADA,
        PLACAEJECUTADA: r => r.PLACA_EJECUTADA, TURNOEJECUTADO: r => r.TURNO_EJECUTADO,
        HORAEJECUTADA: r => r.HORA_EJECUTADA, CONDUCTOR: r => r.CONDUCTOR, ESTADO: r => r.ESTADO,
        COINCIDEPLACA: r => r.COINCIDE_PLACA, TIPODELECTURA: r => r.LECTURA || "DIGITADO", IDREGISTRO: r => r.ID
      };
      const values = nuevos.map(r => head.map(h => { const f = map[norm(h)]; return f ? (f(r) ?? "") : ""; }));
      await graph(tbl + "/rows", { method: "POST", body: JSON.stringify({ values }) });
    }
    res.json({ ok: true, agregados: nuevos.length, reemplazados: borrar.length });
  } catch (err) { fail(res, err); }
};
