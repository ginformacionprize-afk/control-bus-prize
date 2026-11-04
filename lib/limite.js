// Límite de intentos fallidos de ingreso (en memoria de cada servidor).
// - 5 fallos con la misma placa desde la misma IP -> bloqueo de 10 min
// - 30 fallos desde la misma IP (cualquier placa)  -> bloqueo de 10 min
// Al ingresar bien se borra el contador de esa placa. No es un bloqueo global: no deja fuera a la placa real
// desde otra IP, solo frena a quien prueba claves por fuerza bruta.
const VENTANA_MS = 10 * 60 * 1000;
const MAX_PLACA = 5, MAX_IP = 30;
const fallos = new Map(); // clave -> { n, desde }

const ipDe = req => String((req.headers["x-forwarded-for"] || req.headers["x-real-ip"] || "")).split(",")[0].trim() || "desconocida";
function estado(k, max) {
  const e = fallos.get(k);
  if (!e) return 0;
  const resta = e.desde + VENTANA_MS - Date.now();
  if (resta <= 0) { fallos.delete(k); return 0; }
  return e.n >= max ? resta : 0; // milisegundos que faltan, o 0 si no está bloqueado
}
function suma(k) {
  const e = fallos.get(k);
  if (!e || e.desde + VENTANA_MS <= Date.now()) fallos.set(k, { n: 1, desde: Date.now() });
  else e.n++;
}
function limpia() { // evita que el mapa crezca sin fin
  if (fallos.size < 2000) return;
  const ahora = Date.now();
  for (const [k, e] of fallos) if (e.desde + VENTANA_MS <= ahora) fallos.delete(k);
}

// Devuelve los segundos de espera si está bloqueado, o 0 si puede intentar.
function bloqueado(req, placa) {
  const ip = ipDe(req);
  const ms = Math.max(estado(`p|${ip}|${placa}`, MAX_PLACA), estado(`i|${ip}`, MAX_IP));
  return Math.ceil(ms / 1000);
}
function falla(req, placa) { limpia(); const ip = ipDe(req); suma(`p|${ip}|${placa}`); suma(`i|${ip}`); }
function acierta(req, placa) { fallos.delete(`p|${ipDe(req)}|${placa}`); }

module.exports = { bloqueado, falla, acierta };
