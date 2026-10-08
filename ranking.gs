// Pasapalabra IA — registro de partidas, resumen por estudiante y ranking
// Pegar COMPLETO en Extensiones → Apps Script de la planilla, guardar y
// publicar como aplicación web (ver instrucciones).

const SHEET = "partidas";
const RESUMEN = "resumen";

// ---- hojas ---------------------------------------------------------------

function partidas_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SHEET);
  if (!s) s = ss.insertSheet(SHEET);
  if (s.getLastRow() === 0) {
    s.appendRow(["fecha","nombre","seccion","avatar","modo",
                 "aciertos","errores","pasadas","total",
                 "porcentaje","tiempo_seg","inicio"]);
  }
  return s;
}

function resumen_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(RESUMEN);
  if (!s) s = ss.insertSheet(RESUMEN);
  if (s.getLastRow() === 0) {
    s.appendRow(["nombre","seccion","avatar","partidas","ultima_vez",
                 "aciertos","errores","pasadas","mejor_partida",
                 "letras_jugadas","porcentaje","tiempo_seg"]);
  }
  return s;
}

// Normaliza el nombre para reconocer al mismo estudiante aunque escriba
// distinto: ignora mayúsculas, acentos/ñ y espacios de más.
function norm_(t) {
  return String(t||"").toLowerCase().normalize("NFD")
    .replace(/[\u0300-\u036f]/g,"")
    .replace(/ñ/g,"n")
    .replace(/\s+/g," ").trim();
}

// ---- escritura: cada partida suma una fila y actualiza el resumen ---------

function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  const nombre  = String(d.nombre||"").slice(0,40);
  const seccion = String(d.seccion||"").slice(0,20);
  const avatar  = Math.max(0,Math.min(9,+d.avatar||0));
  const modo    = String(d.modo||"").slice(0,20);
  const aciertos= Math.max(0,Math.min(99,+d.aciertos||0));
  const errores = Math.max(0,Math.min(99,+d.errores||0));
  const pasadas = Math.max(0,Math.min(99,+d.pasadas||0));
  const total   = +d.total||16;
  const porcent = Math.max(0,Math.min(100,+d.porcentaje||0));
  const tiempo  = d.tiempo_seg===""||d.tiempo_seg===undefined?"":Math.max(0,+d.tiempo_seg||0);
  const inicio  = String(d.inicio||"").slice(0,4);
  const ahora   = new Date();

  partidas_().appendRow([ahora, nombre, seccion, avatar, modo,
                         aciertos, errores, pasadas, total,
                         porcent, tiempo, inicio]);

  // resumen: una fila por estudiante (nombre + sección)
  const s = resumen_();
  const clave = norm_(nombre)+"|"+norm_(seccion);
  const filas = s.getDataRange().getValues();
  for (let i = 1; i < filas.length; i++) {
    const r = filas[i];
    if (norm_(r[0])+"|"+norm_(r[1]) === clave) {
      const partidas = (+r[3]||0) + 1;
      const ac = (+r[5]||0) + aciertos;
      const er = (+r[6]||0) + errores;
      const pa = (+r[7]||0) + pasadas;
      const mj = Math.max(+r[8]||0, aciertos);
      const lj = (+r[9]||0) + total;
      const ts = (+r[11]||0) + (+tiempo||0);
      s.getRange(i+1, 1, 1, 12).setValues([[
        nombre, seccion, avatar, partidas, ahora,
        ac, er, pa, mj, lj,
        lj ? Math.round(ac/lj*100) : 0, ts
      ]]);
      return ok_();
    }
  }
  s.appendRow([nombre, seccion, avatar, 1, ahora,
               aciertos, errores, pasadas, aciertos, total,
               total ? Math.round(aciertos/total*100) : 0, +tiempo||0]);
  return ok_();
}

function ok_() {
  return ContentService.createTextOutput(JSON.stringify({ok:true}))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---- lectura: ranking para el juego, ya acumulado por estudiante ----------
// GET ?top=N → [{nombre, seccion, aciertos, partidas, avatar, total}]

function doGet(e) {
  const top = Math.min(+(e && e.parameter && e.parameter.top || 15), 200);
  const filas = resumen_().getDataRange().getValues().slice(1);
  const lista = filas.map(r => ({
    nombre:String(r[0]), seccion:String(r[1]), avatar:+r[2]||0,
    partidas:+r[3]||0, aciertos:+r[5]||0, total:16
  })).sort((a,b)=>b.aciertos-a.aciertos || a.partidas-b.partidas).slice(0,top);
  return ContentService.createTextOutput(JSON.stringify(lista))
    .setMimeType(ContentService.MimeType.JSON);
}
