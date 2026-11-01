// Pasapalabra IA — registro de partidas, resumen por estudiante y ranking
// Pegar COMPLETO en Extensiones → Apps Script de la planilla, guardar y
// publicar como aplicación web (ver instrucciones).
//
// Además escribe la nota por tramos (10-19 aciertos = 1, 20-29 = 2, ...,
// 100+ = 10) en la columna "JUEGO EDUCATIVO: PASAPALABRA" de la planilla
// de registro del curso (REGISTRO_ID). Reconoce al estudiante aunque
// escriba un solo nombre/apellido, sin acentos o con letra de más/menos.

const SHEET = "partidas";
const RESUMEN = "resumen";
const REGISTRO_ID = "PEGAR_ID_DE_LA_PLANILLA_DE_NOTAS";
const HOJA_REGISTRO = ""; // opcional: nombre de la pestaña de la lista (ej. "5° 1°")

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
      escribirPasapalabra_(nombre, ac);
      return ok_();
    }
  }
  s.appendRow([nombre, seccion, avatar, 1, ahora,
               aciertos, errores, pasadas, aciertos, total,
               total ? Math.round(aciertos/total*100) : 0, +tiempo||0]);
  escribirPasapalabra_(nombre, aciertos);
  return ok_();
}

function ok_() {
  return ContentService.createTextOutput(JSON.stringify({ok:true}))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---- escritura en la planilla de registro del curso -----------------------
// Columna "JUEGO EDUCATIVO: PASAPALABRA": nota por tramos de aciertos
// acumulados: 10-19 → 1, 20-29 → 2, ..., 100+ → 10.

function tokensReg_(t) {
  const n = norm_(t);
  return n ? n.split(" ") : [];
}
function distReg_(a, b) { // Levenshtein acotado
  if (Math.abs(a.length - b.length) > 2) return 99;
  let prev = [], cur = [];
  for (let j = 0; j <= b.length; j++) prev[j] = j;
  for (let i = 1; i <= a.length; i++) {
    cur[0] = i;
    for (let j = 1; j <= b.length; j++)
      cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1,
        prev[j - 1] + (a.charAt(i - 1) === b.charAt(j - 1) ? 0 : 1));
    const tmp = prev; prev = cur; cur = tmp;
  }
  return prev[b.length];
}
function nivelReg_(filaTok, t) {
  if (filaTok === t || filaTok.indexOf(t) === 0 || t.indexOf(filaTok) === 0 ||
      filaTok.indexOf(t) >= 0) return 2;
  const d = distReg_(filaTok, t);
  const tol = Math.max(filaTok.length, t.length) >= 7 ? 2 : 1;
  return d <= tol ? 1 : 0;
}
function encontrarRegistro_() {
  const ss = SpreadsheetApp.openById(REGISTRO_ID);
  const hojas = HOJA_REGISTRO ? [ss.getSheetByName(HOJA_REGISTRO)] : ss.getSheets();
  for (const h of hojas) {
    if (!h) continue;
    const datos = h.getDataRange().getValues();
    for (let i = 0; i < Math.min(datos.length, 15); i++) {
      let colPasap = -1, colApe = -1, colNom = -1;
      for (let j = 0; j < datos[i].length; j++) {
        const t = norm_(datos[i][j]);
        if (t.indexOf("pasapalabra") >= 0) colPasap = j;
        else if (t.indexOf("apellido") >= 0) colApe = j;
        else if (t.indexOf("nombre") >= 0) colNom = j;
      }
      if (colPasap >= 0) {
        // Encabezado de dos filas: si la columna de estudiantes
        // (apellido/nombre) está en la fila siguiente, mirarla.
        let filaEnc = i;
        if (colApe < 0 && colNom < 0 && i + 1 < datos.length) {
          for (let j2 = 0; j2 < datos[i + 1].length; j2++) {
            const t2 = norm_(datos[i + 1][j2]);
            if (t2.indexOf("apellido") >= 0) colApe = j2;
            else if (t2.indexOf("nombre") >= 0) colNom = j2;
          }
          if (colApe >= 0 || colNom >= 0) filaEnc = i + 1;
        }
        return { h, filaEnc, filaEncGrados: i, datos, colPasap, colApe, colNom };
      }
    }
  }
  return null;
}
function filaDeEstudiante_(reg, nombreCompleto) {
  const buscado = tokensReg_(nombreCompleto);
  if (!buscado.length) return -1;
  let mejor = -1, mejorScore = 0, empate = false;
  for (let r = reg.filaEnc + 1; r < reg.datos.length; r++) {
    let filaTxt = "";
    if (reg.colApe >= 0) filaTxt += " " + reg.datos[r][reg.colApe];
    if (reg.colNom >= 0) filaTxt += " " + reg.datos[r][reg.colNom];
    if (reg.colApe < 0 && reg.colNom < 0) return -1;
    const filaTokens = tokensReg_(filaTxt);
    if (!filaTokens.length) continue;
    let score = 0, ok = true;
    for (const tk of buscado) {
      let n = 0;
      for (const ft of filaTokens) n = Math.max(n, nivelReg_(ft, tk));
      if (n === 0) { ok = false; break; }
      score += n;
    }
    if (!ok) continue;
    if (score > mejorScore) { mejor = r; mejorScore = score; empate = false; }
    else if (score === mejorScore) empate = true;
  }
  return (mejor < 0 || empate) ? -2 : mejor;
}
function escribirPasapalabra_(nombre, aciertosAcum) {
  try {
    if (!REGISTRO_ID || REGISTRO_ID.indexOf("PEGAR") === 0) return;
    // 1) mantener el resumen en la planilla de registro (lo lee también el quiz)
    const ss = SpreadsheetApp.openById(REGISTRO_ID);
    let hp = ss.getSheetByName("PasapalabraResumen");
    if (!hp) { hp = ss.insertSheet("PasapalabraResumen"); hp.appendRow(["nombre","aciertos"]); }
    const prev = hp.getDataRange().getValues();
    let escrito = false;
    for (let i = 1; i < prev.length; i++) {
      if (norm_(prev[i][0]) === norm_(nombre)) {
        hp.getRange(i + 1, 2).setValue(+aciertosAcum || 0);
        escrito = true; break;
      }
    }
    if (!escrito) hp.appendRow([nombre, +aciertosAcum || 0]);
    // 2) nota por tramos en la columna PASAPALABRA
    const reg = encontrarRegistro_();
    if (reg) {
      const f = filaDeEstudiante_(reg, nombre);
      if (f >= 0) {
        const nota = Math.min(10, Math.floor((+aciertosAcum || 0) / 10));
        reg.h.getRange(f + 1, reg.colPasap + 1).setValue(nota);
      }
    }
    // 3) top 5 → 10 en MEJOR DESEMPEÑO (unión con el quiz)
    actualizarMejorDesempeno_(ss);
  } catch (e) { /* no cortar el registro de la partida si falla la planilla */ }
}

// MEJOR DESEMPEÑO DE ACTIVIDADES VIRTUALES: un 10 para quien esté en el
// top 5 de cualquiera de las dos apps (pasapalabra: aciertos acumulados;
// quiz: mejor intento por aciertos y tiempo). Si está en ambas, un solo 10;
// si sale del top 5, se borra.
function quizTop5_(ss) {
  const h = ss.getSheetByName("Intentos");
  if (!h) return [];
  const datos = h.getDataRange().getValues();
  const mejores = {};
  for (let i = 1; i < datos.length; i++) {
    const v = datos[i];
    if (v[7] !== "entregado") continue;
    const clave = (String(v[3]) + "|" + String(v[4])).toLowerCase();
    const corr = Number(v[11] || 0), seg = v[10] === "" ? 999999 : v[10];
    if (!mejores[clave] || corr > mejores[clave].correctas ||
        (corr === mejores[clave].correctas && seg < mejores[clave].segundos))
      mejores[clave] = { nombre: (String(v[3]) + " " + String(v[4])).trim(), correctas: corr, segundos: seg };
  }
  const lista = Object.values(mejores);
  lista.sort((x, y) => y.correctas - x.correctas || x.segundos - y.segundos);
  return lista;
}
function pasapResumen_(ss) {
  const h = ss.getSheetByName("PasapalabraResumen");
  if (!h) return [];
  const datos = h.getDataRange().getValues();
  const lista = [];
  for (let i = 1; i < datos.length; i++)
    lista.push({ nombre: String(datos[i][0] || ""), aciertos: Number(datos[i][1] || 0) });
  lista.sort((a, b) => b.aciertos - a.aciertos || (a.nombre < b.nombre ? -1 : a.nombre > b.nombre ? 1 : 0));
  return lista;
}
function encontrarDesemp_(reg) {
  let col = -1;
  const filaH = reg.filaEncGrados !== undefined ? reg.filaEncGrados : reg.filaEnc;
  for (let j = 0; j < reg.datos[filaH].length; j++) {
    const t = norm_(reg.datos[filaH][j]);
    if (t.indexOf("mejor") === 0 && t.indexOf("desempen") >= 0) { col = j; break; }
  }
  return col;
}
function actualizarMejorDesempeno_(ss) {
  const reg = encontrarRegistro_();
  if (!reg) return;
  const colDesemp = encontrarDesemp_(reg);
  if (colDesemp < 0) return;
  const merecen = {};
  quizTop5_(ss).slice(0, 5).forEach(e => {
    const f = filaDeEstudiante_(reg, e.nombre);
    if (f >= 0) merecen[f] = 1;
  });
  pasapResumen_(ss).slice(0, 5).forEach(e => {
    const f = filaDeEstudiante_(reg, e.nombre);
    if (f >= 0) merecen[f] = 1;
  });
  const activas = {};
  quizTop5_(ss).forEach(e => {
    const f = filaDeEstudiante_(reg, e.nombre);
    if (f >= 0) activas[f] = 1;
  });
  pasapResumen_(ss).forEach(e => {
    const f = filaDeEstudiante_(reg, e.nombre);
    if (f >= 0) activas[f] = 1;
  });
  for (const fr in activas) {
    const fila = Number(fr);
    const celda = reg.h.getRange(fila + 1, colDesemp + 1);
    const actual = reg.datos[fila][colDesemp];
    const actN = actual === "" || actual === null ? -1 : Number(String(actual).replace(",", "."));
    if (merecen[fila]) { if (actN !== 10) celda.setValue(10); }
    else if (actN === 10) celda.setValue("");
  }
}

// Sincroniza de una vez a la planilla de notas TODOS los aciertos ya
// acumulados en "resumen" (para los que jugaron antes de pegar este código).
// Ejecutarla UNA vez desde el editor (▶ Ejecutar) y listo.
function sincronizar() {
  const filas = resumen_().getDataRange().getValues().slice(1);
  let ok = 0;
  const sinFila = [];
  const reg = encontrarRegistro_();
  for (const r of filas) {
    const nombre = String(r[0] || "");
    if (!nombre) continue;
    escribirPasapalabra_(nombre, +r[5] || 0);
    if (reg) {
      const f = filaDeEstudiante_(reg, nombre);
      if (f >= 0) ok++;
      else sinFila.push(nombre);
    } else {
      ok++;
    }
  }
  const msg = "Sincronizados " + ok + " de " + (ok + sinFila.length) +
    (sinFila.length ? " | SIN FILA EN LA PLANILLA: " + sinFila.join(", ") : "");
  Logger.log(msg);
  return msg;
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
