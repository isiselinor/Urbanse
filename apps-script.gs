/**
 * URBANSE — Landing ↔ Google Sheet
 * ------------------------------------------------------------
 * Este script hace dos cosas:
 *  1) GET  → entrega a la landing las pestañas Fechas, Videos, Merch y Redes (JSON).
 *  2) POST → guarda cada solicitud de booking en la pestaña "Booking" (con aviso por email)
 *            y cada pedido "Pídelo en tu ciudad" de los fans en la pestaña "Ciudades pedidas".
 *
 * INSTALACIÓN (una sola vez):
 *  1. En el Sheet: Extensiones → Apps Script. Borra lo que haya y pega este archivo.
 *  2. Ejecuta la función `setup` (botón ▶). Acepta los permisos.
 *     Crea las pestañas con sus columnas si aún no existen.
 *  3. Implementar → Nueva implementación → Tipo: Aplicación web.
 *       Ejecutar como: Yo · Quién tiene acceso: Cualquier persona.
 *  4. Copia la URL que termina en /exec y pégala en index.html → CONFIG.SHEET_API_URL.
 *  Si cambias el script después: Implementar → Gestionar implementaciones → Editar → Nueva versión
 *  (así la URL no cambia).
 */

const NOTIFY_EMAIL = "Buenosaifers@gmail.com"; // deja "" para no recibir avisos por email
const CACHE_SECONDS = 300;                       // la landing se actualiza como máximo cada 5 min

/* Bandsintown: las fechas cargadas allí se suman a las de la pestaña Fechas (una sola lista).
   BIT_APP_ID: se obtiene en Bandsintown for Artists › Settings › General › API (o pidiéndolo a Bandsintown).
   Si una fecha está en los dos lados (mismo día y ciudad), manda la fila del Sheet. */
const BIT_ARTIST = "Urbanse";   // nombre exacto del artista en Bandsintown
const BIT_APP_ID = "";          // app_id de Bandsintown (va en la copia pegada en Apps Script, no en el repositorio)

const TABS = {
  Fechas:  ["fecha", "hora", "ciudad", "pais", "venue", "evento", "link_entradas", "estado", "publicar"],
  Videos:  ["titulo", "youtube_url", "tipo", "fecha", "descripcion", "link_escuchar", "destacado", "publicar", "ft", "verso"],
  Merch:   ["producto", "detalle", "tipo", "precio", "imagen_url", "imagen_url_2", "talles", "color", "tinta", "link_compra", "estado", "publicar", "lanzamiento", "stock"],
  Redes:   ["red", "url"],
  Booking: ["recibido", "motivo", "nombre", "organizacion", "email", "whatsapp", "ciudad", "pais", "tipo_evento",
            "formato", "fecha_tentativa", "venue_capacidad", "presupuesto", "mensaje", "origen", "estado_seguimiento", "notas"],
  "Ciudades pedidas": ["recibido", "ciudad", "pais", "email", "origen"],
  Suscriptores: ["recibido", "email", "ciudad", "intereses", "origen", "estado"]
};

const REDES_DEFAULT = ["instagram", "tiktok", "youtube", "spotify", "x", "facebook"];

/* ---------- Configuración inicial ---------- */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(TABS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const headers = TABS[name];
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers])
        .setFontWeight("bold").setBackground("#D71D22").setFontColor("#000000");
      sh.setFrozenRows(1);
    } else {
      // agrega al final las columnas nuevas que falten, sin tocar las existentes
      const have = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
      headers.filter(h => have.indexOf(h) < 0).forEach(h => {
        sh.getRange(1, sh.getLastColumn() + 1).setValue(h)
          .setFontWeight("bold").setBackground("#D71D22").setFontColor("#000000");
      });
    }
  });
  // Redes: precarga la lista de redes para que solo pegues las URLs
  const redes = ss.getSheetByName("Redes");
  if (redes.getLastRow() === 1) {
    redes.getRange(2, 1, REDES_DEFAULT.length, 2)
      .setValues(REDES_DEFAULT.map(r => [r, ""]));
  }
  // Validaciones útiles
  const siNo = SpreadsheetApp.newDataValidation().requireValueInList(["si", "no"], true).build();
  ss.getSheetByName("Fechas").getRange("I2:I").setDataValidation(siNo);
  ss.getSheetByName("Fechas").getRange("H2:H").setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(["Próximamente", "A la venta", "Últimas entradas", "Agotado"], true).build());
  ss.getSheetByName("Fechas").getRange("A2:A").setNumberFormat("yyyy-mm-dd");
  ss.getSheetByName("Videos").getRange("G2:H").setDataValidation(siNo);
  ss.getSheetByName("Videos").getRange("C2:C").setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(["Videoclip", "Sesión en vivo", "Película", "Nuevo lanzamiento", "Colaboración"], true).build());
  ss.getSheetByName("Merch").getRange("L2:L").setDataValidation(siNo);
  ss.getSheetByName("Merch").getRange("N2:N").setDataValidation(
    SpreadsheetApp.newDataValidation().requireNumberGreaterThanOrEqualTo(0).setAllowInvalid(false).build());
  ss.getSheetByName("Merch").getRange("C2:C").setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(["remera", "buzo", "cd", "vinilo", "gorra", "tote", "otro"], true).build());
  ss.getSheetByName("Merch").getRange("K2:K").setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(["Disponible", "Pronto", "Agotado"], true).build());
  ss.getSheetByName("Suscriptores").getRange("F2:F").setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(["Activo", "Baja"], true).build());
  ss.getSheetByName("Booking").getRange("O2:O").setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(["Nuevo", "Contactado", "Negociando", "Confirmado", "Descartado"], true).build());
  const def = ss.getSheetByName("Sheet1") || ss.getSheetByName("Hoja 1");
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(def);
}

/* ---------- GET: datos para la landing ---------- */
function doGet() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get("data");
  if (hit) return json_(hit);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const redes = {};
  rows_(ss, "Redes").forEach(r => { if (r.red && r.url) redes[String(r.red).toLowerCase().trim()] = r.url; });

  const out = JSON.stringify({
    ok: true,
    actualizado: new Date().toISOString(),
    fechas: fechas_(ss),
    videos: rows_(ss, "Videos"),
    merch:  rows_(ss, "Merch"),
    redes:  redes,
    demanda: demanda_(ss)   // solo ciudad y cantidad; nunca se publican emails
  });
  cache.put("data", out, CACHE_SECONDS);
  return json_(out);
}

/* ---------- POST: booking y pedidos de fans ---------- */
function doPost(e) {
  const p = (e && e.parameter) || {};
  if (p.website) return json_(JSON.stringify({ ok: true })); // anti-spam (campo oculto)
  if (p.tipo === "demanda") return demandaPost_(p);
  if (p.tipo === "suscripcion") return suscripcionPost_(p);
  if (!p.nombre || !p.email) return json_(JSON.stringify({ ok: false, error: "faltan datos" }));
  if (!p.motivo) p.motivo = "Solicitar fecha";

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Booking");
    const row = TABS.Booking.map(h => {
      if (h === "recibido") return new Date();
      if (h === "estado_seguimiento") return "Nuevo";
      if (h === "notas") return "";
      return clean_(p[h]);
    });
    sh.appendRow(row);
  } finally {
    lock.releaseLock();
  }

  if (NOTIFY_EMAIL) {
    const body = TABS.Booking.filter(h => p[h]).map(h => h.replace(/_/g, " ").toUpperCase() + ": " + clean_(p[h])).join("\n");
    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      replyTo: clean_(p.email),
      subject: clean_(p.motivo) + " — Urbanse" + (p.ciudad ? " — " + clean_(p.ciudad) + ", " + clean_(p.pais) : ""),
      body: "Llegó una nueva solicitud desde la landing de Urbanse:\n\n" + body +
            "\n\nVer todas: " + SpreadsheetApp.getActiveSpreadsheet().getUrl()
    });
  }
  return json_(JSON.stringify({ ok: true }));
}

/* "Pídelo en tu ciudad": una fila por pedido en la pestaña "Ciudades pedidas" */
function demandaPost_(p) {
  if (!p.ciudad || !p.pais) return json_(JSON.stringify({ ok: false, error: "faltan datos" }));
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Ciudades pedidas")
      .appendRow([new Date(), clean_(p.ciudad).slice(0, 80), clean_(p.pais).slice(0, 60), clean_(p.email).slice(0, 120), clean_(p.origen)]);
  } finally {
    lock.releaseLock();
  }
  return json_(JSON.stringify({ ok: true }));
}

/* "Entra a la lista": un email por fila en la pestaña Suscriptores (no duplica) */
function suscripcionPost_(p) {
  const email = clean_(p.email).toLowerCase().slice(0, 120);
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return json_(JSON.stringify({ ok: false, error: "email" }));
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Suscriptores");
    const last = sh.getLastRow();
    const ya = last > 1 && sh.getRange(2, 2, last - 1, 1).getValues().some(r => String(r[0]).toLowerCase() === email);
    if (!ya) sh.appendRow([new Date(), email, clean_(p.ciudad).slice(0, 80), clean_(p.intereses).slice(0, 80), clean_(p.origen), "Activo"]);
  } finally {
    lock.releaseLock();
  }
  return json_(JSON.stringify({ ok: true }));
}

/* Suma pedidos por ciudad (ignora mayúsculas y tildes) y devuelve las 8 con más pedidos */
function demanda_(ss) {
  const norm = s => String(s || "").trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  const count = {};
  rows_(ss, "Ciudades pedidas").forEach(r => {
    if (!r.ciudad) return;
    const k = norm(r.ciudad) + "|" + norm(r.pais);
    if (!count[k]) count[k] = { ciudad: String(r.ciudad).trim(), pais: String(r.pais || "").trim(), n: 0 };
    count[k].n++;
  });
  return Object.keys(count).map(k => count[k]).sort((a, b) => b.n - a.n).slice(0, 8);
}

/* Fechas del Sheet + Bandsintown, sin duplicados */
function fechas_(ss) {
  const sheet = rows_(ss, "Fechas").map(r => Object.assign({ fuente: "sheet" }, r));
  const bit = bandsintown_();
  const norm = s => String(s || "").trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  const ymd = v => { const d = v instanceof Date ? v : new Date(v); return isNaN(d) ? String(v).slice(0, 10) : Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd"); };
  const ya = {}; sheet.forEach(r => { ya[ymd(r.fecha) + "|" + norm(r.ciudad)] = true; });
  return sheet.concat(bit.filter(e => !ya[e.fecha + "|" + norm(e.ciudad)]));
}

const PAISES_ = { "Spain": "España", "Mexico": "México", "Peru": "Perú", "United States": "Estados Unidos", "Brazil": "Brasil", "Germany": "Alemania", "France": "Francia", "Italy": "Italia", "United Kingdom": "Reino Unido", "Dominican Republic": "República Dominicana", "Panama": "Panamá" };

function bandsintown_() {
  if (!BIT_APP_ID || !BIT_ARTIST) return [];
  try {
    const url = "https://rest.bandsintown.com/artists/" + encodeURIComponent(BIT_ARTIST) + "/events?date=upcoming&app_id=" + encodeURIComponent(BIT_APP_ID);
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { accept: "application/json" } });
    if (res.getResponseCode() !== 200) return [];
    const list = JSON.parse(res.getContentText());
    if (!Array.isArray(list)) return [];
    return list.map(e => {
      const v = e.venue || {}, dt = String(e.datetime || ""), offer = (e.offers || []).find(o => o && o.url) || {};
      const status = String(offer.status || "").toLowerCase();
      return {
        fuente: "bandsintown",
        fecha: dt.slice(0, 10),
        hora: dt.slice(11, 16),
        ciudad: String(v.city || "").trim(),
        pais: PAISES_[v.country] || v.country || "",
        venue: String(v.name || "").trim() === String(e.title || "").trim() ? "" : String(v.name || "").trim(),
        evento: String(e.title || "").trim(),
        link_entradas: offer.url || e.url || "",
        estado: (e.sold_out || status === "sold out") ? "Agotado" : offer.url ? "A la venta" : "Próximamente",
        publicar: "si"
      };
    }).filter(e => e.fecha);
  } catch (err) {
    return [];
  }
}

/* Borra la caché al editar el Sheet, para que los cambios se vean enseguida */
function onEdit() { CacheService.getScriptCache().remove("data"); }

/* ---------- utilidades ---------- */
function rows_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getDataRange().getValues();
  const headers = values.shift().map(h => String(h).trim());
  const tz = ss.getSpreadsheetTimeZone();
  return values
    .filter(r => r.some(c => c !== ""))
    .map(r => {
      const o = {};
      headers.forEach((h, i) => {
        let v = r[i];
        if (v instanceof Date) v = (h === "hora") ? Utilities.formatDate(v, tz, "HH:mm") : Utilities.formatDate(v, tz, "yyyy-MM-dd");
        o[h] = v;
      });
      return o;
    });
}
function clean_(v) {
  let s = String(v == null ? "" : v).slice(0, 2000).trim();
  if (/^[=+\-@]/.test(s)) s = "'" + s; // evita fórmulas inyectadas en el Sheet
  return s;
}
function json_(str) {
  return ContentService.createTextOutput(str).setMimeType(ContentService.MimeType.JSON);
}
