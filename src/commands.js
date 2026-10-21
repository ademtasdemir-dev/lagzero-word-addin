/* LAG|ZERØ – Ribbon-Befehle (ohne Seitenleiste) */
/* global Office, Word, LZ */
var LZ_DEBUG = true; // TEMP: Fehler ins Dokument schreiben, damit man sie ohne Seitenleiste sieht

function lzReport(msg) {
  if (!LZ_DEBUG) return Promise.resolve();
  return Word.run(function (ctx) {
    ctx.document.body.insertParagraph("[LAG|ZERØ DEBUG] " + msg, "End");
    return ctx.sync();
  }).catch(function () {});
}

function lz01Tiefstellen(event) {
  var done = function () { try { event.completed(); } catch (e) {} };
  try {
    LZ.tiefstellen()
      .then(function (n) { return lzReport("OK: " + n + " ersetzt (v" + LZ.version + ")"); })
      .catch(function (e) {
        var m = (e && (e.code ? e.code + ": " : "") + (e.message || e)) + (e && e.debugInfo ? " | " + JSON.stringify(e.debugInfo) : "");
        return lzReport("FEHLER: " + m);
      })
      .then(done, done);
  } catch (e) {
    lzReport("FEHLER (sync): " + (e && e.message ? e.message : e)).then(done, done);
  }
}

Office.onReady(function () {
  if (Office.actions && Office.actions.associate) {
    try { Office.actions.associate("lz01Tiefstellen", lz01Tiefstellen); } catch (e) {}
  }
});
window.lz01Tiefstellen = lz01Tiefstellen;
