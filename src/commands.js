/* LAG|ZERØ – Ribbon-Befehle (ohne Seitenleiste) */
/* global Office, LZ */
Office.onReady(function () {});

function lz01Tiefstellen(event) {
  LZ.tiefstellen()
    .catch(function (e) { console.error("LAG|ZERØ #01:", e); })
    .then(function () { event.completed(); });
}

// Für die Ribbon-Funktion global bereitstellen
if (typeof window !== "undefined") window.lz01Tiefstellen = lz01Tiefstellen;
if (Office.actions && Office.actions.associate) {
  Office.onReady(function () { Office.actions.associate("lz01Tiefstellen", lz01Tiefstellen); });
}
