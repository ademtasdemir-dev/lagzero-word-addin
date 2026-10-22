/* global Office, LZ */
Office.onReady(function (info) {
  document.getElementById("ver").textContent = "v" + LZ.version;
  function wire(btnId, outId, fn, fmt) {
    var btn = document.getElementById(btnId), out = document.getElementById(outId);
    if (info.host !== Office.HostType.Word) { out.textContent = "Dieses Add-In läuft nur in Word."; return; }
    btn.disabled = false;
    btn.addEventListener("click", function () {
      btn.disabled = true; out.className = "out"; out.textContent = "Läuft…";
      fn().then(function (r) { out.className = "out ok"; out.textContent = fmt(r); })
        .catch(function (e) { out.className = "out err"; out.textContent = "Fehler: " + (e && e.message ? e.message : e); })
        .then(function () { btn.disabled = false; });
    });
  }
  function pl(n, one, many) { return n + " " + (n === 1 ? one : many); }
  wire("btn01", "out01", LZ.tiefstellen, function (n) {
    return n === 0 ? "Nichts gefunden. Schreib z. B. 1010(2), ohne Leerzeichen vor der Klammer." : pl(n, "Basis-Angabe", "Basis-Angaben") + " tiefgestellt ✔";
  });
  wire("btn02", "out02", LZ.hochstellen, function (n) {
    return n === 0 ? "Nichts gefunden. Schreib z. B. 2^10." : pl(n, "Potenz", "Potenzen") + " hochgestellt ✔";
  });
  wire("btnAll", "outAll", LZ.alles, function (r) {
    return pl(r.tief, "Basis", "Basen") + " tief, " + pl(r.hoch, "Potenz", "Potenzen") + " hoch ✔";
  });
});
