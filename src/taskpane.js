/* global Office, LZ */
Office.onReady(function (info) {
  var btn = document.getElementById("btn01");
  var out = document.getElementById("out01");
  document.getElementById("ver").textContent = "v" + LZ.version;
  if (info.host !== Office.HostType.Word) {
    out.textContent = "Dieses Add-In läuft nur in Word.";
    return;
  }
  btn.disabled = false;
  btn.addEventListener("click", function () {
    btn.disabled = true;
    out.className = "out";
    out.textContent = "Läuft…";
    LZ.tiefstellen().then(function (n) {
      out.className = "out ok";
      out.textContent = n === 0
        ? "Nichts gefunden. Schreib z. B. 1010(2), ohne Leerzeichen vor der Klammer."
        : n + " Basis-Angabe" + (n === 1 ? "" : "n") + " tiefgestellt ✔";
    }).catch(function (e) {
      out.className = "out err";
      out.textContent = "Fehler: " + (e && e.message ? e.message : e);
    }).then(function () { btn.disabled = false; });
  });
});
