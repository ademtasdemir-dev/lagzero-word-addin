/* LAG|ZERØ Word Tools – gemeinsame Logik
 * #01 Zahlensysteme: Basis tiefstellen
 *   1010(2) -> 1010 + tiefgestellte 2,  FF(16) -> FF + tiefgestellte 16
 *   Nur wenn direkt vor "(" eine Ziffer 0-9 oder A-F steht und in der Klammer 1-2 Ziffern sind.
 *   "Aufgabe (2)" (Leerzeichen), "(eier)" und "11(100)" bleiben unverändert.
 *   Wirkt im Haupttext, in Tabellen sowie in allen Kopf- und Fußzeilen.
 */
/* global Word */
var LZ = (function () {
  // Wildcard-Muster ohne {n,m} – das bricht in deutschem Word (dort {n;m}).
  var WILDCARD = "[0-9A-F]\\([0-9]@\\)";
  var VALID = /^[0-9A-F]\((\d{1,2})\)$/;

  // Alle Textbereiche: Haupttext + Kopf-/Fußzeilen jeder Sektion
  function collectBodies(context) {
    var sections = context.document.sections;
    sections.load("items");
    return context.sync().then(function () {
      var bodies = [context.document.body];
      var types = ["Primary", "FirstPage", "EvenPages"];
      sections.items.forEach(function (s) {
        types.forEach(function (t) {
          bodies.push(s.getHeader(t));
          bodies.push(s.getFooter(t));
        });
      });
      return bodies;
    });
  }

  function processBody(context, body) {
    var found = body.search(WILDCARD, { matchWildcards: true });
    found.load("items/text");
    return context.sync().then(function () {
      var hits = found.items.filter(function (r) { return VALID.test(r.text); });
      if (hits.length === 0) return 0;
      var parts = hits.map(function (r) {
        var open = r.search("(", { matchCase: true }).getFirst();
        var close = r.search(")", { matchCase: true }).getFirst();
        // "End" von "(" bis "Start" von ")" = nur die Ziffern
        // (getRange("After"/"Before") wirft in Word Desktop InvalidArgument)
        var digits = open.getRange("End").expandTo(close.getRange("Start"));
        return { open: open, close: close, digits: digits };
      });
      parts.forEach(function (p) { p.digits.font.subscript = true; });
      return context.sync().then(function () {
        parts.forEach(function (p) { p.close.delete(); p.open.delete(); });
        return context.sync().then(function () { return hits.length; });
      });
    });
  }

  function tiefstellen() {
    return Word.run(function (context) {
      return collectBodies(context).then(function (bodies) {
        var total = 0;
        return bodies.reduce(function (chain, b, i) {
          return chain.then(function () {
            return processBody(context, b)
              .then(function (n) { total += n; })
              .catch(function (e) {
                if (i === 0) throw e; // Fehler im Haupttext weitergeben
                /* leere oder nicht vorhandene Kopf-/Fußzeile ignorieren */
              });
          });
        }, Promise.resolve()).then(function () { return total; });
      });
    });
  }

  return { tiefstellen: tiefstellen, version: "1.0.4" };
})();
