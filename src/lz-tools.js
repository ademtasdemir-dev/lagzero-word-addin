/* LAG|ZERØ Word Tools – gemeinsame Logik
 * #01 Zahlensysteme: Basis tiefstellen
 *   1010(2) -> 1010 + tiefgestellte 2,  FF(16) -> FF + tiefgestellte 16
 *   Nur wenn direkt vor "(" eine Ziffer 0-9 oder A-F steht und in der Klammer 1-2 Ziffern sind.
 *   "Aufgabe (2)" (Leerzeichen), "(eier)" und "11(100)" bleiben unverändert.
 */
/* global Word */
var LZ = (function () {
  // Wildcard-Muster ohne {n,m} – das bricht in deutschem Word (dort {n;m}).
  var WILDCARD = "[0-9A-F]\\([0-9]@\\)";
  var VALID = /^[0-9A-F]\((\d{1,2})\)$/;

  function tiefstellen() {
    return Word.run(function (context) {
      var found = context.document.body.search(WILDCARD, { matchWildcards: true });
      found.load("items/text");
      return context.sync().then(function () {
        var hits = [];
        found.items.forEach(function (r) {
          if (VALID.test(r.text)) hits.push(r);
        });
        if (hits.length === 0) return 0;

        var parts = hits.map(function (r) {
          var open = r.search("(", { matchCase: true }).getFirst();
          var close = r.search(")", { matchCase: true }).getFirst();
          var digits = open.getRange("After").expandTo(close.getRange("Before"));
          return { open: open, close: close, digits: digits };
        });
        parts.forEach(function (p) { p.digits.font.subscript = true; });
        return context.sync().then(function () {
          parts.forEach(function (p) {
            p.close.delete();
            p.open.delete();
          });
          return context.sync().then(function () { return hits.length; });
        });
      });
    });
  }

  return { tiefstellen: tiefstellen, version: "1.0.1" };
})();
