# LAG|ZERØ Word Add-In: Projektnotizen für Claude

## Was das ist
Office-Web-Add-In (Office.js, XML-Manifest) für Word. Eigener Reiter „LAG|ZERØ" + Seitenleiste.
Gehostet über GitHub Pages: https://sikkkilaci.github.io/lagzero-word-addin/
Das ist das Zuhause der **LAG|ZERØ Word Add-In Familie**: Neue Tools werden als weitere Gruppe/Button im
Reiter und als weitere Karte in der Seitenleiste ergänzt. Die Logik kommt nach `src/lz-tools.js`.

## Regeln für neue Tools
- Nummerierung: #01, #02, …; Funktionsnamen `lz02…`, IDs `LZ.Btn.…`, Icons `assets/lz02-16/32/64/80.png`
- Jeder Ribbon-Button mit `ExecuteFunction` braucht eine globale Funktion in `src/commands.js`, die `event.completed()` aufruft
- Nach Manifest-Änderungen: `Version` hochzählen und `npx office-addin-manifest validate manifest.xml`
- Alles in einem `Word.run`-Batch, damit Strg+Z möglichst wenig Schritte braucht

## Tool #01 Basis tiefstellen
- Wildcard `[0-9A-F]\([0-9]@\)`, danach JS-Regex `^[0-9A-F]\((\d{1,2})\)$`
- Nicht `{1,2}` im Wildcard benutzen: Deutsches Word erwartet `{1;2}` und wirft sonst einen Fehler
- `Aufgabe (2)`, `(eier)`, `11(100)` bleiben unverändert

## Design
Schwarz `#0b0b0d`, Panel `#141418`, Silber `#c9ccd2`, Rot `#e11d27`. Header „LAG|ZERØ" + Badge VERIFIED,
Slogan „TESTED. MEASURED. OPTIMIZED." Discord: https://discord.gg/wUFMj2dGX8

## Store (später)
Partner-Center-Konto (kostenlos), Einreichung mit privacy.html + support.html, Icon `assets/logo-300.png`.

## Arbeitsweise mit Adem
Direkt, einfache Sprache, wenig Fachchinesisch. Anleitungen als HTML im LAG|ZERØ-Design + ZIP.
Bei Dateisammlungen immer ein anklickbares HTML-Inhaltsverzeichnis.
