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

## Verteilung (Stand v1.0.4)
- Paket `download/LAGZERO_Word_Tools.zip`: ANLEITUNG.html, Installieren.bat, Deinstallieren.bat, lagzero-word-addin.xml (= manifest.xml), Testdokument.docx, media/
- Installieren.bat: kopiert Manifest nach `%LOCALAPPDATA%\LAGZERO\WordKatalog` und trägt ihn als
  TrustedCatalog `{212c4ebf-75af-4458-b496-34fd6f3ecf0b}` mit UNC-Pfad `\\localhost\C$\...` ein (Flags=1).
  Danach in Word: Add-Ins → Weitere Add-Ins → GETEILTER ORDNER → Hinzufügen. Parameter `/quiet` = ohne pause.
- Privates Microsoft-Konto hat in Word Desktop KEIN "Mein Add-In hochladen" → deshalb geteilter Ordner.
- Word cached Add-In-Dateien sehr hartnäckig: bei Updates `?v=` hochzählen, Version im Manifest erhöhen;
  beim Entwickeln half nur eine neue `<Id>` (dann neu hinzufügen). Manifest im Paket immer mit neu bauen.
- Debug: in `src/commands.js` `LZ_DEBUG = true` schreibt Ergebnis/Fehler als Zeile ins Dokument.
