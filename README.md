# LAG|ZERØ Tools für Microsoft Word

**TESTED. MEASURED. OPTIMIZED.**

Office-Add-In (Office.js) mit eigenem Reiter **LAG|ZERØ** und Seitenleiste in Word.

## Tools
| # | Tool | Was es macht |
|---|------|--------------|
| 01 | Basis tiefstellen | `1010(2)` → 1010₂, `FF(16)` → FF₁₆ im ganzen Dokument. `Aufgabe (2)` (mit Leerzeichen), `(eier)`, `11(100)` bleiben unverändert. |

## Installieren
Anleitung: **https://ademtasdemir-dev.github.io/lagzero-word-addin/**

Kurz: [`manifest.xml`](https://ademtasdemir-dev.github.io/lagzero-word-addin/manifest.xml) herunterladen → Word → Start → Add-Ins → Weitere Add-Ins → Meine Add-Ins → Benutzerdefiniertes Add-In hochladen.

## Aufbau
```
manifest.xml        Add-In-Beschreibung (Reiter, Buttons, URLs)
taskpane.html       Seitenleiste
commands.html       Ribbon-Befehle ohne Seitenleiste
src/lz-tools.js     Logik aller Tools
src/taskpane.*      Seitenleiste (JS/CSS)
src/site.css        Design der Webseiten
assets/             Icons
index.html          Startseite / Anleitung
support.html        Hilfe
privacy.html        Datenschutz
```
Gehostet über GitHub Pages (Branch `main`, Ordner `/`).

---
LAG|ZERØ · [Discord](https://discord.gg/wUFMj2dGX8)
