# Prompt für Claude Code: "Use Case Pulse – Overview"-Slide

Baue eine einzelne PowerPoint-Slide (16:9, 13.333 × 7.5 inch), die eine Übersicht über
mehrere Catena-X Use Cases zeigt – eine Zeile pro Use Case. Die Slide muss **optisch
identisch** mit der angehängten/referenzierten "Use Case Pulse"-Einzel-Slide aussehen
(gleiche Farben, Schriften, Formensprache, Ampelfarben, Hintergrund-/Vordergrundfarben).
Hänge zur Referenz eine der echten ausgefüllten Use Case PULSE-Decks an
(z. B. `UseCase_Qualtiy_CW40_26.pptx` oder `PCF_UseCase.pptx`) und lies die exakten
Farbwerte/Schriften direkt aus deren `theme1.xml` und `slide1.xml` aus, statt sie zu
schätzen.

## 1. Inhalt der Slide

**Kopfbereich:**
- Eyebrow: "USE CASE PULSE · CATENA-X"
- Headline: "Overview"
- Datumsbox oben rechts: "CW 40 / 2026"

**Tabelle / Raster, sortiert nach Overall Status: Escalation needed → Problem solving → On track**

Spalten: Use case | Overall status | Phase | Next milestone | Dimensions | Suppliers | (Decision-Flag, keine Kopfzeilen-Beschriftung)

| Use case (Abk.) | Overall status | Phase | Next milestone | Dimensions (Status · Name) | Suppliers | Decision |
|---|---|---|---|---|---|---|
| Quality (quality-cx) | 🔴 Escalation needed | Piloting | 🟡 Domain expertise aligned · 30 Sep 2026 | 🟢 External · CX Association, 🟡 Business, 🟢 Data Provisioning, 🟡 Technical, 🟡 Supplier Activation | 0 / 1 · 0% · resc. | 🚩 ¹ |
| Battery Passport (Batt Pass) | 🟡 Problem solving | Pilot | 🟢 Tech requirements done · 31 Oct 2026 | 🟡 External · CX Association, 🟡 Business, 🟢 Data Provisioning, 🟢 Technical, 🟢 Supplier Activation | 0 / 1 · 0% | – |
| Product Passes (pass-cx) | 🟡 Problem solving | Definition | 🟡 Klärung & Prio Produktpässe · 31 Oct 2026 | 🟡 External · CX Association, 🟡 Business, 🟢 Data Provisioning, 🟢 Technical, 🟢 Supplier Activation | 0 / 1 · 0% | – |
| PURIS (PURIS) | 🟡 Problem solving | Piloting | 🟡 Plan for existing challenges · 9 Oct 2026 | 🟢 External · CX Association, 🟡 Business, 🟡 Data Provisioning, 🟡 Technical, 🟡 Supplier Activation | 1 / 6 · 17% | 🚩 ² |
| Business Partner Data Mgmt (BPDM) | 🟡 Problem solving | Implementation | 🟢 Define BPDM target picture · 31 Oct 2026 | 🟡 External · CX Association, 🟡 Business, 🟢 Data Provisioning, 🟢 Technical, ⚪ Supplier Activation (not started) | 0 / 1 · 0% | – |
| Certificate Management (cert-mgmt) | 🟡 Problem solving | Scaling | 🟢 Official CX CCM release · Sep 2026 | 🟡 External · CX Association, 🟢 Business, 🟢 Data Provisioning, 🟢 Technical, 🟢 Supplier Activation | 72 / 100 · 72% | – |
| Product Carbon Footprint (pcf-cx) | 🟢 On track | Piloting | 🟢 Supplier activation pilot 1 · 15 Oct 2026 | 🟢 External · CX Association, 🟢 Business, 🟢 Data Provisioning, 🟢 Technical, 🟢 Supplier Activation | 0 / 1 · 0% | – |

**Dimensions-Layout:** pro Zeile 2 Dimensionen nebeneinander (External·CX Association + Business /
Data Provisioning + Technical), "Supplier Activation" als 5., optisch abgesetzte Dimension
darunter (volle Breite, da sie inhaltlich am engsten mit der Suppliers-Spalte zusammenhängt).
Jede Dimension ausgeschrieben, nicht abgekürzt.

**Footer unterhalb der Tabelle:** "Decisions required:" als fette Überschrift, darunter:
1. Quality: Rescope von 15 auf 1 Lieferant, bis Early Warning Production MVP live ist.
2. PURIS: Wie und wann wird über die WINGS-Connectivity-Rollout-Roadmap entschieden?

## 2. Stilanforderung – MUSS exakt wie die Use Case Pulse-Einzel-Slide aussehen

**Nicht als native PowerPoint-Tabelle bauen** – stattdessen mit einzelnen Shapes/Textboxen
(wie auf der Referenz-Slide), damit Status-Badges, Kreise und Karten pixelgenau wie im
Original aussehen, nicht wie eine generische Tabelle.

**Theme (aus `theme1.xml` der echten Decks, Theme-Name "Volkswagen Group 2023"):**
- `dk1`/Text dunkel: `000000`
- `lt1`/Hintergrund hell: `FFFFFF`
- `dk2` / `accent1` (dunkles Petrol): `002733`
- `accent2` (Teal): `008C82`
- `accent3` (helles Mint/Teal): `99D1CD`
- `accent4` (Blaugrau): `809399`
- `accent5` (helles Grau): `CCD3D6`
- `accent6` (Schiefer-Teal): `4C6870`

**Schriften:**
- Headline-Font: `The Group HEAD Light` (für große Titel wie "Overview")
- Fließtext-Font: `The Group TEXT` (für Fließtext, Labels, Zahlen)
- Badge-/Pill-Text (Status-Badges): `Segoe UI`, bold, ca. 8pt, weiß

**Ampelfarben (1:1 aus den echten Decks ausgelesen, nicht neu erfinden):**
- On Track: `53BE70` (grün)
- Problem Solving: `F0B135` (amber)
- Escalation needed: `E64343` (rot)
- Not Started: `CFD6DB` (hellgrau)
- Completed: `001E50` (dunkles Navy)

**Formensprache:**
- **Overall-Status-Badge** = gefüllte **Pille** (abgerundetes Rechteck, großer Radius),
  Vollfarbe je Ampelstatus, weißer fetter Text (Segoe UI) – **kein** einfacher Punkt+Text,
  so wie im "Status Overview"-Kasten der Referenz-Slide.
- **Dimensionen & Milestone-Timeline** = kleiner farbiger **Kreis** (Punkt) + ausgeschriebenes
  Label daneben (so wie im Original).
- Karten/Container: weiße, abgerundete Flächen auf hellem Seitenhintergrund (Seitenhintergrund
  ca. `F0F4F9`, Kartenhintergrund `F6F8FA`/`FFFFFF`), dünne helle Border (`E2E8EC`/`CFD6DB`).
- Keine Farbverläufe, keine Schatten, keine dekorativen Effekte – ausschließlich Flächenfarben,
  genau wie im VW-Corporate-Design der Referenz-Slide.
- Textfarben: Haupttext sehr dunkel (`292D31`/`151B21`), sekundärer Text `63707A`, gedämpfter
  Text/Captions `9AA6AD`.

**Layout-Vorbild:** Kopfzeile mit kleinem Eyebrow-Label + großer fetter Headline, dann Kacheln/
Karten für den Inhalt – exakt wie bei der einzelnen Use Case Pulse-Slide, nur eben als
Mehrzeilen-Übersicht statt Einzel-Use-Case-Ansicht.

## 3. Technischer Hinweis
Nach dem Bauen: `python scripts/office/validate.py <datei>.pptx --original <referenz>.pptx`
laufen lassen und die Slide als Bild rendern (LibreOffice → PDF → `pdftoppm`), um Overflow,
Zeilenhöhen und Farben visuell zu prüfen, bevor die Datei geliefert wird.
