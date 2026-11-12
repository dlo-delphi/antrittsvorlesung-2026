# Antrittsvorlesung Hochschule Mainz – 12.11.2026

**Fliegen, wenn die Systeme ausfallen**
Was Cyber-Resilienz von der Luftfahrt lernen kann – und was nicht

48 Minuten Vortrag, danach Diskussion (insgesamt eine Stunde).
Requisiten: WinCtrl-A320-Fahrwerkpanel und FCU, Powerbank.

## PowerPoint bauen

```bash
pip install python-pptx pillow matplotlib
python scripts/diagramme.py bilder     # eigene Schaubilder (Rückbau, Formcodierung)
python scripts/build_pptx.py           # -> _output/antrittsvorlesung.pptx
```

Inhalt und Sprechernotizen: `vortrag.qmd`. Design: `vorlage/hs-mainz.pptx`.
`quarto render --to pptx` nicht verwenden – PowerPoint stürzt mit der Vorlage ab.

## Noch fehlende Bilder (in `bilder/` ablegen, dann neu bauen)

| Datei | Motiv | Quelle / Lizenz |
|---|---|---|
| `crowdstrike_lga.jpg` | Bluescreens an der Gepäckausgabe, LaGuardia | Wikimedia Commons „CrowdStrike_BSOD_at_LGA.jpg“ (Smishra1, CC BY-SA) |
| `hawaii_menue.jpg` | Nachbildung des Alarm-Menüs | Hawaii EMA, z. B. via Business Insider / Star-Advertiser – als „Nachbildung der Behörde“ kennzeichnen |
| `citi_flexcube.jpg` | Flexcube-Eingabemaske | Gerichtsentscheidung SDNY 2021 (In re Citibank August 11, 2020 Wire Transfers), z. B. via Core77 |
| `hudson_1549.jpg` | US Airways 1549 auf dem Hudson | Wikimedia Commons „Miracle on the Hudson.jpg“ (Greg L, CC BY 2.0) |

Fehlt eine Datei, zeigt die Folie einen grauen Platzhalter mit Dateinamen.

## Handout (zwei Seiten, PDF)

```bash
quarto render handout.qmd --to typst    # -> _output/handout.pdf
```

Layout (Logo, Fußzeile, Schrift) in `vorlage/page.typ`, kompakter Titelblock in `vorlage/typst-template.typ`.
Vorlage übernommen aus dem LUX-Vortrag. Das Handout muss auf genau zwei Seiten bleiben (Doppelseite).
