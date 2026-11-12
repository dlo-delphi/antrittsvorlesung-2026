// Seitenlayout für die PDF-Skripte (Quarto-Format typst).
// Nachbildung der Vorlesungsskripte WS 2025/26: HSM-Logo oben rechts,
// Fußzeile mit blauer Linie, Autor und Dokumenttitel links, Seitenzahl rechts.
// Übernommen aus csm-cyber_resilienz-quarto; eingebunden über template-partials in handout.qmd (Vorlage aus LUX-Gesundheitswirtschaft).

#let hsm-blau = rgb("#2f5496")
#let hsm-linie = rgb("#4472c4")

#set page(
  paper: "a4",
  margin: (top: 2.9cm, bottom: 2.3cm, left: 2.2cm, right: 2.2cm),
  header: align(right, image("/vorlage/hsm-wirtschaft-logo.png", height: 1.25cm)),
  header-ascent: 25%,
  footer: context [
    #line(length: 100%, stroke: 2.5pt + hsm-linie)
    #v(-0.35em)
    #set text(size: 7.5pt, fill: rgb("#404040"), tracking: 0.3pt)
    #grid(
      columns: (1fr, auto),
      align: (left, right),
      upper[Dirk Loomans · $title$],
      counter(page).display("1"),
    )
  ],
  footer-descent: 30%,
)

#set text(font: ("Calibri", "Carlito", "DejaVu Sans"), size: 10pt)
#set par(justify: false)
#show heading: set text(fill: hsm-blau, font: ("Calibri", "Carlito", "DejaVu Sans"))
#show heading.where(level: 1): set text(size: 12pt)
#show heading.where(level: 2): set text(size: 10.5pt)
#show heading.where(level: 3): set text(size: 10pt)
#show link: set text(fill: hsm-blau)
#show table.cell: set par(justify: false)
#show table.cell: set text(hyphenate: false)
#show figure.where(kind: table): set block(breakable: true)
