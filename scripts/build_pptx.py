"""Baut die PowerPoint aus vortrag.qmd mit der Originalvorlage der Hochschule.

Aufruf (im Repo-Verzeichnis):  python scripts/build_pptx.py
Ergebnis:                      _output/vortrag.pptx

Warum nicht `quarto render --to pptx`? Pandoc erzeugt mit dieser Vorlage
Dateien, die PowerPoint nicht öffnen kann. Dieses Skript nutzt die Vorlage
direkt (python-pptx), entfernt deren Beispielfolien und baut die Folien
mit den Layouts der Vorlage neu auf.

Unterstütztes Markdown (bewusst schlank):
  ## Titel                  neue Folie
  **Satz**                  Kernaussage (groß), "\\" am Zeilenende = Zeilenumbruch
  *Text*                    kursiver Hinweis (z. B. Bildplatzhalter)
  - Punkt / 1. Punkt        Aufzählung
  :::: {.columns} / ::: {.column}   zwei Spalten
  ::: notes ... :::         Sprechernotizen (auch ::: {.notes})
  > Zitat                   große These, Folgezeilen = Zeilenumbruch
  ::: {.fragment}           Folie wird doppelt erzeugt: erst ohne, dann mit diesem Teil
  ## Titel {background-color="#1a1a2e"}   dunkle Folie mit weißer Schrift
  ## Titel {.section}       Kapiteltrenner (Layout der Vorlage)
  ![Beschriftung](bilder/x.jpg)   Bild; fehlt die Datei, erscheint ein Platzhalter
"""

import copy
import re
import sys
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Pt

ROOT = Path(__file__).resolve().parent.parent
QMD = ROOT / "vortrag.qmd"
VORLAGE = ROOT / "vorlage" / "hs-mainz.pptx"
ZIEL = ROOT / "_output" / "antrittsvorlesung.pptx"

FUSSZEILE = "Fachbereich Wirtschaft | School of Business"
DATUM = "12. November 2026"
KERN_PT = 32      # Kernaussagen
THESEN_PT = 24    # Aufzählung auf der Thesen-Folie
SPALTE_KOPF_PT = 22  # Überschrift einer Spalte
SPALTE_PT = 20       # Text in Spalten
LISTE_PT = 24        # Aufzählungen und Fließtext
HINWEIS_PT = 18      # kursive Hinweiszeile

MONATE = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli",
          "August", "September", "Oktober", "November", "Dezember"]


# --------------------------------------------------------------------------
# qmd lesen
# --------------------------------------------------------------------------
def parse_qmd(text):
    fm_match = re.match(r"^---\n(.*?)\n---\n", text, re.S)
    meta = {}
    for line in fm_match.group(1).splitlines():
        m = re.match(r'^([\w-]+):\s*"?(.*?)"?\s*$', line)
        if m:
            meta[m.group(1)] = m.group(2)
    body = text[fm_match.end():]

    def new_slide(title, attrs=""):
        bg = re.search(r'background-color="(#[0-9a-fA-F]{6})"', attrs)
        return {"title": title, "blocks": [[]], "notes": [], "fragment": [],
                "background": bg.group(1) if bg else None,
                "section": ".section" in attrs, "image": None}

    slides = [new_slide(None)]  # Titelfolie
    cur = slides[0]
    stack = []  # offene Divs: columns | column | notes | fragment | other
    for raw in body.splitlines():
        line = raw.rstrip()
        if line.startswith("## "):
            m = re.match(r"^## (.*?)\s*(\{.*\})?\s*$", line)
            cur = new_slide(m.group(1).strip(), m.group(2) or "")
            slides.append(cur)
            stack = []
            continue
        div = re.match(r"^:{3,}\s*(.*)$", line)
        if div:
            spec = div.group(1).strip()
            if not spec:                      # Div schließen
                if stack:
                    stack.pop()
                continue
            if "columns" in spec:
                kind = "columns"
                cur["blocks"] = []
            elif "column" in spec:
                kind = "column"
                cur["blocks"].append([])
            elif "notes" in spec:
                kind = "notes"
            elif "fragment" in spec:
                kind = "fragment"
            else:
                kind = "other"
            stack.append(kind)
            continue
        img = re.match(r"^!\[(.*?)\]\((.*?)\)", line.strip())
        if img and "notes" not in stack:
            cur["image"] = (img.group(2), img.group(1))
            continue
        if "notes" in stack:
            cur["notes"].append(line)
        elif "fragment" in stack:
            cur["fragment"].append(line)
        elif "columns" in stack and "column" not in stack:
            continue  # Leerzeilen zwischen Spalten
        else:
            cur["blocks"][-1].append(line)
    return meta, slides


def paragraphs(lines):
    """Zeilen -> Absätze: (art, [zeilen])  art: kern | hinweis | punkt | nummer | text"""
    out, buf = [], []

    def flush(kind):
        if buf:
            out.append((kind, list(buf)))
            buf.clear()

    kind = None
    for line in lines:
        if line.startswith(">"):              # Zitat: Folgezeilen = Zeilenumbruch
            zitat = line[1:].strip()
            if kind != "zitat":
                flush(kind)
                kind = "zitat"
            if zitat:
                buf.append(zitat)
            continue
        if kind == "zitat":
            flush(kind); kind = None
        if not line.strip():
            flush(kind); kind = None
            continue
        cont = line.endswith("\\")
        clean = line[:-1].rstrip() if cont else line
        if buf:                       # Fortsetzung nach "\"
            buf.append(clean)
        else:
            m_b = re.match(r"^- (.*)$", clean)
            m_n = re.match(r"^\d+\. (.*)$", clean)
            if m_b:
                kind, clean = "punkt", m_b.group(1)
            elif m_n:
                kind, clean = "nummer", m_n.group(1)
            elif clean.startswith("**"):
                kind = "kern"
            elif clean.startswith("*"):
                kind = "hinweis"
            else:
                kind = "text"
            buf.append(clean)
        if not cont:
            flush(kind); kind = None
    flush(kind)
    return out


def add_runs(paragraph, text, size=None, force_bold=None, force_italic=None):
    """**fett** und *kursiv* inline umsetzen; \\n wird zum Zeilenumbruch."""
    for part in re.split(r"(\*\*.+?\*\*|\*.+?\*)", text, flags=re.S):
        if not part:
            continue
        bold = part.startswith("**")
        italic = not bold and part.startswith("*") and len(part) > 1
        part = part.strip("*") if (bold or italic) else part
        part = part.replace('"', "“")  # deutsches schließendes Anführungszeichen
        for i, piece in enumerate(part.split("\n")):
            if i:
                paragraph.add_line_break()
            if not piece:
                continue
            run = paragraph.add_run()
            run.text = piece
            if bold or force_bold:
                run.font.bold = True
            if italic or force_italic:
                run.font.italic = True
            if size:
                run.font.size = Pt(size)


def fill(text_frame, lines, thesen=False, spalte=False):
    text_frame.clear()
    first = True
    first_para = True
    nr = 0
    for kind, parts in paragraphs(lines):
        p = text_frame.paragraphs[0] if first else text_frame.add_paragraph()
        first = False
        if spalte:
            size = SPALTE_KOPF_PT if kind == "kern" else SPALTE_PT
        elif kind in ("kern", "zitat"):
            size = KERN_PT
        elif kind == "hinweis":
            size = HINWEIS_PT
        else:
            size = THESEN_PT if thesen else LISTE_PT
        if kind == "nummer":
            # Nummer in den Text schreiben – unabhängig vom Aufzählungsstil der Vorlage
            nr += 1
            parts = [f"{nr}. {parts[0]}"] + parts[1:]
        add_runs(p, "\n".join(parts), size=size, force_bold=(kind == "zitat"))
        if kind in ("kern", "zitat", "hinweis", "text", "nummer"):
            _no_bullet(p)
        if kind == "zitat":
            p.space_after = Pt(18)
        if kind in ("hinweis", "kern") and not first_para:
            p.space_before = Pt(14)
        if kind in ("punkt", "nummer"):
            p.space_after = Pt(6)
        first_para = False


def _no_bullet(p):
    pPr = p._p.get_or_add_pPr()
    for tag in ("a:buNone", "a:buChar", "a:buAutoNum"):
        for el in pPr.findall("{http://schemas.openxmlformats.org/drawingml/2006/main}" + tag.split(":")[1]):
            pPr.remove(el)
    pPr.set("marL", "0")
    pPr.set("indent", "0")
    pPr.insert(0, pPr.makeelement("{http://schemas.openxmlformats.org/drawingml/2006/main}buNone", {}))


# --------------------------------------------------------------------------
# Vorlage vorbereiten
# --------------------------------------------------------------------------
def leere_vorlage(path):
    prs = Presentation(path)
    sld_ids = prs.slides._sldIdLst
    for sld_id in list(sld_ids):
        prs.part.drop_rel(sld_id.rId)
        sld_ids.remove(sld_id)
    # Überarbeitungs-Infos der Vorlage verweisen auf alte Folien -> entfernen
    for rel in list(prs.part.rels.values()):
        if rel.reltype.endswith(("/revisionInfo", "/changesInfo")):
            prs.part.drop_rel(rel.rId)
    return prs


def layout(prs, name):
    return next(l for l in prs.slide_layouts if l.name == name)


def footer_placeholders(slide, nr):
    """Datum, Fußzeile und Foliennummer aus dem Layout auf die Folie übernehmen."""
    for ph in slide.slide_layout.placeholders:
        t = ph.placeholder_format.type
        if t is None or t not in (15, 16, 13):  # FOOTER, DATE, SLIDE_NUMBER
            continue
        el = copy.deepcopy(ph._element)
        slide.shapes._spTree.append(el)
        shape = slide.placeholders[ph.placeholder_format.idx]
        tf = shape.text_frame
        if t == 15:
            _set_text(tf, FUSSZEILE)
        elif t == 16:
            _set_text(tf, DATUM)
        # Foliennummer: Feld aus dem Layout bleibt erhalten, PowerPoint zählt selbst


def _set_text(tf, text):
    p = tf.paragraphs[0]
    for r in list(p.runs):
        r._r.getparent().remove(r._r)
    for extra in tf.paragraphs[1:]:
        extra._p.getparent().remove(extra._p)
    p.add_run().text = text


def remove_unused(slide, keep_idx):
    for ph in list(slide.placeholders):
        if ph.placeholder_format.idx not in keep_idx and ph.placeholder_format.type not in (13, 15, 16):
            ph._element.getparent().remove(ph._element)


def set_notes(slide, notes):
    text = "\n".join(l for l in notes).strip()
    if text:
        slide.notes_slide.notes_text_frame.text = text


# --------------------------------------------------------------------------
def main():
    meta, slides = parse_qmd(QMD.read_text(encoding="utf-8"))
    prs = leere_vorlage(VORLAGE)

    # Titelfolie
    s = prs.slides.add_slide(layout(prs, "Title Slide"))
    s.shapes.title.text = meta.get("title", "")
    sub = s.placeholders[1].text_frame
    sub.clear()
    lines = [meta.get("subtitle", ""), "", meta.get("author", ""), DATUM]
    for i, l in enumerate(lines):
        p = sub.paragraphs[0] if i == 0 else sub.add_paragraph()
        p.text = l
    set_notes(s, slides[0]["notes"])

    for sd in slides[1:]:
        # {.fragment}: Folie zweimal – erst ohne, dann mit dem nachgereichten Text.
        # Das funktioniert auf jedem Laptop, ohne Animationen.
        stufen = [sd["blocks"]]
        if any(l.strip() for l in sd["fragment"]):
            stufen.append([sd["blocks"][0] + [""] + sd["fragment"]] + sd["blocks"][1:])
        for blocks in stufen:
            add_content_slide(prs, sd, blocks)

    ZIEL.parent.mkdir(exist_ok=True)
    prs.save(ZIEL)
    print(f"{ZIEL.relative_to(ROOT)}: {len(prs.slides)} Folien")


BILD_GRAU = RGBColor(0xE6, 0xE9, 0xEE)


def add_picture_fit(slide, path, alt, left, top, width, height):
    """Bild proportional in den Rahmen einpassen; fehlt es, grauen Platzhalter zeichnen."""
    from pptx.enum.shapes import MSO_SHAPE
    datei = ROOT / path
    if datei.exists():
        from PIL import Image
        w, h = Image.open(datei).size
        scale = min(width / w, height / h)
        pw, ph = int(w * scale), int(h * scale)
        pic = slide.shapes.add_picture(str(datei), left + (width - pw) // 2, top + (height - ph) // 2, pw, ph)
        pic.name = alt
        return
    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    box.fill.solid(); box.fill.fore_color.rgb = BILD_GRAU
    box.line.color.rgb = RGBColor(0x99, 0x99, 0x99)
    tf = box.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run(); r.text = f"BILD FEHLT: {alt}\n({path})"
    r.font.size = Pt(16); r.font.color.rgb = RGBColor(0x55, 0x55, 0x55)


def add_section_slide(prs, sd):
    s = prs.slides.add_slide(layout(prs, "Kapiteltrenner (hellblau)"))
    s.shapes.title.text = sd["title"]
    remove_unused(s, {0})
    footer_placeholders(s, None)
    set_notes(s, sd["notes"])


def add_content_slide(prs, sd, blocks):
    if sd["section"]:
        return add_section_slide(prs, sd)
    if sd["image"]:
        return add_image_slide(prs, sd, blocks)
    two = len(blocks) == 2
    s = prs.slides.add_slide(layout(prs, "Two Content" if two else "Title and Content"))
    s.shapes.title.text = sd["title"]
    thesen = sd["title"] == "Debriefing"
    if two:
        fill(s.placeholders[1].text_frame, blocks[0], spalte=True)
        fill(s.placeholders[14].text_frame, blocks[1], spalte=True)
        remove_unused(s, {0, 1, 14})
    else:
        fill(s.placeholders[1].text_frame, blocks[0], thesen=thesen)
        remove_unused(s, {0, 1})
    footer_placeholders(s, None)
    if sd["background"]:
        dunkel(s, sd["background"])
    set_notes(s, sd["notes"])


def add_image_slide(prs, sd, blocks):
    """Bild rechts, Text links; ohne Text: Bild über die ganze Breite."""
    hat_text = any(l.strip() for l in blocks[0])
    path, alt = sd["image"]
    breit = False
    if hat_text and (ROOT / path).exists():
        from PIL import Image
        w, h = Image.open(ROOT / path).size
        breit = w / h > 1.8
    if breit:
        # Breites Schaubild: Text oben, Bild darunter über die ganze Breite
        s = prs.slides.add_slide(layout(prs, "Title and Content"))
        s.shapes.title.text = sd["title"]
        ph = s.placeholders[1]
        ph.left, ph.top, ph.height, ph.width = 371475, 1324947, 1300000, 11449050
        fill(ph.text_frame, blocks[0])
        remove_unused(s, {0, 1})
        add_picture_fit(s, path, alt, 371475, 2700000, 11449050, 3450000)
        footer_placeholders(s, None)
        set_notes(s, sd["notes"])
        return
    s = prs.slides.add_slide(layout(prs, "Two Content" if hat_text else "Nur Titel"))
    s.shapes.title.text = sd["title"]
    if hat_text:
        fill(s.placeholders[1].text_frame, blocks[0])
        ph = s.placeholders[14]
        l, t, w, h = ph.left, ph.top, ph.width, ph.height
        ph._element.getparent().remove(ph._element)
        remove_unused(s, {0, 1})
        add_picture_fit(s, path, alt, l, t, w, h)
    else:
        remove_unused(s, {0})
        add_picture_fit(s, path, alt, 371475, 1324947, 11449050, 4767878)
    footer_placeholders(s, None)
    set_notes(s, sd["notes"])


def dunkel(slide, farbe):
    """Hintergrundfarbe setzen und allen Text weiß machen."""
    fill_ = slide.background.fill
    fill_.solid()
    fill_.fore_color.rgb = RGBColor.from_string(farbe.lstrip("#"))
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for p in shape.text_frame.paragraphs:
            for r in p.runs:
                r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            # Feldinhalte (z. B. Foliennummer) und Text ohne Runs
            for fld in p._p.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}fld"):
                rpr = fld.find("{http://schemas.openxmlformats.org/drawingml/2006/main}rPr")
                if rpr is None:
                    rpr = fld.makeelement("{http://schemas.openxmlformats.org/drawingml/2006/main}rPr", {})
                    fld.insert(0, rpr)
                for old in list(rpr):
                    rpr.remove(old)
                sf = rpr.makeelement("{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill", {})
                clr = sf.makeelement("{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr", {"val": "FFFFFF"})
                sf.append(clr)
                rpr.append(sf)


if __name__ == "__main__":
    sys.exit(main())
