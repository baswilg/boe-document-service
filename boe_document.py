#!/usr/bin/env python3
"""
Vult een BOE Word-sjabloon met content van een AI-agent.

Het sjabloon levert de stijl: logo, header/footer, marges, kleuren, stijlbladen.
De content (dict) levert alleen structuur en tekst. Het script koppelt die twee.

Sjabloon bevat placeholders:
    {{titel}} {{jobnummer}} {{klant}} {{contactpersoon}} {{opdracht}} {{datum}}
    {{content}}   <- hier worden de contentblokken ingevoegd
"""

import re

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Twips

# --- Mapping: naam die de agent gebruikt -> styleId in het sjabloon ------------
STIJLEN = {
    "titel":     "Titel",
    "kop1":      "Kop1",
    "kop2":      "Kop2",
    "kop3":      "Kop3",
    "accent":    "Kop4",
    "standaard": "Standaard",
    "citaat":    "Citaat",
    "bullet":    "Lijstalinea",
}

TABELSTIJL   = "BOEtabel"      # eigen stijl "BOE tabel"
KOPRIJSTIJL  = "Koppentabel"   # eigen stijl "Koppen tabel"
BULLET_NUMID = 1
PAGINABREEDTE_DXA = 9072       # A4 minus marges van 1417 dxa links en rechts
CONTENT_PLACEHOLDER = "{{content}}"


def schrijf_tekst(alinea, tekst):
    """Plaatst tekst in een alinea en maakt **stukken** echt vet."""
    delen = re.split(r"(\*\*[^*]+\*\*)", tekst)
    for deel in delen:
        if not deel:
            continue
        if deel.startswith("**") and deel.endswith("**"):
            run = alinea.add_run(deel[2:-2])
            run.bold = True
        else:
            alinea.add_run(deel)


def vul_velden(doc, velden):
    """Vervangt {{sleutel}} door de waarde, in body, headers en footers."""
    bronnen = list(doc.paragraphs)
    for sectie in doc.sections:
        for deel in (sectie.header, sectie.footer,
                     sectie.first_page_header, sectie.first_page_footer):
            if deel is not None:
                bronnen.extend(deel.paragraphs)

    for alinea in bronnen:
        for run in alinea.runs:
            for sleutel, waarde in velden.items():
                merk = "{{" + sleutel + "}}"
                if merk in run.text:
                    run.text = run.text.replace(merk, str(waarde))


def zoek_content_alinea(doc):
    for alinea in doc.paragraphs:
        if CONTENT_PLACEHOLDER in alinea.text:
            return alinea
    raise ValueError(f"{CONTENT_PLACEHOLDER} niet gevonden in het sjabloon")


def zet_stijl(alinea, style_id):
    pPr = alinea._p.get_or_add_pPr()
    for oud in pPr.findall(qn("w:pStyle")):
        pPr.remove(oud)
    el = OxmlElement("w:pStyle")
    el.set(qn("w:val"), style_id)
    pPr.insert(0, el)


def maak_bullet(alinea):
    pPr = alinea._p.get_or_add_pPr()
    numPr = OxmlElement("w:numPr")
    ilvl = OxmlElement("w:ilvl")
    ilvl.set(qn("w:val"), "0")
    numId = OxmlElement("w:numId")
    numId.set(qn("w:val"), str(BULLET_NUMID))
    numPr.append(ilvl)
    numPr.append(numId)
    pPr.append(numPr)


def zet_tabelstijl(tabel, style_id):
    tblPr = tabel._tbl.tblPr
    for oud in tblPr.findall(qn("w:tblStyle")):
        tblPr.remove(oud)
    el = OxmlElement("w:tblStyle")
    el.set(qn("w:val"), style_id)
    tblPr.insert(0, el)

    for oud in tblPr.findall(qn("w:tblLook")):
        tblPr.remove(oud)
    look = OxmlElement("w:tblLook")
    for k, v in [("firstRow", "1"), ("lastRow", "0"), ("firstColumn", "0"),
                 ("lastColumn", "0"), ("noHBand", "0"), ("noVBand", "1")]:
        look.set(qn("w:" + k), v)
    look.set(qn("w:val"), "04A0")
    tblPr.append(look)


def bouw_tabel(doc, kop, rijen, breedtes=None):
    kolommen = len(kop)
    if not breedtes:
        breedtes = [PAGINABREEDTE_DXA // kolommen] * kolommen
        breedtes[-1] += PAGINABREEDTE_DXA - sum(breedtes)

    tabel = doc.add_table(rows=1 + len(rijen), cols=kolommen)
    zet_tabelstijl(tabel, TABELSTIJL)

    for i, tekst in enumerate(kop):
        cel = tabel.cell(0, i)
        cel.text = str(tekst)
        cel.width = Twips(breedtes[i])
        zet_stijl(cel.paragraphs[0], KOPRIJSTIJL)
    for r, rij in enumerate(rijen, start=1):
        for i, tekst in enumerate(rij[:kolommen]):
            cel = tabel.cell(r, i)
            cel.text = "" if tekst is None else str(tekst)
            cel.width = Twips(breedtes[i])

    trPr = tabel.rows[0]._tr.get_or_add_trPr()
    trPr.append(OxmlElement("w:tblHeader"))   # koprij herhalen bij paginaovergang
    return tabel


def bouw(sjabloon_pad, content, uitvoer):
    """
    sjabloon_pad : pad naar het .docx-sjabloon
    content      : dict met 'velden' en 'blokken'
    uitvoer      : bestandspad (str) OF een open bytes-buffer
    """
    doc = Document(sjabloon_pad)

    vul_velden(doc, content.get("velden", {}))
    anker = zoek_content_alinea(doc)

    for blok in content.get("blokken", []):
        soort = blok.get("type", "standaard").lower()

        if soort == "tabel":
            tabel = bouw_tabel(doc, blok["kop"], blok.get("rijen", []),
                               blok.get("breedtes"))
            anker._p.addprevious(tabel._tbl)
            witregel = doc.add_paragraph()
            zet_stijl(witregel, STIJLEN["standaard"])
            anker._p.addprevious(witregel._p)
            continue

        if soort not in STIJLEN:
            soort = "standaard"

        alinea = doc.add_paragraph()
        schrijf_tekst(alinea, blok.get("tekst", ""))
        zet_stijl(alinea, STIJLEN[soort])
        if soort == "bullet":
            maak_bullet(alinea)
        anker._p.addprevious(alinea._p)

    anker._p.getparent().remove(anker._p)
    doc.save(uitvoer)
    return uitvoer
