#!/usr/bin/env python3
"""
Zet de briefing_content (Markdown zoals de agent die levert) om naar
'blokken' die boe_document.bouw() begrijpt.

Herkent:
    # tekst        -> kop1
    ## tekst       -> kop2
    ### tekst      -> kop3
    _Kopje         -> kop1 (eerste) / kop2 (daarna), BOE-stijl
    * of - bullet  -> bullet
    ---            -> genegeerd (scheidingslijn)
    **vet**        -> vet in de tekst (character run)
    lege regel     -> alineascheiding
    overige tekst  -> standaard

Markers die je in kleur wilt behouden, blijven als tekst staan; kleuren
regelen we later eventueel via character-stijlen.
"""

import re


def _schoon(regel):
    # verwijder harde regeleinde-spaties die Markdown gebruikt (twee spaties)
    return regel.rstrip()


def parse_markdown(md):
    blokken = []
    kopteller = 0

    # buffer voor doorlopende bodytekst (meerdere regels -> één alinea)
    body_buffer = []

    def flush_body():
        if body_buffer:
            tekst = " ".join(s.strip() for s in body_buffer if s.strip())
            if tekst:
                blokken.append({"type": "standaard", "tekst": tekst})
            body_buffer.clear()

    for ruwe_regel in md.split("\n"):
        regel = _schoon(ruwe_regel)
        kaal = regel.strip()

        if kaal == "" or kaal == "---":
            flush_body()
            continue

        # ATX-koppen: ###, ##, #
        m = re.match(r"^(#{1,3})\s+(.*)$", kaal)
        if m:
            flush_body()
            niveau = len(m.group(1))
            soort = {1: "kop1", 2: "kop2", 3: "kop3"}[niveau]
            blokken.append({"type": soort, "tekst": m.group(2).strip()})
            continue

        # BOE _Kopje: eerste wordt kop1 (documenttitel), daarna kop2
        m = re.match(r"^_(\S.*)$", kaal)
        if m and not kaal.startswith("__"):
            flush_body()
            kopteller += 1
            soort = "kop1" if kopteller == 1 else "kop2"
            blokken.append({"type": soort, "tekst": "_" + m.group(1).strip()})
            continue

        # bullets: * of -
        m = re.match(r"^[\*\-]\s+(.*)$", kaal)
        if m:
            flush_body()
            blokken.append({"type": "bullet", "tekst": m.group(1).strip()})
            continue

        # gewone tekstregel -> aan de bodybuffer toevoegen
        body_buffer.append(kaal)

    flush_body()
    return blokken


def bouw_content(payload):
    """
    Zet de n8n-payload (met briefing_content) om naar het content-formaat
    dat boe_document.bouw() verwacht: {'velden': {...}, 'blokken': [...]}.
    """
    md = payload.get("briefing_content", "") or ""
    blokken = parse_markdown(md)

    # De titel bovenaan het sjabloon: eerste kop uit de content, of val terug
    # op onderwerp/klant. De agent stuurt geen losse titel mee.
    titel = ""
    for b in blokken:
        if b["type"] in ("kop1", "kop2"):
            titel = b["tekst"].lstrip("_").strip()
            break
    if not titel:
        titel = payload.get("onderwerp", "") or payload.get("klant", "Document")

    velden = {
        "titel":         titel,
        "jobnummer":     payload.get("jobnummer", ""),
        "klant":         payload.get("klant", ""),
        "contactpersoon": payload.get("contactpersoon", ""),
        "opdracht":      payload.get("onderwerp", ""),
        "datum":         payload.get("datum", ""),
    }

    return {"velden": velden, "blokken": blokken}
