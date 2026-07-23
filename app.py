#!/usr/bin/env python3
"""
BOE document-service.

Ontvangt JSON van n8n, vult het juiste sjabloon, geeft een .docx terug.

Welk sjabloon gebruikt wordt, hangt af van de KLANT. Staat de klant niet in
KLANT_SJABLOON, dan wordt het standaard BOE-sjabloon gebruikt.

Alle sjablonen gebruiken dezelfde stijlnamen (Titel, Kop1-Kop4, Standaard,
Citaat, Lijstalinea). Alleen de tabelstijl verschilt per klant.

Endpoints:
    GET  /gezond      -> {"status": "ok"}
    GET  /sjablonen   -> overzicht van sjablonen en klantkoppeling
    POST /genereer    -> vult sjabloon, geeft .docx terug
        header:  x-api-key: <jouw sleutel>
"""

import io
import os
import re

from fastapi import FastAPI, Header, HTTPException
from fastapi.responses import StreamingResponse

import boe_document
import markdown_parser

SJABLOON_MAP = os.environ.get("SJABLOON_MAP", "templates")
API_KEY = os.environ.get("API_KEY", "")

# ── Uniforme stijlmapping: geldt voor ALLE sjablonen ───────────────────────
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

# ── Sjablonen ──────────────────────────────────────────────────────────────
# Nieuwe klant toevoegen? Zet het .docx in templates/, voeg hier een regel
# toe, en koppel de klantnaam in KLANT_SJABLOON hieronder.
SJABLONEN = {
    "boe": {"bestand": "BOE_sjabloon.docx", "tabelstijl": "BOEtabel"},
    "nxt": {"bestand": "NXT_sjabloon.docx", "tabelstijl": "NXTtabel"},
}

STANDAARD = "boe"

# ── Klant -> sjabloon ──────────────────────────────────────────────────────
# De sleutel wordt genormaliseerd vergeleken (kleine letters, geen leestekens)
# en er wordt gekeken of hij VOORKOMT in de klantnaam. Zo matcht "nxtgen" ook
# op "NXTGEN Hightech" en "NXTGEN Hightech Agrifood".
KLANT_SJABLOON = {
    "nxtgen": "nxt",
}

# Vaste instellingen die voor alle sjablonen gelden
KOPRIJSTIJL   = "Koppentabel"
BULLET_NUMID  = 1
PAGINABREEDTE = 9072


def normaliseer(tekst):
    return re.sub(r"[^a-z0-9]", "", (tekst or "").lower())


def kies_sjabloon(klant):
    k = normaliseer(klant)
    for sleutel, naam in KLANT_SJABLOON.items():
        if normaliseer(sleutel) in k:
            return naam
    return STANDAARD


def maak_config(naam):
    s = SJABLONEN[naam]
    return {
        "stijlen": STIJLEN,
        "tabelstijl": s["tabelstijl"],
        "koprijstijl": KOPRIJSTIJL,
        "bullet_numid": BULLET_NUMID,
        "paginabreedte": PAGINABREEDTE,
    }


app = FastAPI(title="BOE document-service")


def controleer_sleutel(sleutel):
    if not API_KEY:
        return  # geen sleutel ingesteld = beveiliging uit (alleen voor testen)
    if sleutel != API_KEY:
        raise HTTPException(status_code=401, detail="Ongeldige of ontbrekende API-key")


@app.get("/gezond")
def gezond():
    return {"status": "ok"}


@app.get("/sjablonen")
def sjablonen():
    overzicht = {}
    for naam, s in SJABLONEN.items():
        pad = os.path.join(SJABLOON_MAP, s["bestand"])
        overzicht[naam] = {
            "bestand": s["bestand"],
            "tabelstijl": s["tabelstijl"],
            "gevonden": os.path.exists(pad),
        }
    return {
        "sjablonen": overzicht,
        "klantkoppeling": KLANT_SJABLOON,
        "standaard": STANDAARD,
        "stijlen": STIJLEN,
    }


@app.post("/genereer")
def genereer(payload: dict, x_api_key: str = Header(default="")):
    controleer_sleutel(x_api_key)

    naam = kies_sjabloon(payload.get("klant", ""))
    cfg = maak_config(naam)

    sjabloon_pad = os.path.join(SJABLOON_MAP, SJABLONEN[naam]["bestand"])
    if not os.path.exists(sjabloon_pad):
        raise HTTPException(
            status_code=500,
            detail=f"Sjabloon niet gevonden: {sjabloon_pad} (set '{naam}')",
        )

    # Levert de agent Markdown in plaats van blokken? Dan eerst vertalen.
    if "briefing_content" in payload and "blokken" not in payload:
        content = markdown_parser.bouw_content(payload)
    else:
        content = payload

    try:
        buffer = io.BytesIO()
        boe_document.bouw(sjabloon_pad, content, buffer, config=cfg)
        buffer.seek(0)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Genereren mislukt: {e}")

    titel = (content.get("velden") or {}).get("titel", "document")
    veilig = "".join(c for c in str(titel) if c.isalnum() or c in " -_").strip()

    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={
            "Content-Disposition": f'attachment; filename="{veilig or "document"}.docx"',
            "X-Sjabloon": naam,
        },
    )
