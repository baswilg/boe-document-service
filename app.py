#!/usr/bin/env python3
"""
BOE document-service.

Ontvangt JSON van n8n, vult het juiste BOE-sjabloon, geeft een .docx terug.

Endpoint:
    POST /genereer
        header:  x-api-key: <jouw sleutel>
        body:    { "documenttype": "planning",
                   "velden": {...},
                   "blokken": [...] }
        return:  het .docx-bestand (binair)

    GET /gezond      -> {"status": "ok"}   (voor een healthcheck)
    GET /sjablonen   -> lijst met beschikbare documenttypes
"""

import io
import os

from fastapi import FastAPI, Header, HTTPException
from fastapi.responses import StreamingResponse, JSONResponse

import boe_document
import markdown_parser

# --- Sjabloon-registry ----------------------------------------------------------
# Zet je .docx-sjablonen in de map templates/ en koppel ze hier aan een naam.
# De agent kiest via "documenttype": "<naam>".
SJABLOON_MAP = os.environ.get("SJABLOON_MAP", "templates")
# Alle doc_types uit je n8n-flow. Nu wijzen ze naar hetzelfde basissjabloon;
# vervang de bestandsnaam zodra je per type een eigen sjabloon hebt.
SJABLONEN = {
    "briefing_copy":    "BOE_sjabloon.docx",
    "briefing_creatie": "BOE_sjabloon.docx",
    "briefing_studio":  "BOE_sjabloon.docx",
    "com_strategie":    "BOE_sjabloon.docx",
    "debriefing":       "BOE_sjabloon.docx",
    "planning":         "BOE_sjabloon.docx",
    "rationale":        "BOE_sjabloon.docx",
    "copy":             "BOE_sjabloon.docx",
}
STANDAARD_TYPE = "briefing_studio"

API_KEY = os.environ.get("API_KEY", "")   # zet deze in je hosting-omgeving

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
    return {"documenttypes": sorted(SJABLONEN.keys())}


@app.post("/genereer")
def genereer(payload: dict, x_api_key: str = Header(default="")):
    controleer_sleutel(x_api_key)

    # doc_type (zoals de agent stuurt) of documenttype; val terug op standaard
    doctype = (payload.get("doc_type")
               or payload.get("documenttype")
               or STANDAARD_TYPE)
    if doctype not in SJABLONEN:
        raise HTTPException(
            status_code=400,
            detail=f"Onbekend documenttype '{doctype}'. "
                   f"Beschikbaar: {sorted(SJABLONEN.keys())}",
        )

    sjabloon_pad = os.path.join(SJABLOON_MAP, SJABLONEN[doctype])
    if not os.path.exists(sjabloon_pad):
        raise HTTPException(status_code=500,
                            detail=f"Sjabloon niet gevonden: {sjabloon_pad}")

    # Als de agent briefing_content (Markdown) stuurt: vertaal naar blokken.
    # Stuurt hij al 'blokken', dan gebruiken we die direct.
    if "briefing_content" in payload and "blokken" not in payload:
        content = markdown_parser.bouw_content(payload)
    else:
        content = payload

    try:
        buffer = io.BytesIO()
        boe_document.bouw(sjabloon_pad, content, buffer)
        buffer.seek(0)
    except Exception as e:  # nette foutmelding terug naar n8n
        raise HTTPException(status_code=500, detail=f"Genereren mislukt: {e}")

    titel = payload.get("velden", {}).get("titel", "document")
    veilig = "".join(c for c in titel if c.isalnum() or c in " -_").strip() or "document"
    bestandsnaam = f"{veilig}.docx"

    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f'attachment; filename="{bestandsnaam}"'},
    )
