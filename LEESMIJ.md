# BOE document-service

Kleine webservice die een BOE Word-sjabloon vult met content van een AI-agent
en het resultaat als `.docx` teruggeeft. Bedoeld om vanuit n8n Cloud aan te
roepen met een HTTP Request-node.

## Wat zit erin

    app.py              de webservice (endpoints)
    boe_document.py     de logica die het sjabloon vult
    templates/          hier staan je .docx-sjablonen
    requirements.txt    Python-afhankelijkheden
    Dockerfile          om te hosten als container

## Endpoints

    GET  /gezond        -> {"status":"ok"}          (healthcheck)
    GET  /sjablonen     -> lijst met documenttypes
    POST /genereer      -> vult sjabloon, geeft .docx terug
        header:  x-api-key: <jouw sleutel>
        body:    zie content_voorbeeld.json

## Meerdere sjablonen

Zet extra sjablonen in templates/ en koppel ze in app.py in de dict SJABLONEN:

    SJABLONEN = {
        "planning":           "BOE_sjabloon.docx",
        "functioneel_design": "FD_sjabloon.docx",
        "briefing":           "briefing_sjabloon.docx",
    }

De agent kiest met "documenttype": "briefing" in de JSON.

## Beveiliging

Zet in je hosting-omgeving een omgevingsvariabele API_KEY met een geheime
waarde. De service weigert dan elke aanvraag zonder de juiste x-api-key header.

## Lokaal draaien (test)

    pip install -r requirements.txt
    API_KEY=test uvicorn app:app --port 8080
