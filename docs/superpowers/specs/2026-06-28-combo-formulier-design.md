# Combo Formulier (Per Birdie + Eenmalig) — Design Spec

**Datum:** 2026-06-28

## Samenvatting

Het bestaande aanmeldformulier (`index.html`) krijgt een tab-switcher bovenaan waarmee een bezoeker kiest tussen "Per birdie" (bestaand) of "Eenmalig" (nieuw). De webhook (`doPost` in `Code.gs`) routeert op basis van een `type`-veld naar het juiste tabblad.

## HTML Formulier (`index.html`)

### Tab-switcher

Twee tabs bovenaan het formulier:
- **"Per birdie"** (standaard geselecteerd)
- **"Eenmalig"**

Actieve tab is visueel onderscheiden (zelfde stijl als bestaande UI — paarse accentkleur).

### Per birdie tab (bestaand, ongewijzigd)

- Naam, email, telefoon
- Bedrag per birdie (€2 / €3,50 / €5 + eigen bedrag)
- Seizoenscap (optioneel)
- WhatsApp-groep checkbox
- Submit: "Aanmelden als Birdie Vriend 🐦"

### Eenmalig tab (nieuw)

- Naam, email, telefoon
- Eenmalig bedrag (vrij invoerveld, euro)
- Geen seizoenscap, geen WhatsApp
- Submit: "Eenmalig sponsoren"

### Submit logica

- Het formulier voegt een veld `type` toe: `"per_birdie"` of `"eenmalig"` afhankelijk van actieve tab.
- Bestaande validatie en foutafhandeling blijft intact.
- Na succesvolle submit: bestaande bevestigingsmelding, aangepast per type.

## Webhook (`Code.gs` — `doPost`)

### Routing op `type`

```
if (data.type === "eenmalig") → schrijf naar "Eenmalige Sponsors" sheet
else                          → schrijf naar "Aanmeldingen" sheet (bestaand gedrag)
```

### Eenmalige Sponsors rij

Kolommen in "Eenmalige Sponsors" (aangemaakt via `setupEenmaligSponsors`):
`NAAM | EMAIL | TELEFOON | BEDRAG (€) | BETAALD`

Bij inkomende eenmalige aanmelding:
- NAAM = `data.naam`
- EMAIL = `data.email`
- TELEFOON = `data.telefoon`
- BEDRAG = `data.bedrag` (getal)
- BETAALD = `"Nee"` (standaard)

### Bestaand gedrag

Alle aanmeldingen zonder `type` of met `type !== "eenmalig"` gaan naar "Aanmeldingen" — geen wijziging.

## Buiten scope

- Geen bevestigingsmail.
- Geen wijziging aan de "Eenmalige Sponsors" sheet-structuur (al geïmplementeerd).
- Geen aanpassing aan Overzicht, Betalingen of syncToernooien.
