# Eenmalige Sponsors — Design Spec

**Datum:** 2026-06-28

## Samenvatting

Een nieuw tabblad "Eenmalige Sponsors" voor sponsors die een vast eenmalig bedrag doneren, los van de bestaande per-birdie sponsors. Volledig zelfstandig — geen koppeling met Aanmeldingen, Betalingen of het betaalverzoek.

## Tabblad: "Eenmalige Sponsors"

### Kolommen

| # | Kolomnaam       | Type         | Opmaak / Validatie              |
|---|-----------------|--------------|----------------------------------|
| A | NAAM            | Tekst        | —                                |
| B | EMAIL           | Tekst        | —                                |
| C | TELEFOON        | Tekst        | —                                |
| D | BEDRAG (€)      | Getal        | `€#,##0.00`                      |
| E | BETAALD         | Dropdown     | Ja / Nee                         |

### Opmaak

- **Header (rij 1):** achtergrond `#9D174D` (donkerpaars), witte vette tekst — zelfde stijl als Betalingen.
- **Bevroren rij:** rij 1 bevroren.
- **Kolombreedtes:** NAAM 180px, EMAIL 200px, TELEFOON 130px, BEDRAG 130px, BETAALD 100px.
- **Voorwaardelijke opmaak op BETAALD:**
  - "Ja" → groene achtergrond `#d1fae5`, groene tekst `#065f46`
  - "Nee" → rode achtergrond `#fee2e2`, rode tekst `#991b1b`
- **Totaalrij** onderaan (direct na laatste datarij):
  - Kolom A: "TOTAAL OPEN"
  - Kolom D: som van BEDRAG waar BETAALD = "Nee"
  - Achtergrond `#fdf2f8`, vette tekst, euro-opmaak

### Gedrag setup-functie

- Functie `setupEenmaligSponsors()` maakt het tabblad aan.
- Als het tabblad al bestaat: bevestigingsdialoog ("Wil je het opnieuw aanmaken? Alle data gaat verloren.") — zelfde patroon als `setupBetalingen`.
- Na aanmaken: alert "Eenmalige Sponsors-tabblad aangemaakt."

### Menu-item

Toegevoegd aan het bestaande "Birdie Vrienden"-menu (na "Setup Betalingen"):

```
Setup Eenmalige Sponsors
```

## Buiten scope

- Geen koppeling met Aanmeldingen, Betalingen of Overzicht.
- Niet opgenomen in het betaalverzoek-popup.
- Geen automatische sync of formules — volledig handmatig invullen.
