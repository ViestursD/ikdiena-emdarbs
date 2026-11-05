# OECD — Ūdens resursu noturības stiprināšana Latvijā

Latviešu tulkojuma projekts OECD ziņojumam **Fostering Water Resilience in Latvia** (2026. gada 4. novembra projekts apspriešanai ar iesaistītajām pusēm).

## Struktūra

- `source/` — oriģinālais angļu valodas ziņojums un iepriekšējā tulkojuma materiāli.
- `translation/` — katra ziņojuma sadaļa atsevišķā Markdown failā.
- `scripts/` — rīki Markdown failu apvienošanai vienā DOCX dokumentā.
- `AGENTS.md` — Codex pastāvīgās darba instrukcijas.
- `CHANGELOG.md` — paveiktais un tulkojuma progress.

## Pašreizējais stāvoklis

Mapē `translation/` jau ievietotas iepriekš tulkotās sadaļas:
1. Notekūdeņu apsaimniekošana.
2. Ūdensapgādes un kanalizācijas nozares investīciju vajadzības.
3. Pakalpojumu sniegšanu Latvijā apgrūtinošās megatendences.
4. Terminu piezīmes.

Šīs sadaļas ir jāpārbauda pret oriģinālu un jāintegrē kopējā dokumenta struktūrā.

## Darba sākšana ar Codex

1. Lejupielādē oriģinālo OECD DOCX failu un ievieto to mapē `source/` ar nosaukumu `oecd-water-resilience-latvia-en.docx`.
2. Ja ir iepriekšējā tulkojuma 1. daļa, ievieto to mapē `source/`.
3. Atver šo repozitoriju Codex.
4. Uzdod: **“Izlasi AGENTS.md un README.md. Pārbaudi esošos tulkojuma failus pret oriģinālo ziņojumu un turpini visa ziņojuma tulkošanu patstāvīgi līdz pilnīgai pabeigšanai. Neapstājies pēc vienas sadaļas.”**

## Pamatprincips

Tulkojumam jāatbilst oriģinālajam dokumentam. Nedrīkst izlaist rindkopas, tabulas, attēlu parakstus, piezīmes, atsauces vai pielikumus.
