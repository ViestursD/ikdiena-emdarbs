# OECD — Ūdens resursu noturības stiprināšana Latvijā

Latviešu tulkojuma projekts OECD ziņojumam **Fostering Water Resilience in Latvia** (2026. gada 4. novembra projekts apspriešanai ar iesaistītajām pusēm).

## Struktūra

- `source/` — oriģinālais angļu valodas ziņojums un iepriekšējā tulkojuma materiāli.
- `translation/` — tulkojums latviešu valodā, katra ziņojuma sadaļa atsevišķā Markdown failā.
- `scripts/` — rīki Markdown failu apvienošanai vienā DOCX dokumentā.
- `CHANGELOG.md` — paveiktais un tulkojuma progress.

## Darba sākšana ar Codex

1. Ievieto oriģinālo OECD DOCX failu mapē `source/` ar nosaukumu `oecd-water-resilience-latvia-en.docx`.
2. Ja pieejams iepriekšējais tulkojums, ievieto to mapē `source/` kā atsevišķu failu.
3. Atver šo repozitoriju Codex un uzdod: **“Izlasi AGENTS.md un turpini pilna ziņojuma tulkošanu. Strādā patstāvīgi līdz visu sadaļu pabeigšanai.”**
4. Codex jāveic tulkojums pa nodaļām, saglabājot katru sadaļu atsevišķā `.md` failā un regulāri atjauninot `CHANGELOG.md`.

## Pamatprincips

Tulkojumam jāatbilst oriģinālajam dokumentam. Nedrīkst izlaist rindkopas, tabulas, attēlu parakstus, piezīmes, atsauces vai pielikumus. Ja kādu elementu nav iespējams korekti pārnest uz Markdown, tas jāatzīmē un jāpaskaidro progresa žurnālā.
