# OECD — Ūdens resursu noturības stiprināšana Latvijā

Latviešu tulkojuma projekts OECD ziņojumam **Fostering Water Resilience in Latvia** (2026. gada 4. novembra projekts apspriešanai ar iesaistītajām pusēm).

## Struktūra

- `source/` — oriģinālais angļu valodas ziņojums un iepriekšējā tulkojuma materiāli.
- `translation/` — katra ziņojuma sadaļa atsevišķā Markdown failā.
- `scripts/` — rīki Markdown failu apvienošanai vienā DOCX dokumentā.
- `AGENTS.md` — Codex pastāvīgās darba instrukcijas.
- `CHANGELOG.md` — paveiktais, pārbaudes un zināmās avota neatbilstības.
- `output/OECD_udens_resursu_noturiba_Latvija.docx` — apvienotais gala DOCX dokuments.

## Pašreizējais stāvoklis

Oriģinālais OECD DOCX dokuments ir pieejams mapē `source/`. Tā pilnā struktūra ir pārbaudīta: 409 rindkopas, 16 tabulas, 25 attēlu/tabulu parakstu rindkopas, 15 zemsvītras piezīmes, galvenes/kājenes un iegulti attēli.

Pilnais tulkojums ir pabeigts un sadalīts secīgos Markdown failos:

1. `translation/00-saīsinājumi-un-akronīmi.md` — saīsinājumi un akronīmi.
2. `translation/01-udensapgade-un-kanalizacija.md` — 1. nodaļa par pakalpojumu sniegšanu, investīciju vajadzībām un megatendencēm.
3. `translation/02-udensapgades-un-kanalizacijas-parvaldiba.md` — 2. nodaļa par pārvaldību.
4. `translation/03-ilgtermina-strategija.md` — 3. nodaļa ar četrām rekomendācijām un A pielikumu.

Iepriekšējie tulkojuma fragmenti ir pārbaudīti pret oriģinālu un integrēti šajos pilnajos failos. Terminu konsekvenci uztur `translation/GLOSSARY.md`.

DOCX dokumentu izveido `scripts/build_docx.py`. Tas saglabā virsrakstu hierarhiju, sarakstus, Markdown tabulas, aktīvas ārējās saites un piezīmes failu numerācijas secībā.

Veikta gala kvalitātes kontrole: pārbaudīta avota struktūra un tulkojuma sadaļu secība, attēlu un tabulu paraksti, skaitļi, procenti, datumi, mērvienības, atsauces, terminoloģija, Markdown UTF-8 kodējums, DOCX ZIP struktūra un renderētas dokumenta lappuses. Avota iekšējās atsauces uz attēliem, kas savstarpēji nesakrīt, ir saglabātas un aprakstītas `CHANGELOG.md`.

## Turpmāka atjaunināšana ar Codex

1. Atjaunini avota DOCX tikai tad, ja ir pieejama jaunāka OECD projekta versija.
2. Salīdzini izmaiņas ar `translation/` failiem un atjaunini `translation/GLOSSARY.md`.
3. Palaid `scripts/build_docx.py`, lai pārbūvētu gala dokumentu.
4. Pārbaudi renderēto DOCX un atjaunini `CHANGELOG.md`.

## Pamatprincips

Tulkojumam jāatbilst oriģinālajam dokumentam. Nedrīkst izlaist rindkopas, tabulas, attēlu parakstus, piezīmes, atsauces vai pielikumus.
