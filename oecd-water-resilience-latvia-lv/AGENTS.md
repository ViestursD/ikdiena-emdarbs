# Codex darba instrukcijas

## Mērķis

Patstāvīgi iztulkot visu OECD ziņojumu “Fostering Water Resilience in Latvia” (2026. gada 4. novembra projekts apspriešanai ar iesaistītajām pusēm) no angļu valodas latviešu valodā. Rezultāts jāsaglabā Markdown formātā — katra ziņojuma nodaļa vai loģiski patstāvīga apakšnodaļa savā `.md` failā.

## Autonomais darba režīms

- Neapstājies pēc vienas nodaļas un negaidi lietotāja ziņu “turpinām”.
- Vispirms izpēti visu oriģinālo DOCX, tā satura rādītāju, virsrakstu hierarhiju, tabulas, attēlu parakstus, zemsvītras piezīmes un pielikumus.
- Izveido pilnu tulkojamo sadaļu sarakstu un atzīmē to `CHANGELOG.md`.
- Tulko secīgi, pa vienai sadaļai/failam. Pēc katras sadaļas saglabā rezultātu un atjaunini progresu.
- Pēc faila saglabāšanas turpini ar nākamo sadaļu automātiski.
- Ja konteksta vai izpildes ierobežojumu dēļ darbs jāpārtrauc, atjaunini `CHANGELOG.md` ar precīzu pēdējo pabeigto sadaļu, nākamo veicamo sadaļu un visām neatrisinātajām vietām. Nākamajā palaišanas reizē turpini no šīs vietas, nedublējot pabeigto.
- Neizdomā trūkstošu avota tekstu un neizlaid neskaidras vietas. Atzīmē tās ar `[PĀRBAUDĀMS: ...]` un ieraksti progresa žurnālā.

## Avota faili

Primārais avots:
`source/Draft%20report%20-%20Fostering%20water%20resilience%20in%20Latvia%20-%20For%20stakeholders%204%20Nov%202026.docx`

Ja šī faila nav, neizdomā tulkojumu. Pārbaudi `source/` saturu un skaidri norādi, ka nepieciešams ievietot oriģinālo ziņojumu. Ja ir pieejams iepriekšējais tulkojums, izmanto to kā darba turpinājuma pamatu, bet pārbaudi pret angļu oriģinālu.

## Tulkošanas kvalitāte

- Tulko pilnā apjomā, precīzi un profesionālā, skaidrā latviešu valodā.
- Saglabā OECD ziņojuma struktūru, argumentāciju, nozīmes nianses un neitrālo institucionālo stilu.
- Nekoriģē avota faktus pēc saviem ieskatiem un nepievieno ārēju informāciju.
- Saglabā visus skaitļus, mērvienības, procentus, gadus, atsauces, avotu norādes un juridisko aktu nosaukumus.
- Skaitļus un decimāldaļas noformē atbilstoši latviešu valodas rakstībai, nemainot vērtību.
- Saglabā attēlu un tabulu numerāciju un parakstus. Tabulas veido Markdown tabulu formā, ja tas ir tehniski iespējams.
- Zemsvītras piezīmes pārnes uz atbilstošu piezīmju sadaļu vai Markdown footnote formātu, saglabājot saites ar tekstu.
- Bibliogrāfijā saglabā oriģinālos publikāciju nosaukumus, DOI un URL; tulko tikai paskaidrojošo tekstu, ja nepieciešams.
- Ja avotā ir kļūda vai pretruna, nepārraksti to klusējot. Atzīmē progresa žurnālā.

## Terminoloģija

Sākotnēji izmanto šādu konsekventu terminoloģiju; ja avota konteksts prasa citu risinājumu, fiksē to terminoloģijas tabulā:

- water resilience — ūdens resursu noturība
- water supply and sanitation (WSS) — ūdensapgāde un kanalizācija (ŪK)
- water supply and sanitation services — ūdensapgādes un kanalizācijas pakalpojumi
- wastewater — notekūdeņi
- wastewater treatment — notekūdeņu attīrīšana
- wastewater treatment plant (WWTP) — notekūdeņu attīrīšanas iekārta (NAI)
- water utility / WSS utility — ūdenssaimniecības pakalpojumu sniedzējs
- agglomeration — aglomerācija
- population equivalent (PE) — cilvēkekvivalents (CE)
- sludge — dūņas / notekūdeņu dūņas
- stormwater — lietusūdens
- runoff — notece
- non-revenue water — neuzskaitītais ūdens
- affordability — pakalpojumu pieejamība izmaksu ziņā
- water security — ūdensdrošība (izvērtē kontekstu un lieto konsekventi)
- resilience — noturība
- tertiary treatment — terciārā attīrīšana
- secondary treatment — sekundārā attīrīšana
- quaternary treatment — ceturtās pakāpes attīrīšana
- energy neutrality — enerģētiskā neitralitāte

Izveido `translation/GLOSSARY.md` un papildini to ar terminiem, kuru tulkojums jālieto konsekventi.

## Failu organizācija

- `translation/00-saīsinājumi-un-akronīmi.md`
- `translation/01-udensapgade-un-kanalizacija.md`
- `translation/02-udensapgades-un-kanalizacijas-parvaldiba.md`
- `translation/03-ilgtermina-strategija.md` — ietver 1.–4. rekomendāciju un A pielikumu.
- `translation/GLOSSARY.md` — terminoloģija.
- Katram failam sāc ar virsrakstu, kas atbilst avota sadaļai. Neiekļauj mākslīgus kopsavilkumus vai tulkotāja komentārus pamattekstā.

## Pārbaude pirms pabeigšanas

1. Salīdzini avota satura rādītāju ar izveidoto Markdown failu sarakstu.
2. Pārbaudi, vai nav izlaista neviena nodaļa, apakšnodaļa, rindkopa, tabula, attēla paraksts, piezīme vai pielikums.
3. Pārbaudi skaitļus, procentus, datumus, atsauces un terminoloģijas konsekvenci.
4. Pārliecinies, ka Markdown faili ir UTF-8 un korekti attēlo latviešu diakritiskās zīmes.
5. Atjaunini `README.md` un `CHANGELOG.md` ar pabeigšanas statusu un zināmajiem ierobežojumiem.

## DOCX sagatavošana

Kad viss tulkojums ir pabeigts, sagatavo Python skriptu `scripts/build_docx.py`, kas apvieno Markdown failus pareizā secībā vienā DOCX dokumentā. Saglabā virsrakstu līmeņus, sarakstus, tabulas, saites un piezīmes, cik iespējams. Neiekļauj `GLOSSARY.md` pamatdokumentā, ja vien lietotājs to nav prasījis; terminoloģiju vari pievienot kā atsevišķu pielikumu.
