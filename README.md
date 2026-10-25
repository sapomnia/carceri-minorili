# Carceri minorili in Italia

Dati e analisi sugli istituti penali per i minorenni (IPM): presenze, affollamento, trasferimenti, posizione giuridica, ingressi e uscite. Le rilevazioni vanno dal 31/12/2025 al 30/09/2026; le serie annuali dal 2016.

## Contenuto

| File | Contenuto |
|---|---|
| `fonti/` | PDF originali delle fonti ufficiali |
| `screenshot/` | Screenshot di alcune pagine di documenti presenti nelle fonti |
| `script/` | Script Python di estrazione e costruzione di un file Excel con i dati. Per quelli sull'affollamento fare riferimento a 'Tasso di occupazione carceri minorili.csv', ricostruito manualmente|

## Fonti

- Dipartimento per la giustizia minorile e di comunità (DGMC), *Minorenni e giovani adulti in carico ai Servizi minorili – Analisi statistica dei dati*, edizioni quindicinali dal 31/12/2025 al 30/09/2026 ([pagina delle statistiche](https://www.giustizia.it/giustizia/it/mg_1_14.page?selectedNode=0_6)). I dati 2026 sono provvisori.
- DGMC, *Flussi di utenza degli Istituti penali per i minorenni – Anno 2025*.
- Garante nazionale dei diritti delle persone private della libertà personale, *Osservatorio penitenziario minori e giovani adulti*, report al 31/03/2025 (`20250331_osservatorio_ipm.pdf`): capienze DGMC per istituto, presenti al 31 dicembre 2020-2024.
- Ministero della Giustizia, risposta all'interrogazione 4-06110 del 18/11/2025 (capienze di L'Aquila, Lecce e Rovigo), e Carta dei servizi 2026 IPM e CPA di Cagliari.
- Antigone, Osservatorio sulla giustizia minorile, schede istituto su [ragazzidentro.it](https://www.ragazzidentro.it/istituti/istituti-penali-per-minorenni/) (capienze rilevate nelle visite, per confronto).

Il DGMC non pubblica la capienza degli IPM nei report statistici. Il foglio sull'affollamento contenuto nel file Excel usa le capienze DGMC al 30/03/2025 riportate dal Garante; per gli istituti aperti dopo usa i posti indicati dal ministero. Le capienze di altre fonti sono riportate a confronto. Il file .csv sull'affollamento riporta invece le capienze al 31/120/2025 riportate dal Garante integrate con notizie di stampa relative all'apertura del carcere di Rovigo, alla chiusura di quello di Treviso e al completamento dei lavori in quelli di L'Aquila e Lecce avvenuti nel corso del 2026.

## Rifare l'estrazione

```bash
pip install pdfplumber pandas openpyxl
python3 script/dump.py fonti txt
python3 script/parse.py fonti txt out
python3 script/parse_static.py txt/Analisi_Servizi_minorili_30.09.2026G.pdf.json out
```

`build_xlsx.py` usa anche un modulo locale per i codici Istat (`geo_lookup`) e i dati 2025 estratti con `parse_static.py` dall'edizione al 31/12/2025.
