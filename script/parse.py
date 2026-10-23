"""Estrae le tabelle IPM dai report quindicinali DGMC.

Il testo (JSON) serve a trovare le pagine; i valori si leggono dalle parole del PDF con la loro
posizione, perché la fonte lascia vuote le celle a zero e senza posizione non si sa di che colonna sia
un numero.
"""
import glob, json, os, re, sys
import pandas as pd
import pdfplumber

PDF_DIR, TXT_DIR, OUT = sys.argv[1:4]

MESI = {m: i for i, m in enumerate(['gennaio', 'febbraio', 'marzo', 'aprile', 'maggio', 'giugno', 'luglio',
                                    'agosto', 'settembre', 'ottobre', 'novembre', 'dicembre'], 1)}
NUM = re.compile(r'^\d{1,3}(?:\.\d{3})*(?:,\d+)?$')
SEDI = ['Torino', 'Pontremoli (MS)', 'Milano', 'Rovigo', 'Treviso', 'Bologna', 'Bologna Sezione Dozza', 'Firenze', 'Roma', "L'Aquila", 'Lecce',
        'Nisida (NA)', 'Airola (BN)', 'Bari', 'Potenza', 'Catanzaro', 'Palermo', 'Catania', 'Acireale (CT)',
        'Caltanissetta', 'Quartucciu (CA)']


def n(s):
    s = s.replace('.', '').replace(',', '.')
    return float(s) if '.' in s else int(s)


def linee(page):
    """Righe della pagina: (etichetta, [(x1, valore)])."""
    righe = []
    for w in sorted(page.extract_words(), key=lambda w: w['top']):
        if righe and w['top'] - righe[-1][0]['top'] <= 3:
            righe[-1].append(w)
        else:
            righe.append([w])
    out = []
    for r in righe:
        ws = sorted(r, key=lambda w: w['x0'])
        nums = [(w['x1'], n(w['text'])) for w in ws if NUM.match(w['text'])]
        label = ' '.join(w['text'] for w in ws if not NUM.match(w['text']))
        out.append((label, nums))
    return out


def blocco(page, titolo):
    """Righe dalla riga del titolo della tabella fino alla tabella successiva."""
    rr = linee(page)
    start = next(i for i, (l, _) in enumerate(rr) if re.search(titolo, l))
    end = next((i for i in range(start + 1, len(rr)) if rr[i][0].startswith('Tabella')), len(rr))
    return rr[start:end]


def tabella(righe, etichette, ncol, check):
    """Valori per etichetta; i numeri mancanti (zeri omessi) si collocano per posizione orizzontale."""
    piene = [nums for _, nums in righe if len(nums) == ncol]
    colx = [sorted(r[i][0] for r in piene)[len(piene) // 2] for i in range(ncol)]
    out = {}
    for k, (lab, nums) in enumerate(righe):
        lab = re.sub(r'\s*\(\*+\)$', '', lab).strip()
        # etichetta con note (*) a volte leggermente sfalsata: i numeri finiscono su una riga a sé
        if not nums and k + 1 < len(righe) and righe[k + 1][0] == '':
            nums = righe[k + 1][1]
        if lab not in etichette or lab in out or not nums:
            continue
        v = [0] * ncol
        for x, val in nums:
            v[min(range(ncol), key=lambda i: abs(colx[i] - x))] = val
        if not check(v):
            raise ValueError(f'{lab}: {nums} -> {v}')
        out[lab] = v
    return out


def ok9(v):
    return v[0] + v[1] == v[2] and v[3] + v[4] == v[5] and v[6] + v[7] == v[8] and v[2] + v[5] == v[8]


def ok_sedi(v):
    return v[0] + v[1] == v[2]


def ok_eta(v):
    return sum(v[:4]) == v[4]


def ok3(v):
    return v[0] + v[1] == v[2]


def trova(pages, pattern):
    return next(i for i, t in enumerate(pages) if re.search(pattern, t, re.S))


riep, sedi, eta_sede, eta_naz, pos = [], [], [], [], []
for f in sorted(glob.glob(os.path.join(TXT_DIR, '*.json'))):
    P = json.load(open(f))
    fonte = os.path.basename(f)[:-5]
    m = re.search(r'riferiti alla data del (\d{1,2}) (\w+) (\d{4})', P[0], re.I)
    d = pd.Timestamp(int(m.group(3)), MESI[m.group(2).lower()], int(m.group(1)))
    print(fonte, d.date(), file=sys.stderr)
    pdf = pdfplumber.open(os.path.join(PDF_DIR, fonte))

    pg = pdf.pages[trova(P, r'Dati di riepilogo')]
    voci = {'Centri di prima accoglienza': 'cpa', 'Istituti penali per i minorenni': 'ipm',
            'Comunità ministeriali': 'com_min', 'Comunità private': 'com_priv',
            'Totale presenti alla data considerata': 'tot_res',
            'Totale soggetti in carico alla data considerata': 'ussm'}
    t = tabella(linee(pg), set(voci), 3, ok3)
    row = {'data': d, 'fonte': fonte}
    for lab, key in voci.items():
        row[key + '_m'], row[key + '_f'], row[key] = t[lab]
    riep.append(row)

    pg = pdf.pages[trova(P, r'Ingressi e presenze negli Istituti penali per i minorenni')]
    t = tabella(blocco(pg, r'Ingressi e presenze negli Istituti penali'), set(SEDI + ['Totale']), 5, ok_sedi)
    for s, v in t.items():
        sedi.append({'data': d, 'ipm': s, 'ingressi_stabili': v[0], 'trasferimenti_da_ipm': v[1],
                     'ingressi_totali': v[2], 'presenza_media': v[3], 'presenti': v[4]})

    pg = pdf.pages[trova(P, r'secondo la sede dell.IPM e\s*l.età')]
    t = tabella(blocco(pg, r'secondo la sede dell'), set(SEDI + ['Totale']), 5, ok_eta)
    for s, v in t.items():
        eta_sede.append({'data': d, 'ipm': s, 'eta_14_15': v[0], 'eta_16_17': v[1], 'eta_18_20': v[2],
                         'eta_21_24': v[3], 'totale': v[4]})

    pg = pdf.pages[trova(P, r'Detenuti presenti negli IPM alla data del .{0,60}secondo\s+l.età,\s+la\s+nazionalità')]
    labs = ['14-15 anni', '16-17 anni', '18-20 anni', '21-24 anni', 'Totale']
    t = tabella(blocco(pg, r'^Tabella.*l.età, la'), set(labs), 9, ok9)
    for lab in labs:
        v = t[lab]
        eta_naz.append({'data': d, 'eta': lab.replace(' anni', ''), 'ita_m': v[0], 'ita_f': v[1], 'ita': v[2],
                        'str_m': v[3], 'str_f': v[4], 'str': v[5], 'tot_m': v[6], 'tot_f': v[7], 'tot': v[8]})

    pg = pdf.pages[trova(P, r'posizione giuridica complessiva')]
    rr = linee(pg)
    gruppi = [i for i, (l, _) in enumerate(rr) if l in ('Minorenni.', 'Giovani adulti.', 'Totale.')]
    labs = ['In attesa di 1° giudizio', 'Appellanti', 'Ricorrenti', 'Mista senza definitivo',
            'Definitivi senza altri procedimenti a carico', 'Definitivi con altri procedimenti a carico', 'Totale']
    for gi, start in enumerate(gruppi):
        end = gruppi[gi + 1] if gi + 1 < len(gruppi) else len(rr)
        t = tabella(rr[start + 1:end], set(labs), 9, ok9)
        for lab in labs:
            v = t[lab]
            pos.append({'data': d, 'gruppo': rr[start][0].rstrip('.'), 'posizione': lab,
                        'ita': v[2], 'str': v[5], 'tot': v[8]})
    pdf.close()

os.makedirs(OUT, exist_ok=True)
for name, rows in [('riepilogo', riep), ('sedi', sedi), ('eta_sede', eta_sede), ('eta_naz', eta_naz),
                   ('posizione', pos)]:
    df = pd.DataFrame(rows).sort_values('data', kind='stable')
    df.to_csv(os.path.join(OUT, name + '.csv'), index=False)
    print(name, len(df), df['data'].nunique())
