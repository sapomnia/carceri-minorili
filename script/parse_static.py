"""Tabelle IPM dell'ultima edizione DGMC (30/09/2026): motivi di ingresso, età, provenienza, delitti, uscite,
serie storiche 2016-2025. Le righe con etichetta su due linee vengono ricomposte."""
import json, re, sys
import pandas as pd

TXT, OUT = sys.argv[1:3]
P = json.load(open(TXT))
NUM = r'\d{1,3}(?:\.\d{3})*'


def n(s):
    return int(s.replace('.', ''))


def tab(numero, ncol):
    t = next(p for p in P if re.search(rf'Tabella {numero} ', p))
    t = t[t.index(f'Tabella {numero} '):]
    nxt = re.search(r'\nTabella \d+ |\nGrafico |\nLa tabella riporta|\nI dati sono', t[10:])
    t = t[:nxt.start() + 10] if nxt else t
    righe, sospesa, sezione = [], '', ''
    for line in t.split('\n')[1:]:
        m = re.match(rf'^(.*?)\s*((?:{NUM}\s+){{{ncol - 1}}}{NUM})\s*$', line)
        if m:
            lab = (sospesa + ' ' + m.group(1)).strip()
            v = [n(x) for x in m.group(2).split()]
            righe.append((sezione, lab, v))
            sospesa = ''
        elif righe and line.strip() == 'misura':
            # coda di un'etichetta andata a capo dopo i numeri
            righe[-1] = (righe[-1][0], righe[-1][1] + ' misura', righe[-1][2])
        elif re.match(r'^(Per custodia cautelare|Per esecuzione di pena|Da custodia cautelare|Da espiazione pena|'
                      r'Trasferimenti a strutture per adulti|DELITTI|CONTRAVVENZIONI)$', line.strip()):
            sezione = line.strip()
        elif not re.search(r'\d', line) and not re.match(r'^(Italiani|Motivo|Età|Anni|Sesso|Paesi|maschi|m f|REATI|'
                                                         r'Motivi|Anno|nazionalità|e il sesso|- fino)', line.strip()):
            sospesa = (sospesa + ' ' + line.strip()).strip()
    return righe


def controlla9(righe, nome):
    for s, lab, v in righe:
        ok = v[0] + v[1] == v[2] and v[3] + v[4] == v[5] and v[6] + v[7] == v[8] and v[2] + v[5] == v[8]
        if not ok:
            print(f'ATTENZIONE {nome}: totali non coerenti nella fonte -> {lab} {v}', file=sys.stderr)


out = {}
for num, nome in [(25, 'motivi_ingresso'), (26, 'eta_ingressi'), (28, 'delitti_ingressi'), (29, 'uscite'),
                  (30, 'serie_ingressi'), (31, 'serie_presenza_media')]:
    r = tab(num, 9)
    controlla9(r, nome)
    out[nome] = pd.DataFrame([{'sezione': s, 'voce': l, 'ita_m': v[0], 'ita_f': v[1], 'ita': v[2], 'str_m': v[3],
                               'str_f': v[4], 'str': v[5], 'tot_m': v[6], 'tot_f': v[7], 'tot': v[8]}
                              for s, l, v in r])
r = tab(23, 3)
out['provenienza'] = pd.DataFrame([{'voce': l, 'm': v[0], 'f': v[1], 'tot': v[2]} for _, l, v in r])
for k, df in out.items():
    df.to_csv(f'{OUT}/{k}.csv', index=False)
    print(k, len(df))
    print(df[[c for c in df.columns if c in ('sezione', 'voce', 'tot')]].to_string(), '\n')
