import pdfplumber, sys, glob, os, json
src = sys.argv[1]; out = sys.argv[2]
os.makedirs(out, exist_ok=True)
for f in sorted(glob.glob(os.path.join(src, 'Analisi_*.pdf'))):
    with pdfplumber.open(f) as p:
        pages = [(pg.extract_text() or '') for pg in p.pages]
    json.dump(pages, open(os.path.join(out, os.path.basename(f) + '.json'), 'w'))
    print(os.path.basename(f), len(pages))
