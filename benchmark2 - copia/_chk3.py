# -*- coding: utf-8 -*-
import re, sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import Benchmark as B

print("############ NUFRI: objetos energyPrice / powerPrice ############")
raw = B.get_web_text('https://www.energianufri.com/es/tarifas-luz')
u = raw.encode().decode('unicode_escape', 'ignore') if '\\u' in raw else raw
for pat in [r'energyPrice"?:\{[^}]{0,120}\}', r'powerPrice"?:\{[^}]{0,120}\}']:
    ms = list(re.finditer(pat, raw))
    print(f"\n  '{pat[:20]}': {len(ms)}")
    for m in ms[:6]:
        print("   ", m.group(0).replace('\\"','"'))

print("\n############ LUMISA: contexto de 0,125141 ############")
raw = B.get_web_text('https://lumisa.es/es/tarifa-fija')
t = B.strip_tags(raw)
for m in re.finditer(r'0[.,]125141', t):
    a=max(0,m.start()-90); b=min(len(t),m.end()+30)
    print("  ...", t[a:b].strip()); break
# ver si Simple 24h y su precio unico estan juntos
m = re.search(r'Simple\s*24h[\s\S]{0,500}?([0-9][.,]125141|0[.,]125141|[0-9][.,][0-9]{5,6})\s*.{0,5}/?\s*kWh', t, re.I)
print("  Simple24h-first-kWh:", m.group(1) if m else "no")

print("\n############ ENDESA: hay precio CON descuento? ############")
raw = B.get_web_text('https://www.endesa.com/es/luz-y-gas/luz/one/tarifa-one-luz')
for kw in ['0,163097','descuento','0,1282','0,128','con descuento','dto']:
    i = raw.find(kw)
    if i>=0:
        print(f"  [{kw}] ...{re.sub(chr(92)+'s+',' ',raw[max(0,i-70):i+50])}...")
