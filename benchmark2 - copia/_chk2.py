# -*- coding: utf-8 -*-
import re, sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import Benchmark as B

print("############ HSC /tarifas_precios ############")
raw = B.get_web_text('https://www.hscomercializadora.com/tarifas_precios')
if raw:
    t = B.strip_tags(raw)
    print("  len_txt=", len(t))
    for kw in ['kWh','kW']:
        for i,m in enumerate(re.finditer(kw, t)):
            if i>=5: break
            a=max(0,m.start()-70); b=min(len(t),m.end()+18)
            print(f"  [{kw}] ...{t[a:b].strip()}...")
else:
    print("  no descarga")

print("\n############ ENDESA: todos los bloques de precio ############")
raw = B.get_web_text('https://www.endesa.com/es/luz-y-gas/luz/one/tarifa-one-luz')
for m in re.finditer(r'T\.\s*Potencia[\s\S]{0,140}?T\.\s*Energ.a:\s*[0-9][.,][0-9]+', raw):
    print("  BLOQUE:", re.sub(r'\s+',' ', m.group(0)))
for m in re.finditer(r'[0-9][.,][0-9]{3,6}\s*\\?u20ac?\s*/?\s*kWh', raw):
    a=max(0,m.start()-60); print("   kWh:", re.sub(r'\s+',' ',raw[a:m.end()]))

print("\n############ NUFRI: potencia numerica en JSON ############")
raw = B.get_web_text('https://www.energianufri.com/es/tarifas-luz')
for pat in [r'powerPrice[^,}]{0,40}', r'"power[^"]*":[^,}]{0,30}', r'potencia[^,}]{0,40}[0-9][.,][0-9]{2,6}',
            r'[0-9][.,][0-9]{4,6}[^0-9]{0,8}(?:kW|/mes|/d)']:
    ms=list(re.finditer(pat, raw, re.I))
    print(f"  '{pat}': {len(ms)}")
    for m in ms[:4]:
        print("     ", re.sub(r'\s+',' ',m.group(0))[:70])

print("\n############ LUMISA: Simple 24h precio unico ############")
raw = B.get_web_text('https://lumisa.es/es/tarifa-fija')
t = B.strip_tags(raw)
for m in re.finditer(r'Simple\s*24h', t, re.I):
    a=max(0,m.start()-20); b=min(len(t),m.end()+160)
    print("  ...", t[a:b].strip()[:190]); break
# buscar un precio unico energia (mismo valor) cerca de 'Simple 24h'
for m in re.finditer(r'24\s*h[\s\S]{0,120}?([0-9][.,][0-9]{4,6})\s*.{0,4}/?\s*kWh', t, re.I):
    print("  24h->kWh:", m.group(1)); break
