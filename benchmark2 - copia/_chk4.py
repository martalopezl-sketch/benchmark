# -*- coding: utf-8 -*-
import re, sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import Benchmark as B

raw = B.get_web_text('https://www.energianufri.com/es/tarifas-luz')
print("### NUFRI powerPrice / energyPrice (raw, 160 chars) ###")
for m in re.finditer(r'powerPrice', raw):
    print("  POWER:", raw[m.start():m.start()+150]);
for m in re.finditer(r'energyPrice\\?":', raw):
    print("  ENERGY:", raw[m.start():m.start()+120])
# tambien buscar los productos con su nombre + precios
for m in re.finditer(r'(sin mirar el reloj|reloj|24)', raw, re.I):
    a=m.start()
    seg = raw[a:a+400]
    if 'ric' in seg.lower():
        print("  SEG:", re.sub(r'\s+',' ',seg)[:220]); break
