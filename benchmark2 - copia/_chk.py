# -*- coding: utf-8 -*-
import re, sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import Benchmark as B

def look(name, url, needles):
    print("="*80)
    print(f"### {name}  {url}")
    raw = B.get_web_text(url)
    if not raw:
        print("  NO DESCARGA"); return None
    t = B.strip_tags(raw)
    print(f"  len_raw={len(raw)} len_txt={len(t)}")
    # kWh / kW contexts
    for kw in ['kWh', 'kW']:
        seen=0
        for m in re.finditer(kw, t):
            a=max(0,m.start()-60); b=min(len(t),m.end()+18)
            print(f"  [{kw}] ...{t[a:b].strip()}...")
            seen+=1
            if seen>=4: break
    for n in needles:
        c=raw.count(n)
        if c:
            i=raw.find(n); print(f"  RAW '{n}' x{c}: ...{raw[max(0,i-50):i+40]}...")
    # links a posibles paginas de tarifa
    links = set(re.findall(r'href="([^"]*(?:tarifa|precio|luz|hogar)[^"]*)"', raw, re.I))
    print("  LINKS tarifa:", list(links)[:10])
    return raw, t

look("HSC", "https://www.hscomercializadora.com/", ['kWh','0,1','0.1'])
look("LUMISA", "https://lumisa.es/es/tarifa-fija", ['Energía','Potencia','24','único','unico'])
look("NUFRI", "https://www.energianufri.com/es/tarifas-luz", ['mirar el reloj','0,099','potencia','Potencia','/mes','/día'])
look("ENDESA", "https://www.endesa.com/es/luz-y-gas/luz/one/tarifa-one-luz", ['0,1','precio','Precio','kWh','taxless','sinImp','energyPrice'])
