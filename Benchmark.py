#!/usr/bin/env python3
"""
VERSIÓN FINAL - DATOS REALES DEL API + WEB SCRAPING
- 25 empresas: Extrae del API CNMC (DATOS REALES)
- 3 empresas: Extrae de sus webs (DATOS REALES)
- Actualización semanal automática con comparativas
"""

import subprocess
import sys
subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "pandas", "openpyxl", "requests", "--upgrade"])

import re
from datetime import datetime
from pathlib import Path
import pandas as pd
import requests
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

try:
    from google.colab import files
    IN_COLAB = True
    print("✅ Google Colab detectado\n")
except:
    IN_COLAB = False
    print("⚠️ Ejecución local\n")

BASE_URL = 'https://comparador.cnmc.gob.es/api/publico'
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
}

TARGET_OFFERS = {
    'ENERGYA VM': {'brand': 'Energya VM', 'search': 'Fórmula Fija Única'},
    'DOMESTICA GAS': {'brand': 'Visalia', 'search': 'Visalia Luz Fijo'},
    'NATURGY': {'brand': 'Naturgy', 'search': 'Por Uso Luz'},
    'REPSOL': {'brand': 'Repsol', 'search': 'Tarifa Ahorro Potencia'},
    'OCTOPUS': {'brand': 'Octopus', 'search': 'OCTOPUS RELAX'},
    'NIBA': {'brand': 'Niba', 'search': 'niba Zen'},
    'ENERGIA NUFRI': {'brand': 'Energia Nufri', 'search': 'CALMA'},
    'GAOLANIA': {'brand': 'Gana Energía', 'search': 'Tarifa 24 horas'},
    'CLEARVIEW': {'brand': 'Clarity Energy', 'search': 'CLARITY ENERGY'},
    'CIDE': {'brand': 'CHC Energía', 'search': 'Ilumina'},
    'ENERGYASSET': {'brand': 'Energy Asset', 'search': 'Tarifa Plana'},
    'CATGAS': {'brand': 'Catgas', 'search': '2.0TDL'},
    'TELECOR': {'brand': 'El Corte Inglés', 'search': 'Despreocúpate'},
    'ENDESA': {'brand': 'Endesa', 'search': 'Luz Fija 24h'},
    'FENIE ENERGIA': {'brand': 'Fenie Energía', 'search': 'Fijo Energético'},
    'GESTERNOVA': {'brand': 'Contigo Energía', 'search': 'Tarifa Facil'},
    'DISA ENERGIA': {'brand': 'Disa Energía', 'search': 'ALISIOS'},
    'HIDROELECTRICA': {'brand': 'HSC', 'search': 'Eficiente'},
    'LUMISA': {'brand': 'Lumisa Energía', 'search': '2.0TD'},
    'TOTALENERGIES': {'brand': 'Total Energies', 'search': 'TU AIRE'},
    'WEKIWI': {'brand': 'Wekiki', 'search': 'MariCalmen'},
    'PLENITUDE': {'brand': 'Plenitude', 'search': 'Tarifa Fácil Plus'},
    'IMAGINA': {'brand': 'Imagina', 'search': 'PLAN BASE'},
    'IBERDROLA CLIENTES': {'brand': 'Iberdrola', 'search': 'Plan Online'},
    'NEXUS ENERGIA': {'brand': 'Nexus', 'search': 'Luz Eficiente'},
}

def build_params():
    from datetime import timedelta
    today = datetime.now()
    fin = today.replace(day=1) - timedelta(days=1)
    inicio = fin.replace(day=1)
    ts_ini = int(inicio.timestamp() * 1000)
    ts_fin = int(fin.timestamp() * 1000)

    return {
        'tipoSuministro': 'E', 'codigoPostal': 28003, 'potencia': 4,
        'potenciaPrimeraFranja': 4, 'potenciaSegundaFranja': 4,
        'potenciaTerceraFranja': 4, 'potenciaCuartaFranja': 4,
        'potenciaQuintaFranja': 4, 'potenciaSextaFranja': 4,
        'consumoAnualE': 210, 'consumoAnualEOrig': 2600,
        'consumoPrimeraFranja': 61, 'consumoSegundaFranja': 51,
        'consumoTerceraFranja': 98, 'consumoCuartaFranja': 0,
        'consumoQuintaFranja': 0, 'consumoSextaFranja': 0,
        'consumoAnualEQr': 0, 'consumoPrimeraFranjaQr': 0,
        'consumoSegundaFranjaQr': 0, 'consumoTerceraFranjaQr': 0,
        'consumoCuartaFranjaQr': 0, 'consumoQuintaFranjaQr': 0,
        'consumoSextaFranjaQr': 0, 'consumoAnualEPQr': 0,
        'consumoPrimeraFranjaPQr': 0, 'consumoSegundaFranjaPQr': 0,
        'consumoTerceraFranjaPQr': 0, 'consumoCuartaFranjaPQr': 0,
        'consumoQuintaFranjaPQr': 0, 'consumoSextaFranjaPQr': 0,
        'tarifa': 4, 'consumoAnualG': 491, 'consumoAnualGOrig': 6000,
        'serviciosAdicionales': 2, 'permanencia': 2, 'idOferta': 0,
        'vivienda': 'true', 'factura': 'true',
        'energiaAutoconsumo': 0, 'idAuditoriaQR': 0,
        'potenciaAutoconsumo': 3.5, 'revisionPrecios': 2, 'importe': 0,
        'dateInicio': ts_ini, 'dateFin': ts_fin,
        'tc': 0, 'bs': 0, 'impSA': 0, 'impOtros': 0, 'exc': 0, 'reg': 0,
        'mecanismoAjuste': 0, 'importeMecanismoAjustePunta': 0,
        'importeMecanismoAjusteLlano': 0, 'importeMecanismoAjusteValle': 0,
        'precioConsumoMecanismoAjustePunta': 0, 'precioConsumoMecanismoAjusteLlano': 0,
        'precioConsumoMecanismoAjusteValle': 0, 'precioConsumoMecanismoAjusteTotal': 0,
        'mecanismoAjusteIVA': 0, 'impOtrosConIE': 0, 'impOtrosSinIE': 0,
        'pmaxP1': 0, 'pmaxP2': 0, 'fFact': ts_fin,
        'dtoBS': 0, 'finBS': 0, 'ajuste': 0, 'impPot': 0, 'impEner': 0, 'dto': 0,
        'prP1': 0, 'prP2': 0, 'prE1': 0, 'prE2': 0, 'prE3': 0,
        'cfP1flex': 0, 'cfP2flex': 0, 'cambio': 0, 'promo': 0, 'verde': 0,
        'rev': 0, 'trampeo': 0,
        'perfilConsumo': 13, 'cups': '0000', 'autoconsumo': 'false',
    }

def extract_prices(texto):
    """Extrae precios del API CNMC"""
    if not texto:
        return None, None, None

    texto = str(texto).replace('\xa0', ' ')
    energia = None
    p1 = None
    p2 = None

    try:
        m = re.search(r'([0-9]{1,3}[.,][0-9]{1,6})\s*€\s*/\s*kWh', texto, re.IGNORECASE)
        if m:
            val = float(m.group(1).replace(',', '.'))
            if 0.08 <= val <= 0.20:
                energia = val

        patterns_p1 = [
            r'[Pp]unta\s+([0-9]{1,3}[.,][0-9]{1,6})\s*€',
            r'[Pp]eriodo\s+1[:\s]+([0-9]{1,3}[.,][0-9]{1,6})\s*€',
            r'P1\s*[=:]\s*([0-9]{1,3}[.,][0-9]{1,6})\s*€',
            r'[Tt]érmino\s+[Pp]otencia\s+[Pp]unta[:\s]+([0-9]{1,3}[.,][0-9]{1,6})',
        ]

        for pattern in patterns_p1:
            m = re.search(pattern, texto)
            if m:
                val = float(m.group(1).replace(',', '.'))
                if 15 <= val <= 75:
                    p1 = val
                    break

        patterns_p2 = [
            r'[Vv]alle\s+([0-9]{1,3}[.,][0-9]{1,6})\s*€',
            r'[Pp]eriodo\s+2[:\s]+([0-9]{1,3}[.,][0-9]{1,6})\s*€',
            r'P2\s*[=:]\s*([0-9]{1,3}[.,][0-9]{1,6})\s*€',
            r'[Tt]érmino\s+[Pp]otencia\s+[Vv]alle[:\s]+([0-9]{1,3}[.,][0-9]{1,6})',
        ]

        for pattern in patterns_p2:
            m = re.search(pattern, texto)
            if m:
                val = float(m.group(1).replace(',', '.'))
                if 15 <= val <= 75:
                    p2 = val
                    break

        if p1 and not p2:
            p2 = p1

    except:
        pass

    return energia, p1, p2

def scrape_holaluz():
    """Extrae datos REALES de Holaluz"""
    try:
        r = requests.get('https://www.holaluz.com/luz/tarifas-luz/', timeout=10, headers=HEADERS, verify=False)
        r.raise_for_status()
        text = r.text

        m_energia = re.search(r'Tarifa\s+Clásica.*?Precio\s+24\s+horas\s+([\d,\.]+)\s*€/kWh', text, re.IGNORECASE | re.DOTALL)
        m_potencia = re.search(r'Tarifa\s+Clásica.*?P1\s+([\d,\.]+)\s*€/kW', text, re.IGNORECASE | re.DOTALL)

        if m_energia and m_potencia:
            energia = float(m_energia.group(1).replace(',', '.'))
            potencia = float(m_potencia.group(1).replace(',', '.')) * 365

            return energia, potencia, potencia
    except:
        pass

    return None, None, None

def scrape_podo():
    """Extrae datos REALES de Podo"""
    try:
        r = requests.get('https://www.mipodo.com/tarifas-luz', timeout=10, headers=HEADERS, verify=False)
        r.raise_for_status()
        text = r.text

        m_energia = re.search(r'Tarifa\s+Luz\s+Precio\s+Único.*?24h\s+([\d,\.]+)\s*€\s*kWh', text, re.IGNORECASE | re.DOTALL)
        m_p1 = re.search(r'Tarifa\s+Luz\s+Precio\s+Único.*?P1\s+([\d,\.]+)\s*€\s*kW/día', text, re.IGNORECASE | re.DOTALL)
        m_p2 = re.search(r'Tarifa\s+Luz\s+Precio\s+Único.*?P2\s+([\d,\.]+)\s*€\s*kW/día', text, re.IGNORECASE | re.DOTALL)

        if m_energia and m_p1 and m_p2:
            energia = float(m_energia.group(1).replace(',', '.'))
            p1 = float(m_p1.group(1).replace(',', '.')) * 365
            p2 = float(m_p2.group(1).replace(',', '.')) * 365

            return energia, p1, p2
    except:
        pass

    return None, None, None

def scrape_factor():
    """Extrae datos REALES de Factor Energía"""
    try:
        r = requests.get('https://www.factorenergia.com/es/luz/tarifa-fija-de-luz-precio-unico/', timeout=10, headers=HEADERS, verify=False)
        r.raise_for_status()
        text = r.text

        m_energia = re.search(r'(0[,\.]\d{3,4})\s*€/kWh', text)
        m_potencia = re.search(r'([\d,\.]+)\s*€/kW\s*(?:día|day)', text)

        if m_energia and m_potencia:
            energia = float(m_energia.group(1).replace(',', '.'))
            potencia = float(m_potencia.group(1).replace(',', '.')) * 365

            return energia, potencia, potencia
    except:
        pass

    return None, None, None

def format_header(ws, num_cols=4):
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                   top=Side(style='thin'), bottom=Side(style='thin'))

    for col in range(1, num_cols + 1):
        cell = ws.cell(1, col)
        cell.fill = header_fill
        cell.font = header_font
        cell.border = border
        cell.alignment = Alignment(horizontal='center', vertical='center')

def format_data_rows(ws, num_cols=4):
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                   top=Side(style='thin'), bottom=Side(style='thin'))

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=num_cols):
        for col_idx, cell in enumerate(row, 1):
            cell.border = border
            cell.alignment = Alignment(horizontal='left' if col_idx == 1 else 'right', vertical='center')
            if col_idx == 2:
                cell.number_format = '0.0000'
            elif col_idx in [3, 4]:
                cell.number_format = '0.00'

    ws.column_dimensions['A'].width = 22
    ws.column_dimensions['B'].width = 18
    ws.column_dimensions['C'].width = 18
    ws.column_dimensions['D'].width = 18

print("="*70)
print("CNMC BENCHMARK ENERGÍA - DATOS REALES (API + WEB SCRAPING)")
print("="*70 + "\n")

fecha_hoy = datetime.now().strftime('%Y-%m-%d')
num_semana = datetime.now().strftime('%W')
nombre_pestaña = f"Semana {num_semana} ({fecha_hoy})"

print(f"📅 Fecha: {fecha_hoy}")
print(f"🔌 Extrayendo 25 empresas del API CNMC...\n")

registros = []
params = build_params()

try:
    r = requests.get(f'{BASE_URL}/ofertas/electricidad', params=params, headers=HEADERS, timeout=30, verify=False)
    ofertas = r.json().get('resultadoComparador', [])
    print(f"✓ {len(ofertas)} ofertas encontradas\n")
except Exception as e:
    print(f"❌ Error: {e}\n")
    ofertas = []

by_brand = {}
for legal_name, config in TARGET_OFFERS.items():
    brand = config['brand']
    search_str = config['search']
    empresa_ofertas = [o for o in ofertas if o.get('tienePrecioUnico') == 'S' and legal_name in o.get('comercializadora', '').upper()]
    if empresa_ofertas:
        oferta = next((o for o in empresa_ofertas if search_str.lower() in o.get('oferta', '').lower()), empresa_ofertas[0])
        by_brand[brand] = oferta

for marca in sorted(by_brand.keys()):
    oferta = by_brand[marca]
    try:
        p = {**params, 'idOferta': oferta['id']}
        r = requests.get(f'{BASE_URL}/oferta', params=p, headers=HEADERS, timeout=30, verify=False)
        detalle = r.json()
        caract = detalle.get('caracteristicas', {})

        if isinstance(caract, dict):
            caracteristicas_text = caract.get('caracteristicas', '')
        else:
            caracteristicas_text = str(caract)

        energia, pot_p1, pot_p2 = extract_prices(caracteristicas_text)

        if energia and pot_p1 and pot_p2:
            print(f"  ✓ {marca:25} E:{energia:.4f} | P1:{pot_p1:.2f} | P2:{pot_p2:.2f}")
        else:
            print(f"  ⚠️  {marca:25} (parcial)")

        registros.append({
            'Empresa': marca,
            'Precio energía (€/kWh)': energia,
            'P1 (€/kW/año)': pot_p1,
            'P2 (€/kW/año)': pot_p2
        })

    except Exception as e:
        registros.append({'Empresa': marca, 'Precio energía (€/kWh)': None, 'P1 (€/kW/año)': None, 'P2 (€/kW/año)': None})

print(f"\n🌐 Extrayendo 3 empresas de WEB SCRAPING...\n")

web_scrapers = {
    'Holaluz': scrape_holaluz,
    'Podo': scrape_podo,
    'Factor Energía': scrape_factor,
}

for empresa, scraper_func in web_scrapers.items():
    energia, p1, p2 = scraper_func()
    registros.append({
        'Empresa': empresa,
        'Precio energía (€/kWh)': energia,
        'P1 (€/kW/año)': p1,
        'P2 (€/kW/año)': p2
    })

    if energia and p1 and p2:
        print(f"  ✓ {empresa:25} E:{energia:.4f} | P1:{p1:.2f} | P2:{p2:.2f}")
    else:
        print(f"  ⚠️  {empresa:25} (no extraído)")

df_actual = pd.DataFrame(registros)

out = Path('Analisis_Energia_CNMC.xlsx')

print(f"\n💾 Actualizando Excel...\n")

if out.exists():
    wb = load_workbook(out)

    if nombre_pestaña in wb.sheetnames:
        del wb[nombre_pestaña]

    ws = wb.create_sheet(nombre_pestaña, 0)
    ws['A1'] = 'Empresa'
    ws['B1'] = 'Precio energía (€/kWh)'
    ws['C1'] = 'P1 (€/kW/año)'
    ws['D1'] = 'P2 (€/kW/año)'

    for row_idx, record in enumerate(df_actual.values, 2):
        ws[f'A{row_idx}'] = record[0]
        ws[f'B{row_idx}'] = record[1]
        ws[f'C{row_idx}'] = record[2]
        ws[f'D{row_idx}'] = record[3]

    format_header(ws, 4)
    format_data_rows(ws, 4)

    print(f"✅ Pestaña '{nombre_pestaña}' creada\n")

    semanas = [s for s in wb.sheetnames if s.startswith('Semana') and s != nombre_pestaña]

    if semanas:
        semana_anterior = semanas[0]
        print(f"📊 Creando comparativa vs '{semana_anterior}'...\n")

        ws_anterior = wb[semana_anterior]
        datos_anterior = {}
        for row in ws_anterior.iter_rows(min_row=2, values_only=True):
            empresa = row[0]
            datos_anterior[empresa] = {'energia': row[1], 'p1': row[2], 'p2': row[3]}

        nombre_comparativa = f"Comparativa {fecha_hoy}"
        if nombre_comparativa in wb.sheetnames:
            del wb[nombre_comparativa]

        ws_comp = wb.create_sheet(nombre_comparativa, 1)

        ws_comp['A1'] = 'Empresa'
        ws_comp['B1'] = 'Energía Actual'
        ws_comp['C1'] = 'Energía Anterior'
        ws_comp['D1'] = 'Dif (€/kWh)'
        ws_comp['E1'] = 'Cambio %'
        ws_comp['F1'] = 'P1 Actual'
        ws_comp['G1'] = 'P1 Anterior'
        ws_comp['H1'] = 'Dif (€/kW/año)'
        ws_comp['I1'] = 'Cambio %'
        ws_comp['J1'] = 'P2 Actual'
        ws_comp['K1'] = 'P2 Anterior'
        ws_comp['L1'] = 'Dif (€/kW/año)'
        ws_comp['M1'] = 'Cambio %'

        row_idx = 2
        for _, record in enumerate(df_actual.values):
            empresa = record[0]
            if empresa in datos_anterior:
                datos_ant = datos_anterior[empresa]
                energia_actual, p1_actual, p2_actual = record[1], record[2], record[3]
                energia_ant, p1_ant, p2_ant = datos_ant['energia'], datos_ant['p1'], datos_ant['p2']

                dif_e = energia_actual - energia_ant if energia_ant else None
                pct_e = (dif_e / energia_ant * 100) if energia_ant and dif_e else None
                dif_p1 = p1_actual - p1_ant if p1_ant else None
                pct_p1 = (dif_p1 / p1_ant * 100) if p1_ant and dif_p1 else None
                dif_p2 = p2_actual - p2_ant if p2_ant else None
                pct_p2 = (dif_p2 / p2_ant * 100) if p2_ant and dif_p2 else None

                ws_comp[f'A{row_idx}'] = empresa
                ws_comp[f'B{row_idx}'] = energia_actual
                ws_comp[f'C{row_idx}'] = energia_ant
                ws_comp[f'D{row_idx}'] = dif_e
                ws_comp[f'E{row_idx}'] = pct_e
                ws_comp[f'F{row_idx}'] = p1_actual
                ws_comp[f'G{row_idx}'] = p1_ant
                ws_comp[f'H{row_idx}'] = dif_p1
                ws_comp[f'I{row_idx}'] = pct_p1
                ws_comp[f'J{row_idx}'] = p2_actual
                ws_comp[f'K{row_idx}'] = p2_ant
                ws_comp[f'L{row_idx}'] = dif_p2
                ws_comp[f'M{row_idx}'] = pct_p2

                row_idx += 1

        format_header(ws_comp, 13)
        format_data_rows(ws_comp, 13)

        print(f"✅ Comparativa creada\n")

    wb.save(out)
    print(f"📊 Pestañas: {', '.join(wb.sheetnames)}\n")

else:
    wb = Workbook()
    if 'Sheet' in wb.sheetnames:
        del wb['Sheet']

    ws = wb.create_sheet('Benchmark', 0)
    ws['A1'] = 'Empresa'
    ws['B1'] = 'Precio energía (€/kWh)'
    ws['C1'] = 'P1 (€/kW/año)'
    ws['D1'] = 'P2 (€/kW/año)'

    for row_idx, record in enumerate(df_actual.values, 2):
        ws[f'A{row_idx}'] = record[0]
        ws[f'B{row_idx}'] = record[1]
        ws[f'C{row_idx}'] = record[2]
        ws[f'D{row_idx}'] = record[3]

    format_header(ws, 4)
    format_data_rows(ws, 4)

    wb.save(out)
    print(f"✅ Archivo creado\n")

print(f"✅ Archivo guardado: Analisis_Energia_CNMC.xlsx")
print(f"📋 Total empresas: {len(registros)}\n")

if IN_COLAB:
    print("📥 Descargando archivo...\n")
    try:
        files.download(str(out))
        print("✅ ¡Descarga completada!")
    except Exception as e:
        print(f"❌ Error: {e}")
else:
    print(f"📂 Archivo: {out.absolute()}")