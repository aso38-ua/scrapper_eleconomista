import requests
from bs4 import BeautifulSoup
import openpyxl
import time
import random
from concurrent.futures import ThreadPoolExecutor

API_KEY = 'TU_API_KEY'
SECTOR = 'Sector'
PAGINA_INICIAL = 1 #La pagina donde empieza
NUM_PAGINAS = 8 #El número de páginas que scrappea
MAX_HILOS = 5 #Subir o bajar para ir más rápido (consume recursos del PC)

wb = openpyxl.Workbook()
ws = wb.active
ws.append(['Posición', 'Evolución', 'Empresa', 'Web', 'Sector', 'Facturación', 'Provincia', 'Enlace'])

def delay_aleatorio():
    time.sleep(random.uniform(0.6, 1.4))

def usar_scraperapi(url):
    try:
        r = requests.get('http://api.scraperapi.com/', params={
            'api_key': API_KEY,
            'url': url
        }, timeout=30)
        return r.text if r.status_code == 200 else None
    except:
        return None

def construir_url_lista(pagina):
    # ✅ Enlace con paginación
    return f"https://ranking-empresas.eleconomista.es/ranking_empresas_nacional.html?qSectorNorm={SECTOR}&qPagina={pagina}"

def extraer_empresas(html):
    soup = BeautifulSoup(html, 'html.parser')
    filas = soup.select('tr[itemprop="itemListElement"], tr.even')
    datos = []

    for fila in filas:
        try:
            nombre_elem = fila.select_one('td.tal a')
            datos.append({
                'posicion': fila.select_one('td[align="center"] span').text.strip(),
                'evolucion': fila.select_one('td.col_responsive1 span.inline').text.strip(),
                'nombre': nombre_elem.text.strip(),
                'enlace': nombre_elem['href'],
                'facturacion': fila.find_all('td')[3].text.strip(),
                'sector': fila.select_one('abbr')['title'],
                'provincia': fila.select_one('div[itemprop="addressRegion"]').text.strip()
            })
        except:
            continue
    return datos

def obtener_web_empresa(datos):
    html = usar_scraperapi(datos['enlace'])
    if not html:
        return None

    try:
        soup = BeautifulSoup(html, 'html.parser')
        web_elem = soup.select_one("td:contains('Página Web') + td a") or \
                   soup.select_one("table tr:has(td:contains('Página Web')) a")

        if web_elem and web_elem.get("href"):
            web = web_elem["href"].strip()
            print(f"✔ {datos['posicion']} - {datos['nombre']} → {web}")
            return [
                datos['posicion'],
                datos['evolucion'],
                datos['nombre'],
                web,
                datos['sector'],
                datos['facturacion'],
                datos['provincia'],
                datos['enlace']
            ]
    except:
        pass
    return None

# ──────────────────────── Bucle principal ────────────────────────
for pagina in range(PAGINA_INICIAL, PAGINA_INICIAL + NUM_PAGINAS):
    print(f"\n📄 Procesando página {pagina}")
    html_lista = usar_scraperapi(construir_url_lista(pagina))
    if not html_lista:
        print("⚠ No se pudo obtener la lista de empresas.")
        continue

    empresas = extraer_empresas(html_lista)

    with ThreadPoolExecutor(max_workers=MAX_HILOS) as executor:
        resultados = list(executor.map(obtener_web_empresa, empresas))

    for fila in resultados:
        if fila:
            ws.append(fila)
            wb.save(f'empresas_sector_{SECTOR}_solo_con_web.xlsx')

    delay_aleatorio()

print("\n✅ Scraping completado: solo empresas con web.")
