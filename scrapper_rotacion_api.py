import requests
from bs4 import BeautifulSoup
import openpyxl
import time
import random
from concurrent.futures import ThreadPoolExecutor
import tkinter as tk
from tkinter import ttk, messagebox

API_KEY = 'TU_API_KEY'

def delay_aleatorio():
    time.sleep(random.uniform(0.6, 1.4))

def usar_scraperapi(url):
    try:
        r = requests.get('http://api.scraperapi.com/', params={
            'api_key': API_KEY,
            'url': url
        }, timeout=60)  # Timeout aumentado
        return r.text if r.status_code == 200 else None
    except Exception as e:
        print(f"Error al realizar la solicitud: {str(e)}")
        return None

def construir_url_lista(sector, pagina):
    return f"https://ranking-empresas.eleconomista.es/ranking_empresas_nacional.html?qSectorNorm={sector}&qPagina={pagina}"

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
        web_elem = soup.select_one("td:-soup-contains('Página Web') + td a") or \
                   soup.select_one("table tr:has(td:-soup-contains('Página Web')) a")

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
    except Exception as e:
        print(f"Error al obtener la web de {datos['nombre']}: {str(e)}")
    return None

def ejecutar_scraper():
    try:
        sector = entry_sector.get().strip()
        pagina_inicial = int(entry_pagina_inicial.get())
        num_paginas = int(entry_num_paginas.get())
        max_hilos = int(entry_hilos.get())

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(['Posición', 'Evolución', 'Empresa', 'Web', 'Sector', 'Facturación', 'Provincia', 'Enlace'])

        for pagina in range(pagina_inicial, pagina_inicial + num_paginas):
            output_text.insert(tk.END, f"\n📄 Procesando página {pagina}\n")
            output_text.see(tk.END)

            html_lista = usar_scraperapi(construir_url_lista(sector, pagina))
            if not html_lista:
                output_text.insert(tk.END, f"⚠ No se pudo obtener la lista de empresas.\n")
                continue

            empresas = extraer_empresas(html_lista)

            with ThreadPoolExecutor(max_workers=max_hilos) as executor:
                resultados = list(executor.map(obtener_web_empresa, empresas))

            for fila in resultados:
                if fila:
                    ws.append(fila)
                    wb.save(f'empresas_sector_{sector}_solo_con_web.xlsx')

            delay_aleatorio()

        output_text.insert(tk.END, "\n✅ Scraping completado. Archivo guardado.\n")
        messagebox.showinfo("Completado", "Scraping finalizado y guardado en Excel.")
    except Exception as e:
        messagebox.showerror("Error", str(e))

root = tk.Tk()
root.title("Scraper de Empresas")

frame = ttk.Frame(root, padding=20)
frame.grid()

ttk.Label(frame, text="Sector:").grid(column=0, row=0, sticky=tk.W)
entry_sector = ttk.Entry(frame, width=20)
entry_sector.insert(0, "5221")
entry_sector.grid(column=1, row=0)

ttk.Label(frame, text="Página inicial:").grid(column=0, row=1, sticky=tk.W)
entry_pagina_inicial = ttk.Entry(frame, width=20)
entry_pagina_inicial.insert(0, "1")
entry_pagina_inicial.grid(column=1, row=1)

ttk.Label(frame, text="Número de páginas:").grid(column=0, row=2, sticky=tk.W)
entry_num_paginas = ttk.Entry(frame, width=20)
entry_num_paginas.insert(0, "5")
entry_num_paginas.grid(column=1, row=2)

ttk.Label(frame, text="Máx. hilos:").grid(column=0, row=3, sticky=tk.W)
entry_hilos = ttk.Entry(frame, width=20)
entry_hilos.insert(0, "3")  # Ajusta el número de hilos
entry_hilos.grid(column=1, row=3)

ttk.Button(frame, text="Iniciar scraping", command=ejecutar_scraper).grid(column=0, row=4, columnspan=2, pady=10)

output_text = tk.Text(root, height=20, width=80)
output_text.grid(padx=20, pady=10)

root.mainloop()

