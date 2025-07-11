import requests
from bs4 import BeautifulSoup
import openpyxl
import time
import random
from concurrent.futures import ThreadPoolExecutor
import tkinter as tk
from tkinter import ttk, messagebox
import threading

API_KEY = 'TU_API_TOKEN'
parar_scraper = False  # Variable para controlar si se debe parar el scraper

def delay_aleatorio():
    time.sleep(random.uniform(0.6, 1.4))

def usar_scraperapi(url):
    try:
        r = requests.get('http://api.scraperapi.com/', params={
            'api_key': API_KEY,
            'url': url,
            'retry': '3'  # Reintentar en caso de fallo
        }, timeout=60)
        
        if r.status_code != 200:
            print(f"Error ScraperAPI: {r.status_code} - {r.text[:100]}...")
            return None
            
        return r.text
    except requests.exceptions.RequestException as e:
        print(f"Error de conexión con ScraperAPI: {str(e)}")
        return None
    except Exception as e:
        print(f"Error inesperado en ScraperAPI: {str(e)}")
        return None

def construir_url_lista(sector, pagina):
    return f"https://ranking-empresas.eleconomista.es/ranking_empresas_nacional.html?qSectorNorm={sector}&qPagina={pagina}"

def extraer_empresas(html):
    if not html:
        return []
        
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
        except Exception as e:
            print(f"Error extrayendo datos de empresa: {str(e)}")
            continue
    return datos

def obtener_web_empresa(datos):
    if not datos or not datos.get('enlace'):
        return None

    html = usar_scraperapi(datos['enlace'])
    if not html:
        return None

    try:
        soup = BeautifulSoup(html, 'html.parser')
        web_elem = soup.select_one("td:-soup-contains('Página Web') + td a") or \
                   soup.select_one("table tr:has(td:-soup-contains('Página Web')) a")

        if web_elem and web_elem.get("href"):
            web = web_elem["href"].strip()
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
        print(f"Error al obtener la web de {datos.get('nombre', 'empresa desconocida')}: {str(e)}")
    return None

def ejecutar_scraper():
    global parar_scraper
    try:
        sector = entry_sector.get().strip()
        pagina_inicial = int(entry_pagina_inicial.get())
        num_paginas = int(entry_num_paginas.get())
        max_hilos = int(entry_hilos.get())

        # Configuración inicial
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(['Posición', 'Evolución', 'Empresa', 'Web', 'Sector', 'Facturación', 'Provincia', 'Enlace'])
        nombre_archivo = f'empresas_sector_{sector}_solo_con_web.xlsx'
        guardar_cada = 10  # Guardar cada 10 empresas
        contador_guardado = 0
        total_empresas = 0

        # Actualizar interfaz inicial
        root.after(0, actualizar_interfaz, f"🚀 Iniciando scraping para sector {sector}")
        root.after(0, lambda: progress_bar.config(value=0))

        # Verificar conexión inicial
        root.after(0, actualizar_interfaz, "🔍 Verificando conexión con ScraperAPI...")
        if not usar_scraperapi("https://google.com"):
            root.after(0, lambda: messagebox.showerror("Error", "No se pudo conectar con ScraperAPI"))
            return

        # Procesar cada página
        for pagina in range(pagina_inicial, pagina_inicial + num_paginas):
            if parar_scraper:
                root.after(0, actualizar_interfaz, "🔴 Scraping detenido por el usuario")
                break

            root.after(0, actualizar_interfaz, f"\n📄 Procesando página {pagina} ({(pagina-pagina_inicial+1)/num_paginas*100:.1f}%)")
            root.after(0, lambda v=(pagina-pagina_inicial+1)/num_paginas*100: progress_bar.config(value=v))

            html_lista = usar_scraperapi(construir_url_lista(sector, pagina))
            if not html_lista:
                root.after(0, actualizar_interfaz, f"⚠ No se pudo obtener la página {pagina}")
                continue

            empresas = extraer_empresas(html_lista)
            if not empresas:
                root.after(0, actualizar_interfaz, f"ℹ️ No se encontraron empresas en la página {pagina}")
                continue

            # Procesar empresas en paralelo
            with ThreadPoolExecutor(max_workers=max_hilos) as executor:
                resultados = list(executor.map(obtener_web_empresa, empresas))

            # Guardar resultados
            for fila in resultados:
                if fila:
                    ws.append(fila)
                    total_empresas += 1
                    contador_guardado += 1
                    
                    if contador_guardado % guardar_cada == 0:
                        wb.save(nombre_archivo)
                        root.after(0, actualizar_interfaz, f"💾 Guardado temporal ({total_empresas} empresas)")

            delay_aleatorio()

        # Guardar final y mostrar resultados
        wb.save(nombre_archivo)
        root.after(0, lambda: [
            actualizar_interfaz(f"\n✅ Scraping completado. {total_empresas} empresas guardadas en {nombre_archivo}"),
            progress_bar.config(value=100),
            messagebox.showinfo("Completado", f"Scraping finalizado.\n{total_empresas} empresas guardadas en {nombre_archivo}")
        ])

    except Exception as e:
        root.after(0, lambda: messagebox.showerror("Error", f"Error inesperado: {str(e)}"))

def actualizar_interfaz(mensaje):
    output_text.insert(tk.END, mensaje + "\n")
    output_text.see(tk.END)
    root.update_idletasks()

def iniciar_scraper_en_hilo():
    global parar_scraper
    parar_scraper = False
    output_text.delete(1.0, tk.END)  # Limpiar el output
    threading.Thread(target=ejecutar_scraper, daemon=True).start()

def parar_scraper_func():
    global parar_scraper
    parar_scraper = True
    actualizar_interfaz("\n🛑 Solicitando parada... Por favor espere.")

# Configuración de la interfaz gráfica
root = tk.Tk()
root.title("Scraper de Empresas - Mejorado")
root.geometry("800x600")

frame = ttk.Frame(root, padding=20)
frame.pack(fill=tk.BOTH, expand=True)

# Controles de entrada
ttk.Label(frame, text="Sector:").grid(column=0, row=0, sticky=tk.W, pady=5)
entry_sector = ttk.Entry(frame, width=25)
entry_sector.insert(0, "5221")
entry_sector.grid(column=1, row=0, sticky=tk.W, pady=5)

ttk.Label(frame, text="Página inicial:").grid(column=0, row=1, sticky=tk.W, pady=5)
entry_pagina_inicial = ttk.Entry(frame, width=25)
entry_pagina_inicial.insert(0, "1")
entry_pagina_inicial.grid(column=1, row=1, sticky=tk.W, pady=5)

ttk.Label(frame, text="Número de páginas:").grid(column=0, row=2, sticky=tk.W, pady=5)
entry_num_paginas = ttk.Entry(frame, width=25)
entry_num_paginas.insert(0, "5")
entry_num_paginas.grid(column=1, row=2, sticky=tk.W, pady=5)

ttk.Label(frame, text="Máx. hilos:").grid(column=0, row=3, sticky=tk.W, pady=5)
entry_hilos = ttk.Entry(frame, width=25)
entry_hilos.insert(0, "3")
entry_hilos.grid(column=1, row=3, sticky=tk.W, pady=5)

# Botones
btn_frame = ttk.Frame(frame)
btn_frame.grid(column=0, row=4, columnspan=2, pady=10)

ttk.Button(btn_frame, text="Iniciar scraping", command=iniciar_scraper_en_hilo).pack(side=tk.LEFT, padx=5)
ttk.Button(btn_frame, text="Parar scraping", command=parar_scraper_func).pack(side=tk.LEFT, padx=5)

# Barra de progreso
progress_bar = ttk.Progressbar(frame, orient='horizontal', length=300, mode='determinate')
progress_bar.grid(column=0, row=5, columnspan=2, pady=10)

# Área de texto para output
output_frame = ttk.Frame(root)
output_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0,20))

output_text = tk.Text(output_frame, height=20, width=90, wrap=tk.WORD)
output_scroll = ttk.Scrollbar(output_frame, orient=tk.VERTICAL, command=output_text.yview)
output_text.configure(yscrollcommand=output_scroll.set)

output_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
output_scroll.pack(side=tk.RIGHT, fill=tk.Y)

root.mainloop()
