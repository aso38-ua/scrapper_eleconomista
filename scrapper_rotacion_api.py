import requests
from bs4 import BeautifulSoup
import openpyxl
import time
import random
from concurrent.futures import ThreadPoolExecutor
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import os

API_KEY = 'TU_TOKEN_API'
parar_scraper = False

def delay_aleatorio():
    time.sleep(random.uniform(0.6, 1.2))

def usar_scraperapi(url):
    try:
        r = requests.get('http://api.scraperapi.com/', params={
            'api_key': API_KEY,
            'url': url,
            'retry': '3'
        }, timeout=60)
        return r.text if r.status_code == 200 else None
    except:
        return None

def construir_url_lista(modo, valor, pagina):
    if modo == "sector":
        return f"https://ranking-empresas.eleconomista.es/ranking_empresas_nacional.html?qSectorNorm={valor}&qPagina={pagina}"
    else:
        return f"https://ranking-empresas.eleconomista.es/ranking_empresas_nacional.html?qProvNorm={valor}&qPagina={pagina}"

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
    global parar_scraper
    if parar_scraper:
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
            mensaje = f"✔ {datos['posicion']} - {datos['nombre']} → {web}"
            print(mensaje)
            root.after(0, actualizar_interfaz, mensaje)
            return [
                datos['posicion'], datos['evolucion'], datos['nombre'], web,
                datos['sector'], datos['facturacion'], datos['provincia'], datos['enlace']
            ]
    except:
        return None

def ejecutar_scraper():
    global parar_scraper
    try:
        modo = modo_scrapeo.get()
        valor = entry_valor.get().strip()
        pagina_inicial = int(entry_pagina_inicial.get())
        num_paginas = int(entry_num_paginas.get())
        max_hilos = int(entry_hilos.get())

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(['Posición', 'Evolución', 'Empresa', 'Web', 'Sector', 'Facturación', 'Provincia', 'Enlace'])

        nombre_base = f"empresas_{modo}_{valor}_solo_con_web"
        nombre_archivo = f"{nombre_base}.xlsx"
        contador = 1
        while os.path.exists(nombre_archivo):
            nombre_archivo = f"{nombre_base}_{contador}.xlsx"
            contador += 1

        root.after(0, actualizar_interfaz, f"🔍 Iniciando scraping por {modo}: {valor}")
        total_empresas = 0

        for pagina in range(pagina_inicial, pagina_inicial + num_paginas):
            if parar_scraper:
                root.after(0, actualizar_interfaz, "🛑 Scraping detenido por el usuario")
                break

            porcentaje = (pagina - pagina_inicial + 1) / num_paginas * 100
            root.after(0, actualizar_interfaz, f"\n📄 Página {pagina} ({porcentaje:.1f}%)")
            root.after(0, lambda v=porcentaje: progress_bar.config(value=v))

            url = construir_url_lista(modo, valor, pagina)
            html_lista = usar_scraperapi(url)
            if not html_lista:
                root.after(0, actualizar_interfaz, f"⚠ Error al cargar página {pagina}")
                continue

            empresas = extraer_empresas(html_lista)
            if not empresas:
                continue

            with ThreadPoolExecutor(max_workers=max_hilos) as executor:
                resultados = list(executor.map(obtener_web_empresa, empresas))

            for fila in resultados:
                if fila:
                    ws.append(fila)
                    total_empresas += 1
                    if total_empresas % 10 == 0:
                        wb.save(nombre_archivo)
                        root.after(0, actualizar_interfaz, f"💾 Guardado temporal: {total_empresas} empresas")

            delay_aleatorio()

        wb.save(nombre_archivo)
        root.after(0, actualizar_interfaz, f"\n✅ Scraping finalizado. Total: {total_empresas}")
        messagebox.showinfo("Finalizado", f"Se guardaron {total_empresas} empresas en {nombre_archivo}")
        progress_bar.config(value=100)

    except Exception as e:
        messagebox.showerror("Error", f"Fallo inesperado: {str(e)}")

def actualizar_interfaz(mensaje):
    output_text.insert(tk.END, mensaje + "\n")
    output_text.see(tk.END)
    root.update_idletasks()

def iniciar_scraper_en_hilo():
    global parar_scraper
    parar_scraper = False
    output_text.delete(1.0, tk.END)
    threading.Thread(target=ejecutar_scraper, daemon=True).start()

def parar_scraper_func():
    global parar_scraper
    parar_scraper = True
    actualizar_interfaz("🛑 Deteniendo...")

def cambiar_label_input(*args):
    tipo = modo_scrapeo.get()
    if tipo == "sector":
        label_valor.config(text="Sector (código):")
        entry_valor.delete(0, tk.END)
        entry_valor.insert(0, "5221")
    else:
        label_valor.config(text="Provincia (nombre web):")
        entry_valor.delete(0, tk.END)
        entry_valor.insert(0, "SANTA-CRUZ-TENERIFE")

# ─────────── Interfaz ───────────
root = tk.Tk()
root.title("Scraper de Empresas por Sector o Provincia")
root.geometry("800x600")

frame = ttk.Frame(root, padding=20)
frame.pack(fill=tk.BOTH, expand=True)

# Opción de scrapeo
ttk.Label(frame, text="Modo:").grid(row=0, column=0, sticky=tk.W)
modo_scrapeo = tk.StringVar(value="sector")
modo_dropdown = ttk.Combobox(frame, textvariable=modo_scrapeo, values=["sector", "provincia"], state="readonly", width=20)
modo_dropdown.grid(row=0, column=1)
modo_dropdown.bind("<<ComboboxSelected>>", cambiar_label_input)

# Valor del sector o provincia
label_valor = ttk.Label(frame, text="Sector (código):")
label_valor.grid(row=1, column=0, sticky=tk.W)
entry_valor = ttk.Entry(frame, width=25)
entry_valor.insert(0, "5221")
entry_valor.grid(row=1, column=1)

# Paginación y hilos
ttk.Label(frame, text="Página inicial:").grid(row=2, column=0, sticky=tk.W)
entry_pagina_inicial = ttk.Entry(frame, width=25)
entry_pagina_inicial.insert(0, "1")
entry_pagina_inicial.grid(row=2, column=1)

ttk.Label(frame, text="Número de páginas:").grid(row=3, column=0, sticky=tk.W)
entry_num_paginas = ttk.Entry(frame, width=25)
entry_num_paginas.insert(0, "5")
entry_num_paginas.grid(row=3, column=1)

ttk.Label(frame, text="Máx. hilos:").grid(row=4, column=0, sticky=tk.W)
entry_hilos = ttk.Entry(frame, width=25)
entry_hilos.insert(0, "3")
entry_hilos.grid(row=4, column=1)

# Botones
btn_frame = ttk.Frame(frame)
btn_frame.grid(row=5, column=0, columnspan=2, pady=10)
ttk.Button(btn_frame, text="Iniciar scraping", command=iniciar_scraper_en_hilo).pack(side=tk.LEFT, padx=5)
ttk.Button(btn_frame, text="Parar scraping", command=parar_scraper_func).pack(side=tk.LEFT, padx=5)

# Progreso
progress_bar = ttk.Progressbar(frame, orient='horizontal', mode='determinate', length=300)
progress_bar.grid(row=6, column=0, columnspan=2, pady=10)

# Output
output_frame = ttk.Frame(root)
output_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0,20))
output_text = tk.Text(output_frame, wrap=tk.WORD, height=20)
scrollbar = ttk.Scrollbar(output_frame, orient=tk.VERTICAL, command=output_text.yview)
output_text.configure(yscrollcommand=scrollbar.set)
output_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

root.mainloop()

