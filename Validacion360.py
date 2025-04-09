import os
import shutil
import cv2
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, UnidentifiedImageError
from openpyxl import Workbook
import hashlib

# ----- CONFIGURACIÓN AJUSTABLE -----
UMBRAL_NITIDEZ_PROMEDIO = 200
UMBRAL_NITIDEZ_MINIMA = 100
UMBRAL_OSCURIDAD_MIN = 50
UMBRAL_SOBREEXPOSICION_MAX = 205
RESOLUCION_MINIMA = (3000, 1500)  # ancho, alto
# -----------------------------------

# Detección por bloques (desenfoque 360)
def esta_borrosa(imagen):
    gris = cv2.cvtColor(imagen, cv2.COLOR_BGR2GRAY)
    alto, ancho = gris.shape
    bloques_x, bloques_y = 4, 3
    bloque_ancho = ancho // bloques_x
    bloque_alto = alto // bloques_y
    focos = []

    for i in range(bloques_y):
        for j in range(bloques_x):
            bloque = gris[i*bloque_alto:(i+1)*bloque_alto, j*bloque_ancho:(j+1)*bloque_ancho]
            var_lap = cv2.Laplacian(bloque, cv2.CV_64F).var()
            focos.append(var_lap)

    promedio = np.mean(focos)
    minimo = np.min(focos)
    return promedio < UMBRAL_NITIDEZ_PROMEDIO or minimo < UMBRAL_NITIDEZ_MINIMA

# Oscuridad o sobreexposición
def oscuridad_o_sobreexposicion(imagen):
    gris = cv2.cvtColor(imagen, cv2.COLOR_BGR2GRAY)
    media = np.mean(gris)
    return media < UMBRAL_OSCURIDAD_MIN or media > UMBRAL_SOBREEXPOSICION_MAX

# Tamaño mínimo
def resolucion_inadecuada(imagen):
    alto, ancho = imagen.shape[:2]
    return ancho < RESOLUCION_MINIMA[0] or alto < RESOLUCION_MINIMA[1]

# Hash de imagen
def hash_imagen(ruta):
    with open(ruta, 'rb') as f:
        return hashlib.md5(f.read()).hexdigest()

# Lógica principal de validación
def validar_fotos():
    carpeta_origen = entry_origen.get()
    carpeta_destino = entry_destino.get()
    ruta_excel = entry_excel.get()

    if not os.path.isdir(carpeta_origen) or not os.path.isdir(carpeta_destino) or not ruta_excel:
        messagebox.showerror("Error", "Debes completar todas las rutas.")
        return

    os.makedirs(carpeta_destino, exist_ok=True)
    reporte = []
    hashes = set()

    for archivo in os.listdir(carpeta_origen):
        if not archivo.lower().endswith(('.jpg', '.jpeg', '.png')):
            continue

        ruta = os.path.join(carpeta_origen, archivo)
        razones = []

        if os.path.getsize(ruta) == 0:
            razones.append("Archivo vacío (0 KB o no descargado)")
            mover = True
        else:
            try:
                with Image.open(ruta) as img:
                    img.verify()
            except (UnidentifiedImageError, IOError):
                razones.append("Archivo corrupto (no abre con PIL)")

        if not razones:
            imagen = cv2.imread(ruta)
            if imagen is None:
                razones.append("No legible por OpenCV")
            else:
                if esta_borrosa(imagen):
                    razones.append("Imagen distorsionada o poco nítida")
                if oscuridad_o_sobreexposicion(imagen):
                    razones.append("Oscuridad o sobreexposición")
                if resolucion_inadecuada(imagen):
                    razones.append("Resolución inadecuada")
                hash_val = hash_imagen(ruta)
                if hash_val in hashes:
                    razones.append("Duplicado (hash)")
                else:
                    hashes.add(hash_val)

        if razones:
            try:
                shutil.copy(ruta, os.path.join(carpeta_destino, archivo))
            except Exception as e:
                razones.append(f"No se pudo copiar: {e}")
            reporte.append((archivo, "; ".join(razones)))

    # Guardar Excel
    wb = Workbook()
    ws = wb.active
    ws.title = "ReporteErrores"
    ws.append(["Archivo", "Razones de Falla"])
    for archivo, razon in reporte:
        ws.append([archivo, razon])
    wb.save(ruta_excel)

    messagebox.showinfo("✅ Finalizado", f"{len(reporte)} imágenes con errores.\nReporte generado.")

# ---------------- INTERFAZ ----------------
def seleccionar_origen():
    carpeta = filedialog.askdirectory(title="Selecciona carpeta de fotos")
    if carpeta:
        entry_origen.delete(0, tk.END)
        entry_origen.insert(0, carpeta)

def seleccionar_destino():
    carpeta = filedialog.askdirectory(title="Selecciona carpeta destino")
    if carpeta:
        entry_destino.delete(0, tk.END)
        entry_destino.insert(0, carpeta)

def seleccionar_excel():
    archivo = filedialog.asksaveasfilename(title="Guardar reporte Excel", defaultextension=".xlsx",
                                           filetypes=[("Excel files", "*.xlsx")])
    if archivo:
        entry_excel.delete(0, tk.END)
        entry_excel.insert(0, archivo)

root = tk.Tk()
root.title("Validador de Fotos - 360° / OneDrive")

frame = tk.Frame(root, padx=15, pady=15)
frame.pack()

tk.Label(frame, text="📁 Carpeta de Fotos:").grid(row=0, column=0, sticky="w")
entry_origen = tk.Entry(frame, width=60)
entry_origen.grid(row=1, column=0)
tk.Button(frame, text="Seleccionar", command=seleccionar_origen).grid(row=1, column=1)

tk.Label(frame, text="📂 Carpeta para errores:").grid(row=2, column=0, sticky="w")
entry_destino = tk.Entry(frame, width=60)
entry_destino.grid(row=3, column=0)
tk.Button(frame, text="Seleccionar", command=seleccionar_destino).grid(row=3, column=1)

tk.Label(frame, text="📊 Reporte Excel:").grid(row=4, column=0, sticky="w")
entry_excel = tk.Entry(frame, width=60)
entry_excel.grid(row=5, column=0)
tk.Button(frame, text="Seleccionar", command=seleccionar_excel).grid(row=5, column=1)

tk.Button(frame, text="🚀 Ejecutar Validación", bg="#0b7dda", fg="white", font=('Arial', 10, 'bold'),
          command=validar_fotos).grid(row=6, column=0, columnspan=2, pady=20)

root.mainloop()
