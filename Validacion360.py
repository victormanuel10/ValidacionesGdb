import os
import shutil
import cv2
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, UnidentifiedImageError
from openpyxl import Workbook
import hashlib
import torch
import torchvision.models as models
import torchvision.transforms as transforms
import urllib.request
import exifread
import pandas as pd
import pandas as pd
import geopandas as gpd
import math
from shapely.geometry import Point

"""
RadioBUtton
zona=01
RUral= 00

"""

def descargar_archivos_places():
    if not os.path.exists("resnet18_places365.pth.tar"):
        urllib.request.urlretrieve("http://places2.csail.mit.edu/models_places365/resnet18_places365.pth.tar", "resnet18_places365.pth.tar")
    if not os.path.exists("categories_places365.txt"):
        urllib.request.urlretrieve("https://raw.githubusercontent.com/csailvision/places365/master/categories_places365.txt", "categories_places365.txt")

def cargar_modelo_places365():
    model = models.resnet18(num_classes=365)
    checkpoint = torch.load("resnet18_places365.pth.tar", map_location=torch.device('cpu'))
    state_dict = {k.replace('module.', ''): v for k, v in checkpoint['state_dict'].items()}
    model.load_state_dict(state_dict)
    model.eval()
    return model

def cargar_clases_places365():
    with open("categories_places365.txt") as f:
        classes = [line.strip().split(' ')[0][3:] for line in f.readlines()]
    return classes

def transformar_imagen_places365(img):
    transform = transforms.Compose([
        transforms.Resize((256, 256)),
        transforms.CenterCrop(224),
        transforms.ToTensor(),
        transforms.Normalize(mean=[0.485, 0.456, 0.406], std=[0.229, 0.224, 0.225])
    ])
    return transform(img).unsqueeze(0)

def es_orientacion_valida(ruta_imagen, modelo, clases):
    try:
        img = Image.open(ruta_imagen).convert('RGB')
        input_tensor = transformar_imagen_places365(img)
        with torch.no_grad():
            logit = modelo(input_tensor)
            h_x = torch.nn.functional.softmax(logit, 1).squeeze()
            probs, idx = h_x.sort(0, True)
            top5 = [clases[idx[i]] for i in range(5)]
        clases_validas = ['street', 'road', 'highway', 'alley', 'residential_neighborhood']
        for c in clases_validas:
            if c in top5:
                return True, top5[0]
        return False, top5[0]
    except Exception as e:
        return False, f"Error: {e}"

def esta_borrosa(imagen, promedio_umbral, minimo_umbral):
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
    return np.mean(focos) < promedio_umbral or np.min(focos) < minimo_umbral

def oscuridad_o_sobreexposicion(imagen, min_oscuridad, max_sobreexp):
    gris = cv2.cvtColor(imagen, cv2.COLOR_BGR2GRAY)
    media = np.mean(gris)
    return media < min_oscuridad or media > max_sobreexp

def resolucion_inadecuada(imagen, resolucion_minima):
    alto, ancho = imagen.shape[:2]
    return ancho < resolucion_minima[0] or alto < resolucion_minima[1]

def hash_imagen(ruta):
    with open(ruta, 'rb') as f:
        return hashlib.md5(f.read()).hexdigest()

def validar_fotos():
    carpeta_origen = entry_origen.get()
    carpeta_errores = entry_errores.get()
    carpeta_validas = entry_validas.get()
    ruta_excel = entry_excel.get()

    try:
        prom = int(entry_nitidez_prom.get())
        mini = int(entry_nitidez_min.get())
        oscuro = int(entry_oscuro.get())
        sobrex = int(entry_sobreexp.get())
        resol = (int(entry_res_ancho.get()), int(entry_res_alto.get()))
    except ValueError:
        messagebox.showerror("Error", "Los parámetros deben ser enteros.")
        return

    if not all(map(os.path.isdir, [carpeta_origen, carpeta_errores, carpeta_validas])) or not ruta_excel:
        messagebox.showerror("Error", "Faltan rutas.")
        return

    os.makedirs(carpeta_errores, exist_ok=True)
    os.makedirs(carpeta_validas, exist_ok=True)
    reporte, hashes = [], set()

    descargar_archivos_places()
    modelo = cargar_modelo_places365()
    clases = cargar_clases_places365()

    for archivo in os.listdir(carpeta_origen):
        if not archivo.lower().endswith(('jpg', 'jpeg', 'png')):
            continue

        ruta = os.path.join(carpeta_origen, archivo)
        razones = []

        if os.path.getsize(ruta) == 0:
            razones.append("Archivo vacío")
        else:
            try:
                with Image.open(ruta) as img:
                    img.verify()
            except:
                razones.append("Archivo corrupto")

        escena = "No detectada"
        if not razones:
            try:
                img = cv2.imread(ruta)
                if img is None:
                    try:
                        img_data = np.fromfile(ruta, dtype=np.uint8)
                        img = cv2.imdecode(img_data, cv2.IMREAD_COLOR)
                    except:
                        img = None

                if img is None:
                    try:
                        with Image.open(ruta) as pil_img:
                            img = np.array(pil_img.convert('RGB'))
                    except:
                        razones.append("No legible por OpenCV ni PIL")

                if img is not None:
                    if esta_borrosa(img, prom, mini): razones.append("Borroso")
                    if oscuridad_o_sobreexposicion(img, oscuro, sobrex): razones.append("Oscuro o sobreexpuesto")
                    if resolucion_inadecuada(img, resol): razones.append("Resolución baja")
                    h = hash_imagen(ruta)
                    if h in hashes:
                        razones.append("Duplicado")
                    else:
                        hashes.add(h)
                    valida, escena = es_orientacion_valida(ruta, modelo, clases)
                    if not valida:
                        razones.append(f"No orientado ({escena})")
            except Exception as e:
                razones.append(f"Error procesamiento: {str(e)}")

        destino = carpeta_errores if razones else carpeta_validas
        shutil.copy(ruta, os.path.join(destino, archivo))
        reporte.append((archivo, escena, "; ".join(razones) if razones else "OK"))

    wb = Workbook()
    ws = wb.active
    ws.title = "Reporte"
    ws.append(["Archivo", "Escena", "Estado"])
    for row in reporte:
        ws.append(row)
    wb.save(ruta_excel)
    messagebox.showinfo("✅ Validación completa", f"Se procesaron {len(reporte)} imágenes.")

# ------------------- INTERFAZ -------------------
root = tk.Tk()
root.title("Validador de Fotos Cartografia")
notebook = ttk.Notebook(root)
notebook.pack(fill='both', expand=True)

# TAB 1: Validación Visual
tab1 = ttk.Frame(notebook)
notebook.add(tab1, text="Validación fotos")
frame = tk.Frame(tab1, padx=10, pady=10)
frame.pack()

entry_origen = tk.Entry(frame, width=60); entry_origen.grid(row=1, column=0)
tk.Button(frame, text="Seleccionar", command=lambda: entry_origen.insert(0, filedialog.askdirectory())).grid(row=1, column=1)
entry_errores = tk.Entry(frame, width=60); entry_errores.grid(row=3, column=0)
tk.Button(frame, text="Seleccionar", command=lambda: entry_errores.insert(0, filedialog.askdirectory())).grid(row=3, column=1)
entry_validas = tk.Entry(frame, width=60); entry_validas.grid(row=5, column=0)
tk.Button(frame, text="Seleccionar", command=lambda: entry_validas.insert(0, filedialog.askdirectory())).grid(row=5, column=1)
entry_excel = tk.Entry(frame, width=60); entry_excel.grid(row=7, column=0)
tk.Button(frame, text="Seleccionar", command=lambda: entry_excel.insert(0, filedialog.asksaveasfilename(defaultextension=".xlsx"))).grid(row=7, column=1)

for i, text in enumerate(["📁 Carpeta de Fotos:", "📂 Carpeta para errores:", "📂 Carpeta para válidas:", "📊 Reporte Excel:"]):
    tk.Label(frame, text=text).grid(row=i*2, column=0, sticky="w")

def crear_parametro(texto, fila, valor):
    tk.Label(frame, text=texto).grid(row=fila, column=0, sticky="e")
    entry = tk.Entry(frame, width=10)
    entry.insert(0, str(valor))
    entry.grid(row=fila, column=1, sticky="w")
    return entry

entry_nitidez_prom = crear_parametro("Nitidez promedio:", 9, 100)
entry_nitidez_min = crear_parametro("Nitidez mínima:", 10, 50)
entry_oscuro = crear_parametro("Oscuridad mínima:", 11, 20)
entry_sobreexp = crear_parametro("Sobreexposición máx:", 12, 240)
entry_res_ancho = crear_parametro("Resolución mínima (ancho):", 13, 1500)
entry_res_alto = crear_parametro("Resolución mínima (alto):", 14, 800)

tk.Button(frame, text="🚀 Ejecutar Validación", bg="#0b7dda", fg="white", font=('Arial', 10, 'bold'), command=validar_fotos).grid(row=15, column=0, columnspan=2, pady=20)

# TAB 2: Extracción GPS

# TAB 3: Coordenadas de Shapefile
import geopandas as gpd

tab3 = ttk.Frame(notebook)
notebook.add(tab3, text="Shapefile a Excel")

frame3 = tk.Frame(tab3, padx=15, pady=15)
frame3.pack()

shapefile_var = tk.StringVar()
excel_out_var = tk.StringVar()

tk.Label(frame3, text="🗂 Archivo SHP:").grid(row=0, column=0, sticky="e", padx=10, pady=10)
tk.Entry(frame3, textvariable=shapefile_var, width=50).grid(row=0, column=1)
tk.Button(frame3, text="Seleccionar", command=lambda: shapefile_var.set(filedialog.askopenfilename(filetypes=[("Shapefiles", "*.shp")]))).grid(row=0, column=2)

tk.Label(frame3, text="📄 Guardar archivo Excel:").grid(row=1, column=0, sticky="e", padx=10)
tk.Entry(frame3, textvariable=excel_out_var, width=50).grid(row=1, column=1)
tk.Button(frame3, text="Seleccionar", command=lambda: excel_out_var.set(filedialog.asksaveasfilename(defaultextension=".xlsx"))).grid(row=1, column=2)
def exportar_coords_shapefile():
    shp = shapefile_var.get()
    salida = excel_out_var.get()
    if not shp or not salida:
        messagebox.showwarning("Faltan datos", "Debes seleccionar el archivo shapefile y la ruta de guardado.")
        return

    try:
        gdf = gpd.read_file(shp)
        if gdf.geometry.geom_type.unique().tolist() == ['Point']:
            gdf['Latitud'] = gdf.geometry.y
            gdf['Longitud'] = gdf.geometry.x
        else:
            gdf['Latitud'] = gdf.geometry.centroid.y
            gdf['Longitud'] = gdf.geometry.centroid.x

        # ✅ Mostrar solo el nombre del archivo en la columna 'path'
        if 'Path' in gdf.columns:
            gdf['Path'] = gdf['Path'].apply(lambda x: os.path.basename(str(x)))

        # ✅ Eliminar columnas duplicadas de lat/lon y excluir geometría
        columnas = ['Latitud', 'Longitud'] + [col for col in gdf.columns if col not in ['geometry', 'Latitud', 'Longitud']]
        gdf = gdf[columnas]

        gdf.to_excel(salida, index=False)
        messagebox.showinfo("¡Éxito!", f"Coordenadas exportadas a:\n{salida}")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo leer el archivo SHP:\n{str(e)}")



tk.Button(frame3, text="📌 Exportar Coordenadas", command=exportar_coords_shapefile, bg="#0A8754", fg="white", font=("Arial", 12, "bold")).grid(row=3, column=1, pady=30)
tab2 = ttk.Frame(notebook)
notebook.add(tab2, text="Extracción GPS")
frame2 = tk.Frame(tab2, padx=15, pady=15)
frame2.pack()

entrada_fotos_var = tk.StringVar()
entrada_guardar_var = tk.StringVar()

tk.Label(frame2, text="📁 Carpeta de Fotos:").grid(row=0, column=0, sticky="e", padx=10, pady=10)
tk.Entry(frame2, textvariable=entrada_fotos_var, width=50).grid(row=0, column=1)
tk.Button(frame2, text="Seleccionar", command=lambda: entrada_fotos_var.set(filedialog.askdirectory())).grid(row=0, column=2)

tk.Label(frame2, text="📄 Guardar archivo Excel:").grid(row=1, column=0, sticky="e", padx=10)
tk.Entry(frame2, textvariable=entrada_guardar_var, width=50).grid(row=1, column=1)
tk.Button(frame2, text="Seleccionar", command=lambda: entrada_guardar_var.set(filedialog.asksaveasfilename(defaultextension=".xlsx"))).grid(row=1, column=2)

def ejecutar_proceso():
    carpeta = entrada_fotos_var.get()
    ruta_salida_excel = entrada_guardar_var.get()
    if not carpeta or not ruta_salida_excel:
        messagebox.showwarning("Campos incompletos", "Debes seleccionar carpeta de fotos y ruta de guardado.")
        return
    datos = []
    for archivo in os.listdir(carpeta):
        if archivo.lower().endswith((".jpg", ".jpeg", ".png")):
            ruta_completa = os.path.join(carpeta, archivo)
            coordenadas = extraer_coordenadas_gps(ruta_completa)
            if coordenadas:
                datos.append({"NombreFoto": archivo, "Latitud": coordenadas["Latitud"], "Longitud": coordenadas["Longitud"]})
    if datos:
        df = pd.DataFrame(datos)
        df.to_excel(ruta_salida_excel, index=False)
        ruta_txt = ruta_salida_excel.replace(".xlsx", ".txt")
        df.to_csv(ruta_txt, sep="\t", index=False)
        messagebox.showinfo("¡Éxito!", f"Archivo generado:\n{ruta_salida_excel}")
    else:
        messagebox.showwarning("Sin coordenadas", "No se encontraron fotos con coordenadas GPS.")

tk.Button(frame2, text="📌 Ejecutar Extracción", command=ejecutar_proceso, bg="#0078D7", fg="white", font=("Arial", 12, "bold")).grid(row=3, column=1, pady=30)


# TAB 4: Verificación Dirección
tab_direccion_shape = ttk.Frame(notebook)
notebook.add(tab_direccion_shape, text="Verificación con Shape")
frame_ds = tk.Frame(tab_direccion_shape, padx=15, pady=15)
frame_ds.pack()

archivo_excel_fotos_var = tk.StringVar()
archivo_shapefile_var = tk.StringVar()
archivo_guardar_reporte_var = tk.StringVar()

tk.Label(frame_ds, text="📄 Archivo Excel (fotos):").grid(row=0, column=0, sticky="e", padx=10, pady=10)
tk.Entry(frame_ds, textvariable=archivo_excel_fotos_var, width=50).grid(row=0, column=1)
tk.Button(frame_ds, text="Seleccionar", command=lambda: archivo_excel_fotos_var.set(filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")]))).grid(row=0, column=2)

tk.Label(frame_ds, text="🗂 Shapefile (puntos objetivo):").grid(row=1, column=0, sticky="e", padx=10, pady=10)
tk.Entry(frame_ds, textvariable=archivo_shapefile_var, width=50).grid(row=1, column=1)
tk.Button(frame_ds, text="Seleccionar", command=lambda: archivo_shapefile_var.set(filedialog.askopenfilename(filetypes=[("Shapefiles", "*.shp")]))).grid(row=1, column=2)

tk.Label(frame_ds, text="📄 Guardar reporte Excel:").grid(row=2, column=0, sticky="e", padx=10)
tk.Entry(frame_ds, textvariable=archivo_guardar_reporte_var, width=50).grid(row=2, column=1)
tk.Button(frame_ds, text="Seleccionar", command=lambda: archivo_guardar_reporte_var.set(filedialog.asksaveasfilename(defaultextension=".xlsx"))).grid(row=2, column=2)

def calcular_rumbo(lat1, lon1, lat2, lon2):
    lat1, lon1, lat2, lon2 = map(math.radians, [lat1, lon1, lat2, lon2])
    dlon = lon2 - lon1
    x = math.sin(dlon) * math.cos(lat2)
    y = math.cos(lat1) * math.sin(lat2) - (math.sin(lat1) * math.cos(lat2) * math.cos(dlon))
    bearing = (math.degrees(math.atan2(x, y)) + 360) % 360
    return bearing


tab_direccion_excel = ttk.Frame(notebook)
notebook.add(tab_direccion_excel, text="Verificación Dirección (Excel)")
frame_de = tk.Frame(tab_direccion_excel, padx=15, pady=15)
frame_de.pack()

excel_fotos_var = tk.StringVar()
excel_objetivos_var = tk.StringVar()
excel_salida_var = tk.StringVar()

tk.Label(frame_de, text="📄 Excel de Fotos:").grid(row=0, column=0, sticky="e", padx=10, pady=10)
tk.Entry(frame_de, textvariable=excel_fotos_var, width=50).grid(row=0, column=1)
tk.Button(frame_de, text="Seleccionar", command=lambda: excel_fotos_var.set(filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")]))).grid(row=0, column=2)

tk.Label(frame_de, text="📄 Excel de Puntos Objetivo:").grid(row=1, column=0, sticky="e", padx=10, pady=10)
tk.Entry(frame_de, textvariable=excel_objetivos_var, width=50).grid(row=1, column=1)
tk.Button(frame_de, text="Seleccionar", command=lambda: excel_objetivos_var.set(filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")]))).grid(row=1, column=2)

tk.Label(frame_de, text="📄 Guardar reporte:").grid(row=2, column=0, sticky="e", padx=10)
tk.Entry(frame_de, textvariable=excel_salida_var, width=50).grid(row=2, column=1)
tk.Button(frame_de, text="Seleccionar", command=lambda: excel_salida_var.set(filedialog.asksaveasfilename(defaultextension=".xlsx"))).grid(row=2, column=2)

def calcular_rumbo(lat1, lon1, lat2, lon2):
    lat1, lon1, lat2, lon2 = map(math.radians, [lat1, lon1, lat2, lon2])
    dlon = lon2 - lon1
    x = math.sin(dlon) * math.cos(lat2)
    y = math.cos(lat1) * math.sin(lat2) - (math.sin(lat1) * math.cos(lat2) * math.cos(dlon))
    bearing = (math.degrees(math.atan2(x, y)) + 360) % 360
    return bearing

def ejecutar_verificacion_excel():
    archivo_fotos = excel_fotos_var.get()
    archivo_objetivos = excel_objetivos_var.get()
    salida = excel_salida_var.get()

    if not archivo_fotos or not archivo_objetivos or not salida:
        messagebox.showwarning("Faltan datos", "Selecciona ambos archivos Excel y la ruta de guardado.")
        return

    try:
        df_fotos = pd.read_excel(archivo_fotos)
        df_objetivos = pd.read_excel(archivo_objetivos)

        columnas_fotos = ['Path', 'LatitudFoto', 'LongitudFoto']
        columnas_obj = ['Path', 'LatitudObjetivo', 'LongitudObjetivo']
        for col in columnas_fotos + columnas_obj:
            if col not in (df_fotos.columns.tolist() + df_objetivos.columns.tolist()):
                raise Exception(f"Falta la columna '{col}' en uno de los archivos.")

        df_merged = pd.merge(df_fotos, df_objetivos, on='Path', how='inner')

        resultados = []
        for _, row in df_merged.iterrows():
            rumbo = calcular_rumbo(row['LatitudFoto'], row['LongitudFoto'], row['LatitudObjetivo'], row['LongitudObjetivo'])
            resultados.append({
                'Path': row['Path'],
                'LatitudFoto': row['LatitudFoto'],
                'LongitudFoto': row['LongitudFoto'],
                'LatitudObjetivo': row['LatitudObjetivo'],
                'LongitudObjetivo': row['LongitudObjetivo'],
                'RumboEsperado': round(rumbo, 2)
            })

        df_resultado = pd.DataFrame(resultados)
        df_resultado.to_excel(salida, index=False)
        messagebox.showinfo("✅ Listo",f"Reporte generado en:\n {salida}")
    except Exception as e:
        messagebox.showerror("Error", str(e))

tk.Button(frame_de, text="📌 Verificar Dirección", command=ejecutar_verificacion_excel, bg="#117864", fg="white", font=("Arial", 12, "bold")).grid(row=4, column=1, pady=30)


root.mainloop()
