# -*- coding: utf-8 -*-
# ArcGIS Desktop 10.8 - Python 2.7

import arcpy
import os
import pandas as pd
import Tkinter as tk
import tkFileDialog
import tkMessageBox

# ============================================================
# FUNCIÓN PRINCIPAL DE VALIDACIÓN
# ============================================================
nombres_hojas = {

    # MUST NOT OVERLAP (nombres truncados)
    "u_lc_terreno_NO_OVERLAP_u_lc_te": "Traslape_Terrenos_Urbanos",
    "r_lc_terreno_NO_OVERLAP_r_lc_te": "Traslape_Terrenos_Rurales",
    "r_lc_terreno_NO_OVERLAP_u_lc_te": "Traslape_Terreno_Rural_vs_Urbano",

    "Manzanas_NO_OVERLAP_Manzanas": "Traslape_Manzanas",
    "Veredas_NO_OVERLAP_Veredas": "Traslape_Veredas",
    "Barrios_NO_OVERLAP_Barrios": "Traslape_Barrios",

    # MUST BE COVERED BY
    "u_lc_terreno_COVERED_BY_Manzana": "Terrenos_Urbanos_fuera_de_Manzana",
    "r_lc_terreno_COVERED_BY_Veredas": "Terrenos_Rurales_fuera_de_Vereda",

    # GAPS
    "Manzanas_GAPS": "Huecos_en_Manzanas",
    "Veredas_GAPS": "Huecos_en_Veredas",

    # GEOMETRÍA
    "u_lc_terreno_GEOMETRIA": "Errores_Geometria_Terrenos_Urbanos",
    "r_lc_terreno_GEOMETRIA": "Errores_Geometria_Terrenos_Rurales",
    "Manzanas_GEOMETRIA": "Errores_Geometria_Manzanas",
    "Veredas_GEOMETRIA": "Errores_Geometria_Veredas",
    "Barrios_GEOMETRIA": "Errores_Geometria_Barrios"
}


def validar_topologia_completa(gdb_rural, gdb_urbana, output_excel):
    def asegurar_unicode(df):
        for col in df.columns:
            if df[col].dtype == object:
                df[col] = df[col].apply(
                    lambda x: unicode(x, 'utf-8', 'ignore')
                    if isinstance(x, str) else x
                )
        return df
    arcpy.env.overwriteOutput = True
    resultados = {}

    # --------------------------------------------------------
    # Resolver capa según reglas reales del proyecto
    # --------------------------------------------------------
    def obtener_capa(nombre):
        if nombre.startswith("r_"):
            ruta = os.path.join(gdb_rural, nombre)
            return ruta if arcpy.Exists(ruta) else None

        if nombre.startswith("u_"):
            ruta = os.path.join(gdb_urbana, nombre)
            return ruta if arcpy.Exists(ruta) else None

        if nombre == "Manzanas":
            ruta = os.path.join(gdb_urbana, nombre)
            return ruta if arcpy.Exists(ruta) else None

        if nombre == "Veredas":
            ruta = os.path.join(gdb_rural, nombre)
            return ruta if arcpy.Exists(ruta) else None

        if nombre == "Barrios":
            ruta_u = os.path.join(gdb_urbana, nombre)
            if arcpy.Exists(ruta_u):
                return ruta_u
            ruta_r = os.path.join(gdb_rural, nombre)
            return ruta_r if arcpy.Exists(ruta_r) else None

        return None
    def cargar_terreno_codigo_dict(capa):
        """
        Carga {OID : TERRENO_CODIGO} en memoria.
        No falla si el campo no existe.
        """
        d = {}

        campos = [f.name for f in arcpy.ListFields(capa)]
        if "TERRENO_CODIGO" not in campos:
            return d

        with arcpy.da.SearchCursor(capa, ["OID@", "TERRENO_CODIGO"]) as cur:
            for oid, codigo in cur:
                d[oid] = codigo

        return d
    terreno_cod_r = {}
    terreno_cod_u = {}

    capa_r = obtener_capa("r_lc_terreno")
    capa_u = obtener_capa("u_lc_terreno")

    if capa_r:
        terreno_cod_r = cargar_terreno_codigo_dict(capa_r)

    if capa_u:
        terreno_cod_u = cargar_terreno_codigo_dict(capa_u)
    # ========================================================
    # A. MUST NOT OVERLAP
    # ========================================================
    reglas_overlap = [
    ("Manzanas", "Manzanas"),
    ("Veredas", "Veredas"),
    ("Barrios", "Barrios"),
    ("r_lc_terreno", "r_lc_terreno"),
    ("u_lc_terreno", "u_lc_terreno"),
    ("r_lc_terreno", "u_lc_terreno"),
    ("Manzanas", "r_lc_terreno"),
    ("Veredas", "u_lc_terreno")
    ]

    def es_condicion_2(codigo):
        """
        Retorna True si el dígito 22 del TERRENO_CODIGO es '2'
        """
        if not codigo or len(codigo) < 22:
            return False
        return codigo[21] == "2"

    for l1, l2 in reglas_overlap:

        hoja = "{}_NO_OVERLAP_{}".format(l1, l2)[:31]

        c1 = obtener_capa(l1)
        c2 = obtener_capa(l2)

        if not c1 or not c2:
            resultados[hoja] = pd.DataFrame(
                [{"Error": "Capa faltante: {} o {}".format(l1, l2)}]
            )
            continue

        out = os.path.join("in_memory", "int_{}_{}".format(l1, l2))
        arcpy.Intersect_analysis([c1, c2], out, "ONLY_FID")
        arcpy.AddGeometryAttributes_management(
            out, "AREA", Area_Unit="SQUARE_METERS"
        )

        # -----------------------------
        # Resolver nombres FID reales
        # -----------------------------
        base1 = os.path.basename(c1)
        base2 = os.path.basename(c2)

        fid_1 = "FID_" + base1
        fid_2 = fid_1 + "_1" if c1 == c2 else "FID_" + base2

        campos = [fid_1, fid_2, "POLY_AREA"]

        rows = []

        with arcpy.da.SearchCursor(out, campos) as cursor:
            for fid_a, fid_b, area in cursor:

                # 1️⃣ eliminar auto-intersección
                if fid_a == fid_b:
                    continue

                # 2️⃣ eliminar duplicados A–B / B–A
                if fid_a > fid_b:
                    continue

                # 3️⃣ filtrar ruido geométrico
                if area < 1:
                    continue

                # -----------------------------
                # Obtener TERRENO_CODIGO
                # -----------------------------
                cod_a = None
                cod_b = None

                if l1 == "u_lc_terreno":
                    cod_a = terreno_cod_u.get(fid_a)
                elif l1 == "r_lc_terreno":
                    cod_a = terreno_cod_r.get(fid_a)

                if l2 == "u_lc_terreno":
                    cod_b = terreno_cod_u.get(fid_b)
                elif l2 == "r_lc_terreno":
                    cod_b = terreno_cod_r.get(fid_b)

                # -------------------------------------------------
                # REGLA ESPECIAL:
                # OMITIR condición 2 SOLO en urbano vs urbano
                # -------------------------------------------------
                if l1 == l2 and l1 in ("u_lc_terreno", "r_lc_terreno"):
                    if es_condicion_2(cod_a) or es_condicion_2(cod_b):
                        continue

                # -----------------------------
                # Registrar resultado
                # -----------------------------
                rows.append({
                    "FID_{}_A".format(l1): fid_a,
                    "TERRENO_CODIGO_A": cod_a,
                    "FID_{}_B".format(l2): fid_b,
                    "TERRENO_CODIGO_B": cod_b,
                    "Area_m2": round(area, 2)
                })

        resultados[hoja] = pd.DataFrame(
            rows if rows else [{"Resultado": "Sin traslapes"}]
        )

    # ========================================================
    # B. MUST BE COVERED BY
    # ========================================================
    hoja = "Terreno_Urbano_fuera_Manzana"[:31]

    c_terreno_u = obtener_capa("u_lc_terreno")
    c_manzanas = obtener_capa("Manzanas")

    if not c_terreno_u or not c_manzanas:
        resultados[hoja] = pd.DataFrame(
            [{"Error": "Capa faltante: u_lc_terreno o Manzanas"}]
        )
    else:
        # 1. Copiar TERRENO_CODIGO a una capa temporal
        terreno_tmp = os.path.join("in_memory", "u_terreno_tmp")
        arcpy.CopyFeatures_management(c_terreno_u, terreno_tmp)

        # Asegurar que el campo existe
        campos = [f.name for f in arcpy.ListFields(terreno_tmp)]
        if "TERRENO_CODIGO" not in campos:
            resultados[hoja] = pd.DataFrame(
                [{"Error": "El campo TERRENO_CODIGO no existe en u_lc_terreno"}]
            )
        else:
            # 2. ERASE (conserva TERRENO_CODIGO)
            fuera = os.path.join("in_memory", "u_fuera_manzana")
            arcpy.Erase_analysis(terreno_tmp, c_manzanas, fuera)

            # 3. Calcular área
            arcpy.AddGeometryAttributes_management(
                fuera, "AREA", Area_Unit="SQUARE_METERS"
            )

            rows = []

            with arcpy.da.SearchCursor(fuera, ["TERRENO_CODIGO", "POLY_AREA"]) as cur:
                for codigo, area in cur:
                    if area < 1:
                        continue

                    rows.append({
                        "TERRENO_CODIGO_u": codigo,
                        "Area_fuera_m2": round(area, 2)
                    })

            resultados[hoja] = pd.DataFrame(
                rows if rows else [{
                    "Resultado": "Todos los terrenos urbanos están completamente dentro de una manzana"
                }]
            )

    # ========================================================
    # C. MUST NOT HAVE GAPS
    # ========================================================
    for capa_nombre in ["Manzanas", "Veredas"]:
        hoja = "{}_GAPS".format(capa_nombre)[:31]
        capa = obtener_capa(capa_nombre)

        if not capa:
            resultados[hoja] = pd.DataFrame(
                [{"Error": "Capa faltante: {}".format(capa_nombre)}]
            )
            continue

        dis = os.path.join("in_memory", "dis_" + capa_nombre)
        arcpy.Dissolve_management(capa, dis)

        ext = arcpy.Describe(dis).extent
        marco = arcpy.CreateFishnet_management(
            os.path.join("in_memory", "frame_" + capa_nombre),
            "{} {}".format(ext.XMin, ext.YMin),
            "{} {}".format(ext.XMin, ext.YMax),
            0, 0, 1, 1,
            "{} {}".format(ext.XMax, ext.YMax),
            "NO_LABELS", "", "POLYGON"
        )

        gap = os.path.join("in_memory", "gap_" + capa_nombre)
        arcpy.Erase_analysis(marco, dis, gap)

        arcpy.AddGeometryAttributes_management(gap, "AREA", Area_Unit="SQUARE_METERS")

        rows = []
        with arcpy.da.SearchCursor(gap, ["POLY_AREA"]) as cur:
            for area, in cur:
                if area >= 1:
                    rows.append({"Area_gap_m2": round(area, 2)})

        resultados[hoja] = pd.DataFrame(
            rows if rows else [{"Resultado": "Sin huecos"}]
        )

    # ========================================================
    # D. GEOMETRÍA INVÁLIDA
    # ========================================================
    for capa_nombre in ["r_lc_terreno", "u_lc_terreno", "Manzanas", "Veredas", "Barrios"]:
        hoja = "{}_GEOMETRIA".format(capa_nombre)[:31]
        capa = obtener_capa(capa_nombre)

        if not capa:
            resultados[hoja] = pd.DataFrame(
                [{"Error": "Capa faltante: {}".format(capa_nombre)}]
            )
            continue

        out = os.path.join("in_memory", "chk_" + capa_nombre)
        arcpy.CheckGeometry_management(capa, out)

        rows = []
        with arcpy.da.SearchCursor(out, ["OID@", "PROBLEM"]) as cur:
            for oid, prob in cur:
                rows.append({"OID": oid, "Problema": prob})

        resultados[hoja] = pd.DataFrame(
            rows if rows else [{"Resultado": "Geometría válida"}]
        )

    # ========================================================
    # EXPORTAR EXCEL
    # ========================================================
    writer = pd.ExcelWriter(output_excel, engine="xlwt")

    for hoja, df in resultados.items():
        df = asegurar_unicode(df)
        nombre_excel = nombres_hojas.get(hoja, hoja)
        df.to_excel(writer, sheet_name=nombre_excel[:31], index=False)

    writer.save()

    tkMessageBox.showinfo(
        "Validación completada",
        "Proceso finalizado correctamente\n{}".format(output_excel)
    )

# ============================================================
# INTERFAZ TKINTER
# ============================================================

class TopologiaApp(object):

    def __init__(self, master):
        self.master = master
        master.title("Validación Topológica Completa")

        self.gdb_rural = tk.StringVar()
        self.gdb_urbana = tk.StringVar()
        self.excel = tk.StringVar()

        tk.Label(master, text="GDB Rural").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        tk.Entry(master, textvariable=self.gdb_rural, width=60).grid(row=0, column=1)
        tk.Button(master, text="...", command=self.sel_rural).grid(row=0, column=2)

        tk.Label(master, text="GDB Urbana").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        tk.Entry(master, textvariable=self.gdb_urbana, width=60).grid(row=1, column=1)
        tk.Button(master, text="...", command=self.sel_urbana).grid(row=1, column=2)

        tk.Label(master, text="Archivo Excel").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        tk.Entry(master, textvariable=self.excel, width=60).grid(row=2, column=1)
        tk.Button(master, text="...", command=self.sel_excel).grid(row=2, column=2)

        tk.Button(
            master,
            text="EJECUTAR VALIDACIÓN",
            bg="#4CAF50",
            fg="white",
            command=self.ejecutar
        ).grid(row=3, column=1, pady=15)

    def sel_rural(self):
        r = tkFileDialog.askdirectory(title="Seleccionar GDB Rural")
        if r.endswith(".gdb"):
            self.gdb_rural.set(r)

    def sel_urbana(self):
        r = tkFileDialog.askdirectory(title="Seleccionar GDB Urbana")
        if r.endswith(".gdb"):
            self.gdb_urbana.set(r)

    def sel_excel(self):
        r = tkFileDialog.asksaveasfilename(
            defaultextension=".xls",
            filetypes=[("Excel", "*.xls")]
        )
        if r:
            self.excel.set(r)

    def ejecutar(self):
        if not self.gdb_rural.get() or not self.gdb_urbana.get() or not self.excel.get():
            tkMessageBox.showerror("Error", "Debes seleccionar todas las rutas")
            return

        validar_topologia_completa(
            self.gdb_rural.get(),
            self.gdb_urbana.get(),
            self.excel.get()
        )

# ============================================================
# MAIN
# ============================================================

if __name__ == "__main__":
    root = tk.Tk()
    app = TopologiaApp(root)
    root.mainloop()
