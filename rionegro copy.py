# -*- coding: utf-8 -*-
import os
import arcpy
import xlrd
import xlwt
import pandas as pd
import sys
import Tkinter as tk
import tkFileDialog 
import tkMessageBox 
from Tkinter import Frame
import ttk
import re

try:
    reload(sys)
    sys.setdefaultencoding('utf-8')
except:
    pass

sys.path.append(r'C:\Program Files (x86)\ArcGIS\Desktop10.8\arcpy')
os.environ['PATH'] = r"C:\Program Files\ArcGIS\Bin;" + os.environ['PATH']

# -------------------------------------------------------------------------
# UTILIDADES DE CAMPOS EN MINÚSCULAS
# -------------------------------------------------------------------------

def get_fields_lower(fc_path):
    """
    Dada una feature class / tabla, retorna:
      - fields_lower: lista de nombres de campo en minúsculas
      - field_map: dict {nombre_en_minuscula: nombre_original}
    Esto permite trabajar internamente en minúsculas.
    """
    fields = arcpy.ListFields(fc_path)
    field_map = {}
    fields_lower = []
    for f in fields:
        low = f.name.lower()
        field_map[low] = f.name
        fields_lower.append(low)
    return fields_lower, field_map

def fmap(field_map, wanted_lower):
    """
    Devuelve el nombre REAL del campo en la GDB para un nombre lógico en minúsculas.
    Si no existe, retorna None.
    """
    return field_map.get(wanted_lower.lower())

def dividir_npn_en_columnas(npn):
    """Divide el campo NPN en 12 segmentos lógicos."""
    return [
        npn[:2],      # Departamento 
        npn[2:5],     # Municipio 
        npn[5:7],     # Zona 
        npn[7:9],     # Sector 
        npn[9:11],    # Comuna 
        npn[11:13],   # Barrio 
        npn[13:17],   # Manzana o Vereda
        npn[17:21],   # Terreno o Predios
        npn[21:22],   # Condición Predio 
        npn[22:24],   # Edificio 
        npn[24:26],   # Número Piso 
        npn[26:30],   # Unidad Predial
    ]

def extraer_tabla_de_gdb_area_construida(gdb_path, tipo_area):
    """
    Lee la tabla unidadconstruccion (rural/urbano) en DataFrame.
    Convierte nombres de columnas a minúsculas en el DF.
    """
    table_name = "r_lc_unidadconstruccion" if tipo_area == "Rural" else "u_lc_unidadconstruccion"
    table_path = os.path.join(gdb_path, table_name)

    if not arcpy.Exists(table_path):
        raise Exception("La tabla {} no existe en la geodatabase seleccionada.".format(table_name))

    fields_lower, field_map = get_fields_lower(table_path)

    # Excluir geometry y OID
    usable_fields_real = []
    for f in arcpy.ListFields(table_path):
        if f.type not in ["Geometry", "OID"]:
            usable_fields_real.append(f.name)

    data = []
    with arcpy.da.SearchCursor(table_path, usable_fields_real) as cursor:
        for row in cursor:
            row_dict = {}
            for idx, real_name in enumerate(usable_fields_real):
                row_dict[real_name.lower()] = row[idx]
            data.append(row_dict)

    df = pd.DataFrame(data)
    return df  # columnas ya en minúsculas

class GDBExcelValidator(Frame):
    def __init__(self, parent):
        Frame.__init__(self, parent)
        self.parent = parent
        self.parent.winfo_toplevel().title("Validación GDB vs Excel")

        self.gdb_path = tk.StringVar()
        self.excel_path = tk.StringVar()
        self.excel_path_bcgs = tk.StringVar()
        self.output_excel = tk.StringVar()
        self.tipo_area = tk.StringVar(value="Rural")  # default

        self.notebook = ttk.Notebook(parent)
        self.notebook.pack(fill='both', expand=True)

        self.tab_topologia = tk.Frame(self.notebook)
        self.notebook.add(self.tab_topologia, text="Topología")

        btn_validar_topologia = tk.Button(
            self.tab_topologia,
            text="Validar Topología",
            command=self.ejecutar_validacion_topologia
        )
        btn_validar_topologia.pack(pady=10)

        tk.Label(parent, text="Seleccionar GDB:").pack()
        tk.Entry(parent, textvariable=self.gdb_path, width=50).pack()
        tk.Button(parent, text="Buscar GDB", command=self.select_gdb).pack()

        tk.Label(parent, text="Seleccionar Excel:").pack()
        tk.Entry(parent, textvariable=self.excel_path, width=50).pack()
        tk.Button(parent, text="Buscar Excel", command=self.select_excel).pack()

        tk.Label(parent, text="Seleccionar Excel BCGS:").pack()
        tk.Entry(parent, textvariable=self.excel_path_bcgs, width=50).pack()
        tk.Button(parent, text="Buscar Excel BCGS", command=self.select_excel_bcgs).pack()

        tk.Label(parent, text="Seleccione el tipo de área:").pack()
        tk.Radiobutton(parent, text="Rural", variable=self.tipo_area, value="Rural").pack()
        tk.Radiobutton(parent, text="Urbano", variable=self.tipo_area, value="Urbano").pack()

        tk.Button(parent, text="Ejecutar Validación", command=self.run_validation).pack()

    def run_validation(self):
        gdb_path = self.gdb_path.get()
        excel_path = self.excel_path.get()
        excel_path_bcgs = self.excel_path_bcgs.get()

        feature_class_name_terreno = "r_lc_terreno" if self.tipo_area.get() == "Rural" else "u_lc_terreno"
        feature_class_name_unidad_construccion = "r_lc_unidadconstruccion" if self.tipo_area.get() == "Rural" else "u_lc_unidadconstruccion"
        feature_class_name_construccion = "r_lc_construccion" if self.tipo_area.get() == "Rural" else "u_lc_construccion"

        self.select_output_excel()
        output_path = self.output_excel.get()
        if not output_path:
            tkMessageBox.showerror("Error", "Debe seleccionar una ruta de salida válida antes de continuar.")
            return

        if not gdb_path or not excel_path:
            tkMessageBox.showerror("Error", "Debe seleccionar la GDB y el archivo Excel.")
            return

        feature_class_path_terreno = os.path.join(gdb_path, feature_class_name_terreno)

        try:
            # --- Excel (para Npn)
            wb = xlrd.open_workbook(excel_path)
            sheet_comparacion = wb.sheet_by_index(0)

            npn_col_idx = None
            for i in range(sheet_comparacion.ncols):
                if sheet_comparacion.cell_value(0, i).strip().lower() == 'npn':
                    npn_col_idx = i
                    break
            if npn_col_idx is None:
                tkMessageBox.showerror("Error", "La columna 'Npn' no se encuentra en el Excel.")
                return

            npn_excel = [
                str(sheet_comparacion.cell_value(row, npn_col_idx).encode('utf-8'))
                for row in range(1, sheet_comparacion.nrows)
            ]

            # --- Campos de la capa terreno
            arcpy.env.workspace = gdb_path
            fields_lower_terr, fmap_terr = get_fields_lower(feature_class_path_terreno)

            terreno_codigo_real = fmap(fmap_terr, "terreno_codigo")
            if not terreno_codigo_real:
                tkMessageBox.showerror("Error", "La columna 'terreno_codigo' no existe en la Feature Class.")
                return

            # Leer terreno_codigo
            terreno_codigo_gdb = [
                str(row[0]) for row in arcpy.da.SearchCursor(feature_class_path_terreno, [terreno_codigo_real])
            ]

            diff_terreno_codigo = set(terreno_codigo_gdb) - set(npn_excel)
            diff_npn_excel = set(npn_excel) - set(terreno_codigo_gdb)

            def filtrar_omisiones(npn_list):
                salida = []
                for npn in npn_list:
                    if len(npn) >= 30:
                        cp = npn[21]
                        if cp in ('8', '9', '0', '2', '4', '5', '3'):
                            if (cp == '8') or (npn[26:30] == '0000'):
                                salida.append(npn)
                return salida

            self.verificar_geometrias_vacias(os.path.join(gdb_path, feature_class_name_construccion))
            omisiones_filtradas = filtrar_omisiones(diff_npn_excel)

            # --- Crear libro de salida
            workbook = xlwt.Workbook()
            sheet_comisiones = workbook.add_sheet('Comisiones')
            sheet_omisiones = workbook.add_sheet('Omisiones')
            sheet_diferencias = workbook.add_sheet('Diferencia Areas Construidas')
            sheet_df_comparar_areas_por_unidad = workbook.add_sheet('Comparar Areas por Unidad')
            sheet_duplicados = workbook.add_sheet('Npn Duplicados')
            sheet_ph_sin_unidad = workbook.add_sheet('PH sin Unidad Predial')
            sheet_terreno_nro_piso = workbook.add_sheet('Terreno con Nro Piso')
            sheet_duplicados_terreno = workbook.add_sheet('Terreno Duplicados')
            sheet_informalidades_sin_predio_formal = workbook.add_sheet('Informalidades Sin P')
            sheet_df_informalidad_condicion2_vs_formal_area = workbook.add_sheet('Area informal superior a Fomral')
            sheet_npn__unidad_diferente_de_terreno = workbook.add_sheet('Npn Unidad Dif De Terreno')
            sheet_npn__construccion_diferente_de_terreno = workbook.add_sheet('Npn Construccion Dif De Terreno')
            sheet_npn_validacion_informalidad_sobre_predio = workbook.add_sheet('informalidad_sobre_predio')
            sheet_etiqueta = workbook.add_sheet('Etiqueta')

            headers = [
                "Npn", "Departamento", "Municipio", "Zona", "Sector", "Comuna", "Barrio",
                "Manzana o Vereda","Terreno o Predios", "Condición Predio", "Edificio",
                "Número Piso", "Unidad Predial"
            ]

            bold_style = xlwt.XFStyle()
            bold_font = xlwt.Font()
            bold_font.bold = True
            bold_style.font = bold_font

            sheet_comisiones.panes_frozen = True
            sheet_comisiones.horz_split_pos = 1
            sheet_omisiones.panes_frozen = True
            sheet_omisiones.horz_split_pos = 1

            column_widths = [len(h) for h in headers]

            for col_num, header in enumerate(headers):
                sheet_comisiones.write(0, col_num, header.decode('utf-8'), bold_style)
                sheet_omisiones.write(0, col_num, header.decode('utf-8'), bold_style)

            for row_num, npn in enumerate(diff_terreno_codigo, 1):
                cols = dividir_npn_en_columnas(npn)
                fila = [npn] + cols
                for col_num, valor in enumerate(fila):
                    valor_utf = valor.decode('utf-8')
                    sheet_comisiones.write(row_num, col_num, valor_utf)
                    column_widths[col_num] = max(column_widths[col_num], len(valor_utf))

            for row_num, npn in enumerate(omisiones_filtradas, 1):
                cols = dividir_npn_en_columnas(npn)
                fila = [npn] + cols
                for col_num, valor in enumerate(fila):
                    valor_utf = valor.decode('utf-8')
                    sheet_omisiones.write(row_num, col_num, valor_utf)
                    column_widths[col_num] = max(column_widths[col_num], len(valor_utf))

            for col_num, width in enumerate(column_widths):
                sheet_comisiones.col(col_num).width = (width + 2) * 256
                sheet_omisiones.col(col_num).width = (width + 2) * 256

            # --- cálculos auxiliares
            """
            
            df_informalidad_sobre_predio = self.validacion_informalidad_sobre_predio(gdb_path)
            df_validar_terreno_codigo_duplicado_ficha = self.validar_terreno_codigo_duplicado_ficha(gdb_path)
            self.extraer_letras_identificador(gdb_path)

            df_diferencias_areas_construidas = self.calcular_areas_construidas()
            df_npns_duplicados = self.validar_terreno_codigo_duplicado(gdb_path)
            df_ph_sin_unidad = self.calcular_campos_y_filtrar(gdb_path, self.tipo_area.get())
            df_terreno_con_nro_piso = self.validar_terreno_con_piso(gdb_path, self.tipo_area.get())
            df_informalidades_sin_predio_formal = self.copiar_filtrar_buffer_y_join(gdb_path)
            df_npn__unidad_diferente_de_terreno = self.validar_npn__unidad_diferente_de_terreno(gdb_path)
            df_npn__construccion_diferente_de_terreno = self.validar_npn__construccion_diferente_de_terreno(gdb_path)
            df_etiqueta = self.validar_etiqueta(gdb_path)
            df_comparar_areas_por_unidad=self.comparar_areas_por_unidad()
            df_informalidad_condicion2_vs_formal_area=self.informalidades_suman_mas_que_formal(gdb_path)
            """
            reportes_dict = {
                u"Comisiones": diff_terreno_codigo,
                u"Omisiones": omisiones_filtradas,
                #u"Diferencias de area > 2.5": df_diferencias_areas_construidas,
                #u"NPNs Duplicados": df_npns_duplicados,
                #u"Terreno Duplicados": df_validar_terreno_codigo_duplicado_ficha,
                #u"PH sin Unidad Predial": df_ph_sin_unidad,
                #u"Terrenos con Numero de Piso": df_terreno_con_nro_piso,
                #u"Informalidades Sin Predio Formal": df_informalidades_sin_predio_formal,
                #u"NPN Unidad Diferente de Terreno": df_npn__unidad_diferente_de_terreno,
                #u"NPN Construcción Diferente de Terreno": df_npn__construccion_diferente_de_terreno,
                #u"Informalidades Sobre Predio": df_informalidad_sobre_predio,
                #u"Etiqueta": df_etiqueta,
                #u"Comparar Areas por Unidad": df_comparar_areas_por_unidad,
                #u"Area informal superior a Fomral": df_informalidad_condicion2_vs_formal_area,
            }

            # reporte resumen
            self.reporte(workbook, reportes_dict, gdb_path)

            # ---- volcar dfs en las hojas correspondientes ----
            # diferencias áreas
            """
            
            for col_num, column in enumerate(df_diferencias_areas_construidas.columns):
                sheet_diferencias.write(0, col_num, column.decode('utf-8'), bold_style)

            for row_num, row in enumerate(df_diferencias_areas_construidas.itertuples(index=False), 1):
                for col_num, value in enumerate(row):
                    if isinstance(value, float):
                        value_str = ('%.2f' % value).replace('.', ',')
                        sheet_diferencias.write(row_num, col_num, value_str.decode('utf-8'))
                    else:
                        sheet_diferencias.write(row_num, col_num, str(value).decode('utf-8'))

            # terreno con nro piso
            for col_num, column in enumerate(df_terreno_con_nro_piso.columns):
                sheet_terreno_nro_piso.write(0, col_num, column.decode('utf-8'), bold_style)
            for row_num, row in enumerate(df_terreno_con_nro_piso.itertuples(index=False), 1):
                for col_num, value in enumerate(row):
                    sheet_terreno_nro_piso.write(row_num, col_num, str(value).decode('utf-8'))

            # npn duplicados
            for col_num, column in enumerate(df_npns_duplicados.columns):
                sheet_duplicados.write(0, col_num, column.decode('utf-8'), bold_style)
            for row_num, row in enumerate(df_npns_duplicados.itertuples(index=False), 1):
                for col_num, value in enumerate(row):
                    sheet_duplicados.write(row_num, col_num, str(value).decode('utf-8'))

            # ph sin unidad
            for col_num, column in enumerate(df_ph_sin_unidad.columns):
                sheet_ph_sin_unidad.write(0, col_num, column.decode('utf-8'), bold_style)
            for row_num, row in enumerate(df_ph_sin_unidad.itertuples(index=False), 1):
                for col_num, value in enumerate(row):
                    sheet_ph_sin_unidad.write(row_num, col_num, str(value).decode('utf-8'))

            # informalidades sin predio formal
            for col_num, column in enumerate(df_informalidades_sin_predio_formal.columns):
                sheet_informalidades_sin_predio_formal.write(0, col_num, column.decode('utf-8'), bold_style)
            for row_num, row in enumerate(df_informalidades_sin_predio_formal.itertuples(index=False), 1):
                for col_num, value in enumerate(row):
                    sheet_informalidades_sin_predio_formal.write(row_num, col_num, str(value).decode('utf-8'))

            # unidad diferente de terreno
            for col_num, column in enumerate(df_npn__unidad_diferente_de_terreno.columns):
                sheet_npn__unidad_diferente_de_terreno.write(0, col_num, column.decode('utf-8'), bold_style)
            for row_num, row in enumerate(df_npn__unidad_diferente_de_terreno.itertuples(index=False), 1):
                for col_num, value in enumerate(row):
                    sheet_npn__unidad_diferente_de_terreno.write(row_num, col_num, str(value).decode('utf-8'))

            # construccion diferente de terreno
            for col_num, column in enumerate(df_npn__construccion_diferente_de_terreno.columns):
                sheet_npn__construccion_diferente_de_terreno.write(0, col_num, column.decode('utf-8'), bold_style)
            for row_num, row in enumerate(df_npn__construccion_diferente_de_terreno.itertuples(index=False), 1):
                for col_num, value in enumerate(row):
                    sheet_npn__construccion_diferente_de_terreno.write(row_num, col_num, str(value).decode('utf-8'))

            # informalidad sobre predio
            for col_num, column in enumerate(df_informalidad_sobre_predio.columns):
                sheet_npn_validacion_informalidad_sobre_predio.write(0, col_num, column.decode('utf-8'), bold_style)
            for row_num, row in enumerate(df_informalidad_sobre_predio.itertuples(index=False), 1):
                for col_num, value in enumerate(row):
                    sheet_npn_validacion_informalidad_sobre_predio.write(row_num, col_num, str(value).decode('utf-8'))

            # duplicados terreno
            for col_num, column in enumerate(df_validar_terreno_codigo_duplicado_ficha.columns):
                sheet_duplicados_terreno.write(0, col_num, column.decode('utf-8'), bold_style)
            for row_num, row in enumerate(df_validar_terreno_codigo_duplicado_ficha.itertuples(index=False), 1):
                for col_num, value in enumerate(row):
                    sheet_duplicados_terreno.write(row_num, col_num, str(value).decode('utf-8'))

            # etiqueta
            for col_num, column in enumerate(df_etiqueta.columns):
                sheet_etiqueta.write(0, col_num, column.decode('utf-8'), bold_style)
            for row_num, row in enumerate(df_etiqueta.itertuples(index=False), 1):
                for col_num, value in enumerate(row):
                    sheet_etiqueta.write(row_num, col_num, str(value).decode('utf-8'))

            for col_num, column in enumerate(df_comparar_areas_por_unidad.columns):
                sheet_df_comparar_areas_por_unidad.write(0, col_num, column.decode('utf-8'), bold_style)

            # Escribir datos
            for row_num, row in enumerate(df_comparar_areas_por_unidad.itertuples(index=False), 1):
                for col_num, value in enumerate(row):
                    sheet_df_comparar_areas_por_unidad.write(row_num, col_num, str(value).decode('utf-8'))

            for col_num, column in enumerate(df_informalidad_condicion2_vs_formal_area.columns):
                sheet_df_informalidad_condicion2_vs_formal_area.write(0, col_num, column.decode('utf-8'), bold_style)

            # Escribir datos
            for row_num, row in enumerate(df_informalidad_condicion2_vs_formal_area.itertuples(index=False), 1):
                for col_num, value in enumerate(row):
                    sheet_df_informalidad_condicion2_vs_formal_area.write(row_num, col_num, str(value).decode('utf-8'))

            """

            workbook.save(output_path)

            tkMessageBox.showinfo(
                "Éxito".decode('utf-8'),
                u"Proceso finalizado correctamente.\nArchivos guardados en:\n" + output_path.decode('utf-8')
            )

        except Exception as e:
            tkMessageBox.showerror("Error", str(e))

    def ejecutar_validacion_topologia(self):
        gdb_path = self.gdb_path.get()
        if not gdb_path or not gdb_path.endswith(".gdb"):
            tkMessageBox.showerror("Error", "Debe seleccionar una GDB válida.")
            return

        ruta_excel = os.path.join(os.path.dirname(gdb_path), "reporte_topologia.xls")
        try:
            self.validar_topologia(gdb_path, ruta_excel)
        except Exception as e:
            tkMessageBox.showerror("Error", "Fallo la validación de topología:\n{}".format(str(e)))

    def validar_topologia(self, gdb_path, output_excel):
        arcpy.env.workspace = gdb_path
        arcpy.env.overwriteOutput = True

        reglas = [
            {"layer1": "Manzanas", "layer2": "Manzanas", "tipo": "intersect", "regla": "Must Not Overlap"},
            {"layer1": "r_lc_terreno", "layer2": "r_lc_terreno", "tipo": "intersect", "regla": "Must Not Overlap"},
            {"layer1": "u_lc_terreno", "layer2": "u_lc_terreno", "tipo": "intersect", "regla": "Must Not Overlap"},
            {"layer1": "Veredas", "layer2": "Veredas", "tipo": "intersect", "regla": "Must Not Overlap"},
            {"layer1": "Barrios", "layer2": "Barrios", "tipo": "intersect", "regla": "Must Not Overlap"},
            {"layer1": "r_lc_terreno", "layer2": "Veredas", "tipo": "covered_by", "regla": "Must Be Covered By"},
            {"layer1": "u_lc_terreno", "layer2": "Manzanas", "tipo": "covered_by", "regla": "Must Be Covered By"},
            {"layer1": "Manzanas", "layer2": "u_lc_terreno", "tipo": "intersect", "regla": "Must Not Overlap With"},
            {"layer1": "r_lc_terreno", "layer2": "u_lc_terreno", "tipo": "intersect", "regla": "Must Not Overlap With"},
            {"layer1": "Manzanas", "layer2": "r_lc_terreno", "tipo": "intersect", "regla": "Must Not Overlap With"},
            {"layer1": "Veredas", "layer2": "u_lc_terreno", "tipo": "intersect", "regla": "Must Not Overlap With"},
        ]

        resultados = {}

        for regla in reglas:
            lyr1 = regla["layer1"]
            lyr2 = regla["layer2"]
            tipo = regla["tipo"]

            nombre_hoja = "{}_{}_{}".format(
                lyr1,
                regla["regla"].replace(" ", "_"),
                lyr2 if lyr1 != lyr2 else "Interno"
            )[:31]

            try:
                capa1 = os.path.join(gdb_path, lyr1)
                capa2 = os.path.join(gdb_path, lyr2)

                if not arcpy.Exists(capa1) or not arcpy.Exists(capa2):
                    resultados[nombre_hoja] = pd.DataFrame([{
                        "Error": "Capa faltante: {} o {}".format(lyr1, lyr2)
                    }])
                    continue

                if tipo == "intersect":
                    temp_out = os.path.join("in_memory", "intersect_" + lyr1 + "_" + lyr2)
                    arcpy.Intersect_analysis([capa1, capa2], temp_out, output_type="INPUT")
                    arcpy.AddGeometryAttributes_management(
                        temp_out,
                        "AREA",
                        Area_Unit="SQUARE_METERS"
                    )

                    fields_needed = []
                    f1 = "FID_{}".format(lyr1)
                    f2 = "FID_{}".format(lyr2)
                    a_field = "POLY_AREA"
                    fields_needed.extend([f1, f2, a_field])

                    rows = []
                    with arcpy.da.SearchCursor(temp_out, fields_needed) as cursor:
                        for row in cursor:
                            area_val = row[2]
                            if area_val >= 1:
                                rows.append({
                                    f1: row[0],
                                    f2: row[1],
                                    "Area_m2": round(area_val, 2)
                                })

                    if rows:
                        resultados[nombre_hoja] = pd.DataFrame(rows)
                    else:
                        resultados[nombre_hoja] = pd.DataFrame([{
                            "Resultado": "Sin traslapes detectados"
                        }])

                elif tipo == "covered_by":
                    arcpy.MakeFeatureLayer_management(capa2, "lyr_cobertura")
                    arcpy.MakeFeatureLayer_management(capa1, "lyr_objetivo")

                    # invert_spatial_relationship="INVERT" -> selecciona los que NO están cubiertos
                    arcpy.SelectLayerByLocation_management(
                        "lyr_objetivo",
                        "ARE_IDENTICAL_TO",
                        "lyr_cobertura",
                        invert_spatial_relationship="INVERT"
                    )

                    ids = [row[0] for row in arcpy.da.SearchCursor("lyr_objetivo", ["OID@"])]
                    if ids:
                        resultados[nombre_hoja] = pd.DataFrame(
                            ids, columns=["OID_{}".format(lyr1)]
                        )
                    else:
                        resultados[nombre_hoja] = pd.DataFrame([{
                            "Resultado": "Todos cubiertos correctamente"
                        }])

            except Exception as e:
                resultados[nombre_hoja] = pd.DataFrame([{"Error": str(e)}])

        writer = pd.ExcelWriter(output_excel, engine='xlwt')
        for hoja, df in resultados.items():
            df.to_excel(writer, sheet_name=hoja, index=False)
        writer.save()

        tkMessageBox.showinfo("Validación completada", "Archivo guardado:\n{}".format(output_excel))

    def calcular_areas_construidas(self):
        # Usa nombres de columna en minúscula
        reload(sys)
        sys.setdefaultencoding('utf-8')

        gdb_path = self.gdb_path.get()
        excel_path = self.excel_path.get()

        if not excel_path:
            tkMessageBox.showerror("Error", "Debe seleccionar el Excel.")
            return

        # Aceptar carpeta que contiene una .gdb
        if not gdb_path.endswith(".gdb"):
            gdbs = [f for f in os.listdir(gdb_path) if f.endswith(".gdb")]
            if gdbs:
                gdb_path = os.path.join(gdb_path, gdbs[0])
            else:
                tkMessageBox.showerror("Error", "No se encontró ninguna GDB.")
                return

        print("calcular_areas_construidas")
        arcpy.env.workspace = gdb_path
        arcpy.env.overwriteOutput = True

        df_unidadconstruccion = extraer_tabla_de_gdb_area_construida(gdb_path, self.tipo_area.get())
        if df_unidadconstruccion is None or df_unidadconstruccion.empty:
            tkMessageBox.showerror("Error", "No se pudo obtener datos de la GDB.")
            return

        # columnas clave esperadas en minúsculas
        for col in ["codigo_unidad_construccion", "shape_area"]:
            if col not in df_unidadconstruccion.columns:
                tkMessageBox.showerror("Error", "Falta la columna '{}' en datos de GDB.".format(col))
                return

        try:
            xl = pd.ExcelFile(excel_path)
        except Exception as e:
            tkMessageBox.showerror("Error", "No se pudo abrir el Excel:\n{}".format(str(e)))
            return

        hoja_construcciones = None
        hoja_fichas = None
        for h in xl.sheet_names:
            l = h.strip().lower()
            if ("construccion" in l) and hoja_construcciones is None:
                hoja_construcciones = h
            if ("ficha" in l) and hoja_fichas is None:
                hoja_fichas = h

        if not hoja_construcciones or not hoja_fichas:
            tkMessageBox.showerror("Error", "No se encontró hoja 'Construcciones' y/o 'Fichas'.")
            return

        try:
            df_construcciones = pd.read_excel(excel_path, sheetname=hoja_construcciones)
            df_fichas = pd.read_excel(excel_path, sheetname=hoja_fichas)
        except Exception as e:
            tkMessageBox.showerror("Error", "No se pudo leer Excel:\n{}".format(str(e)))
            return

        # normalizar columnas excel a minúsculas para trabajar igual
        df_construcciones.columns = [c.strip().lower() for c in df_construcciones.columns]
        df_fichas.columns = [c.strip().lower() for c in df_fichas.columns]

        if "nroficha" not in df_construcciones.columns and "nroficha" not in df_fichas.columns:
            tkMessageBox.showerror("Error", "Falta la columna 'NroFicha' en el Excel.")
            return

        if "areaconstruida" not in df_construcciones.columns:
            df_construcciones["areaconstruida"] = 0

        df_construcciones["areaconstruida"] = pd.to_numeric(
            df_construcciones["areaconstruida"], errors='coerce'
        ).fillna(0)

        if "nroficha" not in df_construcciones.columns or "nroficha" not in df_fichas.columns:
            tkMessageBox.showerror("Error", "No se encuentra 'NroFicha' en ambas hojas para unir.")
            return

        df_fichas = df_fichas.merge(
            df_construcciones[["nroficha", "areaconstruida"]],
            on="nroficha", how="left"
        )
        df_fichas["areaconstruida"] = pd.to_numeric(
            df_fichas["areaconstruida"], errors='coerce'
        ).fillna(0)

        # Agrupar GDB por prefijo 22
        df_gdb = df_unidadconstruccion[["codigo_unidad_construccion", "shape_area"]].copy()
        df_gdb["clave22"] = df_gdb["codigo_unidad_construccion"].astype(str).str[:22]
        df_gdb["area_gdb"] = pd.to_numeric(df_gdb["shape_area"], errors='coerce').fillna(0)
        df_agr_gdb = df_gdb.groupby("clave22", as_index=False)["area_gdb"].sum()

        # Agrupar Excel por prefijo 22
        if "npn" not in df_fichas.columns:
            tkMessageBox.showerror("Error", "No se encuentra columna 'Npn' en la hoja Fichas.")
            return

        df_excel = df_fichas[["npn", "areaconstruida"]].copy()
        df_excel["clave22"] = df_excel["npn"].astype(str).str[:22]
        df_excel["area_excel"] = pd.to_numeric(df_excel["areaconstruida"], errors='coerce').fillna(0)
        df_agr_excel = df_excel.groupby("clave22", as_index=False)["area_excel"].sum()

        df_comp = df_agr_gdb.merge(df_agr_excel, on="clave22", how="inner")

        df_comp["area_gdb"] = df_comp["area_gdb"].round(2)
        df_comp["area_excel"] = df_comp["area_excel"].round(2)
        df_comp["diferencia"] = (df_comp["area_gdb"] - df_comp["area_excel"]).abs().round(2)

        df_comp.rename(columns={"clave22": "codigo/npn (22)"}, inplace=True)

        df_filtrada = df_comp[df_comp["diferencia"] > 2.5].copy()

        def _calc_pct(gdb_val, xls_val):
            if gdb_val == 0:
                return 0.0
            return round((gdb_val - xls_val) / float(gdb_val), 4)

        df_filtrada["porcentaje"] = [
            _calc_pct(g, e) for g, e in zip(df_filtrada["area_gdb"], df_filtrada["area_excel"])
        ]

        return df_filtrada

    def validar_terreno_codigo_duplicado(self, gdb_path):
        feature_class_name = "r_lc_terreno" if self.tipo_area.get() == "Rural" else "u_lc_terreno"
        feature_class_path = os.path.join(gdb_path, feature_class_name)
        arcpy.env.workspace = gdb_path

        if not arcpy.Exists(feature_class_path):
            raise Exception("La capa {} no existe en la GDB.".format(feature_class_name))

        print("validar_terreno_codigo_duplicado")

        fields_lower, fmap_terr = get_fields_lower(feature_class_path)
        terreno_codigo_real = fmap(fmap_terr, "terreno_codigo")
        if not terreno_codigo_real:
            raise Exception("La columna 'terreno_codigo' no existe en la Feature Class.")

        terreno_codigos = [
            row[0] for row in arcpy.da.SearchCursor(feature_class_path, [terreno_codigo_real]) if row[0]
        ]

        contador = {}
        for codigo in terreno_codigos:
            contador[codigo] = contador.get(codigo, 0) + 1

        duplicados = [codigo for codigo, cnt in contador.items() if cnt > 1]

        if duplicados:
            data = []
            for npn in duplicados:
                cols = dividir_npn_en_columnas(npn)
                data.append([npn] + cols)

            columnas_df = [
                "terreno_codigo", "Departamento", "Municipio", "Zona", "Sector",
                "Comuna", "Barrio", "Manzana/Vereda", "Terreno/Predios",
                "Condición Predio", "Edificio", "Número Piso", "Unidad Predial"
            ]
            df = pd.DataFrame(data, columns=columnas_df)
            return df
        else:
            return pd.DataFrame(columns=["terreno_codigo"])

    def calcular_campos_y_filtrar(self, gdb_path, tipo_area):
        feature_class_name = "r_lc_unidadconstruccion" if tipo_area == "Rural" else "u_lc_unidadconstruccion"
        feature_class_path = os.path.join(gdb_path, feature_class_name)
        filtro_feature_class_path = os.path.join(gdb_path, "filtro_unidadconstruccion_campos_filtrar")
        arcpy.env.workspace = gdb_path

        if not arcpy.Exists(feature_class_path):
            raise Exception("La capa {} no existe en la GDB.".format(feature_class_name))
        print("calcular_campos_y_filtrar")

        fields_lower, fmap_uni = get_fields_lower(feature_class_path)

        cp_real = fmap(fmap_uni, "cp")
        if not cp_real:
            arcpy.AddField_management(feature_class_path, "CP", "TEXT", field_length=1)
            cp_real = "CP"

        unidad_real = fmap(fmap_uni, "unidad")
        if not unidad_real:
            arcpy.AddField_management(feature_class_path, "UNIDAD", "TEXT", field_length=4)
            unidad_real = "UNIDAD"

        codigo_uc_real = fmap(fmap_uni, "codigo_unidad_construccion")
        if not codigo_uc_real:
            raise Exception("No se encontró el campo 'codigo_unidad_construccion' en la capa unidadconstruccion.")

        dominio_real = fmap(fmap_uni, "tipo_dominio")
        if not dominio_real:
            raise Exception("No se encontró el campo 'tipo_dominio' en la capa unidadconstruccion.")

        # CP = Mid([CODIGO_UNIDAD_CONSTRUCCION], 22, 1)
        arcpy.CalculateField_management(
            feature_class_path,
            cp_real,
            "Mid([{0}], 22, 1)".format(codigo_uc_real),
            "VB"
        )
        # UNIDAD = Right([CODIGO_UNIDAD_CONSTRUCCION],4)
        arcpy.CalculateField_management(
            feature_class_path,
            unidad_real,
            "Right([{0}],4)".format(codigo_uc_real),
            "VB"
        )

        query_filtro = "{cp} = '9' AND {uni} = '0000' AND {dom} = 1".format(
            cp=cp_real, uni=unidad_real, dom=dominio_real
        )

        arcpy.MakeFeatureLayer_management(
            feature_class_path,
            "temp_layer_name_calcular_campos_filtrar",
            query_filtro
        )
        arcpy.CopyFeatures_management(
            "temp_layer_name_calcular_campos_filtrar",
            filtro_feature_class_path
        )

        if arcpy.Exists(filtro_feature_class_path):
            f_low_sub, fmap_sub = get_fields_lower(filtro_feature_class_path)

            # leer todos los campos de esa capa intermedia
            all_real = [fmap_sub[x] for x in f_low_sub]

            data = [list(row) for row in arcpy.da.SearchCursor(filtro_feature_class_path, all_real)]
            df = pd.DataFrame(data, columns=f_low_sub)

            if "codigo_unidad_construccion" in df.columns:
                work_col = "codigo_unidad_construccion"
            else:
                # buscar equivalente real
                work_col = None
                for cand in df.columns:
                    if cand.endswith("codigo_unidad_construccion"):
                        work_col = cand
                        break
                if not work_col and len(df.columns) > 0:
                    work_col = df.columns[0]

            df = df[[work_col]]
            df.columns = ["codigo_unidad_construccion"]

            if not df.empty:
                data2 = []
                for npn in df["codigo_unidad_construccion"]:
                    cols = dividir_npn_en_columnas(npn)
                    data2.append([npn] + cols)

                columnas_df = [
                    "codigo_unidad_construccion", "Departamento", "Municipio", "Zona", "Sector",
                    "Comuna", "Barrio", "Manzana/Vereda", "Terreno/Predios",
                    "Condición Predio", "Edificio", "Número Piso", "Unidad Predial"
                ]
                df = pd.DataFrame(data2, columns=columnas_df)

            return df
        else:
            return pd.DataFrame(columns=["codigo_unidad_construccion"])

    def validar_terreno_con_piso(self, gdb_path, tipo_area):
        feature_class_name = "r_lc_terreno" if tipo_area == "Rural" else "u_lc_terreno"
        feature_class_path = os.path.join(gdb_path, feature_class_name)
        filtro_feature_class_path = os.path.join(gdb_path, "filtro_unidadconstruccion_terreno_con_piso")

        arcpy.env.workspace = gdb_path

        if not arcpy.Exists(feature_class_path):
            raise Exception("La capa {} no existe en la GDB.".format(feature_class_name))
        print("validar_terreno_con_piso")

        fields_lower, fmap_terr = get_fields_lower(feature_class_path)

        terreno_codigo_real = fmap(fmap_terr, "terreno_codigo")
        if not terreno_codigo_real:
            raise Exception("La columna 'terreno_codigo' no existe en la Feature Class.")

        cp_real = fmap(fmap_terr, "cp")
        if not cp_real:
            arcpy.AddField_management(feature_class_path, "CP", "TEXT", field_length=1)
            cp_real = "CP"

        ult8_real = fmap(fmap_terr, "ultimos_8")
        if not ult8_real:
            arcpy.AddField_management(feature_class_path, "ULTIMOS_8", "TEXT", field_length=8)
            ult8_real = "ULTIMOS_8"

        arcpy.CalculateField_management(
            feature_class_path,
            cp_real,
            "Mid([{0}], 22, 1)".format(terreno_codigo_real),
            "VB"
        )
        arcpy.CalculateField_management(
            feature_class_path,
            ult8_real,
            "Right([{0}],8)".format(terreno_codigo_real),
            "VB"
        )

        query_filtro = "{cp} <> '8' AND {ult} <> '00000000'".format(
            cp=cp_real, ult=ult8_real
        )

        arcpy.MakeFeatureLayer_management(
            feature_class_path,
            "temp_layer_name_terreno_piso",
            query_filtro
        )
        arcpy.CopyFeatures_management(
            "temp_layer_name_terreno_piso",
            filtro_feature_class_path
        )

        if arcpy.Exists(filtro_feature_class_path):
            f_low_sub, fmap_sub = get_fields_lower(filtro_feature_class_path)
            terr_real_sub = fmap(fmap_sub, "terreno_codigo")

            all_real = [fmap_sub[x] for x in f_low_sub]
            data = [list(row) for row in arcpy.da.SearchCursor(filtro_feature_class_path, all_real)]
            df = pd.DataFrame(data, columns=f_low_sub)

            out_col = "terreno_codigo"
            if "terreno_codigo" not in df.columns and terr_real_sub:
                # intentar renombrar
                for c in df.columns:
                    if c.endswith("terreno_codigo"):
                        out_col = c
                        break

            df = df[[out_col]]
            df.columns = ["terreno_codigo"]

            if not df.empty:
                data2 = []
                for npn in df["terreno_codigo"]:
                    cols = dividir_npn_en_columnas(npn)
                    data2.append([npn] + cols)

                columnas_df = [
                    "terreno_codigo", "Departamento", "Municipio", "Zona", "Sector",
                    "Comuna", "Barrio", "Manzana/Vereda", "Terreno/Predios",
                    "Condición Predio", "Edificio", "Número Piso", "Unidad Predial"
                ]
                df = pd.DataFrame(data2, columns=columnas_df)

            return df
        else:
            return pd.DataFrame(columns=["terreno_codigo"])

    def copiar_filtrar_buffer_y_join(self, gdb_path):
        feature_class_name = "r_lc_terreno" if self.tipo_area.get() == "Rural" else "u_lc_terreno"
        feature_class_path = os.path.join(gdb_path, feature_class_name)

        arcpy.env.workspace = gdb_path

        if not arcpy.Exists(feature_class_path):
            tkMessageBox.showerror("Error", "La capa {} no existe en la GDB.".format(feature_class_name))
            return None
        print("copiar_filtrar_buffer_y_join")

        formal_filtrado_fc = os.path.join(gdb_path, "formal_filtrado")
        informal_filtrado_fc = os.path.join(gdb_path, "informal_filtrado")
        informal_buffer_fc = os.path.join(gdb_path, "informal_buffer")
        join_output_fc = os.path.join(gdb_path, "informal_buffer_joined")

        # nombres de campo reales
        fields_lower, fmap_terr = get_fields_lower(feature_class_path)
        terreno_codigo_real = fmap(fmap_terr, "terreno_codigo")

        try:
            query_formal = "SUBSTRING({0}, 22, 1) <> '2'".format(terreno_codigo_real)
            arcpy.MakeFeatureLayer_management(
                feature_class_path, "formal_layer_copiar_filtrar", query_formal
            )
            arcpy.CopyFeatures_management(
                "formal_layer_copiar_filtrar", formal_filtrado_fc
            )

            query_informal = "SUBSTRING({0}, 22, 1) = '2'".format(terreno_codigo_real)
            arcpy.MakeFeatureLayer_management(
                feature_class_path, "informal_layer_copiar_filtrar", query_informal
            )
            arcpy.CopyFeatures_management(
                "informal_layer_copiar_filtrar", informal_filtrado_fc
            )

            buffer_distancia = "-0.5 Meters"
            arcpy.Buffer_analysis(
                informal_filtrado_fc, informal_buffer_fc, buffer_distancia,
                line_side="FULL", line_end_type="ROUND",
                dissolve_option="NONE", method="PLANAR"
            )

            arcpy.SpatialJoin_analysis(
                target_features=informal_buffer_fc,
                join_features=formal_filtrado_fc,
                out_feature_class=join_output_fc,
                join_operation="JOIN_ONE_TO_ONE",
                join_type="KEEP_ALL",
                match_option="INTERSECT"
            )

            # ahora buscamos en el join_output_fc aquellos sin match
            join_low, fmap_join = get_fields_lower(join_output_fc)

            terr_real_1 = None
            # buscamos TERRENO_CODIGO_1 en minúscula
            for k, v in fmap_join.items():
                if k == "terreno_codigo_1":
                    terr_real_1 = v
                    break
            if terr_real_1 is None:
                # fallback: buscar algo que termine en "_1"
                for k, v in fmap_join.items():
                    if k.endswith("_1"):
                        terr_real_1 = v
                        break

            if terr_real_1:
                query_filtro = "{0} IS NULL".format(terr_real_1)
            else:
                # si no encontramos el campo *_1, devolvemos vacío
                return pd.DataFrame(columns=["terreno_codigo"])

            arcpy.MakeFeatureLayer_management(join_output_fc, "filtered_layer", query_filtro)

            # dataframe
            f_low_sub, fmap_sub = get_fields_lower("filtered_layer")
            all_real = [fmap_sub[x] for x in f_low_sub]
            data = [list(row) for row in arcpy.da.SearchCursor("filtered_layer", all_real)]
            df = pd.DataFrame(data, columns=f_low_sub)

            # elegir col terreno_codigo
            terreno_col = "terreno_codigo"
            if "terreno_codigo" not in df.columns:
                # tratar de encontrar algo parecido
                for c in df.columns:
                    if "terreno" in c and "codigo" in c:
                        terreno_col = c
                        break

            df = df[[terreno_col]]
            df.columns = ["terreno_codigo"]

            if not df.empty:
                data2 = []
                for npn in df["terreno_codigo"]:
                    cols = dividir_npn_en_columnas(npn)
                    data2.append([npn] + cols)

                columnas_df = [
                    "terreno_codigo", "Departamento", "Municipio", "Zona", "Sector",
                    "Comuna", "Barrio", "Manzana/Vereda", "Terreno/Predios",
                    "Condición Predio", "Edificio", "Número Piso", "Unidad Predial"
                ]
                df = pd.DataFrame(data2, columns=columnas_df)

            return df

        except Exception as e:
            tkMessageBox.showerror("Error", "No se pudo procesar: {}".format(str(e)))
            print("Error:", e)
            return None

    def validar_npn__unidad_diferente_de_terreno(self, gdb_path):
        feature_class_name_terreno = "r_lc_terreno" if self.tipo_area.get() == "Rural" else "u_lc_terreno"
        feature_class_name_unidad = "r_lc_unidadconstruccion" if self.tipo_area.get() == "Rural" else "u_lc_unidadconstruccion"

        fc_terr = os.path.join(gdb_path, feature_class_name_terreno)
        fc_unid = os.path.join(gdb_path, feature_class_name_unidad)

        arcpy.env.workspace = gdb_path

        if not arcpy.Exists(fc_terr):
            tkMessageBox.showerror("Error", "La capa {} no existe en la GDB.".format(feature_class_name_terreno))
            return None
        if not arcpy.Exists(fc_unid):
            tkMessageBox.showerror("Error", "La capa {} no existe en la GDB.".format(feature_class_name_unidad))
            return None

        print("validar_npn__unidad_diferente_de_terreno")

        unidad_puntos_fc = os.path.join(gdb_path, "unidad_puntos")
        formal_filtrado_fc = os.path.join(gdb_path, "formal_filtrado")
        informal_filtrado_fc = os.path.join(gdb_path, "informal_filtrado")
        informal_dissolve_fc = os.path.join(gdb_path, "informal_dissolved1")
        erase_output_fc = os.path.join(gdb_path, "unico_clip")
        merge_output_fc = os.path.join(gdb_path, "unico")
        intersect_output_fc = os.path.join(gdb_path, "interseccion_unidad")
        intersect_output_fc_filtro = os.path.join(gdb_path, "interseccion_unidad_filtro")

        try:
            for fc in [
                unidad_puntos_fc, formal_filtrado_fc, informal_filtrado_fc,
                informal_dissolve_fc, erase_output_fc, merge_output_fc,
                intersect_output_fc
            ]:
                if arcpy.Exists(fc):
                    arcpy.Delete_management(fc)

            # reparar geometría en unidades
            check_table = os.path.join(gdb_path, "unidad_checkgeom")
            if arcpy.Exists(check_table):
                arcpy.Delete_management(check_table)

            arcpy.CheckGeometry_management(fc_unid, check_table)
            arcpy.RepairGeometry_management(fc_unid, "DELETE_NULL")

            arcpy.MakeFeatureLayer_management(fc_unid, "unidad_layer")
            arcpy.SelectLayerByAttribute_management(
                "unidad_layer", "NEW_SELECTION", "Shape_Area > 0"
            )

            arcpy.FeatureToPoint_management("unidad_layer", unidad_puntos_fc, "INSIDE")

            # copiar terrenos
            arcpy.CopyFeatures_management(fc_terr, merge_output_fc)

            # obtener nombre real terreno_codigo
            terr_low, fmap_terr = get_fields_lower(fc_terr)
            terreno_codigo_real = fmap(fmap_terr, "terreno_codigo")

            query_formal = "SUBSTRING({0}, 22, 1) <> '2'".format(terreno_codigo_real)
            arcpy.MakeFeatureLayer_management(fc_terr, "formal_layer_npn_unidad", query_formal)
            arcpy.CopyFeatures_management("formal_layer_npn_unidad", formal_filtrado_fc)

            query_informal = "SUBSTRING({0}, 22, 1) = '2'".format(terreno_codigo_real)
            arcpy.MakeFeatureLayer_management(fc_terr, "informal_layer_npn_unidad", query_informal)
            arcpy.CopyFeatures_management("informal_layer_npn_unidad", informal_filtrado_fc)

            inf_low, fmap_inf = get_fields_lower(informal_filtrado_fc)
            dissolve_fields = []
            if "dimension" in inf_low:
                dissolve_fields = [fmap_inf["dimension"]]

            if dissolve_fields:
                arcpy.Dissolve_management(
                    informal_filtrado_fc, informal_dissolve_fc, dissolve_fields
                )
            else:
                arcpy.Dissolve_management(
                    informal_filtrado_fc, informal_dissolve_fc
                )

            # erase merge_output_fc con informal_dissolve_fc
            if arcpy.Exists(informal_dissolve_fc) and int(arcpy.GetCount_management(informal_dissolve_fc)[0]) > 0:
                arcpy.Erase_analysis(merge_output_fc, informal_dissolve_fc, erase_output_fc)
            else:
                erase_output_fc = merge_output_fc

            if arcpy.Exists(informal_filtrado_fc) and int(arcpy.GetCount_management(informal_filtrado_fc)[0]) > 0:
                arcpy.Append_management(informal_filtrado_fc, erase_output_fc, "NO_TEST")

            if arcpy.Exists(fc_terr) and arcpy.Exists(unidad_puntos_fc):
                arcpy.Intersect_analysis(
                    [unidad_puntos_fc, erase_output_fc],
                    intersect_output_fc,
                    "ALL", "", "INPUT"
                )

            # ahora agregamos campos calculados en intersección
            inter_low, fmap_inter = get_fields_lower(intersect_output_fc)

            # nombres reales:
            cod_uc_real = fmap(fmap_inter, "codigo_unidad_construccion")
            terr_cod_real = fmap(fmap_inter, "terreno_codigo")

            arcpy.AddField_management(intersect_output_fc, "CP_U", "TEXT", field_length=1)
            arcpy.AddField_management(intersect_output_fc, "EDIFICIO_UNIDAD", "TEXT", field_length=8)
            arcpy.AddField_management(intersect_output_fc, "TERRENO_22", "TEXT", field_length=22)
            arcpy.AddField_management(intersect_output_fc, "UNIDAD_22", "TEXT", field_length=22)
            arcpy.AddField_management(intersect_output_fc, "DIFERENCIA", "LONG")

            arcpy.CalculateField_management(
                intersect_output_fc,
                "CP_U",
                "Mid([{0}], 22, 1)".format(cod_uc_real),
                "VB"
            )
            arcpy.CalculateField_management(
                intersect_output_fc,
                "EDIFICIO_UNIDAD",
                "Right([{0}],8)".format(cod_uc_real),
                "VB"
            )
            arcpy.CalculateField_management(
                intersect_output_fc,
                "TERRENO_22",
                "Left([{0}],22)".format(terr_cod_real),
                "VB"
            )
            arcpy.CalculateField_management(
                intersect_output_fc,
                "UNIDAD_22",
                "Left([{0}],22)".format(cod_uc_real),
                "VB"
            )
            arcpy.CalculateField_management(
                intersect_output_fc,
                "DIFERENCIA",
                "[TERRENO_22] = [UNIDAD_22]",
                "VB"
            )

            query_filtro = "DIFERENCIA = 0"
            arcpy.MakeFeatureLayer_management(
                intersect_output_fc, "temp_layer_name_unidad_dif", query_filtro
            )
            arcpy.CopyFeatures_management(
                "temp_layer_name_unidad_dif", intersect_output_fc_filtro
            )

            # eliminar CP_U == '2' AND EDIFICIO_UNIDAD <> '00000000'
            arcpy.MakeFeatureLayer_management(
                intersect_output_fc_filtro,
                "temp_layer_name_eliminar_unidad",
                "CP_U = '2' AND EDIFICIO_UNIDAD <> '00000000'"
            )
            arcpy.DeleteRows_management("temp_layer_name_eliminar_unidad")
            arcpy.Delete_management("temp_layer_name_eliminar_unidad")

            # a DataFrame
            f_low_sub, fmap_sub = get_fields_lower(intersect_output_fc_filtro)
            all_real = [fmap_sub[x] for x in f_low_sub]
            data = [list(row) for row in arcpy.da.SearchCursor(intersect_output_fc_filtro, all_real)]
            df = pd.DataFrame(data, columns=f_low_sub)

            # Solo queremos terreno_codigo y codigo_unidad_construccion
            out_cols = []
            for want in ["terreno_codigo", "codigo_unidad_construccion"]:
                if want in df.columns:
                    out_cols.append(want)
                else:
                    # buscar algo parecido
                    for c in df.columns:
                        if want in c:
                            out_cols.append(c)
                            break

            df = df[out_cols]
            # renombrar seguros
            rename_map = {}
            for c in df.columns:
                if "terreno" in c and "codigo" in c:
                    rename_map[c] = "terreno_codigo"
                if "codigo_unidad_construccion" in c:
                    rename_map[c] = "codigo_unidad_construccion"
            df = df.rename(columns=rename_map)

            if df.empty:
                print("Advertencia: el DataFrame está vacío después del filtro.")

            return df

        except Exception as e:
            tkMessageBox.showerror("Error", "No se pudo procesar: {}".format(str(e)))
            print("Error:", e)
            return None

    def validar_npn__construccion_diferente_de_terreno(self, gdb_path):
        feature_class_name_terreno = "r_lc_terreno" if self.tipo_area.get() == "Rural" else "u_lc_terreno"
        feature_class_name_construccion = "r_lc_construccion" if self.tipo_area.get() == "Rural" else "u_lc_construccion"

        fc_terr = os.path.join(gdb_path, feature_class_name_terreno)
        fc_cons = os.path.join(gdb_path, feature_class_name_construccion)

        arcpy.env.workspace = gdb_path

        if not arcpy.Exists(fc_terr):
            tkMessageBox.showerror("Error", "La capa {} no existe en la GDB.".format(feature_class_name_terreno))
            return None
        if not arcpy.Exists(fc_cons):
            tkMessageBox.showerror("Error", "La capa {} no existe en la GDB.".format(feature_class_name_construccion))
            return None

        print("validar_npn__construccion_diferente_de_terreno")

        unidad_puntos_fc = os.path.join(gdb_path, "puntos_construccion")
        formal_filtrado_fc = os.path.join(gdb_path, "formal_filtrado_construccion")
        informal_filtrado_fc = os.path.join(gdb_path, "informal_filtrado_construccion")
        informal_dissolve_fc = os.path.join(gdb_path, "informal_dissolved1_construccion")
        erase_output_fc = os.path.join(gdb_path, "unico_clip__construccion")
        merge_output_fc = os.path.join(gdb_path, "unico__construccion")
        intersect_output_fc = os.path.join(gdb_path, "interseccion_construccion")
        intersect_output_fc_filtro = os.path.join(gdb_path, "interseccion_construccion_filtro")

        try:
            for fc in [
                unidad_puntos_fc, formal_filtrado_fc, informal_filtrado_fc,
                informal_dissolve_fc, erase_output_fc, merge_output_fc,
                intersect_output_fc
            ]:
                if arcpy.Exists(fc):
                    arcpy.Delete_management(fc)

            # copiar terrenos al merge
            arcpy.CopyFeatures_management(fc_terr, merge_output_fc)

            # nombres reales
            terr_low, fmap_terr = get_fields_lower(fc_terr)
            terreno_codigo_real = fmap(fmap_terr, "terreno_codigo")

            query_formal = "SUBSTRING({0}, 22, 1) <> '2'".format(terreno_codigo_real)
            arcpy.MakeFeatureLayer_management(fc_terr, "formal_layer_npn_construccion", query_formal)
            arcpy.CopyFeatures_management("formal_layer_npn_construccion", formal_filtrado_fc)

            query_informal = "SUBSTRING({0}, 22, 1) = '2'".format(terreno_codigo_real)
            arcpy.MakeFeatureLayer_management(fc_terr, "informal_layer_npn_construccion", query_informal)
            arcpy.CopyFeatures_management("informal_layer_npn_construccion", informal_filtrado_fc)

            inf_low, fmap_inf = get_fields_lower(informal_filtrado_fc)
            dissolve_fields = []
            if "dimension" in inf_low:
                dissolve_fields = [fmap_inf["dimension"]]

            if dissolve_fields:
                arcpy.Dissolve_management(
                    informal_filtrado_fc, informal_dissolve_fc, dissolve_fields
                )
            else:
                arcpy.Dissolve_management(
                    informal_filtrado_fc, informal_dissolve_fc
                )

            if arcpy.Exists(informal_dissolve_fc) and int(arcpy.GetCount_management(informal_dissolve_fc)[0]) > 0:
                arcpy.Erase_analysis(merge_output_fc, informal_dissolve_fc, erase_output_fc)
            else:
                erase_output_fc = merge_output_fc

            if arcpy.Exists(informal_filtrado_fc) and int(arcpy.GetCount_management(informal_filtrado_fc)[0]) > 0:
                arcpy.Append_management(informal_filtrado_fc, erase_output_fc, "NO_TEST")

            # Intersect construcción con erase_output_fc
            def verificar_geometria(capa):
                with arcpy.da.SearchCursor(capa, ["SHAPE@"]) as cursor:
                    for row in cursor:
                        geom = row[0]
                        if geom is None:
                            return False
                return True

            if verificar_geometria(fc_cons) and verificar_geometria(erase_output_fc):
                arcpy.Intersect_analysis(
                    [fc_cons, erase_output_fc],
                    intersect_output_fc,
                    "ALL", "", "INPUT"
                )
            else:
                print("Algunas capas no tienen geometrías válidas.")

            inter_low, fmap_inter = get_fields_lower(intersect_output_fc)
            terreno_codigo_real_int = fmap(fmap_inter, "terreno_codigo")

            cod_construccion_real = fmap(fmap_inter, "codigo_construccion")
            if cod_construccion_real is None:
                # fallback heurístico
                for k, v in fmap_inter.items():
                    if "construccion" in k and "codigo" in k:
                        cod_construccion_real = v
                        break

            arcpy.AddField_management(intersect_output_fc, "TERRENO_22", "TEXT", field_length=22)
            arcpy.AddField_management(intersect_output_fc, "CONSTRUCCION_22", "TEXT", field_length=22)
            arcpy.AddField_management(intersect_output_fc, "DIFERENCIA", "LONG")

            arcpy.CalculateField_management(
                intersect_output_fc,
                "TERRENO_22",
                "Left([{0}],22)".format(terreno_codigo_real_int),
                "VB"
            )
            arcpy.CalculateField_management(
                intersect_output_fc,
                "CONSTRUCCION_22",
                "Left([{0}],22)".format(cod_construccion_real),
                "VB"
            )
            arcpy.CalculateField_management(
                intersect_output_fc,
                "DIFERENCIA",
                "[TERRENO_22] = [CONSTRUCCION_22]",
                "VB"
            )

            arcpy.MakeFeatureLayer_management(
                intersect_output_fc,
                "temp_layer_name_const_dif",
                "DIFERENCIA = 0"
            )
            arcpy.CopyFeatures_management(
                "temp_layer_name_const_dif",
                intersect_output_fc_filtro
            )

            # dataframe
            f_low_sub, fmap_sub = get_fields_lower(intersect_output_fc_filtro)
            all_real = [fmap_sub[x] for x in f_low_sub]
            data = [list(row) for row in arcpy.da.SearchCursor(intersect_output_fc_filtro, all_real)]
            df = pd.DataFrame(data, columns=f_low_sub)

            # quedarnos con terreno_codigo y codigo_construccion
            out_cols = []
            for want in ["terreno_codigo", "codigo_construccion"]:
                if want in df.columns:
                    out_cols.append(want)
                else:
                    for c in df.columns:
                        if want in c:
                            out_cols.append(c)
                            break
            df = df[out_cols]

            rename_map = {}
            for c in df.columns:
                if "terreno" in c and "codigo" in c:
                    rename_map[c] = "terreno_codigo"
                if "codigo_construccion" in c:
                    rename_map[c] = "codigo_construccion"
            df = df.rename(columns=rename_map)

            if df.empty:
                print("Advertencia: el DataFrame está vacio después del filtro.")

            return df
        except Exception as e:
            tkMessageBox.showerror("Error", "No se pudo procesar: {}".format(str(e)))
            print("Error:", e)
            return None
    
    def validacion_informalidad_sobre_predio(self, gdb_path):
        feature_class_name_terreno = "r_lc_terreno" if self.tipo_area.get() == "Rural" else "u_lc_terreno"
        fc_terr = os.path.join(gdb_path, feature_class_name_terreno)

        intersect_output = os.path.join(gdb_path, "intersect_informal_formal")

        arcpy.env.workspace = gdb_path

        if not arcpy.Exists(fc_terr):
            tkMessageBox.showerror("Error", u"La capa {} no existe en la GDB.".format(feature_class_name_terreno))
            return None
        print("validacion_informalidad_sobre_predio")

        formal_filtrado_fc = os.path.join(gdb_path, "formal_filtrado")
        informal_filtrado_unidad = os.path.join(gdb_path, "informal_filtrado_unidad")

        for fc in [formal_filtrado_fc, informal_filtrado_unidad, intersect_output]:
            if arcpy.Exists(fc):
                arcpy.Delete_management(fc)

        # obtener campo real terreno_codigo
        terr_low, fmap_terr = get_fields_lower(fc_terr)
        terreno_codigo_real = fmap(fmap_terr, "terreno_codigo")

        try:
            query_formal = "SUBSTRING({0}, 22, 1) <> '2'".format(terreno_codigo_real)
            arcpy.MakeFeatureLayer_management(fc_terr, "formal_layer_predio", query_formal)
            arcpy.CopyFeatures_management("formal_layer_predio", formal_filtrado_fc)

            query_informal = "SUBSTRING({0}, 22, 1) = '2'".format(terreno_codigo_real)
            arcpy.MakeFeatureLayer_management(fc_terr, "informal_layer_predio", query_informal)
            arcpy.CopyFeatures_management("informal_layer_predio", informal_filtrado_unidad)

            arcpy.Intersect_analysis(
                in_features=[formal_filtrado_fc, informal_filtrado_unidad],
                out_feature_class=intersect_output,
                join_attributes="ALL",
                output_type="INPUT"
            )

            # a DataFrame usando nombres minúscula
            inter_low, fmap_inter = get_fields_lower(intersect_output)
            usable_real = []
            usable_low = []
            for low_name, real_name in fmap_inter.items():
                fdesc = [f for f in arcpy.ListFields(intersect_output) if f.name == real_name][0]
                if fdesc.type not in ('Geometry', 'OID'):
                    usable_real.append(real_name)
                    usable_low.append(low_name)

            data = [
                [ ('' if val is None else str(val)) for val in row ]
                for row in arcpy.da.SearchCursor(intersect_output, usable_real)
            ]
            df = pd.DataFrame(data, columns=usable_low)
            df.replace(
                to_replace=["nan", "NaN", float('nan')],
                value='',
                inplace=True
            )

            errores = []

            # mapeos flexibles por minúsculas
            def pick(rowdict, candidatos):
                for c in candidatos:
                    if c in rowdict:
                        return rowdict[c]
                return ""

            for idx, row in df.iterrows():
                # buscar campos por posibles nombres
                nro_ficha_formal = pick(row, ["nroficha", "nroficha_formal", "nroficha_0"])
                nro_ficha_inf = pick(row, ["nroficha_1", "nroficha_informal", "nroficha1"])
                terreno = pick(row, ["terreno_codigo", "terreno_codigo_0", "terreno_codigo_formal"])
                terreno1 = pick(row, ["terreno_codigo_1", "terreno_codigo_informal"])
                modo = pick(row, ["modoreadquisicion", "moadadquisicion", "moadadquisicion_0", "moadadquisicion","modoadquisicion"])
                modo1 = pick(row, ["moadadquisicion_1", "moadadquisicion1", "moadadquisicion_informal","moadadquisicion_inform","modoadquisicion_1"])
                tipo = pick(row, ["prediolctipo", "prediolctipo_0"])
                tipo1 = pick(row, ["prediolctipo_1","prediolctipo1"])
                matricula = pick(row, ["matriculainmobiliaria","matriculainmobiliaria_0"])
                razon = pick(row, ["razonsocial","razonsocial_1","razonsocial_0"])

                # regla negocio
                if matricula :
                    # caso con matrícula
                    if tipo=='Predio.Publico.Uso_Publico' and modo1 != "5|OCUPACIN":
                        errores.append({
                            "Observacion": u"ModoAdquisicion incorrecto para predio informal sobre predio con matricula",
                            "NroFicha_Formal": nro_ficha_formal,
                            "NroFicha_Informal": nro_ficha_inf,
                            "TERRENO_CODIGO_FORMAL": terreno,
                            "TERRENO_CODIGO_INFORMAL": terreno1,
                            "MatriculaInmobiliaria": matricula,
                            "PredioLcTipo_FORMAL": tipo,
                            "PredioLcTipo_INFORMALIDAD": tipo1,
                            "ModoAdquisicion_FORMAL": modo,
                            "ModoAdquisicion_INFORMALIDAD": modo1,
                            "RazonSocial": razon
                        })
                    if tipo=='Predio.Publico.Uso_Publico' and tipo1 != "Predio.Privado.Privado":
                        errores.append({
                            "Observacion": u"PredioLcTipo incorrecto para predio informal sobre predio con matricula",
                            "NroFicha_Formal": nro_ficha_formal,
                            "NroFicha_Informal": nro_ficha_inf,
                            "TERRENO_CODIGO_FORMAL": terreno,
                            "TERRENO_CODIGO_INFORMAL": terreno1,
                            "MatriculaInmobiliaria": matricula,
                            "PredioLcTipo_FORMAL": tipo,
                            "PredioLcTipo_INFORMALIDAD": tipo1,
                            "ModoAdquisicion_FORMAL": modo,
                            "ModoAdquisicion_INFORMALIDAD": modo1,
                            "RazonSocial": razon
                        })
                else:
                    # sin matrícula
                    if tipo=='Predio.Publico.Uso_Publico' and modo1 != "5|OCUPACIN":
                        errores.append({
                            "Observacion": u"Modo de adquisicion incorrecto para predio formal sin matricula",
                            "NroFicha_Formal": nro_ficha_formal,
                            "NroFicha_Informal": nro_ficha_inf,
                            "TERRENO_CODIGO_FORMAL": terreno,
                            "TERRENO_CODIGO_INFORMAL": terreno1,
                            "MatriculaInmobiliaria": matricula,
                            "PredioLcTipo_FORMAL": tipo,
                            "PredioLcTipo_INFORMALIDAD": tipo1,
                            "ModoAdquisicion_FORMAL": modo,
                            "ModoAdquisicion_INFORMALIDAD": modo1,
                            "RazonSocial": razon
                        })
                    if tipo=='Predio.Publico.Uso_Publico' and tipo1!='Predio.Publico.Presunto_Baldio':
                        errores.append({
                            "Observacion": u"PredioLcTipo incorrecto en predio informal sobre predio sin matricula",
                            "NroFicha_Formal": nro_ficha_formal,
                            "NroFicha_Informal": nro_ficha_inf,
                            "TERRENO_CODIGO_FORMAL": terreno,
                            "TERRENO_CODIGO_INFORMAL": terreno1,
                            "MatriculaInmobiliaria": matricula,
                            "PredioLcTipo_FORMAL": tipo,
                            "PredioLcTipo_INFORMALIDAD": tipo1,
                            "ModoAdquisicion_FORMAL": modo,
                            "ModoAdquisicion_INFORMALIDAD": modo1,
                            "RazonSocial": razon
                        })

            if errores:
                df_err = pd.DataFrame(errores)
                df_err.fillna('', inplace=True)
                df_err.replace({
                    "ModoAdquisicion_INFORMALIDAD": {
                        "5|OCUPACIN": "5|OCUPACION",
                        "2|POSESIN": "2|POSESION"
                    },
                    "ModoAdquisicion_FORMAL": {
                        "5|OCUPACIN": "5|OCUPACION",
                        "2|POSESIN": "2|POSESION"
                    }
                }, inplace=True)
                return df_err
            else:
                tkMessageBox.showinfo("Validacion completada", u"No se encontraron errores.")
                return pd.DataFrame()

        except Exception as e:
            tkMessageBox.showerror("Error", u"No se pudo procesar: {}".format(str(e)))
            print("Error:", e)
            return None

    def validar(self, gdb_path):
        # Esta función sigue siendo avanzada; no la toco más allá de normalizar campos si se usa.
        # (Puedes aplicar la misma técnica de fmap si planeas usarla.)
        feature_class_name_terreno = "r_lc_terreno" if self.tipo_area.get() == "Rural" else "u_lc_terreno"
        feature_class_name_unidad = "r_lc_unidadconstruccion" if self.tipo_area.get() == "Rural" else "u_lc_unidadconstruccion"

        fc_terr = os.path.join(gdb_path, feature_class_name_terreno)
        fc_unid = os.path.join(gdb_path, feature_class_name_unidad)

        arcpy.env.workspace = gdb_path

        if not arcpy.Exists(fc_terr):
            tkMessageBox.showerror("Error", u"La capa {} no existe en la GDB.".format(feature_class_name_terreno))
            return None

        try:
            # campo codigo_unidad_construccion
            fields_lower, fmap_uni = get_fields_lower(fc_unid)
            cod_uc_real = fmap(fmap_uni, "codigo_unidad_construccion")
            if not cod_uc_real:
                tkMessageBox.showerror("Error", u"No existe 'codigo_unidad_construccion' en {}".format(feature_class_name_unidad))
                return None

            # agregar npn_22 si hace falta
            npn22_real = fmap(fmap_uni, "npn_22")
            if not npn22_real:
                arcpy.AddField_management(fc_unid, "npn_22", "TEXT", "", "", 22)
                npn22_real = "npn_22"

            arcpy.CalculateField_management(
                fc_unid,
                npn22_real,
                "Left([{0}], 22)".format(cod_uc_real),
                "VB"
            )

            layer_name = "unidadconstruccion_lyr"
            arcpy.MakeFeatureLayer_management(fc_unid, layer_name)
            # PLANTA_UBICACION en minúscula
            planta_real = fmap(fmap_uni, "planta_ubicacion")
            if not planta_real:
                planta_real = "PLANTA_UBICACION"  # fallback

            arcpy.SelectLayerByAttribute_management(
                in_layer_or_view=layer_name,
                selection_type="NEW_SELECTION",
                where_clause="{0} BETWEEN 1 AND 80".format(planta_real)
            )

            conteo = int(arcpy.GetCount_management(layer_name).getOutput(0))

            dissolve_output = os.path.join(gdb_path, "unidadconstruccion_dissolve_npn22")
            if conteo > 0:
                if arcpy.Exists(dissolve_output):
                    arcpy.Delete_management(dissolve_output)

                # estadísticas: MAX_PLANTA_UBICACION
                arcpy.Dissolve_management(
                    in_features=layer_name,
                    out_feature_class=dissolve_output,
                    dissolve_field=npn22_real,
                    statistics_fields=[[planta_real, "MAX"]]
                )
            else:
                print("No hay registros con PLANTA_UBICACION entre 1 y 80.")
                return None

            # capa construccion
            feature_class_name_construccion = "r_lc_construccion" if self.tipo_area.get() == "Rural" else "u_lc_construccion"
            construccion_fc = os.path.join(gdb_path, feature_class_name_construccion)
            buffer_output = os.path.join(gdb_path, feature_class_name_construccion + "_buffer_menos_0_5")

            if not arcpy.Exists(construccion_fc):
                print("La capa construccion no existe en la GDB.")
                return None

            if arcpy.Exists(buffer_output):
                arcpy.Delete_management(buffer_output)

            arcpy.Buffer_analysis(
                in_features=construccion_fc,
                out_feature_class=buffer_output,
                buffer_distance_or_field="-0.5 Meters",
                line_side="FULL",
                line_end_type="ROUND",
                dissolve_option="NONE",
                dissolve_field=""
            )

            intersect_output = os.path.join(gdb_path, "intersect_dissolve_buffer")
            if arcpy.Exists(intersect_output):
                arcpy.Delete_management(intersect_output)

            arcpy.Intersect_analysis(
                in_features=[dissolve_output, buffer_output],
                out_feature_class=intersect_output,
                join_attributes="ALL",
                cluster_tolerance="",
                output_type="INPUT"
            )
            print("Intersect generado correctamente:", intersect_output)

            campos_intersect = [f.name for f in arcpy.ListFields(intersect_output)]
            if "Diferencia" not in campos_intersect:
                arcpy.AddField_management(intersect_output, "Diferencia", "SHORT")

            # NUMERO_PISOS y MAX_PLANTA_UBICACION en minúscula
            inter_low, fmap_inter = get_fields_lower(intersect_output)
            max_planta_real = None
            numero_pisos_real = None
            for low_n, real_n in fmap_inter.items():
                if "max_" in low_n and "planta" in low_n:
                    max_planta_real = real_n
                if "numero_pisos" == low_n or low_n.endswith("numero_pisos"):
                    numero_pisos_real = real_n

            if not max_planta_real:
                max_planta_real = "MAX_PLANTA_UBICACION"
            if not numero_pisos_real:
                numero_pisos_real = "NUMERO_PISOS"

            arcpy.CalculateField_management(
                in_table=intersect_output,
                field="Diferencia",
                expression="[{0}] - [{1}]".format(max_planta_real, numero_pisos_real),
                expression_type="VB"
            )
            print("Campo 'Diferencia' calculado correctamente.")

            fields = []
            if cod_uc_real:
                fields.append(cod_uc_real)
            else:
                fields.append("CODIGO_CONSTRUCCION")
            fields.append(npn22_real)
            fields.append(max_planta_real)
            fields.append(numero_pisos_real)
            fields.append("Diferencia")

            data = [list(row) for row in arcpy.da.SearchCursor(intersect_output, fields)]
            if not data:
                print("Advertencia: el DataFrame está vacío después de Intersect.")
                return None

            # normalizar nombres de salida en minúscula
            cols_norm = [
                "codigo_construccion",
                "npn_22",
                "max_planta_ubicacion",
                "numero_pisos",
                "diferencia"
            ][:len(fields)]

            df = pd.DataFrame(data, columns=cols_norm)
            return df

        except Exception as e:
            tkMessageBox.showerror("Error", u"No se pudo procesar: {}".format(str(e)))
            print("Error:", e)
            return None

    def extraer_letras_identificador(self, gdb_path):
        feature_class_name_unidad = "r_lc_unidadconstruccion" if self.tipo_area.get() == "Rural" else "u_lc_unidadconstruccion"
        fc_unid = os.path.join(gdb_path, feature_class_name_unidad)

        try:
            arcpy.env.workspace = gdb_path

            fields_lower, fmap_uni = get_fields_lower(fc_unid)
            ident_real = fmap(fmap_uni, "identificador")
            if not ident_real:
                tkMessageBox.showerror("Error", u"La capa '{}' no contiene el campo 'IDENTIFICADOR'.".format(feature_class_name_unidad))
                return

            letras_real = fmap(fmap_uni, "letras")
            if not letras_real:
                arcpy.AddField_management(fc_unid, "LETRAS", "TEXT", "", "", 10)
                letras_real = "LETRAS"

            expression = "extraer_letras(!{0}!)".format(ident_real)
            code_block = """def extraer_letras(ident):
    if ident is None:
        return ""
    if ident.startswith("P"):
        num = ""
        i = 1
        while i < len(ident) and ident[i].isdigit():
            num += ident[i]
            i += 1
        letras = ""
        while i < len(ident) and ident[i].isalpha():
            letras += ident[i]
            i += 1
        return letras
    return ""
"""

            arcpy.CalculateField_management(
                in_table=fc_unid,
                field=letras_real,
                expression=expression,
                expression_type="PYTHON_9.3",
                code_block=code_block
            )
            print("Campo 'LETRAS' calculado correctamente.")

        except Exception as e:
            tkMessageBox.showerror("Error", u"No se pudo calcular LETRAS: {}".format(str(e)))
            print("Error:", e)

    def validar_terreno_codigo_duplicado_ficha(self, gdb_path):
        feature_class_name = "r_lc_terreno" if self.tipo_area.get() == "Rural" else "u_lc_terreno"
        fc = os.path.join(gdb_path, feature_class_name)
        arcpy.env.workspace = gdb_path

        if not arcpy.Exists(fc):
            raise Exception("La capa {} no existe en la GDB.".format(feature_class_name))

        fields_lower, fmap_terr = get_fields_lower(fc)

        terreno_codigo_real = fmap(fmap_terr, "terreno_codigo")
        nroficha_real = fmap(fmap_terr, "nroficha")
        cp_real = fmap(fmap_terr, "cp")

        for req_name, real_name in [
            ("terreno_codigo", terreno_codigo_real),
            ("nroficha", nroficha_real),
            ("cp", cp_real)
        ]:
            if not real_name:
                raise Exception("Falta columna requerida '{}'".format(req_name))

        registros = []
        cursor_fields = [terreno_codigo_real, nroficha_real, cp_real]
        with arcpy.da.SearchCursor(fc, cursor_fields) as cur:
            for row in cur:
                codigo = row[0]
                ficha = row[1]
                cp_val = row[2]
                if not codigo or len(codigo) < 22:
                    continue
                condicion = codigo[21]
                prefijo = codigo[:21]
                registros.append({
                    "codigo": codigo,
                    "nroficha": ficha,
                    "condicion": condicion,
                    "prefijo_21": prefijo,
                    "cp": cp_val
                })

        df = pd.DataFrame(registros)

        df_inf = df[df["condicion"] == '2']
        df_otros = df[df["condicion"] != '2']

        emp = []
        for idx, inf in df_inf.iterrows():
            coinc = df_otros[df_otros["prefijo_21"] == inf["prefijo_21"]]
            for jdx, form in coinc.iterrows():
                emp.append({
                    "informalidad": inf["codigo"],
                    "NroFicha_informalidad": inf["nroficha"],
                    "terreno": form["codigo"],
                    "NroFicha_terreno": form["nroficha"],
                })

        df_rep = pd.DataFrame(emp, columns=[
            "informalidad", "NroFicha_informalidad",
            "terreno", "NroFicha_terreno"
        ])

        if df_rep.empty:
            print("No hay informalidades emparejadas.")
            return df_rep

        layer_formal = "lyr_formal"
        where_f = "{0} = '0'".format(cp_real)
        arcpy.MakeFeatureLayer_management(fc, layer_formal, where_f)

        layer_inf = "lyr_informal"
        cods = df_rep["informalidad"].unique().tolist()
        in_clause = ", ".join("'{0}'".format(c) for c in cods)
        where_i = "{0} IN ({1})".format(terreno_codigo_real, in_clause)
        arcpy.MakeFeatureLayer_management(fc, layer_inf, where_i)

        buf_fc = "in_memory/buffer_inf"
        arcpy.Buffer_analysis(layer_inf, buf_fc, "-0.5", dissolve_option="NONE")

        sp_fc = "in_memory/sj_inf_formal"
        arcpy.SpatialJoin_analysis(
            buf_fc, layer_formal, sp_fc,
            join_operation="JOIN_ONE_TO_ONE",
            join_type="KEEP_COMMON"
        )

        # intentar leer el join
        # campos esperados: terreno informal, terreno formal (sufijo _1), ficha formal (_1)
        campos_join = arcpy.ListFields(sp_fc)
        # guess names:
        terr_inf_join = terreno_codigo_real
        terr_for_join = terreno_codigo_real + "_1"
        fic_for_join = nroficha_real + "_1"

        mapping = {}
        with arcpy.da.SearchCursor(sp_fc, [terr_inf_join, terr_for_join, fic_for_join]) as cur2:
            for row2 in cur2:
                infc = row2[0]
                formc = row2[1]
                formf = row2[2]
                mapping[infc] = (formc, formf)

        df_rep["predio_formal"] = df_rep["informalidad"].map(lambda c: mapping.get(c, (None, None))[0])
        df_rep["NroFicha_formal"] = df_rep["informalidad"].map(lambda c: mapping.get(c, (None, None))[1])

        df_rep = df_rep.rename(columns={
            "terreno": "terreno_codigo_duplicado",
            "predio_formal": "ubicacion_espacial_informalidad"
        })

        if "prefijo_21" in df_rep.columns:
            df_rep = df_rep.drop(["prefijo_21"], axis=1)

        print("Informe de informalidad geográfica (columnas ajustadas):")
        print(df_rep)

        return df_rep

    def validar_etiqueta(self, gdb_path):
        feature_class_name = "r_lc_terreno" if self.tipo_area.get() == "Rural" else "u_lc_terreno"
        feature_class_path = os.path.join(gdb_path, feature_class_name)
        arcpy.env.workspace = gdb_path

        if not arcpy.Exists(feature_class_path):
            raise Exception("La capa {} no existe en la GDB.".format(feature_class_name))
        print("validar_etiqueta")

        fields_lower, fmap_terr = get_fields_lower(feature_class_path)

        terreno_codigo_real = fmap(fmap_terr, "terreno_codigo")
        etiqueta_real = fmap(fmap_terr, "etiqueta")

        if not terreno_codigo_real:
            raise Exception("Falta el campo requerido: TERRENO_CODIGO")
        if not etiqueta_real:
            raise Exception("Falta el campo requerido: ETIQUETA")

        errores = []

        with arcpy.da.SearchCursor(feature_class_path, [terreno_codigo_real, etiqueta_real]) as cursor:
            for row in cursor:
                terreno_codigo = row[0]
                etiqueta_val = row[1]
                etiqueta_txt = str(etiqueta_val).strip() if etiqueta_val else ""

                if not terreno_codigo or len(terreno_codigo) < 30:
                    continue

                condicion_predio = terreno_codigo[21]
                unidad_terreno = terreno_codigo[26:30]

                if condicion_predio == '8' and unidad_terreno not in ('0000'):
                    etiqueta_cruda = terreno_codigo[26:30]
                    if etiqueta_cruda.startswith('P'):
                        etiqueta_esperada = etiqueta_cruda
                    else:
                        etiqueta_esperada = etiqueta_cruda.lstrip('0')
                else:
                    etiqueta_esperada = terreno_codigo[17:21]
                    if not etiqueta_esperada.startswith('P'):
                        etiqueta_esperada = etiqueta_esperada.lstrip('0')

                if etiqueta_txt != etiqueta_esperada:
                    errores.append([
                        terreno_codigo,
                        etiqueta_txt,
                        etiqueta_esperada,
                        condicion_predio
                    ])

        columnas_ordenadas = [
            "terreno_codigo",
            "etiqueta",
            "etiqueta_esperada",
            "condicion_predio"
        ]
        return pd.DataFrame(errores, columns=columnas_ordenadas)

    def comparar_areas_por_unidad(self, separador='', solo_diferentes=True):
        import Tkinter as tk
        import tkMessageBox as tkMessageBox

        try:
            reload(sys); sys.setdefaultencoding('utf-8')
        except:
            pass

        SEP = unicode(separador if separador else u"-")

        def _s(x):
            try:
                return unicode("" if x is None else x).strip()
            except:
                try:
                    return unicode(str(x)).strip()
                except:
                    return u""

        gdb_path = self.gdb_path.get()
        excel_path = self.excel_path.get()

        if not excel_path:
            tkMessageBox.showerror("Error", u"Debe seleccionar el Excel."); return
        if not gdb_path:
            tkMessageBox.showerror("Error", u"Debe seleccionar la GDB o carpeta contenedora."); return

        if not gdb_path.lower().endswith(".gdb") and not gdb_path.lower().endswith(".sde"):
            gdbs = [f for f in os.listdir(gdb_path) if f.lower().endswith(".gdb")]
            if gdbs:
                gdb_path = os.path.join(gdb_path, gdbs[0])
            else:
                tkMessageBox.showerror("Error", u"No se encontró ninguna GDB."); return

        arcpy.env.workspace = gdb_path
        arcpy.env.overwriteOutput = True

        try:
            df_uc = extraer_tabla_de_gdb_area_construida(gdb_path, self.tipo_area.get())
        except Exception as e:
            tkMessageBox.showerror("Error", u"No se pudo leer la tabla de la GDB:\n{}".format(e)); return

        df_uc.columns = [c.strip().lower() for c in df_uc.columns]
        req_gdb = ['codigo_unidad_construccion', 'identificador', 'shape_area']
        faltan_gdb = [c for c in req_gdb if c not in df_uc.columns]
        if faltan_gdb:
            tkMessageBox.showerror("Error", u"En la tabla GDB faltan: {}".format(", ".join(faltan_gdb))); return

        df_uc['clave'] = df_uc.apply(
            lambda r: u"{}{}{}".format(
                _s(r['codigo_unidad_construccion']),
                SEP,
                _s(r['identificador'])
            ),
            axis=1
        )
        df_uc['area_gdb'] = pd.to_numeric(df_uc['shape_area'], errors='coerce')
        df_gdb = df_uc.groupby('clave', as_index=False)['area_gdb'].sum()

        try:
            xl = pd.ExcelFile(excel_path)
        except Exception as e:
            tkMessageBox.showerror("Error", u"No se pudo abrir el Excel:\n{}".format(e)); return

        def _find_sheet(xlfile, snippet):
            lst = [h for h in xlfile.sheet_names if snippet in h.strip().lower()]
            return lst[0] if lst else None

        hoja_cons   = _find_sheet(xl, 'construcciones')
        hoja_fichas = _find_sheet(xl, 'fichas')

        if not hoja_cons:
            tkMessageBox.showerror("Error", u"No se encontró la hoja 'Construcciones'."); return

        df_cons = pd.read_excel(excel_path, sheetname=hoja_cons)
        df_cons.columns = [c.strip().lower() for c in df_cons.columns]

        req_cons_base = ['numeroconstruccion', 'areaconstruida']
        faltan_cons_base = [c for c in req_cons_base if c not in df_cons.columns]
        if faltan_cons_base:
            tkMessageBox.showerror("Error", u"En 'Construcciones' faltan: {}".format(", ".join(faltan_cons_base))); return

        if 'npn' in df_cons.columns:
            df_src = df_cons.copy()
        else:
            if not hoja_fichas:
                tkMessageBox.showerror("Error", u"No se encontró 'npn' en 'Construcciones' y no hay hoja 'Fichas'."); return
            df_fich = pd.read_excel(excel_path, sheetname=hoja_fichas)
            df_fich.columns = [c.strip().lower() for c in df_fich.columns]
            if 'npn' not in df_fich.columns or 'nroficha' not in df_fich.columns or 'nroficha' not in df_cons.columns:
                tkMessageBox.showerror("Error", u"No se pudo recuperar 'npn': verifique 'NroFicha' y 'Npn' en 'Fichas'."); return
            df_src = pd.merge(df_cons, df_fich[['nroficha', 'npn']], on='nroficha', how='left')

        df_src['npn'] = df_src['npn'].map(_s)
        df_src['numeroconstruccion'] = df_src['numeroconstruccion'].map(_s)
        df_src['clave'] = df_src.apply(
            lambda r: u"{}{}{}".format(_s(r['npn']), SEP, _s(r['numeroconstruccion'])),
            axis=1
        )
        df_src['area_excel'] = pd.to_numeric(df_src['areaconstruida'], errors='coerce')

        df_out = pd.merge(
            df_src[['clave', 'area_excel']],
            df_gdb,
            on='clave',
            how='outer'
        )

        df_out['area_gdb'] = pd.to_numeric(df_out.get('area_gdb'), errors='coerce')
        df_out['area_excel'] = pd.to_numeric(df_out.get('area_excel'), errors='coerce')
        df_out['diferencia'] = (df_out['area_gdb'] - df_out['area_excel']).abs()

        for col in ['area_gdb', 'area_excel', 'diferencia']:
            if col in df_out.columns:
                df_out[col] = df_out[col].round(2)

        if solo_diferentes:
            mask_dif = df_out['diferencia'].notnull() & (df_out['diferencia'] != 0)
            mask_nocruce = df_out['diferencia'].isnull()
            df_out = df_out[mask_dif | mask_nocruce]

        df_out['npn_nro_construccion'] = df_out['clave']
        return df_out[['npn_nro_construccion', 'area_gdb', 'area_excel', 'diferencia']]
    
    def informalidades_suman_mas_que_formal(self, gdb_path):
        """
        - Informalidad: dígito 22 de TERRENO_CODIGO = '2'
        - Formal:       dígito 22 <> '2'

        Retorna solo diferencias > 0.01
        """

        feature_class_name = "r_lc_terreno" if self.tipo_area.get() == "Rural" else "u_lc_terreno"
        fc = os.path.join(gdb_path, feature_class_name)

        arcpy.env.workspace = gdb_path
        arcpy.env.overwriteOutput = True

        if not arcpy.Exists(fc):
            tkMessageBox.showerror("Error", u"La capa {} no existe en la GDB.".format(feature_class_name))
            return None

        print("informalidades_suman_mas_que_formal")

        # ------------------------------------------------
        # Helpers
        # ------------------------------------------------
        def _fmt_num(a):
            try:
                s = "{:,.2f}".format(float(a))
                return s.replace(",", "X").replace(".", ",").replace("X", ".")
            except:
                return u"0,00"

        def _clean(v):
            if v is None:
                return None
            try:
                return unicode(v).strip()
            except:
                return str(v).strip()

        def _count(fc):
            try:
                return int(arcpy.GetCount_management(fc).getOutput(0))
            except:
                return 0

        def _find_field(fc, name):
            """Encuentra un campo sin importar mayúsculas/minúsculas"""
            for f in arcpy.ListFields(fc):
                if f.name.lower() == name.lower():
                    return f.name
            return None

        # ------------------------------------------------
        # Detectar campo TERRENO_CODIGO dinámicamente
        # ------------------------------------------------
        campo_tc = _find_field(fc, "terreno_codigo")
        if not campo_tc:
            tkMessageBox.showerror("Error", u"No se encontró el campo TERRENO_CODIGO.")
            return None

        # ------------------------------------------------
        # Capas temporales
        # ------------------------------------------------
        inf_fc   = os.path.join(gdb_path, "tmp_inf_d22_2")
        for_fc   = os.path.join(gdb_path, "tmp_for_d22_no2")
        inter_fc = os.path.join(gdb_path, "tmp_inf_for_intersect")
        dis_fc   = os.path.join(gdb_path, "tmp_inf_for_intersect_dis")

        try:
            # 1) Informalidades
            arcpy.MakeFeatureLayer_management(
                fc, "lyr_inf",
                "SUBSTRING({0},22,1) = '2'".format(arcpy.AddFieldDelimiters(fc, campo_tc))
            )
            arcpy.CopyFeatures_management("lyr_inf", inf_fc)

            # 2) Formales
            arcpy.MakeFeatureLayer_management(
                fc, "lyr_for",
                "SUBSTRING({0},22,1) <> '2'".format(arcpy.AddFieldDelimiters(fc, campo_tc))
            )
            arcpy.CopyFeatures_management("lyr_for", for_fc)

            if _count(inf_fc) == 0:
                return pd.DataFrame(columns=[
                    "terreno_codigo",
                    "area_informalidad",
                    "area_interseccion",
                    "diferencia"
                ])

            # 3) Área total por informalidad
            area_inf = {}
            with arcpy.da.SearchCursor(inf_fc, [campo_tc, "SHAPE@AREA"]) as cur:
                for tc, a in cur:
                    tc = _clean(tc)
                    if tc:
                        area_inf[tc] = area_inf.get(tc, 0.0) + float(a)

            # 4) Intersect real
            arcpy.Intersect_analysis([inf_fc, for_fc], inter_fc, "ALL", "", "INPUT")

            # Si no hay intersección
            if _count(inter_fc) == 0:
                data = []
                for tc, a_inf in area_inf.items():
                    if a_inf > 0.01:
                        data.append([
                            tc,
                            _fmt_num(a_inf),
                            "0,00",
                            _fmt_num(a_inf)
                        ])
                return pd.DataFrame(data, columns=[
                    "terreno_codigo",
                    "area_informalidad",
                    "area_interseccion",
                    "diferencia"
                ])

            # 5) Detectar campo del informal en intersect
            campo_inf = _find_field(inter_fc, campo_tc)
            if not campo_inf:
                tkMessageBox.showerror("Error", u"No se pudo identificar el campo TERRENO_CODIGO en intersect.")
                return None

            # 6) Dissolve por informalidad
            arcpy.Dissolve_management(inter_fc, dis_fc, dissolve_field=[campo_inf])

            # 7) Área de intersección
            area_int = {}
            with arcpy.da.SearchCursor(dis_fc, [campo_inf, "SHAPE@AREA"]) as cur:
                for tc, a in cur:
                    tc = _clean(tc)
                    if tc:
                        area_int[tc] = float(a)

            # 8) Resultado final
            data = []
            for tc, a_inf in area_inf.items():
                a_int = area_int.get(tc, 0.0)
                dif = a_inf - a_int

                if dif <= 0.01:
                    continue

                data.append([
                    tc,
                    _fmt_num(a_inf),
                    _fmt_num(a_int),
                    _fmt_num(dif)
                ])

            return pd.DataFrame(data, columns=[
                "terreno_codigo",
                "area_informalidad",
                "area_interseccion",
                "diferencia"
            ])

        except Exception as e:
            tkMessageBox.showerror("Error", u"No se pudo procesar:\n{}".format(str(e)))
            print("Error:", e)
            return None
    def verificar_geometrias_vacias(self, ruta_fc):
        """
        Verifica si hay geometrías vacías (NULL) en una capa.
        """
        contador_vacias = 0
        with arcpy.da.SearchCursor(ruta_fc, ["OID@", "SHAPE@"]) as cursor:
            for oid, geom in cursor:
                if geom is None or (hasattr(geom, "isMultipart") and geom.isMultipart and not geom.partCount):
                    print("Geometría vacía encontrada en el OID:", oid)
                    contador_vacias += 1

        if contador_vacias == 0:
            print("No hay geometrías vacías.")
        else:
            print("Total de geometrías vacías:", contador_vacias)

    def reporte(self, workbook, reportes_dict, gdb_path):
        sheet_reporte = workbook.add_sheet('Reporte')

        style_porcentaje = xlwt.XFStyle()
        style_porcentaje.num_format_str = '0.0%'

        feature_class_name_terreno = "r_lc_terreno" if self.tipo_area.get() == "Rural" else "u_lc_terreno"
        feature_class_path = os.path.join(gdb_path, feature_class_name_terreno)

        try:
            total_fichas = int(arcpy.GetCount_management(feature_class_path)[0])
        except Exception as e:
            print("Error al contar registros de la capa:", e)
            total_fichas = 0

        sheet_reporte.write(0, 0, 'Descripcion')
        sheet_reporte.write(0, 1, 'Cantidad')
        sheet_reporte.write(0, 2, 'Porcentaje de Error')
        sheet_reporte.write(0, 3, 'Porcentaje de Aprobacion')
        sheet_reporte.write(0, 4, 'Concepto')

        row = 1
        total_cantidad = 0
        errores_criticos_keys = (
            u"Comisiones",
            u"Omisiones",
            u"Informalidades Sin Predio Formal",
            u"Terreno Duplicados"
        )
        errores_criticos_presentes = []

        for descripcion, df in reportes_dict.items():
            cantidad = len(df)
            total_cantidad += cantidad

            porcentaje = (float(cantidad) / total_fichas) * 100 if total_fichas else 0
            porcentaje_aprobacion = 100 - porcentaje

            if descripcion in errores_criticos_keys and cantidad > 0:
                errores_criticos_presentes.append(descripcion)

            if porcentaje_aprobacion <= 50:
                concepto = 'NO CUMPLE'
            elif porcentaje_aprobacion <= 87.5:
                concepto = 'CUMPLE PARCIAL'
            else:
                concepto = 'CUMPLE'

            sheet_reporte.write(row, 0, descripcion)
            sheet_reporte.write(row, 1, cantidad)
            sheet_reporte.write(row, 2, porcentaje / 100.0, style_porcentaje)
            sheet_reporte.write(row, 3, porcentaje_aprobacion / 100.0, style_porcentaje)
            sheet_reporte.write(row, 4, concepto)
            row += 1

        porcentaje_total = (float(total_cantidad) / total_fichas) * 100 if total_fichas else 0
        porcentaje_aprob_total = 100 - porcentaje_total

        if porcentaje_aprob_total <= 50:
            evaluacion = 'NO CUMPLE'
        elif porcentaje_aprob_total <= 87.5:
            evaluacion = 'CUMPLE PARCIAL'
        else:
            evaluacion = 'CUMPLE'

        sheet_reporte.write(row, 0, 'TOTAL - ' + evaluacion)

        inicio = 1
        fin = row - 1

        sheet_reporte.write(row, 1, xlwt.Formula("SUM(B{}:B{})".format(inicio + 1, fin + 1)))

        sheet_reporte.write(0, 25, total_fichas)  # Z1
        celda_total_fichas = "$Z$1"

        sheet_reporte.write(row, 2, xlwt.Formula("B{}/{}".format(row + 1, celda_total_fichas)), style_porcentaje)
        sheet_reporte.write(row, 3, xlwt.Formula("1 - C{}".format(row + 1)), style_porcentaje)

        sheet_reporte.write(row, 4, evaluacion)
        row += 1

        if errores_criticos_presentes:
            errores_texto = ', '.join(errores_criticos_presentes)
            sheet_reporte.write(row, 0, u'ERRORES CRÍTICOS: ' + errores_texto)

    def select_gdb(self):
        path = tkFileDialog.askdirectory(title="Seleccionar Geodatabase (GDB)")
        if path:
            self.gdb_path.set(path)

    def select_excel(self):
        path = tkFileDialog.askopenfilename(
            title="Seleccionar Archivo Excel", 
            filetypes=[("Excel files", ".xlsx;.xls")]
        )
        if path:
            self.excel_path.set(path)

    def select_excel_bcgs(self):
        path = tkFileDialog.askopenfilename(
            title="Seleccionar Archivo Excel BCGS", 
            filetypes=[("Excel files", ".xlsx;.xls")]
        )
        if path:
            self.excel_path_bcgs.set(path)

    def select_output_excel(self):
        path = tkFileDialog.asksaveasfilename(
            title="Guardar archivo Excel",
            defaultextension=".xls",
            filetypes=[("Excel files", "*.xls;*.xls")]
        )
        if path:
            self.output_excel.set(path)
        else:
            tkMessageBox.showerror("Error", "Debe seleccionar una ruta de salida válida.")

if __name__ == "__main__":
    root = tk.Tk()
    app = GDBExcelValidator(root)
    root.mainloop()
