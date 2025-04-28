# -- coding: utf-8 --
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

sys.path.append(r'C:\Program Files (x86)\ArcGIS\Desktop10.8\arcpy')
os.environ['PATH'] = r"C:\Program Files\ArcGIS\Bin;" + os.environ['PATH']

def dividir_npn_en_columnas(npn):
    """Divide el campo NPN en las columnas especificadas."""
    return [
        npn[:2],     # Departamento 
        npn[2:5],    # Municipio 
        npn[5:7],    # Zona 
        npn[7:9],    # Sector 
        npn[9:11],    # Comuna 
        npn[11:13],   # Barrio 
        npn[13:17],  # Manzana o Vereda
        npn[17:21],  # Terreno o Predios
        npn[21:22],  # Condición Predio 
        npn[22:24],  # Edificio 
        npn[24:26],  # Número Piso 
        npn[26:30],  # Unidad Predial
    ]


def extraer_tabla_de_gdb_area_construida(gdb_path, tipo_area):
        """
        Extrae la tabla de la GDB y la convierte a un DataFrame.
        
        Parámetros:
            gdb_path (str): Ruta a la geodatabase.
            tipo_area (str): Puede ser "Rural" o cualquier otro valor para determinar la tabla a usar.
        
        Retorna:
            DataFrame de pandas con los datos de la tabla.
        """
        # Determinar la tabla según el tipo de área
        if tipo_area == "Rural":
            table_name = "r_lc_unidadconstruccion"
        else:
            table_name = "u_lc_unidadconstruccion"
        
        # Construir la ruta completa de la tabla
        table_path = os.path.join(gdb_path, table_name)
        
        # Comprobar que la tabla existe en la GDB
        if not arcpy.Exists(table_path):
            raise Exception("La tabla {} no existe en la geodatabase seleccionada.".format(table_name))
        
        # Obtener la lista de campos (excluyendo los de tipo geometría u OID si los hay)
        fields = [f.name for f in arcpy.ListFields(table_path) if f.type not in ["Geometry", "OID"]]
        
        # Extraer los datos utilizando un SearchCursor
        data = []
        with arcpy.da.SearchCursor(table_path, fields) as cursor:
            for row in cursor:
                # Se asocia cada campo a su valor en la fila
                data.append(dict(zip(fields, row)))
        
        # Convertir la lista de diccionarios a un DataFrame
        df = pd.DataFrame(data)
        return df
    
class GDBExcelValidator(Frame):
    def __init__(self, parent):
        Frame.__init__(self, parent)
        self.parent = parent
        self.parent.winfo_toplevel().title("Validación GDB vs Excel")

        # Variables para las rutas de archivos
        self.gdb_path = tk.StringVar()
        self.excel_path = tk.StringVar()
        self.excel_path_bcgs=tk.StringVar()
        # Se define output_excel como StringVar para poder actualizar la ruta de salida
        self.output_excel = tk.StringVar()


        # Variable para seleccionar entre Rural y Urbano
        self.tipo_area = tk.StringVar(value="Rural")  # Valor por defecto: Rural

        # Botón para seleccionar la GDB
        tk.Label(root, text="Seleccionar GDB:").pack()
        tk.Entry(root, textvariable=self.gdb_path, width=50).pack()
        tk.Button(root, text="Buscar GDB", command=self.select_gdb).pack()

        # Botón para seleccionar el archivo Excel de entrada
        tk.Label(root, text="Seleccionar Excel:").pack()
        tk.Entry(root, textvariable=self.excel_path, width=50).pack()
        tk.Button(root, text="Buscar Excel", command=self.select_excel).pack()
        
        tk.Label(root, text="Seleccionar Excel BCGS:").pack()
        tk.Entry(root, textvariable=self.excel_path_bcgs, width=50).pack()
        tk.Button(root, text="Buscar Excel BCGS", command=self.select_excel_bcgs).pack()

        # Radiobuttons para elegir entre Rural y Urbano
        tk.Label(root, text="Seleccione el tipo de área:").pack()
        tk.Radiobutton(root, text="Rural", variable=self.tipo_area, value="Rural").pack()
        tk.Radiobutton(root, text="Urbano", variable=self.tipo_area, value="Urbano").pack()

        # Botón para ejecutar la validación
        tk.Button(root, text="Ejecutar Validación", command=self.run_validation).pack()



    def run_validation(self):
        gdb_path = self.gdb_path.get()
        excel_path = self.excel_path.get()
        excel_path_bcgs = self.excel_path_bcgs.get()
        
        feature_class_name = "r_lc_terreno" if self.tipo_area.get() == "Rural" else "u_lc_terreno"
        self.select_output_excel()
        output_path = self.output_excel.get()
        if not output_path:
            tkMessageBox.showerror("Error", "Debe seleccionar una ruta de salida válida antes de continuar.")
            return
        if not gdb_path or not excel_path:
            tkMessageBox.showerror("Error", "Debe seleccionar la GDB y el archivo Excel.")
            return

        feature_class_path = os.path.join(gdb_path, feature_class_name)

        try:
            wb = xlrd.open_workbook(excel_path)
            sheet_comparacion = wb.sheet_by_index(0)
        
            npn_col_idx = None
            for i in range(sheet_comparacion.ncols):
                if sheet_comparacion.cell_value(0, i) == 'Npn':
                    npn_col_idx = i
                    break

            if npn_col_idx is None:
                tkMessageBox.showerror("Error", "La columna 'Npn' no se encuentra en el Excel.")
                return

            npn_excel = [sheet_comparacion.cell_value(row, npn_col_idx).encode('utf-8') for row in range(1, sheet_comparacion.nrows)]
            
            arcpy.env.workspace = gdb_path
            fields = [field.name for field in arcpy.ListFields(feature_class_path)]

            if 'TERRENO_CODIGO' not in fields:
                tkMessageBox.showerror("Error", "La columna 'TERRENO_CODIGO' no existe en la Feature Class.")
                return

            terreno_codigo_gdb = [row[0] for row in arcpy.da.SearchCursor(feature_class_path, ['TERRENO_CODIGO'])]

            diff_terreno_codigo = set(terreno_codigo_gdb) - set(npn_excel)
            diff_npn_excel = set(npn_excel) - set(terreno_codigo_gdb)

            def filtrar_omisiones(npn_list):
                return [npn for npn in npn_list if len(npn) >= 30 and (npn[21] != '9' or npn[26:30] == '0000')]

            omisiones_filtradas = filtrar_omisiones(diff_npn_excel)

            # Crear archivo Excel
            workbook = xlwt.Workbook()
            sheet_comisiones = workbook.add_sheet('Comisiones')
            sheet_omisiones = workbook.add_sheet('Omisiones')
            sheet_diferencias = workbook.add_sheet('Diferencia Areas Construidas')
            sheet_duplicados = workbook.add_sheet('Npn Duplicados')
            sheet_ph_sin_unidad = workbook.add_sheet('PH sin Unidad Predial')
            sheet_terreno_nro_piso = workbook.add_sheet('Terreno con Nro Piso')
            sheet_informalidades_sin_predio_formal = workbook.add_sheet('Informalidades Sin P')
            sheet_npn__unidad_diferente_de_terreno = workbook.add_sheet('Npn Unidad Dif De Terreno')
            sheet_npn__construccion_diferente_de_terreno = workbook.add_sheet('Npn Construccion Dif De Terreno')
            
            headers = ["Npn", "Departamento", "Municipio", "Zona", "Sector", "Comuna", "Barrio", "Manzana o Vereda",
                    "Terreno o Predios", "Condición Predio", "Edificio", "Número Piso", "Unidad Predial"]
            
            bold_style = xlwt.XFStyle()
            bold_font = xlwt.Font()
            bold_font.bold = True
            bold_style.font = bold_font
            sheet_comisiones.panes_frozen = True  
            sheet_comisiones.horz_split_pos = 1 
            sheet_omisiones.panes_frozen = True  
            sheet_omisiones.horz_split_pos = 1

            column_widths = [len(header) for header in headers]

            for col_num, header in enumerate(headers):
                sheet_comisiones.write(0, col_num, header.decode('utf-8'), bold_style)
                sheet_omisiones.write(0, col_num, header.decode('utf-8'), bold_style)

            for row_num, npn in enumerate(diff_terreno_codigo, 1):
                columnas = dividir_npn_en_columnas(npn)
                valores_fila = [npn] + columnas  

                for col_num, valor in enumerate(valores_fila):
                    valor = valor.decode('utf-8')
                    sheet_comisiones.write(row_num, col_num, valor)
                    column_widths[col_num] = max(column_widths[col_num], len(valor))

            for row_num, npn in enumerate(omisiones_filtradas, 1):
                columnas = dividir_npn_en_columnas(npn)
                valores_fila = [npn] + columnas  

                for col_num, valor in enumerate(valores_fila):
                    valor = valor.decode('utf-8')
                    sheet_omisiones.write(row_num, col_num, valor)
                    column_widths[col_num] = max(column_widths[col_num], len(valor))

            for col_num, width in enumerate(column_widths):
                sheet_comisiones.col(col_num).width = (width + 2) * 256  
                sheet_omisiones.col(col_num).width = (width + 2) * 256  

            df_diferencias_areas_construidas=self.calcular_areas_construidas()
            npn_duplicados = self.validar_terreno_codigo_duplicado(gdb_path)
            ph_sin_unidad = self.validar_ph_sin_unidad_predial(gdb_path)
            terreno_con_nro_piso= self.validar_terreno_con_piso(gdb_path)
            
            df_terreno_con_nro_piso = pd.DataFrame(terreno_con_nro_piso, columns=["TERRENO_CODIGO_Con_Nro_Piso"])
            df_npns_duplicados = pd.DataFrame(npn_duplicados, columns=["TERRENO_CODIGO_Duplicado"])
            df_informalidades_sin_predio_formal=self.copiar_filtrar_buffer_y_join(gdb_path)
            df_npn__unidad_diferente_de_terreno=self.validar_npn__unidad_diferente_de_terreno(gdb_path)
            df_npn__construccion_diferente_de_terreno=self.validar_npn__construccion_diferente_de_terreno(gdb_path)
            
            df_ph_sin_unidad = pd.DataFrame(ph_sin_unidad, columns=["PH sin unidad predial"])
            
            for col_num, column in enumerate(df_diferencias_areas_construidas.columns):
                    sheet_diferencias.write(0, col_num, column.decode('utf-8'), bold_style)

            # Escribir los datos de diferencias de áreas construidas
            for row_num, row in enumerate(df_diferencias_areas_construidas.itertuples(index=False), 1):
                for col_num, value in enumerate(row):
                    sheet_diferencias.write(row_num, col_num, str(value).decode('utf-8'))

            for col_num, column in enumerate(df_terreno_con_nro_piso.columns):
                    sheet_terreno_nro_piso.write(0, col_num, column.decode('utf-8'), bold_style)

            # Escribir los datos de diferencias de áreas construidas
            for row_num, row in enumerate(df_terreno_con_nro_piso.itertuples(index=False), 1):
                for col_num, value in enumerate(row):
                    sheet_terreno_nro_piso.write(row_num, col_num, str(value).decode('utf-8'))

            # Escribir los encabezados de la lista de duplicados
            for col_num, column in enumerate(df_npns_duplicados.columns):
                sheet_duplicados.write(0, len(df_npns_duplicados.columns) + col_num, column.decode('utf-8'), bold_style)

            # Escribir los datos de códigos de terreno duplicados
            for row_num, row in enumerate(df_npns_duplicados.itertuples(index=False), 1):
                for col_num, value in enumerate(row):
                    sheet_duplicados.write(row_num, len(df_npns_duplicados.columns) + col_num, str(value).decode('utf-8'))


            for col_num, column in enumerate(df_ph_sin_unidad.columns):
                sheet_ph_sin_unidad.write(0, len(df_ph_sin_unidad.columns) + col_num, column.decode('utf-8'), bold_style)

            # Escribir los datos de códigos de terreno duplicados
            for row_num, row in enumerate(df_ph_sin_unidad.itertuples(index=False), 1):
                for col_num, value in enumerate(row):
                    sheet_ph_sin_unidad.write(row_num, len(df_ph_sin_unidad.columns) + col_num, str(value).decode('utf-8'))

            for col_num, column in enumerate(df_informalidades_sin_predio_formal.columns): 
                sheet_informalidades_sin_predio_formal.write(0, len(df_informalidades_sin_predio_formal.columns) + col_num, column.decode('utf-8'), bold_style)

            # Escribir los datos de códigos de terreno duplicados
            for row_num, row in enumerate(df_informalidades_sin_predio_formal.itertuples(index=False), 1):
                for col_num, value in enumerate(row):
                    sheet_informalidades_sin_predio_formal.write(row_num, len(df_informalidades_sin_predio_formal.columns) + col_num, str(value).decode('utf-8'))
            

            for col_num, column in enumerate(df_npn__unidad_diferente_de_terreno.columns): 
                sheet_npn__unidad_diferente_de_terreno.write(0, len(df_npn__unidad_diferente_de_terreno.columns) + col_num, column.decode('utf-8'), bold_style)

            # Escribir los datos de códigos de terreno duplicados
            for row_num, row in enumerate(df_npn__unidad_diferente_de_terreno.itertuples(index=False), 1):
                for col_num, value in enumerate(row):
                    sheet_npn__unidad_diferente_de_terreno.write(row_num, len(df_npn__unidad_diferente_de_terreno.columns) + col_num, str(value).decode('utf-8'))
            
            for col_num, column in enumerate(df_npn__construccion_diferente_de_terreno.columns): 
                sheet_npn__construccion_diferente_de_terreno.write(0, len(df_npn__construccion_diferente_de_terreno.columns) + col_num, column.decode('utf-8'), bold_style)

            # Escribir los datos de códigos de terreno duplicados
            for row_num, row in enumerate(df_npn__construccion_diferente_de_terreno.itertuples(index=False), 1):
                for col_num, value in enumerate(row):
                    sheet_npn__construccion_diferente_de_terreno.write(row_num, len(df_npn__construccion_diferente_de_terreno.columns) + col_num, str(value).decode('utf-8'))

            workbook.save(output_path)
            
            tkMessageBox.showinfo("Éxito".decode('utf-8'), u"Proceso finalizado correctamente.\nArchivos guardados en:\n" +
                                output_path.decode('utf-8'))

        except Exception as e:
            tkMessageBox.showerror("Error", str(e))

    
    def calcular_areas_construidas(self):
        # Validar rutas
        gdb_path = self.gdb_path.get()
        excel_path = self.excel_path.get()

        if not excel_path:
            tkMessageBox.showerror("Error", "Debe seleccionar el Excel para calcular áreas.")
            return

        if not gdb_path.endswith(".gdb"):
            gdbs = [f for f in os.listdir(gdb_path) if f.endswith(".gdb")]
            if gdbs:
                gdb_path = os.path.join(gdb_path, gdbs[0])
            else:
                tkMessageBox.showerror("Error", "No se encontró ninguna GDB en la carpeta seleccionada.")
                return

        # Configurar entorno ArcPy
        arcpy.env.workspace = gdb_path
        arcpy.env.overwriteOutput = True
        print("\nGDB seleccionada: {}".format(gdb_path))

        # Extraer datos de la GDB
        df_unidadconstruccion = extraer_tabla_de_gdb_area_construida(gdb_path, self.tipo_area.get())
        #print(df_unidadconstruccion)
        # Leer el archivo Excel
        xl = pd.ExcelFile(excel_path)

        # Buscar la hoja 'Construcciones' y 'Fichas'
        hoja_construcciones = next((h for h in xl.sheet_names if "construcciones" in h.strip().lower()), None)
        hoja_fichas = next((h for h in xl.sheet_names if "fichas" in h.strip().lower()), None)

        if not hoja_construcciones or not hoja_fichas:
            tkMessageBox.showerror("Error", "No se encontró la hoja 'Construcciones' o 'Fichas' en el archivo.")
            return

        try:
            df_construcciones = pd.read_excel(excel_path, sheetname=hoja_construcciones)
            df_fichas = pd.read_excel(excel_path, sheetname=hoja_fichas)

            df_fichas = pd.merge(df_fichas, df_construcciones[['NroFicha', 'AreaConstruida']], on='NroFicha', how='left')
            
            if 'Npn' not in df_fichas.columns or 'AreaConstruida' not in df_fichas.columns:
                tkMessageBox.showerror("Error", "No se encontraron las columnas 'Npn' o 'AreaConstruida' en la hoja Fichas.")
                return


        except Exception as e:
            tkMessageBox.showerror("Error", "No se pudo leer las hojas del Excel:\n{}".format(str(e)))
            return

        # Agrupar datos por los primeros 22 caracteres del identificador
        df_agrupado_gdb = df_unidadconstruccion.groupby(
            df_unidadconstruccion['CODIGO_UNIDAD_CONSTRUCCION'].astype(str).str[:22]
        )['SHAPE_Area'].sum().reset_index().rename(columns={'SHAPE_Area': 'Area_GDB'})

        df_agrupado_excel = df_fichas.groupby(
            df_fichas['Npn'].astype(str).str[:22]
        )['AreaConstruida'].sum().reset_index().rename(columns={'AreaConstruida': 'Area_Excel'})

        # Fusionar los DataFrames agrupados
        df_comparacion = pd.merge(df_agrupado_gdb, df_agrupado_excel, left_on='CODIGO_UNIDAD_CONSTRUCCION', right_on='Npn', how='inner')

        # Calcular diferencia
        df_comparacion['Diferencia'] = abs(df_comparacion['Area_GDB'] - df_comparacion['Area_Excel'])

        # Filtrar diferencias significativas (> 20)
        #df_diferencias = df_comparacion[df_comparacion['Diferencia']]
        return df_comparacion


    def validar_terreno_codigo_duplicado(self, gdb_path):
        feature_class_name = "r_lc_terreno" if self.tipo_area.get() == "Rural" else "u_lc_terreno"
        feature_class_path = os.path.join(gdb_path, feature_class_name)

        arcpy.env.workspace = gdb_path

        if not arcpy.Exists(feature_class_path):
            tkMessageBox.showerror("Error", "La capa {feature_class_name} no existe en la GDB.")
            return

        fields = [field.name for field in arcpy.ListFields(feature_class_path)]
        
        if "TERRENO_CODIGO" not in fields:
            tkMessageBox.showerror("Error", "La columna 'TERRENO_CODIGO' no existe en la Feature Class.")
            return

        # Leer los valores de TERRENO_CODIGO y contar duplicados
        terreno_codigos = [row[0] for row in arcpy.da.SearchCursor(feature_class_path, ["TERRENO_CODIGO"])]

        # Encontrar duplicados
        contador_codigos = {}
        for codigo in terreno_codigos:
            contador_codigos[codigo] = contador_codigos.get(codigo, 0) + 1

        duplicados = [codigo for codigo, count in contador_codigos.items() if count > 1]

        if duplicados:
            #tkMessageBox.showwarning("Advertencia", "Se encontraron {len(duplicados)} códigos de terreno duplicados.")
            return duplicados
        else:
            #tkMessageBox.showinfo("Validación Exitosa", "No se encontraron códigos de terreno duplicados.")
            return []
        

    def validar_ph_sin_unidad_predial(self, gdb_path):
        feature_class_name = "r_lc_unidadconstruccion" if self.tipo_area.get() == "Rural" else "u_lc_unidadconstruccion"
        feature_class_path = os.path.join(gdb_path, feature_class_name)

        arcpy.env.workspace = gdb_path

        if not arcpy.Exists(feature_class_path):
            tkMessageBox.showerror("Error", "La capa {feature_class_name} no existe en la GDB.")
            return

        fields = [field.name for field in arcpy.ListFields(feature_class_path)]
        
        if "CODIGO_UNIDAD_CONSTRUCCION" not in fields:
            tkMessageBox.showerror("Error", "La columna 'CODIGO_UNIDAD_CONSTRUCCION' no existe en la Feature Class.")
            return

        # Leer los valores de CODIGO_UNIDAD_CONSTRUCCION
        UnidadConstruccion = [row[0] for row in arcpy.da.SearchCursor(feature_class_path, ["CODIGO_UNIDAD_CONSTRUCCION"])]

        # Filtrar por los primeros 22 caracteres y verificar el dígito 22
        codigos_filtrados = [codigo[:22] for codigo in UnidadConstruccion if len(codigo) >= 22 and codigo[21] in ('9', '8')]

        # Contar ocurrencias
        contador_codigos = {}
        for codigo in codigos_filtrados:
            contador_codigos[codigo] = contador_codigos.get(codigo, 0) + 1

        # Obtener los códigos sin duplicados
        sin_duplicados = [codigo for codigo, count in contador_codigos.items() if count == 1]
        
        
        return sin_duplicados if sin_duplicados else ["NO HAY REGLAMENTOS"]

    def validar_terreno_con_piso(self, gdb_path):
        feature_class_name = "r_lc_terreno" if self.tipo_area.get() == "Rural" else "u_lc_terreno"
        feature_class_path = os.path.join(gdb_path, feature_class_name)

        arcpy.env.workspace = gdb_path

        if not arcpy.Exists(feature_class_path):
            tkMessageBox.showerror("Error", "La capa {feature_class_name} no existe en la GDB.")
            return []

        fields = [field.name for field in arcpy.ListFields(feature_class_path)]

        if "TERRENO_CODIGO" not in fields:
            tkMessageBox.showerror("Error", "La columna 'TERRENO_CODIGO' no existe en la Feature Class.")
            return []

        try:
            with arcpy.da.SearchCursor(feature_class_path, ["TERRENO_CODIGO"]) as cursor:
                terreno_codigos = [row[0] for row in cursor if row[0] and isinstance(row[0], str) and len(row[0]) >= 30]

            # Filtrar códigos que cumplen ambas condiciones
            codigos_validos = [
                codigo for codigo in terreno_codigos
                if codigo[21] == '0' and sum(int(d) for d in codigo[-4:]) != 0
            ]

            return codigos_validos

        except Exception as e:
            tkMessageBox.showerror("Error", "Error al procesar los datos: {str(e)}")
            return []

    def copiar_filtrar_buffer_y_join(self, gdb_path):
        feature_class_name = "r_lc_terreno" if self.tipo_area.get() == "Rural" else "u_lc_terreno"
        feature_class_path = os.path.join(gdb_path, feature_class_name)

        arcpy.env.workspace = gdb_path

        if not arcpy.Exists(feature_class_path):
            tkMessageBox.showerror("Error", "La capa {} no existe en la GDB.".format(feature_class_name))
            return None  

        # Definir rutas de salida
        formal_filtrado_fc = os.path.join(gdb_path, "formal_filtrado")
        informal_filtrado_fc = os.path.join(gdb_path, "informal_filtrado")
        informal_buffer_fc = os.path.join(gdb_path, "informal_buffer")
        join_output_fc = os.path.join(gdb_path, "informal_buffer_joined")

        try:
            # Copiar y filtrar 'formal' (donde el 22° dígito de TERRENO_CODIGO no es '2')
            query_formal = "SUBSTRING(TERRENO_CODIGO, 22, 1) <> '2'"
            arcpy.MakeFeatureLayer_management(feature_class_path, "formal_layer", query_formal)
            arcpy.CopyFeatures_management("formal_layer", formal_filtrado_fc)
            #print("Capa 'formal' filtrada y guardada.")

            # Copiar y filtrar 'informal' (donde el 22° dígito de TERRENO_CODIGO es '2')
            query_informal = "SUBSTRING(TERRENO_CODIGO, 22, 1) = '2'"
            arcpy.MakeFeatureLayer_management(feature_class_path, "informal_layer", query_informal)
            arcpy.CopyFeatures_management("informal_layer", informal_filtrado_fc)
            #print("Capa 'informal' filtrada y guardada.")

            # Aplicar buffer negativo de -0.5 metros
            buffer_distancia = "-0.5 Meters"
            arcpy.Buffer_analysis(informal_filtrado_fc, informal_buffer_fc, buffer_distancia, 
                                line_side="FULL", line_end_type="ROUND", dissolve_option="NONE", method="PLANAR")
            #print("Buffer de -0.5 metros aplicado con éxito a la capa 'informal'.")

            # Realizar Spatial Join entre informal_buffer y formal_filtrado
            arcpy.SpatialJoin_analysis(target_features=informal_buffer_fc,
                                    join_features=formal_filtrado_fc,
                                    out_feature_class=join_output_fc,
                                    join_operation="JOIN_ONE_TO_ONE",
                                    join_type="KEEP_ALL",
                                    match_option="INTERSECT")
            #print("Spatial Join completado con éxito.")

            # Filtrar la tabla resultante donde TERRENO_CODIGO_1 es NULL
            query_filtro = "TERRENO_CODIGO_1 IS NULL"
            arcpy.MakeFeatureLayer_management(join_output_fc, "filtered_layer", query_filtro)

            # Convertir la tabla filtrada en un DataFrame
            fields = [field.name for field in arcpy.ListFields(join_output_fc)]
            data = [list(row) for row in arcpy.da.SearchCursor("filtered_layer", fields)] if arcpy.Exists("filtered_layer") else []
            df = pd.DataFrame(data, columns=fields)
            df = df[['TERRENO_CODIGO']]
           
            if df.empty:
                print("Advertencia: el DataFrame está vacío después del filtro.")

            return df  

        except Exception as e:
            tkMessageBox.showerror("Error", "No se pudo procesar: {}".format(str(e)))
            print("Error:", e)
            return None
        

    def validar_npn__unidad_diferente_de_terreno(self, gdb_path):
        feature_class_name = "r_lc_terreno" if self.tipo_area.get() == "Rural" else "u_lc_terreno"
        feature_class_path = os.path.join(gdb_path, feature_class_name)
        feature_class_name_unidad = "r_lc_unidadconstruccion" if self.tipo_area.get() == "Rural" else "u_lc_unidadconstruccion"
        feature_class_path_unidad = os.path.join(gdb_path, feature_class_name_unidad)

        arcpy.env.workspace = gdb_path

        if not arcpy.Exists(feature_class_path):
            tkMessageBox.showerror("Error", "La capa {} no existe en la GDB.".format(feature_class_name))
            return None  
        
        if not arcpy.Exists(feature_class_path_unidad):
            tkMessageBox.showerror("Error", "La capa {} no existe en la GDB.".format(feature_class_name_unidad))
            return None  

        # Definir rutas de salida
        unidad_puntos_fc = os.path.join(gdb_path, "unidad_puntos")
        formal_filtrado_fc = os.path.join(gdb_path, "formal_filtrado")
        informal_filtrado_fc = os.path.join(gdb_path, "informal_filtrado")
        informal_dissolve_fc = os.path.join(gdb_path, "informal_dissolved1")
        erase_output_fc = os.path.join(gdb_path, "unico_clip")
        merge_output_fc = os.path.join(gdb_path, "unico")
        intersect_output_fc = os.path.join(gdb_path, "interseccion_unidad")
        intersect_output_fc_filtro= os.path.join(gdb_path, "interseccion_unidad_filtro")
        try:
            # Verificar y eliminar capas existentes
            for fc in [unidad_puntos_fc, formal_filtrado_fc, informal_filtrado_fc, informal_dissolve_fc, erase_output_fc, merge_output_fc, intersect_output_fc]:
                if arcpy.Exists(fc):
                    arcpy.Delete_management(fc)

            # **Convertir la capa de unidad de construcción en puntos**
            arcpy.FeatureToPoint_management(feature_class_path_unidad, unidad_puntos_fc, "INSIDE")
            print("Feature To Point completado: ", unidad_puntos_fc)

            # Ejecutar Merge (terreno y unidades de construcción en puntos)
            arcpy.CopyFeatures_management(feature_class_path, merge_output_fc)
            print("Copia de terreno completada: ", merge_output_fc)

            # Filtrar 'formal' (donde el 22° dígito de TERRENO_CODIGO no es '2')
            query_formal = "SUBSTRING(TERRENO_CODIGO, 22, 1) <> '2'"
            arcpy.MakeFeatureLayer_management(feature_class_path, "formal_layer", query_formal)
            arcpy.CopyFeatures_management("formal_layer", formal_filtrado_fc)

            # Filtrar 'informal' (donde el 22° dígito de TERRENO_CODIGO es '2')
            query_informal = "SUBSTRING(TERRENO_CODIGO, 22, 1) = '2'"
            arcpy.MakeFeatureLayer_management(feature_class_path, "informal_layer", query_informal)
            arcpy.CopyFeatures_management("informal_layer", informal_filtrado_fc)

            # Ejecutar Dissolve si el campo DIMENSION existe
            campos_dissolve = ["DIMENSION"] if "DIMENSION" in [f.name for f in arcpy.ListFields(informal_filtrado_fc)] else None
            arcpy.Dissolve_management(informal_filtrado_fc, informal_dissolve_fc, campos_dissolve)
            print("Dissolve completado: ", informal_dissolve_fc)

            # Ejecutar Erase (descartar áreas de intersección)
            if arcpy.Exists(informal_dissolve_fc) and int(arcpy.GetCount_management(informal_dissolve_fc)[0]) > 0:
                arcpy.Erase_analysis(merge_output_fc, informal_dissolve_fc, erase_output_fc)
                print("Erase completado: ", erase_output_fc)
            else:
                print("Advertencia: 'informal_dissolved1' no tiene datos o no existe, se omite el Erase.")
                erase_output_fc = merge_output_fc  # Si no hay intersecciones, usa el merge original

            # Agregar datos de 'informal_filtrado' al resultado final
            if arcpy.Exists(informal_filtrado_fc) and int(arcpy.GetCount_management(informal_filtrado_fc)[0]) > 0:
                arcpy.Append_management(informal_filtrado_fc, erase_output_fc, "NO_TEST")
                print("Datos de 'informal_filtrado' copiados a 'unico'.")
            else:
                print("Advertencia: 'informal_filtrado' no tiene datos, se omite el Append.")

            # **Ejecutar Intersección entre erase_output_fc y unidad_puntos_fc**
            if arcpy.Exists(feature_class_path) and arcpy.Exists(unidad_puntos_fc):
        # **Ejecutar Intersección entre erase_output_fc y unidad_puntos_fc**
                arcpy.Intersect_analysis([unidad_puntos_fc,erase_output_fc], intersect_output_fc, "ALL", "", "INPUT")
                print("Intersección completada:", intersect_output_fc)
            else:
                print("Error: Una de las capas no existe, intersección omitida.")
            arcpy.AddField_management(intersect_output_fc, "CP_U", "TEXT", field_length = 1)
            arcpy.AddField_management(intersect_output_fc, "EDIFICIO_UNIDAD", "TEXT", field_length = 8)
            arcpy.AddField_management(intersect_output_fc, "TERRENO_22", "TEXT", field_length = 22)
            arcpy.AddField_management(intersect_output_fc, "UNIDAD_22", "TEXT", field_length = 22)
            arcpy.AddField_management(intersect_output_fc, "DIFERENCIA", "LONG")


            arcpy.CalculateField_management(intersect_output_fc, "CP_U", "Mid([CODIGO_UNIDAD_CONSTRUCCION], 22, 1)", "VB")
            arcpy.CalculateField_management(intersect_output_fc, "EDIFICIO_UNIDAD", "Right([CODIGO_UNIDAD_CONSTRUCCION],8)", "VB")
            arcpy.CalculateField_management(intersect_output_fc, "TERRENO_22", "Left([TERRENO_CODIGO],22)", "VB")
            arcpy.CalculateField_management(intersect_output_fc, "UNIDAD_22", "Left([CODIGO_UNIDAD_CONSTRUCCION],22)", "VB")
            arcpy.CalculateField_management(intersect_output_fc, "DIFERENCIA", "[TERRENO_22] = [UNIDAD_22]", "VB")
            query_filtro = "DIFERENCIA = 0"
            arcpy.MakeFeatureLayer_management(intersect_output_fc, "temp_layer_name", query_filtro)
            arcpy.CopyFeatures_management("temp_layer_name", intersect_output_fc_filtro)

            query_filtro_eliminar = "CP_U = '2' AND EDIFICIO_UNIDAD <> '00000000'"
            arcpy.MakeFeatureLayer_management(intersect_output_fc_filtro, "temp_layer_name_eliminar", query_filtro_eliminar)

            # Eliminar los registros de la capa filtrada intersect_output_fc_filtro
            arcpy.DeleteRows_management("temp_layer_name_eliminar")

            # Eliminar la capa temporal después de su uso
            arcpy.Delete_management("temp_layer_name_eliminar")

            fields = [field.name for field in arcpy.ListFields(intersect_output_fc_filtro)]
            data = [list(row) for row in arcpy.da.SearchCursor(intersect_output_fc_filtro, fields)] if arcpy.Exists(intersect_output_fc_filtro) else []
            df = pd.DataFrame(data, columns=fields)
            df = df[['TERRENO_CODIGO','CODIGO_UNIDAD_CONSTRUCCION']]
            print(df)
            if df.empty:
                print("Advertencia: el DataFrame está vacío después del filtro.")

            return df  

        except Exception as e:
            tkMessageBox.showerror("Error", "No se pudo procesar: {}".format(str(e)))
            print("Error:", e)
            return None
    

    def validar_npn__construccion_diferente_de_terreno(self, gdb_path):
        feature_class_name = "r_lc_terreno" if self.tipo_area.get() == "Rural" else "u_lc_terreno"
        feature_class_path = os.path.join(gdb_path, feature_class_name)
        feature_class_name_construccion = "r_lc_construccion" if self.tipo_area.get() == "Rural" else "u_lc_construccion"
        feature_class_path_construccion = os.path.join(gdb_path, feature_class_name_construccion)

        arcpy.env.workspace = gdb_path

        if not arcpy.Exists(feature_class_path):
            tkMessageBox.showerror("Error", "La capa {} no existe en la GDB.".format(feature_class_name))
            return None  
        
        if not arcpy.Exists(feature_class_path_construccion):
            tkMessageBox.showerror("Error", "La capa {} no existe en la GDB.".format(feature_class_name_construccion))
            return None  

        # Definir rutas de salida
        unidad_puntos_fc = os.path.join(gdb_path, "puntos_construccion")
        formal_filtrado_fc = os.path.join(gdb_path, "formal_filtrado_construccion")
        informal_filtrado_fc = os.path.join(gdb_path, "informal_filtrado_construccion")
        informal_dissolve_fc = os.path.join(gdb_path, "informal_dissolved1_construccion")
        erase_output_fc = os.path.join(gdb_path, "unico_clip__construccion")
        merge_output_fc = os.path.join(gdb_path, "unico__construccion")
        intersect_output_fc = os.path.join(gdb_path, "interseccion_construccion")
        intersect_output_fc_filtro = os.path.join(gdb_path, "interseccion_construccion_filtro")
        try:
            # Verificar y eliminar capas existentes
            for fc in [unidad_puntos_fc, formal_filtrado_fc, informal_filtrado_fc, informal_dissolve_fc, erase_output_fc, merge_output_fc, intersect_output_fc]:
                if arcpy.Exists(fc):
                    arcpy.Delete_management(fc)

            # Ejecutar Merge (terreno y unidades de construcción en puntos)
            arcpy.CopyFeatures_management(feature_class_path, merge_output_fc)
            print("Copia de terreno completada: ", merge_output_fc)

            # Filtrar 'formal' (donde el 22° dígito de TERRENO_CODIGO no es '2')
            query_formal = "SUBSTRING(TERRENO_CODIGO, 22, 1) <> '2'"
            arcpy.MakeFeatureLayer_management(feature_class_path, "formal_layer", query_formal)
            arcpy.CopyFeatures_management("formal_layer", formal_filtrado_fc)

            # Filtrar 'informal' (donde el 22° dígito de TERRENO_CODIGO es '2')
            query_informal = "SUBSTRING(TERRENO_CODIGO, 22, 1) = '2'"
            arcpy.MakeFeatureLayer_management(feature_class_path, "informal_layer", query_informal)
            arcpy.CopyFeatures_management("informal_layer", informal_filtrado_fc)

            # Ejecutar Dissolve si el campo DIMENSION existe
            campos_dissolve = ["DIMENSION"] if "DIMENSION" in [f.name for f in arcpy.ListFields(informal_filtrado_fc)] else None
            arcpy.Dissolve_management(informal_filtrado_fc, informal_dissolve_fc, campos_dissolve)
            print("Dissolve completado: ", informal_dissolve_fc)

            # Ejecutar Erase (descartar áreas de intersección)
            if arcpy.Exists(informal_dissolve_fc) and int(arcpy.GetCount_management(informal_dissolve_fc)[0]) > 0:
                arcpy.Erase_analysis(merge_output_fc, informal_dissolve_fc, erase_output_fc)
                print("Erase completado: ", erase_output_fc)
            else:
                print("Advertencia: 'informal_dissolved1' no tiene datos o no existe, se omite el Erase.")
                erase_output_fc = merge_output_fc  # Si no hay intersecciones, usa el merge original

            # Agregar datos de 'informal_filtrado' al resultado final
            if arcpy.Exists(informal_filtrado_fc) and int(arcpy.GetCount_management(informal_filtrado_fc)[0]) > 0:
                arcpy.Append_management(informal_filtrado_fc, erase_output_fc, "NO_TEST")
                print("Datos de 'informal_filtrado' copiados a 'unico'.")
            else:
                print("Advertencia: 'informal_filtrado' no tiene datos, se omite el Append.")

            # **Ejecutar Intersección entre erase_output_fc y unidad_puntos_fc**
            
            def verificar_geometria(capa):
                with arcpy.da.SearchCursor(capa, ["SHAPE@"]) as cursor:
                    for row in cursor:
                        if row[0] is None:
                            print("La capa {capa} tiene geometrías vacías.")
                            return False
                return True

            if verificar_geometria(feature_class_path_construccion) and verificar_geometria(erase_output_fc):
                    arcpy.Intersect_analysis([feature_class_path_construccion, erase_output_fc], intersect_output_fc, "ALL", "", "INPUT")
            else:
                print("Algunas capas no tienen geometrías válidas.")
            
            arcpy.AddField_management(intersect_output_fc, "TERRENO_22", "TEXT", field_length = 22)
            arcpy.AddField_management(intersect_output_fc, "CONSTRUCCION_22", "TEXT", field_length = 22)
            arcpy.AddField_management(intersect_output_fc, "DIFERENCIA", "LONG")


            arcpy.CalculateField_management(intersect_output_fc, "TERRENO_22", "Left([TERRENO_CODIGO],22)", "VB")
            arcpy.CalculateField_management(intersect_output_fc, "CONSTRUCCION_22", "Left([CODIGO_CONSTRUCCION],22)", "VB")
            arcpy.CalculateField_management(intersect_output_fc, "DIFERENCIA", "[TERRENO_22] = [CONSTRUCCION_22]", "VB")
            query_filtro = "DIFERENCIA = 0"
            arcpy.MakeFeatureLayer_management(intersect_output_fc, "temp_layer_name", query_filtro)
            arcpy.CopyFeatures_management("temp_layer_name", intersect_output_fc_filtro)

            fields = [field.name for field in arcpy.ListFields(intersect_output_fc_filtro)]
            data = [list(row) for row in arcpy.da.SearchCursor(intersect_output_fc_filtro, fields)] if arcpy.Exists(intersect_output_fc_filtro) else []
            df = pd.DataFrame(data, columns=fields)
            df = df[['TERRENO_CODIGO','CODIGO_CONSTRUCCION']]
            
            print(df)

            if df.empty:
                print("Advertencia: el DataFrame está vacío después del filtro.")

            return df  
        except Exception as e:
            tkMessageBox.showerror("Error", "No se pudo procesar: {}".format(str(e)))
            print("Error:", e)
            return None

    def select_gdb(self):
        path = tkFileDialog.askdirectory(title="Seleccionar Geodatabase (GDB)")
        if path:
            self.gdb_path.set(path)
    print("")
    def select_excel(self):
        path = tkFileDialog.askopenfilename(title="Seleccionar Archivo Excel", 
                                             filetypes=[("Excel files", ".xlsx;.xls")])
        if path:
            self.excel_path.set(path)
    
    def select_excel_bcgs(self):
        path = tkFileDialog.askopenfilename(title="Seleccionar Archivo Excel BCGS", 
                                             filetypes=[("Excel files", ".xlsx;.xls")])
        if path:
            self.excel_path_bcgs.set(path)

    def select_output_excel(self):
        # Permite al usuario seleccionar la ruta y nombre del archivo de salida.
        path = tkFileDialog.asksaveasfilename(
            title="Guardar archivo Excel",
            defaultextension=".xls",
            filetypes=[("Excel files", "*.xls;*.xlsx")]
        )
        
        if path:  # Solo actualizar si el usuario seleccionó un archivo
            self.output_excel.set(path)
        else:
            tkMessageBox.showerror("Error", "Debe seleccionar una ruta de salida válida.")


        

if __name__ == "__main__":
    root = tk.Tk()
    app = GDBExcelValidator(root)
    root.mainloop()