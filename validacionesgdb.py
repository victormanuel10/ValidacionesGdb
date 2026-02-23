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
import sys
import ttk
import re

try:
    reload(sys)
    sys.setdefaultencoding('utf-8')
except:
    pass
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
        self.excel_path_bcgs = tk.StringVar()
        self.output_excel = tk.StringVar()
        self.tipo_area = tk.StringVar(value="Rural")  # Valor por defecto: Rural

        # ✅ Crear Notebook antes de usarlo
        self.notebook = ttk.Notebook(parent)
        self.notebook.pack(fill='both', expand=True)

        self.tab_topologia = tk.Frame(self.notebook)
        self.notebook.add(self.tab_topologia, text="Topología")

      
        btn_validar_topologia = tk.Button(self.tab_topologia, text="Validar Topología", command=self.ejecutar_validacion_topologia)
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
        
        feature_class_name = "r_lc_terreno" if self.tipo_area.get() == "Rural" else "u_lc_terreno"
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
            npn_excel = [str(n) for n in npn_excel]
            terreno_codigo_gdb = [str(n) for n in terreno_codigo_gdb]
            diff_terreno_codigo = set(terreno_codigo_gdb) - set(npn_excel)
            
            diff_npn_excel = set(npn_excel) - set(terreno_codigo_gdb)

            def filtrar_omisiones(npn_list):
                return [
                    npn for npn in npn_list
                    if len(npn) >= 30
                    and npn[21] in ('8', '9', '0', '2', '4', '5', '3')
                    and (npn[21] == '8' or npn[26:30] == '0000')
                ]
            
            self.verificar_geometrias_vacias(feature_class_name_construccion)
            omisiones_filtradas = filtrar_omisiones(diff_npn_excel)

            # Crear archivo Excel
            workbook = xlwt.Workbook()
            sheet_comisiones = workbook.add_sheet('Comisiones')
            sheet_omisiones = workbook.add_sheet('Omisiones')
            sheet_diferencias = workbook.add_sheet('Diferencia Areas Construidas')
            sheet_df_comparar_areas_por_unidad = workbook.add_sheet('Comparar Areas por Unidad')
            #sheet_df_validar_anio_construccion = workbook.add_sheet('Validar Anio Construccion')
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
            sheet_df_validar_altura_y_anio_unidad = workbook.add_sheet('Altura y Anio Unidad')
            #sheet_validar = workbook.add_sheet('Numero de pisos')
            
            #sheet_reporte=workbook.add_sheet('Reporte')
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
            
            df_informalidad_sobre_predio=self.validacion_informalidad_sobre_predio(gdb_path)
            
            
            
            df_validar_terreno_codigo_duplicado_ficha=self.validar_terreno_codigo_duplicado_ficha(gdb_path)
            self.extraer_letras_identificador(gdb_path)
            
            
            #df_validar=self.validar(gdb_path)
            
            '''
            if df_validar is not None:
                df_filtrado_pisos = df_validar[df_validar["Diferencia"] != 0]
                if df_filtrado_pisos is not None and len(df_filtrado_pisos) > 0:
                    # df_filtrado_pisos tiene filas
                    print(df_filtrado_pisos)
                else:
                    print("Todos tienen diferencia cero.")
            else:
                print("No se generó DataFrame.") 
                '''

            
            
            df_diferencias_areas_construidas=self.calcular_areas_construidas()
            df_npns_duplicados = self.validar_terreno_codigo_duplicado(gdb_path)
            df_ph_sin_unidad = self.calcular_campos_y_filtrar(gdb_path, self.tipo_area.get())
            df_terreno_con_nro_piso= self.validar_terreno_con_piso(gdb_path, self.tipo_area.get())
            #print("df_terreno_con_nro_piso")
            #print (df_terreno_con_nro_piso)
            #df_terreno_con_nro_piso = pd.DataFrame(terreno_con_nro_piso, columns=["TERRENO_CODIGO_Con_Nro_Piso"])
            
            
            #print(df_npns_duplicados)
            df_informalidades_sin_predio_formal=self.informalidad_sin_predio_formal(gdb_path)
            df_npn__unidad_diferente_de_terreno=self.validar_npn__unidad_diferente_de_terreno(gdb_path)
            df_npn__construccion_diferente_de_terreno=self.validar_npn__construccion_diferente_de_terreno(gdb_path)
            df_etiqueta=self.validar_etiqueta(gdb_path)
            df_informalidad_condicion2_vs_formal_area=self.informalidades_suman_mas_que_formal(gdb_path)
            df_comparar_areas_por_unidad=self.comparar_areas_por_unidad()
            df_validar_altura_y_anio_unidad=self.validar_altura_y_anio_unidad(gdb_path)
            #df_validar_anio_construccion=self.validar_anio_construccion()
    
            reportes_dict = {
                u"Comisiones": self.agregar_condicion_predio(pd.DataFrame(list(diff_terreno_codigo), columns=["TERRENO_CODIGO"])),
                u"Omisiones": self.agregar_condicion_predio(pd.DataFrame(list(omisiones_filtradas), columns=["Npn"])),
                u"Diferencias de area > 2.5": self.agregar_condicion_predio(df_diferencias_areas_construidas),
                u"NPNs Duplicados": self.agregar_condicion_predio(df_npns_duplicados),
                u"Terreno Duplicados": self.agregar_condicion_predio(df_validar_terreno_codigo_duplicado_ficha),
                u"PH sin Unidad Predial": self.agregar_condicion_predio(df_ph_sin_unidad),
                u"Terrenos con Numero de Piso": self.agregar_condicion_predio(df_terreno_con_nro_piso),
                u"Informalidades Sin Predio Formal": self.agregar_condicion_predio(df_informalidades_sin_predio_formal),
                u"Area informal superior a Fomral": self.agregar_condicion_predio(df_informalidad_condicion2_vs_formal_area),
                u"NPN Unidad Diferente de Terreno": self.agregar_condicion_predio(df_npn__unidad_diferente_de_terreno),
                u"NPN Construcción Diferente de Terreno": self.agregar_condicion_predio(df_npn__construccion_diferente_de_terreno),
                u"Informalidades Sobre Predio": self.agregar_condicion_predio(df_informalidad_sobre_predio),
                u"Etiqueta": self.agregar_condicion_predio(df_etiqueta),
                u"Comparar Areas por Unidad": self.agregar_condicion_predio(df_comparar_areas_por_unidad),
                u"Validar Altura y Año Unidad": self.agregar_condicion_predio(df_validar_altura_y_anio_unidad),
            }
            
            df_fichas = pd.read_excel(excel_path, sheet_name='Fichas')
            total_fichas = df_fichas['NroFicha'].nunique()
            # Llamar a la función reporte
            self.reporte(workbook, reportes_dict, gdb_path)
            
    

            
            
            
            
            for col_num, column in enumerate(df_diferencias_areas_construidas.columns):
                    sheet_diferencias.write(0, col_num, column.decode('utf-8'), bold_style)

            # Escribir los datos de diferencias de áreas construidas
            for row_num, row in enumerate(df_diferencias_areas_construidas.itertuples(index=False), 1):
                for col_num, value in enumerate(row):
                    if isinstance(value, float):
                        value_str = ('%.2f' % value).replace('.', ',')
                        sheet_diferencias.write(row_num, col_num, value_str.decode('utf-8'))
                    else:
                        sheet_diferencias.write(row_num, col_num, str(value).decode('utf-8'))
            
            for col_num, column in enumerate(df_terreno_con_nro_piso.columns):
                sheet_terreno_nro_piso.write(0, col_num, column.decode('utf-8'), bold_style)

            # Escribir datos
            for row_num, row in enumerate(df_terreno_con_nro_piso.itertuples(index=False), 1):
                for col_num, value in enumerate(row):
                    sheet_terreno_nro_piso.write(row_num, col_num, str(value).decode('utf-8'))


            for col_num, column in enumerate(df_npns_duplicados.columns):
                sheet_duplicados.write(0, col_num, column.decode('utf-8'), bold_style)

            # Escribir datos
            for row_num, row in enumerate(df_npns_duplicados.itertuples(index=False), 1):
                for col_num, value in enumerate(row):
                    sheet_duplicados.write(row_num, col_num, str(value).decode('utf-8'))

            
            for col_num, column in enumerate(df_ph_sin_unidad.columns):
                sheet_ph_sin_unidad.write(0, col_num, column.decode('utf-8'), bold_style)

            # Escribir datos
            for row_num, row in enumerate(df_ph_sin_unidad.itertuples(index=False), 1):
                for col_num, value in enumerate(row):
                    sheet_ph_sin_unidad.write(row_num, col_num, str(value).decode('utf-8'))

            for col_num, column in enumerate(df_informalidades_sin_predio_formal.columns):
                sheet_informalidades_sin_predio_formal.write(0, col_num, column.decode('utf-8'), bold_style)

            # Escribir datos
            for row_num, row in enumerate(df_informalidades_sin_predio_formal.itertuples(index=False), 1):
                for col_num, value in enumerate(row):
                    sheet_informalidades_sin_predio_formal.write(row_num, col_num, str(value).decode('utf-8'))

            
            for col_num, column in enumerate(df_informalidad_condicion2_vs_formal_area.columns):
                sheet_df_informalidad_condicion2_vs_formal_area.write(0, col_num, column.decode('utf-8'), bold_style)

            # Escribir datos
            for row_num, row in enumerate(df_informalidad_condicion2_vs_formal_area.itertuples(index=False), 1):
                for col_num, value in enumerate(row):
                    sheet_df_informalidad_condicion2_vs_formal_area.write(row_num, col_num, str(value).decode('utf-8'))
            
            
            for col_num, column in enumerate(df_npn__unidad_diferente_de_terreno.columns):
                sheet_npn__unidad_diferente_de_terreno.write(0, col_num, column.decode('utf-8'), bold_style)

            # Escribir datos
            for row_num, row in enumerate(df_npn__unidad_diferente_de_terreno.itertuples(index=False), 1):
                for col_num, value in enumerate(row):
                    sheet_npn__unidad_diferente_de_terreno.write(row_num, col_num, str(value).decode('utf-8'))

            
            
            
            
            for col_num, column in enumerate(df_npn__construccion_diferente_de_terreno.columns):
                sheet_npn__construccion_diferente_de_terreno.write(0, col_num, column.decode('utf-8'), bold_style)

            # Escribir datos
            for row_num, row in enumerate(df_npn__construccion_diferente_de_terreno.itertuples(index=False), 1):
                for col_num, value in enumerate(row):
                    sheet_npn__construccion_diferente_de_terreno.write(row_num, col_num, str(value).decode('utf-8'))
            
            
            
            
            
            for col_num, column in enumerate(df_informalidad_sobre_predio.columns):
                sheet_npn_validacion_informalidad_sobre_predio.write(0, col_num, column.decode('utf-8'), bold_style)

            # Escribir datos
            for row_num, row in enumerate(df_informalidad_sobre_predio.itertuples(index=False), 1):
                for col_num, value in enumerate(row):
                    sheet_npn_validacion_informalidad_sobre_predio.write(row_num, col_num, str(value).decode('utf-8'))
            
            
            for col_num, column in enumerate(df_validar_terreno_codigo_duplicado_ficha.columns):
                sheet_duplicados_terreno.write(0, col_num, column.decode('utf-8'), bold_style)

            # Escribir datos
            for row_num, row in enumerate(df_validar_terreno_codigo_duplicado_ficha.itertuples(index=False), 1):
                for col_num, value in enumerate(row):
                    sheet_duplicados_terreno.write(row_num, col_num, str(value).decode('utf-8'))

            for col_num, column in enumerate(df_etiqueta.columns):
                sheet_etiqueta.write(0, col_num, column.decode('utf-8'), bold_style)

            # Escribir datos
            for row_num, row in enumerate(df_etiqueta.itertuples(index=False), 1):
                for col_num, value in enumerate(row):
                    sheet_etiqueta.write(row_num, col_num, str(value).decode('utf-8'))
            
            '''
            for col_num, column in enumerate(df_filtrado_pisos.columns):
                sheet_validar.write(0, col_num, column.decode('utf-8'), bold_style)

            # Escribir datos
            for row_num, row in enumerate(df_filtrado_pisos.itertuples(index=False), 1):
                for col_num, value in enumerate(row):
                    sheet_validar.write(row_num, col_num, str(value).decode('utf-8'))
                    '''
            
            for col_num, column in enumerate(df_comparar_areas_por_unidad.columns):
                sheet_df_comparar_areas_por_unidad.write(0, col_num, column.decode('utf-8'), bold_style)

            # Escribir datos
            for row_num, row in enumerate(df_comparar_areas_por_unidad.itertuples(index=False), 1):
                for col_num, value in enumerate(row):
                    sheet_df_comparar_areas_por_unidad.write(row_num, col_num, str(value).decode('utf-8'))

            
            
            for col_num, column in enumerate(df_validar_altura_y_anio_unidad.columns):
                sheet_df_validar_altura_y_anio_unidad.write(0, col_num, column.decode('utf-8'), bold_style)

            # Escribir datos
            for row_num, row in enumerate(df_validar_altura_y_anio_unidad.itertuples(index=False), 1):
                for col_num, value in enumerate(row):
                    sheet_df_validar_altura_y_anio_unidad.write(row_num, col_num, str(value).decode('utf-8'))
            
                            
            """
            
            for col_num, column in enumerate(df_validar_anio_construccion.columns):
                sheet_df_validar_anio_construccion.write(0, col_num, column.decode('utf-8'), bold_style)

            # Escribir datos
            for row_num, row in enumerate(df_validar_anio_construccion.itertuples(index=False), 1):
                for col_num, value in enumerate(row):
                    sheet_df_validar_anio_construccion.write(row_num, col_num, str(value).decode('utf-8'))
            """
            workbook.save(output_path)
            
            tkMessageBox.showinfo("Éxito".decode('utf-8'), u"Proceso finalizado correctamente.\nArchivos guardados en:\n" +
                                output_path.decode('utf-8'))

        except Exception as e:
            tkMessageBox.showerror("Error", str(e))
    
    def ejecutar_validacion_topologia(self):
        import tkMessageBox
        import os

        gdb_path = self.gdb_path.get()
        if not gdb_path or not gdb_path.endswith(".gdb"):
            tkMessageBox.showerror("Error", "Debe seleccionar una GDB válida.")
            return

        ruta_excel = os.path.join(os.path.dirname(gdb_path), "reporte_topologia.xls")
        try:
            self.validar_topologia(gdb_path, ruta_excel)
        except Exception as e:
            tkMessageBox.showerror("Error", "Fallo la validación de topología:\n{}".format(str(e)))
    
    def agregar_condicion_predio(self, df):
        """
        Agrega la columna CONDICION_PREDIO a cualquier DataFrame
        que tenga TERRENO_CODIGO o Npn.
        """
        if df is None or df.empty:
            return df

        df = df.copy()

        if 'TERRENO_CODIGO' in df.columns:
            df['CONDICION_PREDIO'] = df['TERRENO_CODIGO'].apply(
                lambda x: str(x)[21] if x and len(str(x)) >= 22 else ''
            )

        elif 'Npn' in df.columns:
            df['CONDICION_PREDIO'] = df['Npn'].apply(
                lambda x: str(x)[21] if x and len(str(x)) >= 22 else ''
            )

        return df
    

    def calcular_areas_construidas(self):
        
        sys.setdefaultencoding('utf-8')  # Evitar errores de codificación en Python 2.7

        gdb_path = self.gdb_path.get()
        excel_path = self.excel_path.get()

        if not excel_path:
            tkMessageBox.showerror("Error", "Debe seleccionar el Excel.")
            return

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

        # ---- 1) Traer tabla desde GDB (espera columnas CODIGO_UNIDAD_CONSTRUCCION y SHAPE_Area)
        df_unidadconstruccion = extraer_tabla_de_gdb_area_construida(gdb_path, self.tipo_area.get())
        if df_unidadconstruccion is None or df_unidadconstruccion.empty:
            tkMessageBox.showerror("Error", "No se pudo obtener datos de la GDB.")
            return

        # Asegurar columnas clave
        for col in ["CODIGO_UNIDAD_CONSTRUCCION", "SHAPE_Area"]:
            if col not in df_unidadconstruccion.columns:
                tkMessageBox.showerror("Error", "Falta la columna '{}' en datos de GDB.".format(col))
                return

        # ---- 2) Detectar hojas relevantes en el Excel
        try:
            xl = pd.ExcelFile(excel_path)
        except Exception as e:
            tkMessageBox.showerror("Error", "No se pudo abrir el Excel:\n{}".format(str(e)))
            return

        lower_sheets = [h.strip().lower() for h in xl.sheet_names]
        hoja_construcciones = None
        hoja_fichas = None
        for h in xl.sheet_names:
            l = h.strip().lower()
            if ("construccion" in l) and hoja_construcciones is None:
                hoja_construcciones = h
            if ("ficha" in l) and hoja_fichas is None:
                # acepta "Fichas" o "Ficha"
                hoja_fichas = h

        if not hoja_construcciones or not hoja_fichas:
            tkMessageBox.showerror("Error", "No se encontró la(s) hoja(s) de 'Construcciones' y/o 'Fichas' en el Excel.")
            return

        # ---- 3) Leer Excel y asegurar tipos
        try:
            df_construcciones = pd.read_excel(excel_path, sheetname=hoja_construcciones)
            df_fichas = pd.read_excel(excel_path, sheetname=hoja_fichas)
        except Exception as e:
            tkMessageBox.showerror("Error", "No se pudo leer Excel:\n{}".format(str(e)))
            return

        # Columnas requeridas
        for col in ["NroFicha"]:
            if col not in df_construcciones.columns and col not in df_fichas.columns:
                tkMessageBox.showerror("Error", "Falta la columna '{}' en el Excel.".format(col))
                return

        # Si no existe AreaConstruida en Construcciones, créala en 0
        if "AreaConstruida" not in df_construcciones.columns:
            df_construcciones["AreaConstruida"] = 0

        # Asegurar numérico y llenar NaN con 0 (solicitud recurrente tuya)
        df_construcciones["AreaConstruida"] = pd.to_numeric(df_construcciones["AreaConstruida"], errors='coerce').fillna(0)

        # Agregar AreaConstruida por NroFicha a Fichas
        if "NroFicha" not in df_construcciones.columns or "NroFicha" not in df_fichas.columns:
            tkMessageBox.showerror("Error", "No se encuentra 'NroFicha' en ambas hojas para el merge.")
            return

        df_fichas = df_fichas.merge(
            df_construcciones[["NroFicha", "AreaConstruida"]],
            on="NroFicha", how="left"
        )
        # Si el merge dejó NaN, convertir a 0
        df_fichas["AreaConstruida"] = pd.to_numeric(df_fichas["AreaConstruida"], errors='coerce').fillna(0)

        # ---- 4) Agrupar por prefijo (22) en ambos orígenes
        # GDB
        df_gdb = df_unidadconstruccion[["CODIGO_UNIDAD_CONSTRUCCION", "SHAPE_Area"]].copy()
        df_gdb["clave22"] = df_gdb["CODIGO_UNIDAD_CONSTRUCCION"].astype(str).str[:22]
        df_gdb["Area_GDB"] = pd.to_numeric(df_gdb["SHAPE_Area"], errors='coerce').fillna(0)
        df_agrupado_gdb = df_gdb.groupby("clave22", as_index=False)["Area_GDB"].sum()

        # Excel (desde Fichas con AreaConstruida ya consolidada)
        if "Npn" not in df_fichas.columns:
            tkMessageBox.showerror("Error", "No se encuentra columna 'Npn' en la hoja de Fichas.")
            return
        df_excel = df_fichas[["Npn", "AreaConstruida"]].copy()
        df_excel["clave22"] = df_excel["Npn"].astype(str).str[:22]
        df_excel["Area_Excel"] = pd.to_numeric(df_excel["AreaConstruida"], errors='coerce').fillna(0)
        df_agrupado_excel = df_excel.groupby("clave22", as_index=False)["Area_Excel"].sum()

        # ---- 5) Comparación (inner para claves que existan en ambos lados)
        df_comparacion = df_agrupado_gdb.merge(df_agrupado_excel, on="clave22", how="inner")

        # Redondeos y diferencia
        df_comparacion["Area_GDB"] = df_comparacion["Area_GDB"].round(2)
        df_comparacion["Area_Excel"] = df_comparacion["Area_Excel"].round(2)
        df_comparacion["Diferencia"] = (df_comparacion["Area_GDB"] - df_comparacion["Area_Excel"]).abs().round(2)

        # Renombres para salida
        df_comparacion.rename(columns={
            "clave22": "CODIGO/NPN (22)"
        }, inplace=True)

        # Filtrar > 2.5 m2
        df_filtrada = df_comparacion[df_comparacion["Diferencia"] > 2.5].copy()

        # % diferencia respecto al área GDB (0 si GDB = 0)
        def _calc_pct(gdb_val, xls_val):
            if gdb_val == 0:
                return 0.0
            return round((gdb_val - xls_val) / float(gdb_val), 4)

        df_filtrada["Porcentaje"] = [
            _calc_pct(g, e) for g, e in zip(df_filtrada["Area_GDB"], df_filtrada["Area_Excel"])
        ]

        # ---- 6) (Opcional) Exportar a Excel con formato de porcentaje
        # Descomenta este bloque si quieres exportar el archivo
        """
        try:
            from openpyxl import load_workbook
            output_path = os.path.join(os.path.dirname(excel_path), "comparacion_areas.xlsx")
            writer = pd.ExcelWriter(output_path, engine='openpyxl')

            df_filtrada.to_excel(writer, index=False, sheet_name="Comparacion")

            ws = writer.book["Comparacion"]
            # Ajuste simple de ancho de columnas
            for col_idx, col_name in enumerate(df_filtrada.columns, 1):
                ws.column_dimensions[chr(64 + col_idx)].width = max(12, min(30, len(str(col_name)) + 2))

            # Formato porcentaje (columna 'Porcentaje' está al final)
            from openpyxl.styles import numbers
            pct_col_idx = df_filtrada.columns.get_loc("Porcentaje") + 1  # 1-based
            for row in range(2, ws.max_row + 1):  # omite encabezado
                ws.cell(row=row, column=pct_col_idx).number_format = numbers.FORMAT_PERCENTAGE_00

            writer.save()
            tkMessageBox.showinfo("Exportación completada", "Archivo guardado:\n%s" % output_path)
        except Exception as e:
            tkMessageBox.showwarning("Aviso", "No se pudo exportar a Excel (se devuelve DataFrame):\n{}".format(str(e)))
        """

        return df_filtrada


    
    def validar_terreno_codigo_duplicado(self, gdb_path):
        

        feature_class_name = "r_lc_terreno" if self.tipo_area.get() == "Rural" else "u_lc_terreno"
        feature_class_path = os.path.join(gdb_path, feature_class_name)
        arcpy.env.workspace = gdb_path
        feature_classes = arcpy.ListFeatureClasses()

        if not arcpy.Exists(feature_class_path):
            raise Exception("La capa {feature_class_name} no existe en la GDB.")
        print("validar_terreno_codigo_duplicado")
        fields = [field.name for field in arcpy.ListFields(feature_class_path)]
        if "TERRENO_CODIGO" not in fields:
            raise Exception("La columna 'TERRENO_CODIGO' no existe en la Feature Class.")

        # Leer los valores de TERRENO_CODIGO
        terreno_codigos = [row[0] for row in arcpy.da.SearchCursor(feature_class_path, ["TERRENO_CODIGO"]) if row[0]]

        # Contar duplicados
        contador_codigos = {}
        for codigo in terreno_codigos:
            contador_codigos[codigo] = contador_codigos.get(codigo, 0) + 1

        duplicados = [codigo for codigo, count in contador_codigos.items() if count > 1]

        if duplicados:
            # Desglosar duplicados usando la función auxiliar
            data = []
            for npn in duplicados:
                columnas = dividir_npn_en_columnas(npn)
                data.append([npn] + columnas)

            columnas_df = ["TERRENO_CODIGO", "Departamento", "Municipio", "Zona", "Sector",
                        "Comuna", "Barrio", "Manzana/Vereda", "Terreno/Predios",
                        "Condición Predio", "Edificio", "Número Piso", "Unidad Predial"]

            df = pd.DataFrame(data, columns=columnas_df)
            return df
        else:
            return pd.DataFrame(columns=["TERRENO_CODIGO"])
    

    def calcular_campos_y_filtrar(self, gdb_path, tipo_area):
       

        feature_class_name = "r_lc_unidadconstruccion" if tipo_area == "Rural" else "u_lc_unidadconstruccion"
        feature_class_path = os.path.join(gdb_path, feature_class_name)
        filtro_feature_class_path = os.path.join(gdb_path, "filtro_unidadconstruccion_campos_filtrar")
        arcpy.env.workspace = gdb_path

        if not arcpy.Exists(feature_class_path):
            raise Exception("La capa {feature_class_name} no existe en la GDB.")
        print("calcular_campos_y_filtrar")
        # Agregar campos si no existen
        fields = [f.name for f in arcpy.ListFields(feature_class_path)]
        if "CP" not in fields:
            arcpy.AddField_management(feature_class_path, "CP", "TEXT", field_length=1)
        if "UNIDAD" not in fields:
            arcpy.AddField_management(feature_class_path, "UNIDAD", "TEXT", field_length=4)

        # Calcular campos a partir del código
        arcpy.CalculateField_management(feature_class_path, "CP", "Mid([CODIGO_UNIDAD_CONSTRUCCION], 22, 1)", "VB")
        arcpy.CalculateField_management(feature_class_path, "UNIDAD", "Right([CODIGO_UNIDAD_CONSTRUCCION],4)", "VB")

        # Filtrar unidades con CP=9, UNIDAD=0000 y TIPO_DOMINIO=1
        query_filtro = "CP = '9' AND UNIDAD = '0000' AND TIPO_DOMINIO = 1"
        arcpy.MakeFeatureLayer_management(feature_class_path, "temp_layer_name_calcular_campos_filtrar", query_filtro)
        arcpy.CopyFeatures_management("temp_layer_name_calcular_campos_filtrar", filtro_feature_class_path)

        if arcpy.Exists(filtro_feature_class_path):
            fields = [field.name for field in arcpy.ListFields(filtro_feature_class_path)]
            data = [list(row) for row in arcpy.da.SearchCursor(filtro_feature_class_path, fields)]
            df = pd.DataFrame(data, columns=fields)
            df = df[['CODIGO_UNIDAD_CONSTRUCCION']]

            if not df.empty:
                data = []
                for npn in df['CODIGO_UNIDAD_CONSTRUCCION']:
                    columnas = dividir_npn_en_columnas(npn)
                    data.append([npn] + columnas)

                columnas_df = ["CODIGO_UNIDAD_CONSTRUCCION", "Departamento", "Municipio", "Zona", "Sector",
                            "Comuna", "Barrio", "Manzana/Vereda", "Terreno/Predios",
                            "Condición Predio", "Edificio", "Número Piso", "Unidad Predial"]

                df = pd.DataFrame(data, columns=columnas_df)

            return df

        else:
            return pd.DataFrame(columns=["CODIGO_UNIDAD_CONSTRUCCION"])

    def validar_terreno_con_piso(self, gdb_path, tipo_area):
        feature_class_name = "r_lc_terreno" if tipo_area == "Rural" else "u_lc_terreno"
        feature_class_path = os.path.join(gdb_path, feature_class_name)
        filtro_feature_class_path = os.path.join(gdb_path, "filtro_unidadconstruccion_terreno_con_piso")

        arcpy.env.workspace = gdb_path

        if not arcpy.Exists(feature_class_path):
            raise Exception("La capa {feature_class_name} no existe en la GDB.")
        print("validar_terreno_con_piso")
        fields = [field.name for field in arcpy.ListFields(feature_class_path)]

        if "TERRENO_CODIGO" not in fields:
            raise Exception("La columna 'TERRENO_CODIGO' no existe en la Feature Class.")

        if "CP" not in fields:
            arcpy.AddField_management(feature_class_path, "CP", "TEXT", field_length=1)
        if "UNIDAD" not in fields:
            arcpy.AddField_management(feature_class_path, "ULTIMOS_8", "TEXT", field_length=8)

        arcpy.CalculateField_management(feature_class_path, "CP", "Mid([TERRENO_CODIGO], 22, 1)", "VB")
        arcpy.CalculateField_management(feature_class_path, "ULTIMOS_8", "Right([TERRENO_CODIGO],8)", "VB")
        
        query_filtro = "CP <> '8' AND ULTIMOS_8 <> '00000000'"

        arcpy.MakeFeatureLayer_management(feature_class_path, "temp_layer_name_terreno_piso", query_filtro)
        arcpy.CopyFeatures_management("temp_layer_name_terreno_piso", filtro_feature_class_path)
        
        fields = [field.name for field in arcpy.ListFields(filtro_feature_class_path)]
        data = [list(row) for row in arcpy.da.SearchCursor(filtro_feature_class_path, fields)] if arcpy.Exists(filtro_feature_class_path) else []
        df = pd.DataFrame(data, columns=fields)
        df = df[['TERRENO_CODIGO']]
            
        if not df.empty:
                data = []
                for npn in df['TERRENO_CODIGO']:
                    columnas = dividir_npn_en_columnas(npn)
                    data.append([npn] + columnas)

                columnas_df = ["TERRENO_CODIGO", "Departamento", "Municipio", "Zona", "Sector",
                            "Comuna", "Barrio", "Manzana/Vereda", "Terreno/Predios",
                            "Condición Predio", "Edificio", "Número Piso", "Unidad Predial"]

                df = pd.DataFrame(data, columns=columnas_df)

        return df 


    def informalidad_sin_predio_formal(self, gdb_path):
        feature_class_name = "r_lc_terreno" if self.tipo_area.get() == "Rural" else "u_lc_terreno"
        feature_class_path = os.path.join(gdb_path, feature_class_name)

        arcpy.env.workspace = gdb_path

        if not arcpy.Exists(feature_class_path):
            tkMessageBox.showerror("Error", "La capa {} no existe en la GDB.".format(feature_class_name))
            return None  
        print("informalidad_sin_predio_formal")
        # Definir rutas de salida
        formal_filtrado_fc = os.path.join(gdb_path, "formal_filtrado")
        informal_filtrado_fc = os.path.join(gdb_path, "informal_filtrado")
        informal_buffer_fc = os.path.join(gdb_path, "informal_buffer")
        join_output_fc = os.path.join(gdb_path, "informal_buffer_joined")

        try:
            # Copiar y filtrar 'formal' (donde el 22° dígito de TERRENO_CODIGO no es '2')
            query_formal = "SUBSTRING(TERRENO_CODIGO, 22, 1) <> '2'"
            arcpy.MakeFeatureLayer_management(feature_class_path, "formal_layer_copiar_filtrar", query_formal)
            arcpy.CopyFeatures_management("formal_layer_copiar_filtrar", formal_filtrado_fc)
            #print("Capa 'formal' filtrada y guardada.")

            # Copiar y filtrar 'informal' (donde el 22° dígito de TERRENO_CODIGO es '2')
            query_informal = "SUBSTRING(TERRENO_CODIGO, 22, 1) = '2'"
            arcpy.MakeFeatureLayer_management(feature_class_path, "informal_layer_copiar_filtrar", query_informal)
            arcpy.CopyFeatures_management("informal_layer_copiar_filtrar", informal_filtrado_fc)
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

            if not df.empty:
                data = []
                for npn in df['TERRENO_CODIGO']:
                    columnas = dividir_npn_en_columnas(npn)
                    data.append([npn] + columnas)

                columnas_df = ["TERRENO_CODIGO", "Departamento", "Municipio", "Zona", "Sector",
                            "Comuna", "Barrio", "Manzana/Vereda", "Terreno/Predios",
                            "Condición Predio", "Edificio", "Número Piso", "Unidad Predial"]

                df = pd.DataFrame(data, columns=columnas_df)

            return df  

        except Exception as e:
            tkMessageBox.showerror("Error", "No se pudo procesar: {}".format(str(e)))
            print("Error:", e)
            return None
        
   
    
    
    def informalidades_suman_mas_que_formal(self, gdb_path):
        """
        - Informalidad: dígito 22 de TERRENO_CODIGO = '2'
        - Formal:       dígito 22 <> '2'

        Proceso:
        1) Calcula área total de cada informalidad
        2) Calcula área REAL de intersección (Intersect) con formales
        3) Resta: DIF = AREA_INFORMALIDAD - AREA_INTERSECCION
        4) Retorna SOLO diferencias > 0.01
        """

        import os
        import pandas as pd
        import arcpy
        import tkMessageBox

        feature_class_name = "r_lc_terreno" if self.tipo_area.get() == "Rural" else "u_lc_terreno"
        fc = os.path.join(gdb_path, feature_class_name)

        arcpy.env.workspace = gdb_path
        arcpy.env.overwriteOutput = True

        if not arcpy.Exists(fc):
            tkMessageBox.showerror("Error", "La capa {} no existe en la GDB.".format(feature_class_name))
            return None

        print("validar_informalidad_intersect_diferencia")

        # -------------------------
        # Helpers
        # -------------------------
        def _fmt_num(a):
            try:
                s = "{:,.2f}".format(float(a))
                s = s.replace(",", "X").replace(".", ",").replace("X", ".")
                return s
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

        # -------------------------
        # Capas temporales
        # -------------------------
        inf_fc   = os.path.join(gdb_path, "tmp_inf_d22_2")
        for_fc   = os.path.join(gdb_path, "tmp_for_d22_no2")
        inter_fc = os.path.join(gdb_path, "tmp_inf_for_intersect")
        dis_fc   = os.path.join(gdb_path, "tmp_inf_for_intersect_dis")

        try:
            # 1) Informalidades
            arcpy.MakeFeatureLayer_management(
                fc, "lyr_inf",
                "SUBSTRING(TERRENO_CODIGO,22,1) = '2'"
            )
            arcpy.CopyFeatures_management("lyr_inf", inf_fc)

            # 2) Formales
            arcpy.MakeFeatureLayer_management(
                fc, "lyr_for",
                "SUBSTRING(TERRENO_CODIGO,22,1) <> '2'"
            )
            arcpy.CopyFeatures_management("lyr_for", for_fc)

            if _count(inf_fc) == 0:
                return pd.DataFrame(columns=[
                    "TERRENO_CODIGO",
                    "AREA_INFORMALIDAD",
                    "AREA_INTERSECCION",
                    "DIFERENCIA"
                ])

            # 3) Área total de informalidad
            area_inf = {}
            with arcpy.da.SearchCursor(inf_fc, ["TERRENO_CODIGO", "SHAPE@AREA"]) as cur:
                for tc, a in cur:
                    tc = _clean(tc)
                    if tc:
                        area_inf[tc] = float(a)

            # 4) Intersect real
            arcpy.Intersect_analysis([inf_fc, for_fc], inter_fc, "ALL", "", "INPUT")

            # Si no hay intersección
            if _count(inter_fc) == 0:
                data = []
                for tc, a_inf in area_inf.items():
                    if a_inf > 0.9:
                        data.append([
                            tc,
                            _fmt_num(a_inf),
                            "0,00",
                            _fmt_num(a_inf)
                        ])
                return pd.DataFrame(data, columns=[
                    "TERRENO_CODIGO",
                    "AREA_INFORMALIDAD",
                    "AREA_INTERSECCION",
                    "DIFERENCIA"
                ])

            # 5) Detectar campo del informal en intersect
            fields = [f.name for f in arcpy.ListFields(inter_fc)]
            campo_inf = "TERRENO_CODIGO" if "TERRENO_CODIGO" in fields else None
            if not campo_inf:
                for f in fields:
                    if f.upper().startswith("TERRENO_CODIGO"):
                        campo_inf = f
                        break

            # 6) Dissolve por informalidad
            arcpy.Dissolve_management(inter_fc, dis_fc, dissolve_field=[campo_inf])

            # 7) Área de intersección por informalidad
            area_int = {}
            with arcpy.da.SearchCursor(dis_fc, [campo_inf, "SHAPE@AREA"]) as cur:
                for tc, a in cur:
                    tc = _clean(tc)
                    if tc:
                        area_int[tc] = float(a)

            # 8) Construir resultado final
            data = []
            for tc, a_inf in area_inf.items():
                a_int = area_int.get(tc, 0.0)
                dif = a_inf - a_int

                if dif <= 0.9:
                    continue

                data.append([
                    tc,
                    _fmt_num(a_inf),
                    _fmt_num(a_int),
                    _fmt_num(dif)
                ])

            df = pd.DataFrame(data, columns=[
                "TERRENO_CODIGO",
                "AREA_INFORMALIDAD",
                "AREA_INTERSECCION",
                "DIFERENCIA"
            ])

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
        
        print("validar_npn__unidad_diferente_de_terreno")

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
            for fc in [unidad_puntos_fc, formal_filtrado_fc, informal_filtrado_fc, 
                    informal_dissolve_fc, erase_output_fc, merge_output_fc, intersect_output_fc]:
                if arcpy.Exists(fc):
                    arcpy.Delete_management(fc)

            # -------------------------------------------------------------------
            # 🔹 VALIDACIÓN Y REPARACIÓN DE GEOMETRÍAS EN UNIDADES
            # -------------------------------------------------------------------
            check_table = os.path.join(gdb_path, "unidad_checkgeom")
            if arcpy.Exists(check_table):
                arcpy.Delete_management(check_table)

            print("Ejecutando CheckGeometry...")
            arcpy.CheckGeometry_management(feature_class_path_unidad, check_table)

            print("Ejecutando RepairGeometry...")
            arcpy.RepairGeometry_management(feature_class_path_unidad, "DELETE_NULL")

            # Crear capa temporal y filtrar polígonos con área válida (> 0)
            arcpy.MakeFeatureLayer_management(feature_class_path_unidad, "unidad_layer")
            arcpy.SelectLayerByAttribute_management("unidad_layer", "NEW_SELECTION", "Shape_Area > 0")

            # **Convertir la capa de unidad de construcción en puntos**
            arcpy.FeatureToPoint_management("unidad_layer", unidad_puntos_fc, "INSIDE")
            print("Feature To Point completado: ", unidad_puntos_fc)
            # -------------------------------------------------------------------

            # Ejecutar CopyFeatures sobre terreno
            arcpy.CopyFeatures_management(feature_class_path, merge_output_fc)

            # ... 🔽 aquí sigue el resto de tu lógica original sin cambios ...
            # (filtrar formal/informal, dissolve, erase, append, intersect, etc.)
            # -------------------------------------------------------------------

            # Filtrar 'formal' (donde el 22° dígito de TERRENO_CODIGO no es '2')
            query_formal = "SUBSTRING(TERRENO_CODIGO, 22, 1) <> '2'"
            arcpy.MakeFeatureLayer_management(feature_class_path, "formal_layer_npn_unidad", query_formal)
            arcpy.CopyFeatures_management("formal_layer_npn_unidad", formal_filtrado_fc)

            # Filtrar 'informal' (donde el 22° dígito de TERRENO_CODIGO es '2')
            query_informal = "SUBSTRING(TERRENO_CODIGO, 22, 1) = '2'"
            arcpy.MakeFeatureLayer_management(feature_class_path, "informal_layer_npn_unidad", query_informal)
            arcpy.CopyFeatures_management("informal_layer_npn_unidad", informal_filtrado_fc)

            campos_dissolve = ["DIMENSION"] if "DIMENSION" in [f.name for f in arcpy.ListFields(informal_filtrado_fc)] else None
            arcpy.Dissolve_management(informal_filtrado_fc, informal_dissolve_fc, campos_dissolve)

            if arcpy.Exists(informal_dissolve_fc) and int(arcpy.GetCount_management(informal_dissolve_fc)[0]) > 0:
                arcpy.Erase_analysis(merge_output_fc, informal_dissolve_fc, erase_output_fc)
            else:
                print("Advertencia: 'informal_dissolved1' no tiene datos, se omite el Erase.")
                erase_output_fc = merge_output_fc  

            if arcpy.Exists(informal_filtrado_fc) and int(arcpy.GetCount_management(informal_filtrado_fc)[0]) > 0:
                arcpy.Append_management(informal_filtrado_fc, erase_output_fc, "NO_TEST")
            else:
                print("Advertencia: 'informal_filtrado' no tiene datos, se omite el Append.")

            if arcpy.Exists(feature_class_path) and arcpy.Exists(unidad_puntos_fc):
                arcpy.Intersect_analysis([unidad_puntos_fc, erase_output_fc], intersect_output_fc, "ALL", "", "INPUT")
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
            arcpy.MakeFeatureLayer_management(intersect_output_fc, "temp_layer_name_unidad_dif", query_filtro)
            arcpy.CopyFeatures_management("temp_layer_name_unidad_dif", intersect_output_fc_filtro)

            query_filtro_eliminar = "CP_U = '2' AND EDIFICIO_UNIDAD <> '00000000'"
            arcpy.MakeFeatureLayer_management(intersect_output_fc_filtro, "temp_layer_name_eliminar_unidad", query_filtro_eliminar)
            arcpy.DeleteRows_management("temp_layer_name_eliminar_unidad")
            arcpy.Delete_management("temp_layer_name_eliminar_unidad")

            fields = [field.name for field in arcpy.ListFields(intersect_output_fc_filtro)]
            data = [list(row) for row in arcpy.da.SearchCursor(intersect_output_fc_filtro, fields)] if arcpy.Exists(intersect_output_fc_filtro) else []
            df = pd.DataFrame(data, columns=fields)
            df = df[['TERRENO_CODIGO','CODIGO_UNIDAD_CONSTRUCCION']]

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
        print("validar_npn__construccion_diferente_de_terreno")
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
            #print("Copia de terreno completada: ", merge_output_fc)

            # Filtrar 'formal' (donde el 22° dígito de TERRENO_CODIGO no es '2')
            query_formal = "SUBSTRING(TERRENO_CODIGO, 22, 1) <> '2'"
            arcpy.MakeFeatureLayer_management(feature_class_path, "formal_layer_npn_construccion", query_formal)
            arcpy.CopyFeatures_management("formal_layer_npn_construccion", formal_filtrado_fc)

            # Filtrar 'informal' (donde el 22° dígito de TERRENO_CODIGO es '2')
            query_informal = "SUBSTRING(TERRENO_CODIGO, 22, 1) = '2'"
            arcpy.MakeFeatureLayer_management(feature_class_path, "informal_layer_npn_construccion", query_informal)
            arcpy.CopyFeatures_management("informal_layer_npn_construccion", informal_filtrado_fc)

            # Ejecutar Dissolve si el campo DIMENSION existe
            campos_dissolve = ["DIMENSION"] if "DIMENSION" in [f.name for f in arcpy.ListFields(informal_filtrado_fc)] else None
            arcpy.Dissolve_management(informal_filtrado_fc, informal_dissolve_fc, campos_dissolve)
            #print("Dissolve completado: ", informal_dissolve_fc)

            # Ejecutar Erase (descartar áreas de intersección)
            if arcpy.Exists(informal_dissolve_fc) and int(arcpy.GetCount_management(informal_dissolve_fc)[0]) > 0:
                arcpy.Erase_analysis(merge_output_fc, informal_dissolve_fc, erase_output_fc)
                #print("Erase completado: ", erase_output_fc)
            else:
                print("Advertencia: 'informal_dissolved1' no tiene datos o no existe, se omite el Erase.")
                erase_output_fc = merge_output_fc  # Si no hay intersecciones, usa el merge original

            # Agregar datos de 'informal_filtrado' al resultado final
            if arcpy.Exists(informal_filtrado_fc) and int(arcpy.GetCount_management(informal_filtrado_fc)[0]) > 0:
                arcpy.Append_management(informal_filtrado_fc, erase_output_fc, "NO_TEST")
                #print("Datos de 'informal_filtrado' copiados a 'unico'.")
            else:
                print("Advertencia: 'informal_filtrado' no tiene datos, se omite el Append.")

            # **Ejecutar Intersección entre erase_output_fc y unidad_puntos_fc**
            
            def verificar_geometria(capa):
                with arcpy.da.SearchCursor(capa, ["SHAPE@"]) as cursor:
                    for row in cursor:
                        if row[0] is None:
                            print("La capa {capa} tiene geometrias vacias.")
                            return False
                return True

            if verificar_geometria(feature_class_path_construccion) and verificar_geometria(erase_output_fc):
                    arcpy.Intersect_analysis([feature_class_path_construccion, erase_output_fc], intersect_output_fc, "ALL", "", "INPUT")
            else:
                print("Algunas capas no tienen geometrias validas.")
            
            arcpy.AddField_management(intersect_output_fc, "TERRENO_22", "TEXT", field_length = 22)
            arcpy.AddField_management(intersect_output_fc, "CONSTRUCCION_22", "TEXT", field_length = 22)
            arcpy.AddField_management(intersect_output_fc, "DIFERENCIA", "LONG")


            arcpy.CalculateField_management(intersect_output_fc, "TERRENO_22", "Left([TERRENO_CODIGO],22)", "VB")
            arcpy.CalculateField_management(intersect_output_fc, "CONSTRUCCION_22", "Left([CODIGO_CONSTRUCCION],22)", "VB")
            arcpy.CalculateField_management(intersect_output_fc, "DIFERENCIA", "[TERRENO_22] = [CONSTRUCCION_22]", "VB")
            query_filtro = "DIFERENCIA = 0"
            arcpy.MakeFeatureLayer_management(intersect_output_fc, "temp_layer_name_const_dif", query_filtro)
            arcpy.CopyFeatures_management("temp_layer_name_const_dif", intersect_output_fc_filtro)

            fields = [field.name for field in arcpy.ListFields(intersect_output_fc_filtro)]
            data = [list(row) for row in arcpy.da.SearchCursor(intersect_output_fc_filtro, fields)] if arcpy.Exists(intersect_output_fc_filtro) else []
            df = pd.DataFrame(data, columns=fields)
            df = df[['TERRENO_CODIGO','CODIGO_CONSTRUCCION']]
            
            #print(df)

            if df.empty:
                print("Advertencia: el DataFrame está vacio después del filtro.")

            return df  
        except Exception as e:
            tkMessageBox.showerror("Error", "No se pudo procesar: {}".format(str(e)))
            print("Error:", e)
            return None

    def validacion_informalidad_sobre_predio(self, gdb_path):
        feature_class_name = "r_lc_terreno" if self.tipo_area.get() == "Rural" else "u_lc_terreno"
        feature_class_path = os.path.join(gdb_path, feature_class_name)
        intersect_output = os.path.join(gdb_path, "intersect_informal_formal")

        arcpy.env.workspace = gdb_path

        if not arcpy.Exists(feature_class_path):
            tkMessageBox.showerror("Error", u"La capa {} no existe en la GDB.".format(feature_class_name))
            return None  
        print("validacion_informalidad_sobre_predio")
        # Definir rutas de salida
        formal_filtrado_fc = os.path.join(gdb_path, "formal_filtrado")
        informal_filtrado_unidad = os.path.join(gdb_path, "informal_filtrado_unidad")
        for fc in [formal_filtrado_fc, informal_filtrado_unidad, intersect_output]:
            if arcpy.Exists(fc):
                arcpy.Delete_management(fc)

        try:
            # Filtrar capas formal e informal
            query_formal = "SUBSTRING(TERRENO_CODIGO, 22, 1) <> '2'"
            arcpy.MakeFeatureLayer_management(feature_class_path, "formal_layer_predio", query_formal)
            arcpy.CopyFeatures_management("formal_layer_predio", formal_filtrado_fc)

            query_informal = "SUBSTRING(TERRENO_CODIGO, 22, 1) = '2'"
            arcpy.MakeFeatureLayer_management(feature_class_path, "informal_layer_predio", query_informal)
            arcpy.CopyFeatures_management("informal_layer_predio", informal_filtrado_unidad)

            # Ejecutar Intersect
            arcpy.Intersect_analysis(
                in_features=[formal_filtrado_fc, informal_filtrado_unidad],
                out_feature_class=intersect_output,
                join_attributes="ALL",
                output_type="INPUT"
            )

            fields = [f.name for f in arcpy.ListFields(intersect_output) if f.type not in ('Geometry', 'OID')]
            data = [[str(val) if val is not None else '' for val in row] for row in arcpy.da.SearchCursor(intersect_output, fields)]
            df = pd.DataFrame(data, columns=fields)

            # Reemplazar NaN por string vacío
            df.fillna('', inplace=True)
            df.replace(to_replace=["nan", "NaN", float('nan')], value='', inplace=True)
            errores = []

            for idx, row in df.iterrows():
                NroFicha=row.get("NroFicha","")
                NroFicha1=row.get("NroFicha_1","")
                terreno = row.get("TERRENO_CODIGO", "")
                terreno1 = row.get("TERRENO_CODIGO_1", "")
                modo = row.get("ModoAdquisicion", "")
                modo1 = row.get("ModoAdquisicion_1", "")
                tipo = row.get("PredioLcTipo", "")
                tipo1 = row.get("PredioLcTipo_1", "")
                matricula = row.get("MatriculaInmobiliaria", "")
                RazonSocial = row.get("RazonSocial", "")


                if matricula: # Caso 1: tiene matricula

                    if modo != "1|DOMINIO (TRADICION)":                    
                        errores.append({
                            "Observacion": u"Modo de adquisicion incorrecto para predio formal",
                            "NroFicha_Formal":NroFicha,
                            "NroFicha_Informal":NroFicha1,
                            "TERRENO_CODIGO_FORMAL": terreno,
                            "TERRENO_CODIGO_INFORMAL": terreno1,
                            "MatriculaInmobiliaria": matricula,
                            "PredioLcTipo_FORMAL": tipo,
                            "PredioLcTipo_INFORMALIDAD": tipo1,
                            "ModoAdquisicion_FORMAL": modo,
                            "ModoAdquisicion_INFORMALIDAD": modo1,
                            "RazonSocial": RazonSocial
                        })
                        
                    if modo1 != "2|POSESIN":
                        errores.append({
                            "Observacion": u"ModoAdquisicion incorrecto para predio informal sobre predio con matricula",
                            "NroFicha_Formal": NroFicha,
                            "NroFicha_Informal": NroFicha1,
                            "TERRENO_CODIGO_FORMAL": terreno,
                            "TERRENO_CODIGO_INFORMAL": terreno1,
                            "MatriculaInmobiliaria": matricula,
                            "PredioLcTipo_FORMAL": tipo,
                            "PredioLcTipo_INFORMALIDAD": tipo1,
                            "ModoAdquisicion_FORMAL": modo,
                            "ModoAdquisicion_INFORMALIDAD": modo1,
                            "RazonSocial": RazonSocial
                        })

                    '''if tipo1 != "Predio.Privado.Privado":
                        errores.append({
                            "Observacion": u"PredioLcTipo incorrecto para predio informal sobre predio con matricula",
                            "NroFicha_Formal": NroFicha,
                            "NroFicha_Informal": NroFicha1,
                            "TERRENO_CODIGO_FORMAL": terreno,
                            "TERRENO_CODIGO_INFORMAL": terreno1,
                            "MatriculaInmobiliaria": matricula,
                            "PredioLcTipo_FORMAL": tipo,
                            "PredioLcTipo_INFORMALIDAD": tipo1,
                            "ModoAdquisicion_FORMAL": modo,
                            "ModoAdquisicion_INFORMALIDAD": modo1,
                            "RazonSocial": RazonSocial
                        })'''

                else:  # Caso 2: sin matricula
                    if modo != "1|DOMINIO (TRADICION)":                    
                        errores.append({
                            "Observacion": u"Modo de adquisicion incorrecto para predio formal",
                            "NroFicha_Formal":NroFicha,
                            "NroFicha_Informal":NroFicha1,
                            "TERRENO_CODIGO_FORMAL": terreno,
                            "TERRENO_CODIGO_INFORMAL": terreno1,
                            "MatriculaInmobiliaria": matricula,
                            "PredioLcTipo_FORMAL": tipo,
                            "PredioLcTipo_INFORMALIDAD": tipo1,
                            "ModoAdquisicion_FORMAL": modo,
                            "ModoAdquisicion_INFORMALIDAD": modo1,
                            "RazonSocial": RazonSocial
                        })
                    
                    if modo1 != "5|OCUPACIN":
                        errores.append({
                            "Observacion": u"ModoAdquisicion incorrecto en predio informal sobre predio sin matricula",
                            "NroFicha_Formal":NroFicha,
                            "NroFicha_Informal":NroFicha1,
                            "TERRENO_CODIGO_FORMAL": terreno,
                            "TERRENO_CODIGO_INFORMAL": terreno1,
                            "MatriculaInmobiliaria": matricula,
                            "PredioLcTipo_FORMAL": tipo,
                            "PredioLcTipo_INFORMALIDAD": tipo1,
                            "ModoAdquisicion_FORMAL": modo,
                            "ModoAdquisicion_INFORMALIDAD": modo1,
                            "RazonSocial": RazonSocial
                        })
                    

            if errores:
                df_errores = pd.DataFrame(errores)
                df_errores.fillna('', inplace=True)
                df_errores.replace({
                    "ModoAdquisicion_INFORMALIDAD": {
                        "5|OCUPACIN": "5|OCUPACION",
                        "2|POSESIN": "2|POSESION"
                    },
                    "ModoAdquisicion_FORMAL": {
                        "5|OCUPACIN": "5|OCUPACION",
                        "2|POSESIN": "2|POSESION"
                    }
                }, inplace=True)
                return df_errores
            else:
                tkMessageBox.showinfo("Validacion completada", u"No se encontraron errores.")
                return pd.DataFrame()  # vacío si no hay errores

        except Exception as e:
            tkMessageBox.showerror("Error", u"No se pudo procesar: {}".format(str(e)))
            print("Error:", e)
            return None

    
    def validar(self, gdb_path):
        feature_class_name = "r_lc_terreno" if self.tipo_area.get() == "Rural" else "u_lc_terreno"
        feature_class_path = os.path.join(gdb_path, feature_class_name)
        feature_class_name_unidad = "r_lc_unidadconstruccion" if self.tipo_area.get() == "Rural" else "u_lc_unidadconstruccion"
        feature_class_path_unidad = os.path.join(gdb_path, feature_class_name_unidad)
                
        arcpy.env.workspace = gdb_path

        if not arcpy.Exists(feature_class_path):
            tkMessageBox.showerror("Error", u"La capa {} no existe en la GDB.".format(feature_class_name))
            return None

        try:
            
            campos = [f.name for f in arcpy.ListFields(feature_class_path_unidad)]

            if "CODIGO_UNIDAD_CONSTRUCCION" not in campos:
                tkMessageBox.showerror("Error", u"La capa '{}' no contiene el campo 'CODIGO_UNIDAD_CONSTRUCCION'.".format(feature_class_name_unidad))
                return None

            if "npn_22" not in campos:
                arcpy.AddField_management(feature_class_path_unidad, "npn_22", "TEXT", "", "", 22)

            arcpy.CalculateField_management(
                in_table=feature_class_path_unidad,
                field="npn_22",
                expression="Left([CODIGO_UNIDAD_CONSTRUCCION], 22)",
                expression_type="VB"
            )

            layer_name = "unidadconstruccion_lyr"
            arcpy.MakeFeatureLayer_management(feature_class_path_unidad, layer_name)

            arcpy.SelectLayerByAttribute_management(
                in_layer_or_view=layer_name,
                selection_type="NEW_SELECTION",
                where_clause="PLANTA_UBICACION BETWEEN 1 AND 80"
            )

            conteo = int(arcpy.GetCount_management(layer_name).getOutput(0))
            #print("Registros seleccionados (PLANTA_UBICACION 1-80):", conteo)

            dissolve_output = os.path.join(gdb_path, "unidadconstruccion_dissolve_npn22")
            if conteo > 0:
                if arcpy.Exists(dissolve_output):
                    arcpy.Delete_management(dissolve_output)

                arcpy.Dissolve_management(
                    in_features=layer_name,
                    out_feature_class=dissolve_output,
                    dissolve_field="npn_22",
                    statistics_fields=[["PLANTA_UBICACION", "MAX"]]
                )
                #print("Dissolve completado:", dissolve_output)
            else:
                print("No hay registros con PLANTA_UBICACION entre 1 y 80.")
                return None

            # --- BUFFER ---
            feature_class_name_construccion = "r_lc_construccion" if self.tipo_area.get() == "Rural" else "u_lc_construccion"
            construccion_fc = os.path.join(gdb_path, feature_class_name_construccion)
            buffer_output = os.path.join(gdb_path, feature_class_name_construccion + "_buffer_menos_0_5")

            if not arcpy.Exists(construccion_fc):
                print("La capa 'r_lc_construccion' no existe en la GDB.")
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
            #print("Buffer -0.5 m generado:", buffer_output)

            # --- INTERSECT ---
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

            # --- CALCULO DE DIFERENCIA ---
            campos_intersect = [f.name for f in arcpy.ListFields(intersect_output)]
            if "Diferencia" not in campos_intersect:
                arcpy.AddField_management(intersect_output, "Diferencia", "SHORT")

            arcpy.CalculateField_management(
                in_table=intersect_output,
                field="Diferencia",
                expression="[MAX_PLANTA_UBICACION] - [NUMERO_PISOS]",
                expression_type="VB"
            )
            print("Campo 'Diferencia' calculado correctamente.")

            # --- CONVERTIR A DATAFRAME ---
            fields = ["CODIGO_CONSTRUCCION", "npn_22", "MAX_PLANTA_UBICACION", "NUMERO_PISOS", "Diferencia"]
            data = [list(row) for row in arcpy.da.SearchCursor(intersect_output, fields)]

            if not data:
                print("Advertencia: el DataFrame está vacío después de Intersect.")
                return None

            df = pd.DataFrame(data, columns=fields)
            return df

        except Exception as e:
            tkMessageBox.showerror("Error", u"No se pudo procesar: {}".format(str(e)))
            print("Error:", e)
            return None
    
    
    def extraer_letras_identificador(self, gdb_path):

        feature_class_name_unidad = "r_lc_unidadconstruccion" if self.tipo_area.get() == "Rural" else "u_lc_unidadconstruccion"
        feature_class_path_unidad = os.path.join(gdb_path, feature_class_name_unidad)

        try:
            arcpy.env.workspace = gdb_path

            campos = [f.name for f in arcpy.ListFields(feature_class_path_unidad)]
            if "IDENTIFICADOR" not in campos:
                tkMessageBox.showerror("Error", u"La capa '{}' no contiene el campo 'IDENTIFICADOR'.".format(feature_class_name_unidad))
                return

            if "LETRAS" not in campos:
                arcpy.AddField_management(feature_class_path_unidad, "LETRAS", "TEXT", "", "", 10)

            # Nueva expresión sin re
            expression = "extraer_letras(!IDENTIFICADOR!)"
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
                in_table=feature_class_path_unidad,
                field="LETRAS",
                expression=expression,
                expression_type="PYTHON_9.3",
                code_block=code_block
            )
            print("Campo 'LETRAS' calculado correctamente.")

        except Exception as e:
            tkMessageBox.showerror("Error", u"No se pudo calcular LETRAS: {}".format(str(e)))
            print("Error:", e)

    def validar_terreno_codigo_duplicado_ficha(self, gdb_path):
        
        # 1. Capa según tipo de área
        feature_class_name = "r_lc_terreno" if self.tipo_area.get() == "Rural" else "u_lc_terreno"
        fc = os.path.join(gdb_path, feature_class_name)
        arcpy.env.workspace = gdb_path

        if not arcpy.Exists(fc):
            raise Exception("La capa {} no existe en la GDB.".format(feature_class_name))

        # 2. Campos requeridos
        campos = [f.name for f in arcpy.ListFields(fc)]
        campos_dict = {f.upper(): f for f in campos}
        for req in ("TERRENO_CODIGO", "NROFICHA", "CP"):
            if req not in campos_dict:
                raise Exception("Falta columna requerida: '{}'".format(req))

        fld_codigo = campos_dict["TERRENO_CODIGO"]
        fld_ficha  = campos_dict["NROFICHA"]
        fld_cp     = campos_dict["CP"]

        # 3. Leer todos los registros
        registros = []
        cursor_fields = [fld_codigo, fld_ficha, fld_cp]
        with arcpy.da.SearchCursor(fc, cursor_fields) as cur:
            for row in cur:
                codigo = row[0]
                ficha  = row[1]
                cp_val = row[2]
                if not codigo or len(codigo) < 22:
                    continue
                condicion = codigo[21]
                prefijo   = codigo[:21]
                registros.append({
                    "codigo":     codigo,
                    "nroficha":   ficha,
                    "condicion":  condicion,
                    "prefijo_21": prefijo,
                    "cp":         cp_val
                })

        df = pd.DataFrame(registros)

        # 4. Separar informales (condicion = '2') y resto
        df_inf   = df[df["condicion"] == '2']
        df_otros = df[df["condicion"] != '2']

        # 5. Emparejar cada informalidad con su(s) terreno(s) formal(es)
        emparejamientos = []
        for idx, inf in df_inf.iterrows():
            coinc = df_otros[df_otros["prefijo_21"] == inf["prefijo_21"]]
            for jdx, form in coinc.iterrows():
                emparejamientos.append({
                    "informalidad":          inf["codigo"],
                    "NroFicha_informalidad": inf["nroficha"],
                    "terreno":               form["codigo"],
                    "NroFicha_terreno":      form["nroficha"],
                    "prefijo_21":            inf["prefijo_21"]
                })

        df_rep = pd.DataFrame(emparejamientos, columns=[
            "informalidad", "NroFicha_informalidad",
            "terreno", "NroFicha_terreno", "prefijo_21"
        ])

        if df_rep.empty:
            print("No hay informalidades emparejadas.")
            return df_rep

        # 6. Crear capas temporales
        layer_formal = "lyr_formal"
        where_f = "{0} = '0'".format(fld_cp)
        arcpy.MakeFeatureLayer_management(fc, layer_formal, where_f)

        layer_inf = "lyr_informal"
        cods = df_rep["informalidad"].unique().tolist()
        in_clause = ", ".join("'{0}'".format(c) for c in cods)
        where_i = "{0} IN ({1})".format(fld_codigo, in_clause)
        arcpy.MakeFeatureLayer_management(fc, layer_inf, where_i)

        # 7. Buffer negativo
        buf_fc = "in_memory/buffer_inf"
        arcpy.Buffer_analysis(layer_inf, buf_fc, "-0.5", dissolve_option="NONE")

        # 8. Spatial Join con formales
        sp_fc = "in_memory/sj_inf_formal"
        arcpy.SpatialJoin_analysis(
            buf_fc, layer_formal, sp_fc,
            join_operation="JOIN_ONE_TO_ONE",
            join_type="KEEP_COMMON"
        )

        # 9. Leer el resultado del Spatial Join
        fld_inf_buf  = fld_codigo
        fld_for_join = fld_codigo + "_1"
        fld_fic_join = fld_ficha  + "_1"

        mapping = {}
        with arcpy.da.SearchCursor(sp_fc, [fld_inf_buf, fld_for_join, fld_fic_join]) as cur2:
            for row2 in cur2:
                infc  = row2[0]
                formc = row2[1]
                formf = row2[2]
                mapping[infc] = (formc, formf)

        # 10. Añadir predio formal al reporte
        df_rep["predio_formal"]  = df_rep["informalidad"].map(lambda c: mapping.get(c, (None, None))[0])
        df_rep["NroFicha_formal"] = df_rep["informalidad"].map(lambda c: mapping.get(c, (None, None))[1])

        print("Informe de informalidad geográfica generado:")

         # 11. Renombrar y eliminar columnas según tu petición
        df_rep = df_rep.rename(columns={
            "terreno":                   "terreno_codigo_duplicado",
            "predio_formal":             "ubicacion_espacial_informalidad"
        })

        # Elimina la columna de prefijo de 21 dígitos
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

        campos_necesarios = ["TERRENO_CODIGO", "ETIQUETA"]
        fields = [f.name for f in arcpy.ListFields(feature_class_path)]
        for campo in campos_necesarios:
            if campo not in fields:
                raise Exception("Falta el campo requerido: {}".format(campo))

        errores = []

        with arcpy.da.SearchCursor(feature_class_path, ["TERRENO_CODIGO", "ETIQUETA"]) as cursor:
            for row in cursor:
                terreno_codigo = row[0]
                etiqueta = str(row[1]).strip() if row[1] else ""

                if not terreno_codigo or len(terreno_codigo) < 30:
                    continue

                condicion_predio = terreno_codigo[21]
                unidad_terreno = terreno_codigo[26:30]
                print(terreno_codigo)
                print(unidad_terreno)

                if condicion_predio == '8' and unidad_terreno not in ('0000'):
                    etiqueta_cruda = terreno_codigo[26:30]
                    if etiqueta_cruda.startswith('P'):
                        etiqueta_esperada = etiqueta_cruda  # Se deja tal cual
                    else:
                        etiqueta_esperada = etiqueta_cruda.lstrip('0')  # Se eliminan ceros a la izquierda
                else:
                    etiqueta_esperada = terreno_codigo[17:21]
                    if not etiqueta_esperada.startswith('P'):
                        etiqueta_esperada = etiqueta_esperada.lstrip('0')

                if etiqueta != etiqueta_esperada:
                    errores.append([
                        terreno_codigo,
                        etiqueta,
                        etiqueta_esperada,
                        condicion_predio
                    ])

        columnas_ordenadas = ["TERRENO_CODIGO", "ETIQUETA", "ETIQUETA_ESPERADA", "CONDICION_PREDIO"]
        return pd.DataFrame(errores, columns=columnas_ordenadas)
    
    
    def comparar_areas_por_unidad(self, separador='', solo_diferentes=True):

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
        print("comparar_areas_por_unidad")
        def letras_a_num(col_letras):
            """
            Convierte 'A'->'1', 'B'->'2', ..., 'Z'->'26', 'AA'->'27', etc.
            Devuelve string.
            Si encuentra algo que no son solo letras, devuelve tal cual.
            """
            if not col_letras:
                return u""
            col_up = col_letras.upper()
            # Debe ser solo letras
            if not re.match(r'^[A-Z]+$', col_up):
                return col_letras
            total = 0
            for ch in col_up:
                total = total * 26 + (ord(ch) - ord('A') + 1)
            return unicode(total)

        def extraer_letra_identificador(id_raw):
            """
            De 'P1B1' -> 'B'
            De 'P12C3' -> 'C'
            De 'B2' -> 'B'
            De 'U10AA7' -> 'AA'
            Regla: tomar el bloque de letras que está entre números,
            o el primer bloque de letras si no hay números claros.
            """
            v = _s(id_raw)
            if not v:
                return v

            # 1) Buscar patrón típico piso + letras + resto numérico, ej 'P1B1', 'P12AA7'
            m = re.search(r'\d+([A-Za-z]+)\d+', v)
            if m:
                return m.group(1)

            # 2) Si no, buscar cualquier bloque de letras en el identificador
            m2 = re.search(r'([A-Za-z]+)', v)
            if m2:
                return m2.group(1)

            # 3) Si no encontramos letras, devolvemos vacío
            return u""

        def normalizar_identificador_a_indice(id_raw):
            """
            Toma IDENTIFICADOR completo ('P1B1') y devuelve solo el índice numérico que
            debe compararse con NumeroConstruccion del Excel.
            Ej:
              'P1B1' -> '2'
              'P1A1' -> '1'
              'P1AA1' -> '27'
            Si no puede calcularlo, devuelve el original como string.
            """
            letras = extraer_letra_identificador(id_raw)  # 'B'
            if letras:
                return letras_a_num(letras)               # '2'
            # fallback: si no hay letras, intentemos número puro
            v = _s(id_raw)
            mnum = re.search(r'(\d+)$', v)
            if mnum:
                return unicode(int(mnum.group(1)))
            return v

        gdb_path = self.gdb_path.get()
        excel_path = self.excel_path.get()

        if not excel_path:
            tkMessageBox.showerror("Error", u"Debe seleccionar el Excel.")
            return
        if not gdb_path:
            tkMessageBox.showerror("Error", u"Debe seleccionar la GDB o carpeta contenedora.")
            return

        # Si el usuario apuntó a carpeta en vez de .gdb/.sde, toma la primera .gdb
        if not gdb_path.lower().endswith(".gdb") and not gdb_path.lower().endswith(".sde"):
            gdbs = [f for f in os.listdir(gdb_path) if f.lower().endswith(".gdb")]
            if gdbs:
                gdb_path = os.path.join(gdb_path, gdbs[0])
            else:
                tkMessageBox.showerror("Error", u"No se encontró ninguna GDB.")
                return

        arcpy.env.workspace = gdb_path
        arcpy.env.overwriteOutput = True

        # -------- GDB (AGRUPO AQUÍ) --------
        try:
            df_uc = extraer_tabla_de_gdb_area_construida(gdb_path, self.tipo_area.get())
        except Exception as e:
            tkMessageBox.showerror("Error", u"No se pudo leer la tabla de la GDB:\n{}".format(e))
            return

        df_uc.columns = [c.strip() for c in df_uc.columns]

        # usamos CODIGO_UNIDAD_CONSTRUCCION (confirmado)
        req_gdb = ['CODIGO_UNIDAD_CONSTRUCCION', 'IDENTIFICADOR', 'SHAPE_Area']
        faltan_gdb = [c for c in req_gdb if c not in df_uc.columns]
        if faltan_gdb:
            tkMessageBox.showerror("Error", u"En la tabla GDB faltan: {}".format(", ".join(faltan_gdb)))
            return

        # Paso clave:
        # IDENTIFICADOR -> extraer letra(s) -> convertir letra(s) a número tipo 2,3,4...
        df_uc['IDX_UNIDAD'] = df_uc['IDENTIFICADOR'].apply(normalizar_identificador_a_indice)

        # CLAVE GDB = CODIGO_UNIDAD_CONSTRUCCION + SEP + IDX_UNIDAD
        df_uc['CLAVE'] = df_uc.apply(
            lambda r: u"{}{}{}".format(
                _s(r['CODIGO_UNIDAD_CONSTRUCCION']),
                SEP,
                _s(r['IDX_UNIDAD'])
            ),
            axis=1
        )

        df_uc['Area_GDB'] = pd.to_numeric(df_uc['SHAPE_Area'], errors='coerce').round(2)

        # Agrupar áreas de la GDB por esa clave
        df_gdb = df_uc.groupby('CLAVE', as_index=False)['Area_GDB'].sum()

        # -------- EXCEL (SIN AGRUPAR) --------
        try:
            xl = pd.ExcelFile(excel_path)
        except Exception as e:
            tkMessageBox.showerror("Error", u"No se pudo abrir el Excel:\n{}".format(e))
            return

        def _find_sheet(xlfile, snippet):
            lst = [h for h in xlfile.sheet_names if snippet in h.strip().lower()]
            return lst[0] if lst else None

        hoja_cons   = _find_sheet(xl, 'construcciones')
        hoja_fichas = _find_sheet(xl, 'fichas')

        if not hoja_cons:
            tkMessageBox.showerror("Error", u"No se encontró la hoja 'Construcciones'.")
            return

        df_cons = pd.read_excel(excel_path, sheetname=hoja_cons)
        df_cons.columns = [c.strip() for c in df_cons.columns]

        req_cons_base = ['NumeroConstruccion', 'AreaConstruida']
        faltan_cons_base = [c for c in req_cons_base if c not in df_cons.columns]
        if faltan_cons_base:
            tkMessageBox.showerror("Error", u"En 'Construcciones' faltan: {}".format(", ".join(faltan_cons_base)))
            return

        # Recuperar Npn si hace falta
        if 'Npn' in df_cons.columns:
            df_src = df_cons.copy()
        else:
            if not hoja_fichas:
                tkMessageBox.showerror("Error", u"No se encontró 'Npn' en 'Construcciones' y no hay hoja 'Fichas' para recuperarlo.")
                return

            df_fich = pd.read_excel(excel_path, sheetname=hoja_fichas)
            df_fich.columns = [c.strip() for c in df_fich.columns]

            if 'Npn' not in df_fich.columns or 'NroFicha' not in df_fich.columns or 'NroFicha' not in df_cons.columns:
                tkMessageBox.showerror("Error", u"No se pudo recuperar 'Npn': verifique 'NroFicha' y 'Npn' en 'Fichas'.")
                return

            df_src = pd.merge(
                df_cons,
                df_fich[['NroFicha', 'Npn']],
                on='NroFicha',
                how='left'
            )

        # En Excel:
        # - Npn = mismo código largo
        # - NumeroConstruccion = 1,2,3,...
        df_src['Npn'] = df_src['Npn'].map(_s)
        df_src['NumeroConstruccion'] = df_src['NumeroConstruccion'].map(_s)

        # CLAVE Excel = Npn + SEP + NumeroConstruccion
        df_src['CLAVE'] = df_src.apply(
            lambda r: u"{}{}{}".format(
                _s(r['Npn']),
                SEP,
                _s(r['NumeroConstruccion'])
            ),
            axis=1
        )

        df_src['Area_Excel'] = pd.to_numeric(df_src['AreaConstruida'], errors='coerce').round(2)
        
        # -------- Comparación (outer join) --------
        df_out = pd.merge(
            df_gdb,
            df_src[['CLAVE','Area_Excel']],
            on='CLAVE',
            how='outer'
        )

        # Calcular diferencia de áreas
        df_out['Area_GDB']   = pd.to_numeric(df_out.get('Area_GDB'), errors='coerce')
        df_out['Area_Excel'] = pd.to_numeric(df_out.get('Area_Excel'), errors='coerce')
        df_out['Diferencia'] = (df_out['Area_GDB'] - df_out['Area_Excel']).abs()

        # Redondeo a 2 decimales
        for col in ['Area_GDB', 'Area_Excel', 'Diferencia']:
            if col in df_out.columns:
                df_out[col] = df_out[col].round(2)

        for col in ['Area_GDB', 'Area_Excel', 'Diferencia']:
            if col in df_out.columns:
                df_out[col] = df_out[col].fillna('').replace('nan', '')
        # Si solo_diferentes=True, filtramos para reporte
        umbral = 1.0  # m²

        # Si solo_diferentes=True, filtramos para reporte
        if solo_diferentes:
            # Diferencias estrictamente mayores al umbral
            mask_dif = df_out['Diferencia'].notnull() & (df_out['Diferencia'] > umbral)
            # Registros sin cruce (faltó en GDB o en Excel)
            mask_nocruce = df_out['Diferencia'].isnull()
            df_out = df_out[mask_dif | mask_nocruce]

        df_out['NPN_NRO_CONSTRUCCION'] = df_out['CLAVE']

        return df_out[['NPN_NRO_CONSTRUCCION', 'Area_GDB', 'Area_Excel', 'Diferencia']]
    
    def validar_altura_y_anio_unidad(self, gdb_path):

        import arcpy
        import os
        from datetime import datetime
        import sys

        try:
            reload(sys)
            sys.setdefaultencoding('utf-8')
        except:
            pass

        errores = []
        anio_actual = datetime.now().year

        feature_class_name = "r_lc_unidadconstruccion" if self.tipo_area.get() == "Rural" else "u_lc_unidadconstruccion"
        fc = os.path.join(gdb_path, feature_class_name)

        print("validar_altura_y_anio_unidad")

        if not arcpy.Exists(fc):
            errores.append({
                u'FeatureClass': feature_class_name,
                u'OID': u'',
                u'Campo': u'',
                u'Valor': u'',
                u'Descripcion': u'La capa no existe en la GDB'
            })
            return errores

        campos = [f.name for f in arcpy.ListFields(fc)]

        requeridos = ["ALTURA", "ANIO_CONSTRUCCION"]
        faltantes = [c for c in requeridos if c not in campos]

        if faltantes:
            errores.append({
                u'FeatureClass': feature_class_name,
                u'OID': u'',
                u'Campo': u', '.join(faltantes),
                u'Valor': u'',
                u'Descripcion': u'Campo(s) no existe(n) en la capa'
            })
            return errores

        with arcpy.da.SearchCursor(fc, ["CODIGO_UNIDAD_CONSTRUCCION", "ALTURA", "ANIO_CONSTRUCCION"]) as cursor:
            for row in cursor:

                codigo_unidad = row[0]
                altura = row[1]
                anio = row[2]

                if altura is None:
                    errores.append({
                        u'FeatureClass': feature_class_name,
                        u'CODIGO_UNIDAD_CONSTRUCCION':(codigo_unidad),
                        u'Campo': u'ALTURA',
                        u'Valor': u'',
                        u'Descripcion': u'ALTURA es nula'
                    })
                else:
                    try:
                        altura_val = float(altura)
                        if altura_val <= 0:
                            errores.append({
                                u'FeatureClass': feature_class_name,
                                u'CODIGO_UNIDAD_CONSTRUCCION': codigo_unidad,
                                u'Campo': u'ALTURA',
                                u'Valor': unicode(altura),
                                u'Descripcion': u'ALTURA debe ser mayor a 0'
                            })
                    except:
                        errores.append({
                            u'FeatureClass': feature_class_name,
                            u'CODIGO_UNIDAD_CONSTRUCCION': codigo_unidad,
                            u'Campo': u'ALTURA',
                            u'Valor': unicode(altura),
                            u'Descripcion': u'ALTURA no es numérica'
                        })

                # =============================
                # VALIDAR ANIO_CONSTRUCCION
                # =============================
                if anio is None:
                    errores.append({
                        u'FeatureClass': feature_class_name,
                        u'CODIGO_UNIDAD_CONSTRUCCION': codigo_unidad,
                        u'Campo': u'ANIO_CONSTRUCCION',
                        u'Valor': u'',
                        u'Descripcion': u'ANIO_CONSTRUCCION es nulo'
                    })
                    continue

                try:
                    anio_val = int(anio)
                except:
                    errores.append({
                        u'FeatureClass': feature_class_name,
                        u'CODIGO_UNIDAD_CONSTRUCCION': codigo_unidad,
                        u'Campo': u'ANIO_CONSTRUCCION',
                        u'Valor': unicode(anio),
                        u'Descripcion': u'ANIO_CONSTRUCCION no es numérico'
                    })
                    continue

                if len(str(anio_val)) != 4:
                    errores.append({
                        u'FeatureClass': feature_class_name,
                        u'CODIGO_UNIDAD_CONSTRUCCION': codigo_unidad,
                        u'Campo': u'ANIO_CONSTRUCCION',
                        u'Valor': unicode(anio_val),
                        u'Descripcion': u'ANIO_CONSTRUCCION no tiene 4 dígitos'
                    })
                    continue

                if anio_val < 1800 or anio_val > anio_actual:
                    errores.append({
                        u'FeatureClass': feature_class_name,
                        u'CODIGO_UNIDAD_CONSTRUCCION': codigo_unidad,
                        u'Campo': u'ANIO_CONSTRUCCION',
                        u'Valor': unicode(anio_val),
                        u'Descripcion': u'ANIO_CONSTRUCCION fuera de rango lógico'
                    })

        return pd.DataFrame(errores)
    
    def verificar_geometrias_vacias(self,ruta_fc):
        """
        Verifica si hay geometrías vacías (NULL) en una capa.
        """
        contador_vacias = 0

        with arcpy.da.SearchCursor(ruta_fc, ["OID@", "SHAPE@"]) as cursor:
            for oid, geom in cursor:
                if geom is None or geom.isMultipart and not geom.partCount:
                    print("Geometría vacía encontrada en el OID:", oid)
                    contador_vacias += 1

        if contador_vacias == 0:
            print("No hay geometrias vacias.")
        else:
            print("Total de geometrias vacias:", contador_vacias)

    def reporte(self, workbook, reportes_dict, gdb_path):
        import os
        import arcpy
        import xlwt

        sheet_reporte = workbook.add_sheet('Reporte')

        style_porcentaje = xlwt.XFStyle()
        style_porcentaje.num_format_str = '0.0%'

        feature_class_name = "r_lc_terreno" if self.tipo_area.get() == "Rural" else "u_lc_terreno"
        feature_class_path = os.path.join(gdb_path, feature_class_name)

        try:
            total_fichas = int(arcpy.GetCount_management(feature_class_path)[0])
        except Exception as e:
            print("Error al contar registros de la capa:", e)
            total_fichas = 0

        # Encabezados
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

        # Evaluación final (solo para mostrar texto, la fórmula se hace en Excel)
        porcentaje_total = (float(total_cantidad) / total_fichas) * 100 if total_fichas else 0
        porcentaje_aprob_total = 100 - porcentaje_total

        if porcentaje_aprob_total <= 50:
            evaluacion = 'NO CUMPLE'
        elif porcentaje_aprob_total <= 87.5:
            evaluacion = 'CUMPLE PARCIAL'
        else:
            evaluacion = 'CUMPLE'

        # Celda donde vamos a escribir la evaluación total
        sheet_reporte.write(row, 0, 'TOTAL - ' + evaluacion)

        # Rango de filas con datos
        inicio = 1
        fin = row - 1  # fila anterior a la de total

        # Sumar columna B (Cantidad total)
        sheet_reporte.write(row, 1, xlwt.Formula("SUM(B{}:B{})".format(inicio + 1, fin + 1)))

        # Escribimos el total_fichas en una celda oculta para usar en fórmula
        sheet_reporte.write(0, 25, total_fichas)  # Columna Z (índice 25), fila 0
        celda_total_fichas = "$Z$1"  # Referencia absoluta

        # Fórmula en columna C (porcentaje de error): =B(row+1)/Z1
        sheet_reporte.write(row, 2, xlwt.Formula("B{}/{}".format(row + 1, celda_total_fichas)), style_porcentaje)

        # Fórmula en columna D (porcentaje aprobación): =1 - C(row+1)
        sheet_reporte.write(row, 3, xlwt.Formula("1 - C{}".format(row + 1)), style_porcentaje)

        # Escribimos la evaluación textual también
        sheet_reporte.write(row, 4, evaluacion)
        row += 1

        # Escribir errores críticos (si hay)
        if errores_criticos_presentes:
            errores_texto = ', '.join(errores_criticos_presentes)
            sheet_reporte.write(row, 0, u'ERRORES CRÍTICOS: ' + errores_texto)
    
    def select_gdb(self):
        path = tkFileDialog.askdirectory(title="Seleccionar Geodatabase (GDB)")
        if path:
            self.gdb_path.set(path)

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
            filetypes=[("Excel files", "*.xls;*.xls")]
        )
        
        if path:  # Solo actualizar si el usuario seleccionó un archivo
            self.output_excel.set(path)
        else:
            tkMessageBox.showerror("Error", "Debe seleccionar una ruta de salida válida.")


        

if __name__ == "__main__":
    root = tk.Tk()
    app = GDBExcelValidator(root)
    root.mainloop()