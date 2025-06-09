# -- coding: utf-8 --
import arcpy
import Tkinter as tk
import tkFileDialog
import tkMessageBox
from tkinter import ttk
import os
import pandas as pd


class GDBApp:
    def __init__(self, root):

        self.root = root

        self.root.title("Agregar Campos a Feature Classes")
        self.root.geometry("800x600")

        # Variable para almacenar la ruta de la GDB
        self.gdb_path = tk.StringVar()

        # Variable para almacenar la ruta del archivo Excel
        self.excel_file_path = tk.StringVar()
        self.excel_file_rph_path = tk.StringVar()

        # Configurar la ventana
        window_width = 800
        window_height = 600
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        root.geometry("%dx%d+%d+%d" % (window_width, window_height, x, y))

        # Botón para seleccionar la GDB
        tk.Button(root, text="Seleccionar GDB", command=self.select_gdb).grid(row=2, column=0, pady=10)

        # Mostrar la ruta de la GDB seleccionada
        tk.Label(root, text="GDB seleccionada:").grid(row=2, column=1, pady=10)
        tk.Label(root, textvariable=self.gdb_path).grid(row=2, column=2, pady=10)

        # Botón para seleccionar y cargar el archivo de Excel
        tk.Button(root, text="Seleccionar Excel", command=self.select_excel).grid(row=3, column=0, pady=10)


        # Mostrar la ruta del archivo Excel seleccionado
        tk.Label(root, text="Archivo Excel seleccionado:").grid(row=3, column=1, pady=10)
        tk.Label(root, textvariable=self.excel_file_path).grid(row=3, column=2, pady=10)


        # Botón para ejecutar
        tk.Button(root, text="Ejecutar", command=self.process_all).grid(row=5, column=0, pady=10)

    def select_gdb(self):
        """Abrir un cuadro de diálogo para seleccionar una geodatabase."""
        gdb = tkFileDialog.askdirectory(title="Seleccionar GDB")
        if gdb and os.path.isdir(gdb) and gdb.endswith(".gdb"):
            self.gdb_path.set(gdb)
        else:
            tkMessageBox.showerror("Error", "Seleccione una geodatabase válida")

    def select_excel(self):
        """Seleccionar el archivo Excel para su posterior procesamiento."""
        excel_file = tkFileDialog.askopenfilename(title="Seleccionar archivo Excel",
                                                  filetypes=[("Archivos Excel", "*.xlsx")])
        if excel_file:
            self.excel_file_path.set(excel_file)
        else:
            tkMessageBox.showerror("Error", "Seleccione un archivo Excel válido.")



    def agregar_ficha_terreno(self):
        """Importar un archivo Excel y realizar la comparación con las feature classes r_lc_terreno y u_lc_terreno."""
        try:
            # Verificar si la geodatabase ha sido seleccionada antes de importar Excel
            gdb = self.gdb_path.get()
            if not gdb:
                tkMessageBox.showerror("Error", "Seleccione una geodatabase primero.")
                return

            # Seleccionar el archivo Excel
            excel_file = self.excel_file_path.get()
            if not excel_file:
                tkMessageBox.showerror("Error", "Seleccione un archivo Excel válido.")
                return

            if not excel_file:
                tkMessageBox.showerror("Error", "Seleccione un archivo Excel válido.")
                return

            # Leer el Excel, específicamente el libro "Fichas"
            df = pd.read_excel(excel_file, sheetname='Fichas')

            # Verificar que el archivo contiene las columnas necesarias
            if 'Npn' not in df.columns or 'NroFicha' not in df.columns:
                tkMessageBox.showerror("Error", "El archivo Excel no contiene las columnas 'Npn' y 'NroFicha'.")
                return

            # Crear un diccionario del Excel con Npn como clave y Nficha como valor
            npn_to_ficha = dict(zip(df['Npn'].astype(str), df['NroFicha'].astype(str)))

            # Lista de feature classes a actualizar
            feature_classes = ['r_lc_terreno', 'u_lc_terreno']

            for fc_name in feature_classes:
                # Definir la ruta del feature class
                fc_path = os.path.join(gdb, fc_name)
                if not arcpy.Exists(fc_path):
                    tkMessageBox.showerror("Error", "El feature class"+fc_name+ "no existe en la geodatabase.")
                    continue  # Saltar a la siguiente feature class si no existe

                # Iniciar una sesión de edición
                with arcpy.da.Editor(gdb) as edit_session:
                    # Iterar sobre las filas del feature class
                    with arcpy.da.UpdateCursor(fc_path, ['terreno_codigo', 'NroFicha']) as cursor:
                        for row in cursor:
                            terreno_codigo = str(row[0])  # terreno_codigo es el primer campo
                            if terreno_codigo in npn_to_ficha:
                                row[1] = npn_to_ficha[terreno_codigo]  # Asignar Nficha cuando coincida
                                cursor.updateRow(row)
                                print("Asignado Nficha a terreno_codigo"+ terreno_codigo+ "en "+fc_name)

            #tkMessageBox.showinfo("Éxito", "Datos del archivo Excel importados y actualizados correctamente.")

        except Exception as e:
            tkMessageBox.showerror("Error", "Error al importar o procesar el archivo Excel: "+str(e))

    def agregar_matricula_terreno(self):
        """Importar un archivo Excel y realizar la comparación con las feature classes r_lc_terreno y u_lc_terreno."""
        try:
            # Verificar si la geodatabase ha sido seleccionada antes de importar Excel
            gdb = self.gdb_path.get()
            if not gdb:
                tkMessageBox.showerror("Error", "Seleccione una geodatabase primero.")
                return

            # Seleccionar el archivo Excel
            excel_file = self.excel_file_path.get()
            if not excel_file:
                tkMessageBox.showerror("Error", "Seleccione un archivo Excel válido.")
                return

            # Leer el Excel, específicamente el libro "Fichas" asegurando que MatriculaInmobiliaria sea texto
            df = pd.read_excel(excel_file, sheet_name='Fichas')

            # Verificar que el archivo contiene las columnas necesarias
            if 'Npn' not in df.columns or 'MatriculaInmobiliaria' not in df.columns:
                tkMessageBox.showerror("Error",
                                       "El archivo Excel no contiene las columnas 'Npn' y 'MatriculaInmobiliaria'.")
                return

            # Crear un diccionario del Excel con Npn como clave y MatriculaInmobiliaria como valor
            # Asegurarse de que ambos campos sean tratados como texto
            npn_to_matricula = dict(zip(df['Npn'].astype(str), df['MatriculaInmobiliaria'].astype(str)))

            # Lista de feature classes a actualizar
            feature_classes = ['r_lc_terreno', 'u_lc_terreno']

            for fc_name in feature_classes:
                # Definir la ruta del feature class
                fc_path = os.path.join(gdb, fc_name)
                if not arcpy.Exists(fc_path):
                    tkMessageBox.showerror("Error", "El feature class " + fc_name + " no existe en la geodatabase.")
                    continue  # Saltar a la siguiente feature class si no existe

                # Verificar si el campo MatriculaInmobiliaria existe, si no, crearlo como texto
                fields = [f.name for f in arcpy.ListFields(fc_path)]
                if 'MatriculaInmobiliaria' not in fields:
                    arcpy.AddField_management(fc_path, 'MatriculaInmobiliaria', 'TEXT', field_length=50)

                # Iniciar una sesión de edición
                with arcpy.da.Editor(gdb) as edit_session:
                    # Iterar sobre las filas del feature class
                    with arcpy.da.UpdateCursor(fc_path, ['terreno_codigo', 'MatriculaInmobiliaria']) as cursor:
                        for row in cursor:
                            terreno_codigo = str(row[0])  # terreno_codigo es el primer campo
                            if terreno_codigo in npn_to_matricula:
                                # Asignar MatriculaInmobiliaria como texto
                                matricula = str(npn_to_matricula[terreno_codigo])
                                # Eliminar puntos decimales si existen (ej: "123.0" -> "123")
                                if '.' in matricula:
                                    matricula = matricula.split('.')[0]
                                row[1] = matricula
                                cursor.updateRow(row)
                                print(
                                            "Asignado MatriculaInmobiliaria a terreno_codigo " + terreno_codigo + " en " + fc_name)

            #tkMessageBox.showinfo("Éxito", "Datos del archivo Excel importados y actualizados correctamente.")

        except Exception as e:
            tkMessageBox.showerror("Error", "Error al importar o procesar el archivo Excel: " + str(e))

    def agregar_prediolc_terreno(self):
        """Importar un archivo Excel y realizar la comparación con las feature classes r_lc_terreno y u_lc_terreno."""
        try:
            # Verificar si la geodatabase ha sido seleccionada antes de importar Excel
            gdb = self.gdb_path.get()
            if not gdb:
                tkMessageBox.showerror("Error", "Seleccione una geodatabase primero.")
                return

            # Seleccionar el archivo Excel
            excel_file = self.excel_file_path.get()
            if not excel_file:
                tkMessageBox.showerror("Error", "Seleccione un archivo Excel válido.")
                return

            # Leer el Excel, específicamente el libro "Fichas"
            df = pd.read_excel(excel_file, sheet_name='Fichas')

            # Verificar que el archivo contiene las columnas necesarias
            if 'Npn' not in df.columns or 'PredioLcTipo' not in df.columns:
                tkMessageBox.showerror("Error",
                                       "El archivo Excel no contiene las columnas 'Npn' y 'PredioLcTipo'.")
                return

            # Crear un diccionario del Excel con Npn como clave y PredioLcTipo como valor
            npn_to_prediolc = dict(zip(df['Npn'].astype(str), df['PredioLcTipo'].astype(str)))

            # Lista de feature classes a actualizar
            feature_classes = ['r_lc_terreno', 'u_lc_terreno']

            for fc_name in feature_classes:
                # Definir la ruta del feature class
                fc_path = os.path.join(gdb, fc_name)
                if not arcpy.Exists(fc_path):
                    tkMessageBox.showerror("Error", "El feature class " + fc_name + " no existe en la geodatabase.")
                    continue  # Saltar a la siguiente feature class si no existe

                # Iniciar una sesión de edición
                with arcpy.da.Editor(gdb) as edit_session:
                    # Iterar sobre las filas del feature class
                    with arcpy.da.UpdateCursor(fc_path, ['terreno_codigo', 'PredioLcTipo']) as cursor:
                        for row in cursor:
                            terreno_codigo = str(row[0])  # terreno_codigo es el primer campo
                            if terreno_codigo in npn_to_prediolc:
                                row[1] = npn_to_prediolc[
                                    terreno_codigo]  # Asignar MatriculaInmobiliaria cuando coincida
                                cursor.updateRow(row)
                                print(
                                            "Asignado PredioLcTipo a terreno_codigo " + terreno_codigo + " en " + fc_name)

            #tkMessageBox.showinfo("Éxito", "Datos del archivo Excel importados y actualizados correctamente.")

        except Exception as e:
            tkMessageBox.showerror("Error", "Error al importar o procesar el archivo Excel: " + str(e))

    def agregar_adquisicion_terreno(self):
        """Importar un archivo Excel y realizar la comparación con las feature classes r_lc_terreno y u_lc_terreno."""
        try:
            # Verificar si la geodatabase ha sido seleccionada antes de importar Excel
            gdb = self.gdb_path.get()
            if not gdb:
                tkMessageBox.showerror("Error", "Seleccione una geodatabase primero.")
                return

            # Seleccionar el archivo Excel
            excel_file = self.excel_file_path.get()
            if not excel_file:
                tkMessageBox.showerror("Error", "Seleccione un archivo Excel válido.")
                return

            # Leer el Excel, específicamente el libro "Fichas"
            df = pd.read_excel(excel_file, sheet_name='Fichas',encoding='utf-8')
            df['ModoAdquisicion'] = df['ModoAdquisicion'].str.encode('ascii', 'ignore').str.decode('ascii')


            # Verificar que el archivo contiene las columnas necesarias
            if 'Npn' not in df.columns or 'ModoAdquisicion' not in df.columns:
                tkMessageBox.showerror("Error",
                                       "El archivo Excel no contiene las columnas 'Npn' y 'PredioLcTipo'.")
                return

            # Crear un diccionario del Excel con Npn como clave y PredioLcTipo como valor
            npn_to_adquisicion = dict(zip(df['Npn'].astype(str), df['ModoAdquisicion'].astype(str)))

            # Lista de feature classes a actualizar
            feature_classes = ['r_lc_terreno', 'u_lc_terreno']

            for fc_name in feature_classes:
                # Definir la ruta del feature class
                fc_path = os.path.join(gdb, fc_name)
                if not arcpy.Exists(fc_path):
                    tkMessageBox.showerror("Error", "El feature class " + fc_name + " no existe en la geodatabase.")
                    continue  # Saltar a la siguiente feature class si no existe

                # Iniciar una sesión de edición
                with arcpy.da.Editor(gdb) as edit_session:
                    # Iterar sobre las filas del feature class
                    with arcpy.da.UpdateCursor(fc_path, ['terreno_codigo', 'ModoAdquisicion']) as cursor:
                        for row in cursor:
                            terreno_codigo = str(row[0])  # terreno_codigo es el primer campo
                            if terreno_codigo in npn_to_adquisicion:
                                row[1] = npn_to_adquisicion[
                                    terreno_codigo]  # Asignar MatriculaInmobiliaria cuando coincida
                                cursor.updateRow(row)
                                print(
                                            "Asignado PredioLcTipo a terreno_codigo " + terreno_codigo + " en " + fc_name)

            #tkMessageBox.showinfo("Éxito", "Datos del archivo Excel importados y actualizados correctamente.")

        except Exception as e:
            tkMessageBox.showerror("Error", "Error al importar o procesar el archivo Excel: " + str(e))

    def agregar_ficha_unidad(self):
        """Importar el segundo archivo Excel y realizar la comparación con las feature classes r_lc_unidad y u_lc_unidad."""
        try:
            # Verificar si la geodatabase ha sido seleccionada antes de importar Excel
            gdb = self.gdb_path.get()
            if not gdb:
                tkMessageBox.showerror("Error", "Seleccione una geodatabase primero.")
                return

            # Seleccionar el segundo archivo Excel (unidad)
            excel_file_rph = self.excel_file_path.get()
            if not excel_file_rph:
                tkMessageBox.showerror("Error", "Seleccione un archivo Excel válido.")
                return

            # Leer el Excel (RPH), específicamente el libro "Fichas_RPH"
            df_rph = pd.read_excel(excel_file_rph, sheetname='Fichas')

            # Verificar que el archivo contiene las columnas necesarias
            if 'Npn' not in df_rph.columns or 'NroFicha' not in df_rph.columns:
                tkMessageBox.showerror("Error", "El archivo Excel no contiene las columnas 'Npn' y 'NroFicha'.")
                return

            # Crear un diccionario del Excel (RPH) con Npn_rph como clave y NroFicha_rph como valor
            npn_to_ficha_rph = dict(zip(df_rph['Npn'].astype(str), df_rph['NroFicha'].astype(str)))

            # Lista de feature classes a actualizar (RPH)
            feature_classes_rph = ['u_lc_unidadconstruccion', 'r_lc_unidadconstruccion']

            for fc_name in feature_classes_rph:
                # Definir la ruta del feature class
                fc_path = os.path.join(gdb, fc_name)
                if not arcpy.Exists(fc_path):
                    tkMessageBox.showerror("Error", "El feature class " + fc_name + " no existe en la geodatabase.")
                    continue  # Saltar a la siguiente feature class si no existe

                # Iniciar una sesión de edición
                with arcpy.da.Editor(gdb) as edit_session:
                    # Iterar sobre las filas del feature class
                    with arcpy.da.UpdateCursor(fc_path, ['codigo_unidad_construccion', 'NroFicha']) as cursor:
                        for row in cursor:
                            codigo_unidad_construccion = str(row[0])  # terreno_codigo es el primer campo
                            if codigo_unidad_construccion in npn_to_ficha_rph:
                                row[1] = npn_to_ficha_rph[codigo_unidad_construccion]  # Asignar NroFicha_rph cuando coincida
                                cursor.updateRow(row)
                                print("Asignado NroFicha a terreno_codigo"+ codigo_unidad_construccion+ "en "+ fc_name)

            #tkMessageBox.showinfo("Éxito", "Datos del archivo Excel (unidades) importados y actualizados correctamente.")

        except Exception as e:
            tkMessageBox.showerror("Error", "Error al importar o procesar el archivo Excel (RPH):"+ str(e))

    def agregar_matricula_unidad(self):
        """Importar un archivo Excel y realizar la comparación con las feature classes r_lc_terreno y u_lc_terreno."""
        try:
            # Verificar si la geodatabase ha sido seleccionada antes de importar Excel
            gdb = self.gdb_path.get()
            if not gdb:
                tkMessageBox.showerror("Error", "Seleccione una geodatabase primero.")
                return

            # Seleccionar el archivo Excel
            excel_file = self.excel_file_path.get()
            if not excel_file:
                tkMessageBox.showerror("Error", "Seleccione un archivo Excel válido.")
                return

            # Leer el Excel, específicamente el libro "Fichas" asegurando que MatriculaInmobiliaria sea texto
            df = pd.read_excel(excel_file, sheet_name='Fichas')

            # Verificar que el archivo contiene las columnas necesarias
            if 'Npn' not in df.columns or 'MatriculaInmobiliaria' not in df.columns:
                tkMessageBox.showerror("Error",
                                       "El archivo Excel no contiene las columnas 'Npn' y 'MatriculaInmobiliaria'.")
                return

            # Crear un diccionario del Excel con Npn como clave y MatriculaInmobiliaria como valor
            # Asegurarse de que ambos campos sean tratados como texto
            npn_to_matricula = dict(zip(df['Npn'].astype(str), df['MatriculaInmobiliaria'].astype(str)))

            # Lista de feature classes a actualizar
            feature_classes_rph = ['u_lc_unidadconstruccion', 'r_lc_unidadconstruccion']

            for fc_name in feature_classes_rph:
                # Definir la ruta del feature class
                fc_path = os.path.join(gdb, fc_name)
                if not arcpy.Exists(fc_path):
                    tkMessageBox.showerror("Error", "El feature class " + fc_name + " no existe en la geodatabase.")
                    continue  # Saltar a la siguiente feature class si no existe

                # Verificar si el campo MatriculaInmobiliaria existe, si no, crearlo como texto
                fields = [f.name for f in arcpy.ListFields(fc_path)]
                if 'MatriculaInmobiliaria' not in fields:
                    arcpy.AddField_management(fc_path, 'MatriculaInmobiliaria', 'TEXT', field_length=50)

                # Iniciar una sesión de edición
                with arcpy.da.Editor(gdb) as edit_session:
                    # Iterar sobre las filas del feature class
                    with arcpy.da.UpdateCursor(fc_path, ['codigo_unidad_construccion', 'MatriculaInmobiliaria']) as cursor:
                        for row in cursor:
                            terreno_codigo = str(row[0])  # terreno_codigo es el primer campo
                            if terreno_codigo in npn_to_matricula:
                                # Asignar MatriculaInmobiliaria como texto
                                matricula = str(npn_to_matricula[terreno_codigo])
                                # Eliminar puntos decimales si existen (ej: "123.0" -> "123")
                                if '.' in matricula:
                                    matricula = matricula.split('.')[0]
                                row[1] = matricula
                                cursor.updateRow(row)
                                print(
                                        "Asignado MatriculaInmobiliaria a codigo_unidad_construccion " + terreno_codigo + " en " + fc_name)

            #tkMessageBox.showinfo("Éxito", "Datos del archivo Excel importados y actualizados correctamente.")

        except Exception as e:
            tkMessageBox.showerror("Error", "Error al importar o procesar el archivo Excel: " + str(e))

    def agregar_prediolc_unidad(self):
        """Importar el segundo archivo Excel y realizar la comparación con las feature classes r_lc_unidad y u_lc_unidad."""
        try:
            # Verificar si la geodatabase ha sido seleccionada antes de importar Excel
            gdb = self.gdb_path.get()
            if not gdb:
                tkMessageBox.showerror("Error", "Seleccione una geodatabase primero.")
                return

            # Seleccionar el segundo archivo Excel (RPH)
            excel_file_rph = self.excel_file_path.get()
            if not excel_file_rph:
                tkMessageBox.showerror("Error", "Seleccione un archivo Excel válido.")
                return

            # Leer el Excel (RPH), específicamente el libro "Fichas_RPH"
            df_rph = pd.read_excel(excel_file_rph, sheetname='Fichas')

            # Verificar que el archivo contiene las columnas necesarias
            if 'Npn' not in df_rph.columns or 'PredioLcTipo' not in df_rph.columns:
                tkMessageBox.showerror("Error", "El archivo Excel no contiene las columnas 'Npn' y 'PredioLcTipo'.")
                return

            # Crear un diccionario del Excel (RPH) con Npn_rph como clave y NroFicha_rph como valor
            prediolc_to_ficha = dict(zip(df_rph['Npn'].astype(str), df_rph['PredioLcTipo'].astype(str)))

            # Lista de feature classes a actualizar (RPH)
            feature_classes_rph = ['u_lc_unidadconstruccion', 'r_lc_unidadconstruccion']

            for fc_name in feature_classes_rph:
                # Definir la ruta del feature class
                fc_path = os.path.join(gdb, fc_name)
                if not arcpy.Exists(fc_path):
                    tkMessageBox.showerror("Error", "El feature class " + fc_name + " no existe en la geodatabase.")
                    continue  # Saltar a la siguiente feature class si no existe

                # Iniciar una sesión de edición
                with arcpy.da.Editor(gdb) as edit_session:
                    # Iterar sobre las filas del feature class
                    with arcpy.da.UpdateCursor(fc_path, ['codigo_unidad_construccion', 'PredioLcTipo']) as cursor:
                        for row in cursor:
                            codigo_unidad_construccion = str(row[0])  # terreno_codigo es el primer campo
                            if codigo_unidad_construccion in prediolc_to_ficha:
                                row[1] = prediolc_to_ficha[codigo_unidad_construccion]  # Asignar NroFicha_rph cuando coincida
                                cursor.updateRow(row)
                                print("Asignado PredioLcTipo a unidad de construccion"+ codigo_unidad_construccion+ "en "+ fc_name)

            #tkMessageBox.showinfo("Éxito", "Datos del archivo Excel unidades importados y actualizados correctamente.")

        except Exception as e:
            tkMessageBox.showerror("Error", "Error al importar o procesar el archivo Excel (unidades):"+ str(e))

    def agregar_adquisicion_unidad(self):
        """Importar el segundo archivo Excel y realizar la comparación con las feature classes r_lc_unidad y u_lc_unidad."""
        try:
            # Verificar si la geodatabase ha sido seleccionada antes de importar Excel
            gdb = self.gdb_path.get()
            if not gdb:
                tkMessageBox.showerror("Error", "Seleccione una geodatabase primero.")
                return

            # Seleccionar el segundo archivo Excel (RPH)
            excel_file_rph = self.excel_file_path.get()
            if not excel_file_rph:
                tkMessageBox.showerror("Error", "Seleccione un archivo Excel válido.")
                return

            # Leer el Excel (RPH), específicamente el libro "Fichas_RPH"
            df_rph = pd.read_excel(excel_file_rph, sheetname='Fichas', encoding='utf-8')
            df_rph['ModoAdquisicion'] = df_rph['ModoAdquisicion'].str.encode('ascii', 'ignore').str.decode('ascii')
            # Verificar que el archivo contiene las columnas necesarias
            if 'Npn' not in df_rph.columns or 'ModoAdquisicion' not in df_rph.columns:
                tkMessageBox.showerror("Error", "El archivo Excel no contiene las columnas 'Npn' y 'ModoAdquisicion'.")
                return

            # Crear un diccionario del Excel (RPH) con Npn_rph como clave y NroFicha_rph como valor
            prediolc_to_ModoAdquisicion = dict(zip(df_rph['Npn'].astype(str), df_rph['ModoAdquisicion'].astype(str)))

            # Lista de feature classes a actualizar (RPH)
            feature_classes_rph = ['u_lc_unidadconstruccion', 'r_lc_unidadconstruccion']

            for fc_name in feature_classes_rph:
                # Definir la ruta del feature class
                fc_path = os.path.join(gdb, fc_name)
                if not arcpy.Exists(fc_path):
                    tkMessageBox.showerror("Error", "El feature class " + fc_name + " no existe en la geodatabase.")
                    continue  # Saltar a la siguiente feature class si no existe

                # Iniciar una sesión de edición
                with arcpy.da.Editor(gdb) as edit_session:
                    # Iterar sobre las filas del feature class
                    with arcpy.da.UpdateCursor(fc_path, ['codigo_unidad_construccion', 'ModoAdquisicion']) as cursor:
                        for row in cursor:
                            codigo_unidad_construccion = str(row[0])  # terreno_codigo es el primer campo
                            if codigo_unidad_construccion in prediolc_to_ModoAdquisicion:
                                row[1] = prediolc_to_ModoAdquisicion[codigo_unidad_construccion]  # Asignar NroFicha cuando coincida
                                cursor.updateRow(row)
                                print("Asignado PredioLcTipo a unidad de construccion"+ codigo_unidad_construccion+ "en "+ fc_name)

            #tkMessageBox.showinfo("Éxito", "Datos del archivo Excel unidades importados y actualizados correctamente.")

        except Exception as e:
            tkMessageBox.showerror("Error", "Error al importar o procesar el archivo Excel (unidades):"+ str(e))

    def agregar_propietario_unidad(self):
        """Importar el segundo archivo Excel y realizar la comparación con las feature classes r_lc_unidad y u_lc_unidad."""
        try:
            # Verificar si la geodatabase ha sido seleccionada antes de importar Excel
            gdb = self.gdb_path.get()
            if not gdb:
                tkMessageBox.showerror("Error", "Seleccione una geodatabase primero.")
                return

            # Seleccionar el segundo archivo Excel (RPH)
            excel_file_rph = self.excel_file_path.get()
            if not excel_file_rph:
                tkMessageBox.showerror("Error", "Seleccione un archivo Excel válido.")
                return

            # Leer el Excel (RPH), específicamente el libro "Fichas_RPH"
            df_rph = pd.read_excel(excel_file_rph, sheetname='Propietarios', encoding='utf-8')
            df_rph['RazonSocial'] = df_rph['RazonSocial'].str.encode('ascii', 'ignore').str.decode('ascii')
            # Verificar que el archivo contiene las columnas necesarias
            if 'NroFicha' not in df_rph.columns or 'RazonSocial' not in df_rph.columns:
                tkMessageBox.showerror("Error", "El archivo Excel no contiene las columnas 'Npn' y 'ModoAdquisicion'.")
                return

            # Crear un diccionario del Excel (RPH) con Npn_rph como clave y NroFicha_rph como valor
            prediolc_to_ModoAdquisicion = dict(zip(df_rph['NroFicha'].astype(str), df_rph['RazonSocial'].astype(str)))

            # Lista de feature classes a actualizar (RPH)
            feature_classes = ['u_lc_terreno', 'r_lc_terreno']

            for fc_name in feature_classes:
                # Definir la ruta del feature class
                fc_path = os.path.join(gdb, fc_name)
                if not arcpy.Exists(fc_path):
                    tkMessageBox.showerror("Error", "El feature class " + fc_name + " no existe en la geodatabase.")
                    continue  # Saltar a la siguiente feature class si no existe

                # Iniciar una sesión de edición
                with arcpy.da.Editor(gdb) as edit_session:
                    # Iterar sobre las filas del feature class
                    with arcpy.da.UpdateCursor(fc_path, ['NroFicha', 'RazonSocial']) as cursor:
                        for row in cursor:
                            NroFicha = str(row[0])  # terreno_codigo es el primer campo
                            if NroFicha in prediolc_to_ModoAdquisicion:
                                row[1] = prediolc_to_ModoAdquisicion[NroFicha]  # Asignar NroFicha cuando coincida
                                cursor.updateRow(row)
                                print("Asignado RazonSocial a terreno"+ NroFicha+ "en "+ fc_name)

            #tkMessageBox.showinfo("Éxito", "Datos del archivo Excel unidades importados y actualizados correctamente.")

        except Exception as e:
            tkMessageBox.showerror("Error", "Error al importar o procesar el archivo Excel (unidades):"+ str(e))

    def add_fields(self):
        """Agregar los campos a las feature classes especificadas."""
        try:
            gdb = self.gdb_path.get()
            print (gdb)
            if not gdb:
                tkMessageBox.showerror("Error", "Seleccione una geodatabase primero.")
                return

            # Lista de feature classes a las que se agregarán los campos comunes
            feature_classes = [
                "u_lc_terreno", "r_lc_terreno", "u_lc_construccion", "r_lc_construccion", "u_lc_unidadconstruccion",
                "r_lc_unidadconstruccion"
            ]

            # Campos comunes que se agregarán a todas las feature classes
            common_fields = [
                ('MatriculaInmobiliaria', 'Text', 10),
                ('NroFicha', 'Text', 10),
                ('NpnTerreno', 'Text', 30),
                ('CP', 'Text', 1),
                ('PredioLcTipo','Text',50),
                ('ModoAdquisicion','Text',50)
            ]

            # Campo Npnresumen que se agregará específicamente a dos feature classes

            construccion_fields = [
                ('FHNC', 'Text', 10),
                ]
            terreno_fields = [
                ('RazonSocial','Text',256),
                ('FACHADA','Text', 256),
                ('ACABADOS', 'Text',256),
                ('COCINA', 'Text',256),
                ('BANIO', 'Text',256),
                ('ANEXO_1', 'Text',256),
                ('ANEXO_2','Text',256),
                ('OBSERVACIONES','Text',256),
            ]

            arcpy.env.workspace = gdb

            for fc in feature_classes:
                fc_path = os.path.join(gdb, fc)

                if not arcpy.Exists(fc_path):
                    tkMessageBox.showerror("Error", "La feature class " + fc + " no existe en la geodatabase.")
                    continue

                # Agregar campos comunes
                for field_name, field_type, field_length in common_fields:
                    existing_fields = [f.name for f in arcpy.ListFields(fc_path)]
                    if field_name not in existing_fields:
                        arcpy.AddField_management(fc_path, field_name, field_type, field_length=field_length)
                        print("Campo " + str(field_name) + " agregado a: " + str(fc))
                    else:
                        print("El campo ya existe en " + str(fc) + ", no se agrega.")

                # Si la feature class es u_lc_unidadconstruccion o r_lc_unidadconstruccion, agregar Npnresumen

                if fc in ["u_lc_terreno", "r_lc_terreno"]:
                    arcpy.CalculateField_management(
                        in_table=fc_path,
                        field="NpnTerreno",
                        expression="!terreno_codigo![:22]",
                        expression_type="PYTHON"
                    )

                    arcpy.CalculateField_management(
                        in_table=fc_path,
                        field="CP",
                        expression="!terreno_codigo![21] if len(!terreno_codigo!) >= 22 else ''",
                        expression_type="PYTHON"
                    )
                    for field_name, field_type, field_length in terreno_fields:
                        existing_fields = [f.name for f in arcpy.ListFields(fc_path)]
                        if field_name not in existing_fields:
                            arcpy.AddField_management(fc_path, field_name, field_type, field_length=field_length)
                            print("Campo " + str(field_name) + " agregado a: " + str(fc))
                        else:
                            print("El campo ya existe en " + str(fc) + ", no se agrega.")

                if fc in ["u_lc_construccion", "r_lc_construccion"]:
                    arcpy.CalculateField_management(
                        in_table=fc_path,
                        field="NpnTerreno",
                        expression="!codigo_construccion![:20]",
                        expression_type="PYTHON"
                    )

                    arcpy.CalculateField_management(
                        in_table=fc_path,
                        field="CP",
                        expression="!codigo_construccion![21] if len(!codigo_construccion!) >= 22 else ''",
                        expression_type="PYTHON"
                    )

                if fc in ["u_lc_unidadconstruccion", "r_lc_unidadconstruccion"]:
                    arcpy.CalculateField_management(
                        in_table=fc_path,
                        field="NpnTerreno",
                        expression="!codigo_unidad_construccion![:20]",
                        expression_type="PYTHON"
                    )

                    arcpy.CalculateField_management(
                        in_table=fc_path,
                        field="CP",
                        expression="!codigo_unidad_construccion![21] if len(!codigo_unidad_construccion!) >= 22 else ''",
                        expression_type="PYTHON"
                    )



                if fc in ["r_lc_construccion", "u_lc_construccion"]:
                    for field_name, field_type, field_length in construccion_fields:
                        existing_fields = [f.name for f in arcpy.ListFields(fc_path)]
                        if field_name not in existing_fields:
                            arcpy.AddField_management(fc_path, field_name, field_type, field_length=field_length)
                            print("Campo " + str(field_name) + " agregado a: " + str(fc))
                        else:
                            print("El campo ya existe en " + str(fc) + ", no se agrega.")


            #tkMessageBox.showinfo("Éxito", "Campos agregados correctamente a las feature classes.")

        except Exception as e:
            tkMessageBox.showerror("Error", "Error al agregar campos: " + str(e))

    def import_to_gdb(self):
        """Importar las hojas del archivo Excel a la GDB como tablas."""

        try:
            # Obtener la ruta de la geodatabase y el archivo Excel
            gdb = self.gdb_path.get()
            excel_file = self.excel_file_path.get()

            # Verificar si las rutas son válidas
            if not os.path.exists(gdb) or not gdb.endswith('.gdb'):
                tkMessageBox.showerror("Error", "Seleccione una geodatabase válida.")
                return
            if not os.path.isfile(excel_file) or not excel_file.endswith('.xlsx'):
                tkMessageBox.showerror("Error", "Seleccione un archivo Excel válido.")
                return

            print(gdb)
            print(excel_file)

            # Definir el espacio de trabajo de ArcPy
            arcpy.env.workspace = gdb

            # Listado de las hojas a convertir en tablas
            sheets = [
                "Fichas",
                "Propietarios",
                "Construcciones",
                "CalificacionesConstrucciones",
                "ConstruccionesGenerales",
            ]

            # Convertir cada hoja a una tabla en la geodatabase
            for sheet in sheets:
                output_table = os.path.join(gdb, sheet)  # Definir el nombre de la tabla en la GDB
                print("Importando"+ sheet +" a "+output_table)

                arcpy.ExcelToTable_conversion(
                    Input_Excel_File=excel_file,
                    Output_Table=output_table,
                    Sheet=sheet
                )

            tkMessageBox.showinfo("Éxito", "Todas las hojas se han importado correctamente a la geodatabase.")

        except arcpy.ExecuteError as e:
            # Capturar errores de ArcPy
            tkMessageBox.showerror("Error de ArcPy", "Error al convertir las hojas: {arcpy.GetMessages()}")
        except Exception as e:
            # Capturar otros errores generales
            tkMessageBox.showerror("Error", "Error al importar el archivo Excel: {str(e)}")

    def import_rph_to_gdb(self):
        """Importar las hojas del archivo Excel a la GDB como tablas."""

        try:
            # Obtener la ruta de la geodatabase y el archivo Excel
            gdb = self.gdb_path.get()
            excel_rph_file = self.excel_file_path.get()

            # Verificar si las rutas son válidas
            if not os.path.exists(gdb) or not gdb.endswith('.gdb'):
                tkMessageBox.showerror("Error", "Seleccione una geodatabase válida.")
                return
            if not os.path.isfile(excel_rph_file) or not excel_rph_file.endswith('.xlsx'):
                tkMessageBox.showerror("Error", "Seleccione un archivo Excel válido.")
                return

            # Definir el espacio de trabajo de ArcPy
            arcpy.env.workspace = gdb

            # Listado de las hojas a convertir en tablas
            sheets = [
                "Edificios",
                "FichasPrediales",
                "ConstruccionesFicha",
                "CalificacionesConstrucciones",
                "ConstruccionGeneralFicha",
                "Propietarios",
            ]

            # Convertir cada hoja a una tabla en la geodatabase
            for sheet in sheets:
                output_table = os.path.join(gdb, sheet)  # Definir el nombre de la tabla en la GDB
                print("Importando"+ sheet +" a "+output_table)

                arcpy.ExcelToTable_conversion(
                    Input_Excel_File=excel_rph_file,
                    Output_Table=output_table,
                    Sheet=sheet
                )

            tkMessageBox.showinfo("Éxito", "Todas las hojas se han importado correctamente a la geodatabase.")

        except arcpy.ExecuteError as e:
            # Capturar errores de ArcPy
            tkMessageBox.showerror("Error de ArcPy", "Error al convertir las hojas: {arcpy.GetMessages()}")
        except Exception as e:
            # Capturar otros errores generales
            tkMessageBox.showerror("Error", "Error al importar el archivo Excel: {str(e)}")



    def process_all(self):
        """Función que ejecuta la selección de GDB, agrega campos e importa Excel en orden."""

        if not self.gdb_path.get():
            self.select_gdb()
        if not self.excel_file_path.get():
            self.select_excel()
        if self.gdb_path.get() and self.excel_file_path.get():
            self.add_fields()
            self.agregar_ficha_terreno()
            self.agregar_matricula_terreno()
            self.agregar_prediolc_terreno()
            self.agregar_ficha_unidad()
            self.agregar_matricula_unidad()
            self.agregar_prediolc_unidad()
            self.agregar_adquisicion_terreno()
            self.agregar_adquisicion_unidad()
            self.agregar_propietario_unidad()
            self.import_to_gdb()
            #self.comparar_y_exportar()

# Iniciar la aplicación


if __name__ == "__main__":
    root = tk.Tk()
    app = GDBApp(root)
    root.mainloop()