# -*- coding: utf-8 -*-
import Tkinter as tk
import ttk
import tkFileDialog as filedialog
import tkMessageBox as messagebox
import arcpy
import os

class GDBReaderApp:

    def __init__(self, root):
        self.root = root
        self.root.title("📁 UTR – Procesos Automáticos sobre GDB")
        self.root.geometry("880x550")
        self.root.configure(bg="#ECECEC")

        self._estilos()
        self._interfaz()

    # ------------------------------------------------------------------
    # Estilos visuales
    # ------------------------------------------------------------------
    def _estilos(self):
        style = ttk.Style()
        style.theme_use('clam')

        style.configure("TButton",
                        font=("Segoe UI", 10),
                        padding=6,
                        background="#4F81BD",
                        foreground="white")
        style.map("TButton",
                  background=[("active", "#3567A5")])

        style.configure("TLabel",
                        background="#ECECEC",
                        font=("Segoe UI", 10))

        style.configure("Treeview",
                        font=("Segoe UI", 9),
                        rowheight=22)

    # ------------------------------------------------------------------
    # Interfaz gráfica
    # ------------------------------------------------------------------
    def _interfaz(self):

        frame_top = tk.Frame(self.root, bg="#ECECEC")
        frame_top.pack(fill="x", padx=10, pady=10)

        ttk.Label(frame_top, text="Ruta GDB:").pack(side="left", padx=5)

        self.ruta_gdb = tk.StringVar()
        entry = ttk.Entry(frame_top, textvariable=self.ruta_gdb, width=70)
        entry.pack(side="left", padx=5)

        ttk.Button(frame_top, text="📂 Buscar", command=self.seleccionar_gdb)\
            .pack(side="left", padx=5)

        ttk.Button(frame_top, text="📥 Cargar GDB", command=self.cargar_gdb)\
            .pack(side="left", padx=5)

        ttk.Button(frame_top, text="▶ Ejecutar Proceso", command=self.ejecutar_proceso)\
            .pack(side="left", padx=5)

        # Árbol de capas
        self.tree = ttk.Treeview(self.root)
        self.tree.pack(fill="both", expand=True, padx=10, pady=10)

        self.tree["columns"] = ("Tipo", "Ruta")
        self.tree.heading("#0", text="Nombre")
        self.tree.heading("Tipo", text="Tipo")
        self.tree.heading("Ruta", text="Ruta")
        self.tree.column("Ruta", width=350)

        sb = ttk.Scrollbar(self.tree, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=sb.set)
        sb.pack(side="right", fill="y")

    # ------------------------------------------------------------------
    def seleccionar_gdb(self):
        carpeta = filedialog.askdirectory(title="Selecciona una Geodatabase")
        if carpeta:
            self.ruta_gdb.set(carpeta)

    # ------------------------------------------------------------------
    # Cargar la GDB en el árbol
    # ------------------------------------------------------------------
    def cargar_gdb(self):
        gdb = self.ruta_gdb.get()

        if not gdb or not os.path.exists(gdb):
            messagebox.showerror("Error", "Selecciona una GDB válida.")
            return

        if not gdb.lower().endswith(".gdb"):
            messagebox.showerror("Error", "La ruta seleccionada NO es una .gdb.")
            return

        arcpy.env.workspace = gdb
        self.tree.delete(*self.tree.get_children())

        # Feature Classes raíz
        fc_root = self.tree.insert("", "end", text="Feature Classes", open=True)
        for fc in arcpy.ListFeatureClasses():
            ruta = os.path.join(gdb, fc)
            self.tree.insert(fc_root, "end", text=fc, values=("Feature Class", ruta))

        # Feature Datasets
        for fd in arcpy.ListDatasets():
            nodo_fd = self.tree.insert("", "end", text="FD: " + fd, open=False)
            arcpy.env.workspace = os.path.join(gdb, fd)
            for fc in arcpy.ListFeatureClasses():
                ruta = os.path.join(gdb, fd, fc)
                self.tree.insert(nodo_fd, "end", text=fc, values=("FC en FD", ruta))

        # Tablas
        arcpy.env.workspace = gdb
        tablas_root = self.tree.insert("", "end", text="Tablas", open=True)
        for t in arcpy.ListTables():
            ruta = os.path.join(gdb, t)
            self.tree.insert(tablas_root, "end", text=t, values=("Tabla", ruta))

        messagebox.showinfo("Completado", "GDB cargada exitosamente.")

    # ------------------------------------------------------------------
    # PROCESO COMPLETO (Partes 1 + 2 + 3 del modelo)
    # ------------------------------------------------------------------
    def ejecutar_proceso(self):

        gdb = self.ruta_gdb.get()

        if not gdb or not os.path.exists(gdb):
            messagebox.showerror("Error", "Selecciona una GDB válida.")
            return

        arcpy.env.workspace = gdb
        arcpy.env.overwriteOutput = True

        try:
            # ------------------------------------------------------------
            # CAPAS
            # ------------------------------------------------------------
            sentencias = os.path.join(gdb, "sentencia")
            terreno = os.path.join(gdb, "r_lc_terreno")
            predios = os.path.join(gdb, "PREDIOS_LUR_R")

            if not arcpy.Exists(sentencias):
                messagebox.showerror("Error", "Falta sentencia")
                return
            if not arcpy.Exists(terreno):
                messagebox.showerror("Error", "Falta r_lc_terreno")
                return
            if not arcpy.Exists(predios):
                messagebox.showerror("Error", "Falta PREDIOS_LUR_R")
                return

            # ============================================================
            # 🟦 PARTE 1
            # ============================================================

            arcpy.MakeFeatureLayer_management(sentencias, "sent_lyr")
            arcpy.SelectLayerByLocation_management("sent_lyr", "INTERSECT", terreno)
            count_sent = int(arcpy.GetCount_management("sent_lyr")[0])

            salida_sentencias = os.path.join(gdb, "SENTENCIAS_EXPORT")
            arcpy.CopyFeatures_management("sent_lyr", salida_sentencias)

            # Predios → puntos
            puntos_predios = "in_memory/puntos_predios"
            arcpy.FeatureToPoint_management(predios, puntos_predios, "INSIDE")

            arcpy.MakeFeatureLayer_management(puntos_predios, "puntos_lyr")
            arcpy.SelectLayerByLocation_management("puntos_lyr", "INTERSECT", terreno)
            count_puntos = int(arcpy.GetCount_management("puntos_lyr")[0])

            # ============================================================
            # 🟨 PARTE 2
            # ============================================================

            predios_export = os.path.join(gdb, "PREDIOS_LUR_IN")
            arcpy.CopyFeatures_management(predios, predios_export)

            diferencia_fc = os.path.join(gdb, "Diferencia")
            arcpy.Erase_analysis(predios_export, terreno, diferencia_fc)

            # Si no existe Area_Dif → agregar
            campos = [f.name for f in arcpy.ListFields(diferencia_fc)]
            if "Area_Dif" not in campos:
                arcpy.AddField_management(diferencia_fc, "Area_Dif", "DOUBLE")

            arcpy.CalculateField_management(
                diferencia_fc, "Area_Dif", "!shape.area!", "PYTHON_9.3"
            )

            arcpy.MakeFeatureLayer_management(diferencia_fc, "dif_lyr")
            arcpy.SelectLayerByAttribute_management("dif_lyr", "NEW_SELECTION",
                                                    '"Area_Dif" > 0')

            count_diferencias = int(arcpy.GetCount_management("dif_lyr")[0])

            # ============================================================
            # 🟩 PARTE 3
            # ============================================================

            diferencia_predios = os.path.join(gdb, "Diferencia_predios")
            arcpy.CopyFeatures_management("dif_lyr", diferencia_predios)

            # Export table
            tabla_reportes = os.path.join(gdb, "TABLAREPORTES.dbf")
            arcpy.TableToTable_conversion("dif_lyr", gdb, "TABLAREPORTES.dbf")

            # ============================================================
            # MENSAJE FINAL
            # ============================================================
            messagebox.showinfo("Proceso Finalizado",
                "✔ Sentencias intersectadas: {}\n"
                "✔ Puntos generados en predios: {}\n"
                "✔ Predios con diferencias (Erase): {}\n"
                "✔ Predios exportados con diferencia: {}\n"
                "✔ Tabla generada: TABLAREPORTES.dbf\n"
                "\nProceso completado correctamente."
                .format(count_sent, count_puntos, count_diferencias, count_diferencias)
            )

        except Exception as e:
            messagebox.showerror("Error en el proceso", str(e))


# ----------------------------------------------------------------------
# EJECUCIÓN
# ----------------------------------------------------------------------
if __name__ == "__main__":
    root = tk.Tk()
    app = GDBReaderApp(root)
    root.mainloop()