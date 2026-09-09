from exportar import exportar_excel

from database import Base, engine
import models
import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
from datetime import datetime

import reportes

Base.metadata.create_all(bind=engine)
class App(tk.Tk):

    def __init__(self):
        super().__init__()

        self.title("Sistema de Reportes")
        self.geometry("1100x650")
        self.minsize(900, 550)

        self.registros_actuales = []

        # Materiales cargados para el formulario
        self.materiales = []
        self.modo_edicion_material = False
        self.modo_edicion_reporte = False
        self.reporte_editando_id = None

        self.crear_interfaz()

    # =========================================================
    # INTERFAZ GENERAL
    # =========================================================

    def crear_interfaz(self):

        menu = tk.Frame(self, width=220)
        menu.pack(side="left", fill="y")
        menu.pack_propagate(False)

        tk.Label(
            menu,
            text="SISTEMA DE\nREPORTES",
            font=("Arial", 18, "bold")
        ).pack(pady=30)

        ttk.Button(
            menu,
            text="Inicio",
            command=self.mostrar_inicio
        ).pack(fill="x", padx=20, pady=5)

        ttk.Button(
            menu,
            text="Nuevo reporte",
            command=self.mostrar_nuevo_reporte
        ).pack(fill="x", padx=20, pady=5)

        ttk.Button(
            menu,
            text="Consultar reportes",
            command=self.mostrar_consultas
        ).pack(fill="x", padx=20, pady=5)

        ttk.Separator(
            menu,
            orient="horizontal"
        ).pack(fill="x", padx=20, pady=20)

        tk.Label(
            menu,
            text="CATÁLOGOS",
            font=("Arial", 10, "bold")
        ).pack(pady=5)

        ttk.Button(
            menu,
            text="Materiales",
            command=self.mostrar_materiales
        ).pack(fill="x", padx=20, pady=5)

        ttk.Button(
            menu,
            text="Solicitantes / Líderes",
            command=self.mostrar_relaciones
        ).pack(fill="x", padx=20, pady=5)

        self.contenido = tk.Frame(self)
        self.contenido.pack(
            side="right",
            fill="both",
            expand=True
        )

        self.mostrar_inicio()

    def limpiar_contenido(self):

        for widget in self.contenido.winfo_children():
            widget.destroy()

    def configurar_enter(self, comando):
        self.unbind("<Return>")
        self.bind("<Return>", lambda event: comando())

    # =========================================================
    # INICIO
    # =========================================================

    def mostrar_inicio(self):

        self.limpiar_contenido()
        self.unbind("<Return>")

        tk.Label(
            self.contenido,
            text="Sistema de Reportes",
            font=("Arial", 24, "bold")
        ).pack(pady=80)

        tk.Label(
            self.contenido,
            text=".venv pylace 3,14,14 allmenav c.",
            font=("Arial", 14)
        ).pack()

    # =========================================================
    # MATERIALES
    # =========================================================

    def mostrar_materiales(self):

        self.limpiar_contenido()
        self.configurar_enter(self.guardar_material)
        self.modo_edicion_material = False

        tk.Label(
            self.contenido,
            text="Catálogo de materiales",
            font=("Arial", 22, "bold")
        ).pack(pady=25)

        formulario = tk.Frame(self.contenido)
        formulario.pack()

        tk.Label(formulario, text="Código:").grid(
            row=0, column=0, sticky="e", padx=10, pady=10
        )

        self.material_codigo = ttk.Entry(formulario, width=40)
        self.material_codigo.grid(row=0, column=1, padx=10, pady=10)

        tk.Label(formulario, text="Descripción:").grid(
            row=1, column=0, sticky="e", padx=10, pady=10
        )

        self.material_descripcion = ttk.Entry(formulario, width=40)
        self.material_descripcion.grid(row=1, column=1, padx=10, pady=10)

        tk.Label(formulario, text="Categoría:").grid(
            row=2, column=0, sticky="e", padx=10, pady=10
        )

        categorias = reportes.obtener_categorias()
        self.material_categoria = ttk.Combobox(
            formulario,
            values=categorias,
            state="normal",
            width=37
        )
        self.material_categoria.grid(row=2, column=1, padx=10, pady=10)
        self.material_categoria.set("EPP")

        botones = tk.Frame(formulario)
        botones.grid(row=3, column=0, columnspan=2, pady=20)

        self.boton_guardar_material = ttk.Button(
            botones,
            text="Guardar",
            command=self.guardar_material
        )
        self.boton_guardar_material.pack(side="left", padx=5)

        ttk.Button(
            botones,
            text="Limpiar",
            command=self.limpiar_formulario_material
        ).pack(side="left", padx=5)

        ttk.Button(
            botones,
            text="Cancelar",
            command=self.mostrar_inicio
        ).pack(side="left", padx=5)

        tk.Label(
            self.contenido,
            text="Materiales registrados",
            font=("Arial", 14, "bold")
        ).pack(pady=(20, 10))

        tabla_frame = tk.Frame(self.contenido)
        tabla_frame.pack(fill="both", expand=True, padx=30, pady=10)

        columnas = ("codigo", "descripcion", "categoria")

        self.tabla_materiales = ttk.Treeview(
            tabla_frame,
            columns=columnas,
            show="headings"
        )

        encabezados = {
            "codigo": "Código",
            "descripcion": "Descripción",
            "categoria": "Categoría"
        }

        for columna in columnas:
            self.tabla_materiales.heading(
                columna, text=encabezados[columna]
            )

        self.tabla_materiales.column("codigo", width=150)
        self.tabla_materiales.column("descripcion", width=400)
        self.tabla_materiales.column("categoria", width=150)

        self.tabla_materiales.bind(
            "<Double-1>", self.editar_material_seleccionado
        )

        scrollbar = ttk.Scrollbar(
            tabla_frame,
            orient="vertical",
            command=self.tabla_materiales.yview
        )
        self.tabla_materiales.configure(yscrollcommand=scrollbar.set)

        self.tabla_materiales.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        ttk.Button(
            self.contenido,
            text="Editar material seleccionado",
            command=self.editar_material_seleccionado
        ).pack(pady=(0, 15))

        self.cargar_materiales()

    def limpiar_formulario_material(self):
        self.modo_edicion_material = False
        self.material_codigo.config(state="normal")
        self.material_codigo.delete(0, tk.END)
        self.material_descripcion.delete(0, tk.END)
        self.material_categoria.set("EPP")
        self.boton_guardar_material.config(text="Guardar")

    def editar_material_seleccionado(self, event=None):
        seleccion = self.tabla_materiales.selection()

        if not seleccion:
            if event is not None:
                return
            messagebox.showwarning(
                "Material",
                "Selecciona un material de la tabla."
            )
            return

        valores = self.tabla_materiales.item(seleccion[0], "values")

        self.modo_edicion_material = True

        self.material_codigo.config(state="normal")
        self.material_codigo.delete(0, tk.END)
        self.material_codigo.insert(0, valores[0])
        self.material_codigo.config(state="readonly")

        self.material_descripcion.delete(0, tk.END)
        self.material_descripcion.insert(0, valores[1])

        self.material_categoria.set(valores[2])
        self.boton_guardar_material.config(text="Actualizar")

    def guardar_material(self):
        codigo = self.material_codigo.get().strip()
        descripcion = self.material_descripcion.get().strip()
        categoria = self.material_categoria.get().strip()

        if not codigo or not descripcion or not categoria:
            messagebox.showwarning(
                "Campos incompletos",
                "Completa todos los campos."
            )
            return

        try:
            if self.modo_edicion_material:
                reportes.actualizar_material(
                    codigo,
                    descripcion,
                    categoria
                )
                mensaje = "Material actualizado correctamente."
            else:
                reportes.agregar_material(
                    codigo,
                    descripcion,
                    categoria
                )
                mensaje = "Material agregado correctamente."

            messagebox.showinfo("Material", mensaje)
            self.mostrar_materiales()

        except Exception as e:
            messagebox.showerror(
                "Error",
                f"No se pudo guardar el material.\n\n{e}"
            )

    def cargar_materiales(self):
        for item in self.tabla_materiales.get_children():
            self.tabla_materiales.delete(item)

        materiales = reportes.obtener_materiales()

        for material in materiales:
            self.tabla_materiales.insert(
                "",
                "end",
                values=(
                    material.codigo,
                    material.descripcion,
                    material.categoria
                )
            )

    # =========================================================
    # RELACIONES
    # =========================================================

    def mostrar_relaciones(self):
        self.limpiar_contenido()

        # =========================
        # FORMULARIO
        # =========================

        frame_form = ttk.Frame(self.contenido)
        frame_form.pack(pady=10)

        ttk.Label(
            frame_form,
            text="Solicitante:"
        ).grid(row=0, column=0, padx=5, pady=5)

        self.relacion_solicitante = ttk.Entry(frame_form)
        self.relacion_solicitante.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(
            frame_form,
            text="Líder:"
        ).grid(row=1, column=0, padx=5, pady=5)

        self.relacion_lider = ttk.Entry(frame_form)
        self.relacion_lider.grid(row=1, column=1, padx=5, pady=5)

        ttk.Button(
            frame_form,
            text="Guardar",
            command=self.guardar_relacion
        ).grid(row=2, column=0, padx=5, pady=10)

        ttk.Button(
            frame_form,
            text="Cancelar",
            command=self.mostrar_inicio
        ).grid(row=2, column=1, padx=5, pady=10)

        # =========================
        # TABLA
        # =========================

        ttk.Label(
            self.contenido,
            text="Relaciones registradas",
            font=("Arial", 16, "bold")
        ).pack(pady=10)

        self.tabla_relaciones = ttk.Treeview(
            self.contenido,
            columns=("id", "solicitante", "lider"),
            show="headings"
        )

        self.tabla_relaciones.heading("id", text="ID")
        self.tabla_relaciones.heading("solicitante", text="Solicitante")
        self.tabla_relaciones.heading("lider", text="Líder")

        self.tabla_relaciones.column("id", width=80)
        self.tabla_relaciones.column("solicitante", width=200)
        self.tabla_relaciones.column("lider", width=200)

        self.tabla_relaciones.pack(
            fill="both",
            expand=True
        )

        self.tabla_relaciones.bind(
            "<Double-1>",
            self.editar_relacion
        )

        self.cargar_relaciones()


    def editar_relacion(self, event):

        datos = self.tabla_relaciones.item(
            self.tabla_relaciones.selection()
        )

        valores = datos["values"]

        id_relacion = valores[0]

        relacion = reportes.obtener_relacion(id_relacion)

        ventana = tk.Toplevel(self)
        ventana.title("Editar relación")
        ventana.geometry("300x150")

        formulario = tk.Frame(ventana)
        formulario.pack()
        tk.Label(
        formulario,
        text="Solicitante:"
        ).pack()

        entrada_solicitante = tk.Entry(formulario)
        entrada_solicitante.pack()

        tk.Label(
        formulario,
        text="Líder:"
        ).pack()

        entrada_lider = tk.Entry(formulario)
        entrada_lider.pack()

        entrada_solicitante.insert(0, relacion.solicitante)
        entrada_lider.insert(0, relacion.lider)


        def guardar_cambios():
            nuevo_solicitante = entrada_solicitante.get()
            nuevo_lider = entrada_lider.get()

            reportes.actualizar_relacion(
                id_relacion,
                nuevo_solicitante,
                nuevo_lider
            )

            ventana.destroy()
            self.mostrar_relaciones()
        
        tk.Button(
        formulario,
        text="Guardar cambios",
        command=guardar_cambios
        ).pack(pady=20)
        

        ventana.bind("<Return>", lambda event: guardar_cambios())

    
    def guardar_relacion(self):

        solicitante = self.relacion_solicitante.get().strip()
        lider = self.relacion_lider.get().strip()

        if not solicitante or not lider:

            messagebox.showwarning(
                "Campos incompletos",
                "Completa solicitante y líder."
            )

            return

        try:

            reportes.agregar_relacion(
                solicitante,
                lider
            )

            messagebox.showinfo(
                "Relación",
                "Relación agregada correctamente."
            )

            self.mostrar_relaciones()

        except Exception as e:

            messagebox.showerror(
                "Error",
                f"No se pudo agregar la relación.\n\n{e}"
            )

    def cargar_relaciones(self):

        relaciones = reportes.obtener_relaciones()

        for relacion in relaciones:

            self.tabla_relaciones.insert(
                "",
                "end",
                values=(
                    relacion.id,
                    relacion.solicitante,
                    relacion.lider
                    
                )
            )

    # =========================================================
    # NUEVO REPORTE
    # =========================================================

    def mostrar_nuevo_reporte(self):

        self.limpiar_contenido()
        self.configurar_enter(self.crear_reporte)

        tk.Label(
            self.contenido,
            text="Nuevo reporte",
            font=("Arial", 22, "bold")
        ).pack(pady=25)

        formulario = tk.Frame(self.contenido)
        formulario.pack()

        # -----------------------------------------------------
        # MATERIAL
        # -----------------------------------------------------

        tk.Label(
            formulario,
            text="Material:"
        ).grid(
            row=0,
            column=0,
            sticky="e",
            padx=10,
            pady=10
        )

        self.materiales = reportes.obtener_materiales()

        descripciones = [
            material.descripcion
            for material in self.materiales
        ]

        self.material_descripcion_reporte = ttk.Combobox(
            formulario,
            values=descripciones,
            width=47
        )

        self.material_descripcion_reporte.grid(
            row=0,
            column=1,
            padx=10,
            pady=10
        )

        self.material_descripcion_reporte.bind(
            "<<ComboboxSelected>>",
            self.material_seleccionado
        )

        self.material_descripcion_reporte.bind(
            "<KeyRelease>",
            self.buscar_material
        )

        # -----------------------------------------------------
        # CÓDIGO
        # -----------------------------------------------------

        tk.Label(
            formulario,
            text="Código:"
        ).grid(
            row=1,
            column=0,
            sticky="e",
            padx=10,
            pady=10
        )

        self.codigo_reporte = ttk.Entry(
            formulario,
            width=50,
            state="readonly"
        )

        self.codigo_reporte.grid(
            row=1,
            column=1,
            padx=10,
            pady=10
        )

        # -----------------------------------------------------
        # CATEGORÍA
        # -----------------------------------------------------

        tk.Label(
            formulario,
            text="Categoría:"
        ).grid(
            row=2,
            column=0,
            sticky="e",
            padx=10,
            pady=10
        )

        self.categoria_reporte = ttk.Entry(
            formulario,
            width=50,
            state="readonly"
        )

        self.categoria_reporte.grid(
            row=2,
            column=1,
            padx=10,
            pady=10
        )

        # -----------------------------------------------------
        # CANTIDAD
        # -----------------------------------------------------

        tk.Label(
            formulario,
            text="Cantidad:"
        ).grid(
            row=3,
            column=0,
            sticky="e",
            padx=10,
            pady=10
        )

        self.cantidad_entry = ttk.Entry(
            formulario,
            width=50
        )

        self.cantidad_entry.grid(
            row=3,
            column=1,
            padx=10,
            pady=10
        )

        # -----------------------------------------------------
        # LÍDER
        # -----------------------------------------------------

        tk.Label(
            formulario,
            text="Líder:"
        ).grid(
            row=4, column=0, sticky="e", padx=10, pady=10
        )

        lideres = reportes.obtener_lideres()

        self.lider_reporte = ttk.Combobox(
            formulario,
            values=lideres,
            state="readonly",
            width=47
        )
        self.lider_reporte.grid(
            row=4, column=1, padx=10, pady=10
        )
        self.lider_reporte.bind(
            "<<ComboboxSelected>>",
            self.actualizar_solicitantes
        )

        # -----------------------------------------------------
        # CONTENEDOR EPP (solo solicitante)
        # -----------------------------------------------------

        self.frame_epp = tk.Frame(formulario)
        self.frame_epp.grid(
            row=5, column=0, columnspan=2, pady=5
        )

        tk.Label(
            self.frame_epp,
            text="Solicitante:"
        ).grid(
            row=0, column=0, sticky="e", padx=10, pady=10
        )

        self.solicitante_reporte = ttk.Combobox(
            self.frame_epp,
            state="readonly",
            width=47
        )
        self.solicitante_reporte.grid(
            row=0, column=1, padx=10, pady=10
        )

        # -----------------------------------------------------
        # FABRICACIÓN
        # -----------------------------------------------------

        tk.Label(
            formulario,
            text="Fabricación:"
        ).grid(
            row=6,
            column=0,
            sticky="e",
            padx=10,
            pady=10
        )

        self.fabricacion_entry = ttk.Entry(
            formulario,
            width=50
        )

        self.fabricacion_entry.grid(
            row=6,
            column=1,
            padx=10,
            pady=10
        )

        # -----------------------------------------------------
        # BOTONES
        # -----------------------------------------------------

        botones = tk.Frame(
            formulario
        )

        botones.grid(
            row=7,
            column=0,
            columnspan=2,
            pady=25
        )

        ttk.Button(
            botones,
            text="Crear reporte",
            command=self.crear_reporte
        ).pack(
            side="left",
            padx=5
        )

        ttk.Button(
            botones,
            text="Cancelar",
            command=self.mostrar_inicio
        ).pack(
            side="left",
            padx=5
        )

        # EPP empieza oculto hasta seleccionar material
        self.frame_epp.grid_remove()

    # =========================================================
    # SELECCIONAR MATERIAL
    # =========================================================

    def material_seleccionado(self, event=None):

        descripcion = (
            self.material_descripcion_reporte
            .get()
            .strip()
        )

        material = None

        for m in self.materiales:

            if m.descripcion == descripcion:
                material = m
                break

        if not material:
            return

        # Código
        self.codigo_reporte.config(
            state="normal"
        )

        self.codigo_reporte.delete(
            0,
            tk.END
        )

        self.codigo_reporte.insert(
            0,
            material.codigo
        )

        self.codigo_reporte.config(
            state="readonly"
        )

        # Categoría
        self.categoria_reporte.config(
            state="normal"
        )

        self.categoria_reporte.delete(
            0,
            tk.END
        )

        self.categoria_reporte.insert(
            0,
            material.categoria
        )

        self.categoria_reporte.config(
            state="readonly"
        )

        # EPP
        if material.categoria == "EPP":

            self.frame_epp.grid()

        else:

            self.frame_epp.grid_remove()
            self.solicitante_reporte.set("")

    # =========================================================
    # AUTOCOMPLETE MATERIAL
    # =========================================================

    def buscar_material(self, event=None):

        texto = (
            self.material_descripcion_reporte
            .get()
            .strip()
        )

        # Si no escribió nada, mostrar todos
        if not texto:

            resultados = reportes.obtener_materiales()

        else:

            resultados = reportes.buscar_materiales(
                texto
            )

        # Actualizar opciones del desplegable
        self.material_descripcion_reporte["values"] = [
            material.descripcion
            for material in resultados
        ]

    # =========================================================
    # FILTRAR SOLICITANTES POR LÍDER
    # =========================================================

    def actualizar_solicitantes(self, event=None):

        lider = self.lider_reporte.get()

        if not lider:
            return

        solicitantes = (
            reportes.obtener_solicitantes_por_lider(
                lider
            )
        )

        self.solicitante_reporte["values"] = solicitantes

        self.solicitante_reporte.set("")

    # =========================================================
    # CREAR REPORTE
    # =========================================================

    def crear_reporte(self):

        codigo = self.codigo_reporte.get().strip()
        categoria = self.categoria_reporte.get().strip()
        cantidad = self.cantidad_entry.get().strip()
        fabricacion = self.fabricacion_entry.get().strip()

        lider = self.lider_reporte.get().strip()
        solicitante = None

        # -----------------------------------------------------
        # VALIDACIONES GENERALES
        # -----------------------------------------------------

        if not codigo:

            messagebox.showwarning(
                "Material",
                "Selecciona un material."
            )

            return

        if not cantidad:

            messagebox.showwarning(
                "Cantidad",
                "Ingresa una cantidad."
            )

            return

        if not fabricacion:

            messagebox.showwarning(
                "Fabricación",
                "Ingresa la fabricación."
            )

            return

        try:

            cantidad = float(cantidad)

            if cantidad <= 0:

                raise ValueError(
                    "La cantidad debe ser mayor que cero."
                )

        except ValueError:

            messagebox.showerror(
                "Cantidad inválida",
                "La cantidad debe ser un número entero mayor que cero."
            )

            return

        # -----------------------------------------------------
        # VALIDACIONES EPP
        # -----------------------------------------------------

        if categoria == "EPP":

            lider = self.lider_reporte.get()
            solicitante = self.solicitante_reporte.get()

            if not lider:

                messagebox.showwarning(
                    "Líder",
                    "Selecciona un líder."
                )

                return

            if not solicitante:

                messagebox.showwarning(
                    "Solicitante",
                    "Selecciona un solicitante."
                )

                return

        # -----------------------------------------------------
        # GUARDAR
        # -----------------------------------------------------

        try:

            registro = reportes.agregar_registro(
                codigo,
                cantidad,
                solicitante,
                lider,
                fabricacion
            )


        

            self.mostrar_nuevo_reporte()

        except ValueError as e:

            messagebox.showerror(
                "Error",
                str(e)
            )

        except Exception as e:

            messagebox.showerror(
                "Error",
                f"No se pudo crear el reporte.\n\n{e}"
            )

    # =========================================================
    # EDITAR REPORTE
    # =========================================================

    def editar_reporte_seleccionado(self, event=None):
        seleccion = self.tabla.selection()

        if not seleccion:
            messagebox.showwarning(
                "Reporte",
                "Selecciona un reporte de la tabla."
            )
            return

        valores = self.tabla.item(seleccion[0], "values")
        try:
            id_registro = int(valores[0])
        except (ValueError, IndexError):
            messagebox.showerror("Reporte", "No se pudo obtener el ID del reporte.")
            return

        self.mostrar_editar_reporte(id_registro)

    def mostrar_editar_reporte(self, id_registro):
        registro = reportes.obtener_registro(id_registro)

        if not registro:
            messagebox.showerror("Reporte", "El reporte no existe.")
            return

        ventana = tk.Toplevel(self)
        ventana.title(f"Editar reporte #{registro.id}")
        ventana.geometry("560x500")
        ventana.transient(self)
        ventana.grab_set()

        tk.Label(
            ventana,
            text=f"Editar reporte #{registro.id}",
            font=("Arial", 20, "bold")
        ).pack(pady=20)

        formulario = tk.Frame(ventana)
        formulario.pack()

        tk.Label(formulario, text="Material:").grid(
            row=0, column=0, sticky="e", padx=10, pady=8
        )

        materiales = reportes.obtener_materiales()
        descripciones = [m.descripcion for m in materiales]

        material_var = tk.StringVar()
        material_combo = ttk.Combobox(
            formulario,
            textvariable=material_var,
            values=descripciones,
            width=42,
            state="readonly"
        )
        material_combo.grid(row=0, column=1, padx=10, pady=8)

        tk.Label(formulario, text="Código:").grid(
            row=1, column=0, sticky="e", padx=10, pady=8
        )
        codigo_var = tk.StringVar(value=registro.codigo)
        tk.Entry(
            formulario,
            textvariable=codigo_var,
            width=45,
            state="readonly"
        ).grid(row=1, column=1, padx=10, pady=8)

        tk.Label(formulario, text="Categoría:").grid(
            row=2, column=0, sticky="e", padx=10, pady=8
        )
        categoria_var = tk.StringVar()
        tk.Entry(
            formulario,
            textvariable=categoria_var,
            width=45,
            state="readonly"
        ).grid(row=2, column=1, padx=10, pady=8)

        tk.Label(formulario, text="Cantidad:").grid(
            row=3, column=0, sticky="e", padx=10, pady=8
        )
        cantidad_entry = ttk.Entry(formulario, width=45)
        cantidad_entry.insert(0, str(registro.cantidad))
        cantidad_entry.grid(row=3, column=1, padx=10, pady=8)

        tk.Label(formulario, text="Líder:").grid(
            row=4, column=0, sticky="e", padx=10, pady=8
        )
        lider_combo = ttk.Combobox(
            formulario,
            values=reportes.obtener_lideres(),
            state="readonly",
            width=42
        )
        lider_combo.grid(row=4, column=1, padx=10, pady=8)

        frame_epp = tk.Frame(formulario)
        frame_epp.grid(row=6, column=0, columnspan=2, pady=5)

        tk.Label(frame_epp, text="Solicitante:").grid(
            row=0, column=0, sticky="e", padx=10, pady=8
        )
        solicitante_combo = ttk.Combobox(
            frame_epp,
            state="readonly",
            width=42
        )
        solicitante_combo.grid(row=0, column=1, padx=10, pady=8)

        tk.Label(formulario, text="Fabricación:").grid(
            row=6, column=0, sticky="e", padx=10, pady=8
        )
        fabricacion_entry = ttk.Entry(formulario, width=45)
        fabricacion_entry.insert(0, registro.fabricacion)
        fabricacion_entry.grid(row=6, column=1, padx=10, pady=8)

        def seleccionar_material(event=None):
            descripcion = material_var.get()
            material = next(
                (m for m in materiales if m.descripcion == descripcion),
                None
            )

            if not material:
                return

            codigo_var.set(material.codigo)
            categoria_var.set(material.categoria)

            if material.categoria == "EPP":
                frame_epp.grid()
                if lider_combo.get():
                    solicitante_combo["values"] = (
                        reportes.obtener_solicitantes_por_lider(lider_combo.get())
                    )
            else:
                frame_epp.grid_remove()
                solicitante_combo.set("")

        def actualizar_solicitantes(event=None):
            lider = lider_combo.get()
            if lider:
                solicitante_combo["values"] = (
                    reportes.obtener_solicitantes_por_lider(lider)
                )
                solicitante_combo.set("")

        material_combo.bind("<<ComboboxSelected>>", seleccionar_material)
        lider_combo.bind("<<ComboboxSelected>>", actualizar_solicitantes)

        # Cargar datos actuales
        material_actual = next(
            (m for m in materiales if m.codigo == registro.codigo),
            None
        )
        if material_actual:
            material_var.set(material_actual.descripcion)
            categoria_var.set(material_actual.categoria)

        if registro.lider:
            lider_combo.set(registro.lider)
            solicitante_combo["values"] = (
                reportes.obtener_solicitantes_por_lider(registro.lider)
            )
        if registro.solicitante:
            solicitante_combo.set(registro.solicitante)

        if categoria_var.get() == "EPP":
            frame_epp.grid()
        else:
            frame_epp.grid_remove()

        def guardar_cambios():
            codigo = codigo_var.get().strip()
            categoria = categoria_var.get().strip()
            cantidad_texto = cantidad_entry.get().strip()
            fabricacion = fabricacion_entry.get().strip()
            lider = lider_combo.get().strip() or None
            solicitante = solicitante_combo.get().strip() or None

            if not codigo:
                messagebox.showwarning("Material", "Selecciona un material.", parent=ventana)
                return
            if not cantidad_texto:
                messagebox.showwarning("Cantidad", "Ingresa una cantidad.", parent=ventana)
                return
            if not fabricacion:
                messagebox.showwarning("Fabricación", "Ingresa la fabricación.", parent=ventana)
                return
            if not lider:
                messagebox.showwarning("Líder", "Selecciona un líder.", parent=ventana)
                return

            if categoria != "EPP":
                solicitante = None

            try:
                cantidad = float(cantidad_texto)
                if cantidad <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror(
                    "Cantidad inválida",
                    "La cantidad debe ser un número mayor que cero.",
                    parent=ventana
                )
                return

            if categoria == "EPP" and not solicitante:
                messagebox.showwarning(
                    "Solicitante",
                    "Los materiales EPP requieren un solicitante.",
                    parent=ventana
                )
                return

            try:
                reportes.actualizar_registro(
                    registro.id,
                    codigo,
                    cantidad,
                    solicitante,
                    lider,
                    fabricacion
                )
                messagebox.showinfo(
                    "Reporte actualizado",
                    "El reporte se actualizó correctamente.",
                    parent=ventana
                )
                ventana.destroy()
                self.buscar_reportes()
            except Exception as e:
                messagebox.showerror(
                    "Error",
                    f"No se pudo actualizar el reporte.\n\n{e}",
                    parent=ventana
                )

        botones = tk.Frame(ventana)
        botones.pack(pady=20)

        ttk.Button(
            botones,
            text="Guardar cambios",
            command=guardar_cambios
        ).pack(side="left", padx=5)

        ttk.Button(
            botones,
            text="Cancelar",
            command=ventana.destroy
        ).pack(side="left", padx=5)

    # =========================================================
    # CONSULTAS
    # =========================================================

    def mostrar_consultas(self):
        self.limpiar_contenido()
        self.configurar_enter(self.buscar_reportes)

        tk.Label(
            self.contenido,
            text="Consultar reportes",
            font=("Arial", 22, "bold")
        ).pack(pady=20)

        filtros = tk.Frame(self.contenido)
        filtros.pack()

        tk.Label(filtros, text="ID:").grid(row=0, column=0, padx=5, pady=5)
        self.filtro_id = ttk.Entry(filtros, width=12)
        self.filtro_id.grid(row=0, column=1, padx=5, pady=5)

        tk.Label(filtros, text="Código:").grid(row=0, column=2, padx=5, pady=5)
        self.filtro_codigo = ttk.Entry(filtros, width=16)
        self.filtro_codigo.grid(row=0, column=3, padx=5, pady=5)

        tk.Label(filtros, text="Solicitante:").grid(row=0, column=4, padx=5, pady=5)
        self.filtro_solicitante = ttk.Entry(filtros, width=16)
        self.filtro_solicitante.grid(row=0, column=5, padx=5, pady=5)

        tk.Label(filtros, text="Líder:").grid(row=0, column=6, padx=5, pady=5)
        self.filtro_lider = ttk.Entry(filtros, width=16)
        self.filtro_lider.grid(row=0, column=7, padx=5, pady=5)

        tk.Label(filtros, text="Fabricación:").grid(row=1, column=0, padx=5, pady=5)
        self.filtro_fabricacion = ttk.Entry(filtros, width=16)
        self.filtro_fabricacion.grid(row=1, column=1, padx=5, pady=5)

        tk.Label(filtros, text="Tipo:").grid(row=1, column=2, padx=5, pady=5)
        self.filtro_categoria = ttk.Combobox(
            filtros,
            values=["Todos"] + reportes.obtener_categorias(),
            state="readonly",
            width=14
        )
        self.filtro_categoria.grid(row=1, column=3, padx=5, pady=5)
        self.filtro_categoria.set("Todos")

        tk.Label(filtros, text="Desde:").grid(row=1, column=4, padx=5, pady=5)
        self.filtro_fecha_inicio = DateEntry(
            filtros, width=13, date_pattern="yyyy-mm-dd"
        )
        self.filtro_fecha_inicio.grid(row=1, column=5, padx=5, pady=5)

        tk.Label(filtros, text="Hasta:").grid(row=1, column=6, padx=5, pady=5)
        self.filtro_fecha_fin = DateEntry(
            filtros, width=13, date_pattern="yyyy-mm-dd"
        )
        self.filtro_fecha_fin.grid(row=1, column=7, padx=5, pady=5)

        ttk.Button(
            filtros,
            text="Buscar",
            command=self.buscar_reportes
        ).grid(row=2, column=7, padx=10, pady=10, sticky="e")

        tabla_frame = tk.Frame(self.contenido)
        tabla_frame.pack(fill="both", expand=True, padx=20, pady=15)

        columnas = (
            "id", "codigo", "descripcion", "cantidad",
            "solicitante", "lider", "fabricacion", "fecha"
        )

        self.tabla = ttk.Treeview(
            tabla_frame,
            columns=columnas,
            show="headings"
        )

        encabezados = {
            "id": "ID",
            "codigo": "Código",
            "descripcion": "Descripción",
            "cantidad": "Cantidad",
            "solicitante": "Solicitante",
            "lider": "Líder",
            "fabricacion": "Fabricación",
            "fecha": "Fecha"
        }

        for columna in columnas:
            self.tabla.heading(columna, text=encabezados[columna])
            self.tabla.column(columna, width=110)

        self.tabla.bind("<Double-1>", self.editar_reporte_seleccionado)

        scrollbar = ttk.Scrollbar(
            tabla_frame,
            orient="vertical",
            command=self.tabla.yview
        )
        self.tabla.configure(yscrollcommand=scrollbar.set)

        self.tabla.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        botones = tk.Frame(self.contenido)
        botones.pack(pady=10)

        ttk.Button(
            botones,
            text="Editar seleccionado",
            command=self.editar_reporte_seleccionado
        ).pack(side="left", padx=5)

        ttk.Button(
            botones,
            text="Exportar resultados a Excel",
            command=self.exportar_resultados
        ).pack(side="left", padx=5)

        ttk.Button(
            botones,
            text="Volver",
            command=self.mostrar_inicio
        ).pack(side="left", padx=5)

    def buscar_reportes(self):
        try:
            id_texto = self.filtro_id.get().strip()
            id_registro = int(id_texto) if id_texto else None

            codigo = self.filtro_codigo.get().strip() or None
            solicitante = self.filtro_solicitante.get().strip() or None
            lider = self.filtro_lider.get().strip() or None
            fabricacion = self.filtro_fabricacion.get().strip() or None

            categoria = self.filtro_categoria.get().strip()
            if categoria == "Todos" or not categoria:
                categoria = None

            fecha_inicio = datetime.strptime(
                self.filtro_fecha_inicio.get(), "%Y-%m-%d"
            ).date() if self.filtro_fecha_inicio.get() else None

            fecha_fin = datetime.strptime(
                self.filtro_fecha_fin.get(), "%Y-%m-%d"
            ).date() if self.filtro_fecha_fin.get() else None

            registros = reportes.obtener_registros(
                id_registro=id_registro,
                codigo=codigo,
                solicitante=solicitante,
                lider=lider,
                fabricacion=fabricacion,
                categoria=categoria,
                fecha_inicio=fecha_inicio,
                fecha_fin=fecha_fin
            )

            self.registros_actuales = registros

            for item in self.tabla.get_children():
                self.tabla.delete(item)

            for registro in registros:
                self.tabla.insert(
                    "",
                    "end",
                    values=(
                        registro.id,
                        registro.codigo,
                        registro.descripcion,
                        registro.cantidad,
                        registro.solicitante or "",
                        registro.lider or "",
                        registro.fabricacion,
                        registro.fecha.strftime("%Y-%m-%d")
                    )
                )

            if not registros:
                messagebox.showinfo(
                    "Consulta",
                    "No se encontraron reportes."
                )

        except ValueError:
            messagebox.showerror(
                "Filtro inválido",
                "El ID debe ser un número entero y las fechas deben usar YYYY-MM-DD."
            )

    # =========================================================
    # EXPORTAR
    # =========================================================

    def exportar_resultados(self):

        if not self.registros_actuales:

            messagebox.showwarning(
                "Sin resultados",
                "No hay registros para exportar."
            )

            return

        archivo = exportar_excel(
            self.registros_actuales
        )

        messagebox.showinfo(
            "Excel generado",
            f"Archivo generado correctamente:\n{archivo}"
        )


# =============================================================
# EJECUCIÓN
# =============================================================

if __name__ == "__main__":

    app = App()
    app.mainloop()