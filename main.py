
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
            text="REPORTES",
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

        tk.Label(
            self.contenido,
            text="Catálogo de materiales",
            font=("Arial", 22, "bold")
        ).pack(pady=25)

        formulario = tk.Frame(self.contenido)
        formulario.pack()

        tk.Label(
            formulario,
            text="Código:"
        ).grid(
            row=0,
            column=0,
            sticky="e",
            padx=10,
            pady=10
        )

        self.material_codigo = ttk.Entry(
            formulario,
            width=40
        )

        self.material_codigo.grid(
            row=0,
            column=1,
            padx=10,
            pady=10
        )

        tk.Label(
            formulario,
            text="Descripción:"
        ).grid(
            row=1,
            column=0,
            sticky="e",
            padx=10,
            pady=10
        )

        self.material_descripcion = ttk.Entry(
            formulario,
            width=40
        )

        self.material_descripcion.grid(
            row=1,
            column=1,
            padx=10,
            pady=10
        )

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

        self.material_categoria = ttk.Combobox(
            formulario,
            values=[
                "EPP",
                "DISCOS",
                "SOLDADURA"
            ],
            state="readonly",
            width=37
        )

        self.material_categoria.grid(
            row=2,
            column=1,
            padx=10,
            pady=10
        )

        self.material_categoria.current(0)

        botones = tk.Frame(formulario)
        botones.grid(
            row=3,
            column=0,
            columnspan=2,
            pady=20
        )

        ttk.Button(
            botones,
            text="Guardar",
            command=self.guardar_material
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

        tk.Label(
            self.contenido,
            text="Materiales registrados",
            font=("Arial", 14, "bold")
        ).pack(
            pady=(20, 10)
        )

        # =========================
        # BÚSQUEDA DE MATERIALES
        # =========================

        busqueda_frame = ttk.Frame(self.contenido)
        busqueda_frame.pack(pady=(0, 10))

        ttk.Label(
            busqueda_frame,
            text="Buscar:"
        ).pack(side="left", padx=5)

        self.busqueda_materiales = ttk.Entry(
            busqueda_frame,
            width=45
        )
        self.busqueda_materiales.pack(side="left", padx=5)
        self.busqueda_materiales.bind(
            "<KeyRelease>",
            self.filtrar_materiales
        )

        tabla_frame = tk.Frame(self.contenido)
        tabla_frame.pack(
            fill="both",
            expand=True,
            padx=30,
            pady=10
        )

        columnas = (
            "codigo",
            "descripcion",
            "categoria"
        )

        self.tabla_materiales = ttk.Treeview(
            tabla_frame,
            columns=columnas,
            show="headings"
        )

        self.tabla_materiales.heading(
            "codigo",
            text="Código"
        )

        self.tabla_materiales.heading(
            "descripcion",
            text="Descripción"
        )

        self.tabla_materiales.heading(
            "categoria",
            text="Categoría"
        )

        self.tabla_materiales.column(
            "codigo",
            width=150
        )

        self.tabla_materiales.column(
            "descripcion",
            width=400
        )

        self.tabla_materiales.column(
            "categoria",
            width=150
        )

        scrollbar = ttk.Scrollbar(
            tabla_frame,
            orient="vertical",
            command=self.tabla_materiales.yview
        )

        self.tabla_materiales.configure(
            yscrollcommand=scrollbar.set
        )

        self.tabla_materiales.pack(
            side="left",
            fill="both",
            expand=True
        )

        scrollbar.pack(
            side="right",
            fill="y"
        )

        self.cargar_materiales()

    def guardar_material(self):

        codigo = self.material_codigo.get().strip()
        descripcion = self.material_descripcion.get().strip()
        categoria = self.material_categoria.get()

        if not codigo or not descripcion or not categoria:

            messagebox.showwarning(
                "Campos incompletos",
                "Completa todos los campos."
            )

            return

        try:

            reportes.agregar_material(
                codigo,
                descripcion,
                categoria
            )

            messagebox.showinfo(
                "Material",
                "Material agregado correctamente."
            )

            self.mostrar_materiales()

        except Exception as e:

            messagebox.showerror(
                "Error",
                f"No se pudo agregar el material.\n\n{e}"
            )

    def filtrar_materiales(self, event=None):
        texto = self.busqueda_materiales.get().strip().lower()

        for item in self.tabla_materiales.get_children():
            self.tabla_materiales.delete(item)

        materiales = reportes.obtener_materiales()

        for material in materiales:
            if (
                not texto
                or texto in material.codigo.lower()
                or texto in material.descripcion.lower()
                or texto in material.categoria.lower()
            ):
                self.tabla_materiales.insert(
                    "",
                    "end",
                    values=(
                        material.codigo,
                        material.descripcion,
                        material.categoria
                    )
                )

    def cargar_materiales(self):

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
        self.configurar_enter(self.guardar_relacion)

        tk.Label(
            self.contenido,
            text="Solicitantes y líderes",
            font=("Arial", 22, "bold")
        ).pack(pady=25)

        formulario = tk.Frame(self.contenido)
        formulario.pack()

        tk.Label(
            formulario,
            text="Solicitante:"
        ).grid(
            row=0,
            column=0,
            sticky="e",
            padx=10,
            pady=10
        )

        self.relacion_solicitante = ttk.Entry(
            formulario,
            width=40
        )

        self.relacion_solicitante.grid(
            row=0,
            column=1,
            padx=10,
            pady=10
        )

        tk.Label(
            formulario,
            text="Líder:"
        ).grid(
            row=1,
            column=0,
            sticky="e",
            padx=10,
            pady=10
        )

        self.relacion_lider = ttk.Entry(
            formulario,
            width=40
        )

        self.relacion_lider.grid(
            row=1,
            column=1,
            padx=10,
            pady=10
        )

        botones = tk.Frame(formulario)
        botones.grid(
            row=2,
            column=0,
            columnspan=2,
            pady=20
        )

        ttk.Button(
            botones,
            text="Guardar",
            command=self.guardar_relacion
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

        tk.Label(
            self.contenido,
            text="Relaciones registradas",
            font=("Arial", 14, "bold")
        ).pack(
            pady=(20, 10)
        )

        tabla_frame = tk.Frame(self.contenido)
        tabla_frame.pack(
            fill="both",
            expand=True,
            padx=30,
            pady=10
        )

        columnas = (
            "solicitante",
            "lider"
        )

        self.tabla_relaciones = ttk.Treeview(
            tabla_frame,
            columns=columnas,
            show="headings"
        )

        self.tabla_relaciones.heading(
            "solicitante",
            text="Solicitante"
        )

        self.tabla_relaciones.heading(
            "lider",
            text="Líder"
        )

        self.tabla_relaciones.column(
            "solicitante",
            width=300
        )

        self.tabla_relaciones.column(
            "lider",
            width=300
        )

        self.tabla_relaciones.pack(
            fill="both",
            expand=True
        )

        self.cargar_relaciones()

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
        # CONTENEDOR EPP
        # -----------------------------------------------------

        self.frame_epp = tk.Frame(
            formulario
        )

        self.frame_epp.grid(
            row=4,
            column=0,
            columnspan=2,
            pady=5
        )

        tk.Label(
            self.frame_epp,
            text="Líder:"
        ).grid(
            row=0,
            column=0,
            sticky="e",
            padx=10,
            pady=10
        )

        lideres = reportes.obtener_lideres()

        self.lider_reporte = ttk.Combobox(
            self.frame_epp,
            values=lideres,
            state="readonly",
            width=47
        )

        self.lider_reporte.grid(
            row=0,
            column=1,
            padx=10,
            pady=10
        )

        self.lider_reporte.bind(
            "<<ComboboxSelected>>",
            self.actualizar_solicitantes
        )

        tk.Label(
            self.frame_epp,
            text="Solicitante:"
        ).grid(
            row=1,
            column=0,
            sticky="e",
            padx=10,
            pady=10
        )

        self.solicitante_reporte = ttk.Combobox(
            self.frame_epp,
            state="readonly",
            width=47
        )

        self.solicitante_reporte.grid(
            row=1,
            column=1,
            padx=10,
            pady=10
        )

        # -----------------------------------------------------
        # FABRICACIÓN
        # -----------------------------------------------------

        tk.Label(
            formulario,
            text="Fabricación:"
        ).grid(
            row=5,
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
            row=5,
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
            row=6,
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

            self.lider_reporte.set("")
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

        solicitante = None
        lider = self.lider_reporte.get()

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
        # VALIDACIÓN DE LÍDER
        # -----------------------------------------------------

        if not lider:

            messagebox.showwarning(
                "Líder",
                "Selecciona un líder."
            )

            return

        # -----------------------------------------------------
        # VALIDACIONES EPP
        # -----------------------------------------------------

        if categoria == "EPP":

            solicitante = self.solicitante_reporte.get()

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

            mensaje = (
                f"Reporte creado correctamente.\n\n"
                f"Código: {registro.codigo}\n"
                f"Descripción: {registro.descripcion}\n"
                f"Categoría: {categoria}\n"
                f"Cantidad: {registro.cantidad}\n"
            )

            if registro.solicitante:

                mensaje += (
                    f"Solicitante: {registro.solicitante}\n"
                    f"Líder: {registro.lider}\n"
                )

            mensaje += (
                f"Fabricación: {registro.fabricacion}"
            )

            messagebox.showinfo(
                "Reporte creado",
                mensaje
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

        filtros = tk.Frame(
            self.contenido
        )

        filtros.pack()

        tk.Label(
            filtros,
            text="Código:"
        ).grid(
            row=0,
            column=0,
            padx=5,
            pady=5
        )

        self.materiales_filtro_codigo = reportes.obtener_materiales()

        self.filtro_codigo = ttk.Combobox(
            filtros,
            values=[
                f"{material.codigo} - {material.descripcion}"
                for material in self.materiales_filtro_codigo
            ],
            width=35
        )

        self.filtro_codigo.grid(
            row=0,
            column=1,
            padx=5,
            pady=5
        )

        self.filtro_codigo.bind(
            "<KeyRelease>",
            self.autocompletar_codigo_reporte
        )

        self.filtro_codigo.bind(
            "<<ComboboxSelected>>",
            self.seleccionar_codigo_reporte
        )

        tk.Label(
            filtros,
            text="Solicitante:"
        ).grid(
            row=0,
            column=2,
            padx=5,
            pady=5
        )

        self.filtro_solicitante = ttk.Entry(
            filtros,
            width=20
        )

        self.filtro_solicitante.grid(
            row=0,
            column=3,
            padx=5,
            pady=5
        )

        tk.Label(
            filtros,
            text="Líder:"
        ).grid(
            row=0,
            column=4,
            padx=5,
            pady=5
        )

        self.filtro_lider = ttk.Entry(
            filtros,
            width=20
        )

        self.filtro_lider.grid(
            row=0,
            column=5,
            padx=5,
            pady=5
        )

        tk.Label(
            filtros,
            text="Desde:"
        ).grid(
            row=1,
            column=0,
            padx=5,
            pady=5
        )

        self.filtro_fecha_inicio = DateEntry(
            filtros,
            width=17,
            date_pattern="yyyy-mm-dd"
        )

        self.filtro_fecha_inicio.grid(
            row=1,
            column=1,
            padx=5,
            pady=5
        )

        tk.Label(
            filtros,
            text="Hasta:"
        ).grid(
            row=1,
            column=2,
            padx=5,
            pady=5
        )

        self.filtro_fecha_fin = DateEntry(
            filtros,
            width=17,
            date_pattern="yyyy-mm-dd"
        )

        self.filtro_fecha_fin.grid(
            row=1,
            column=3,
            padx=5,
            pady=5
        )
        tk.Label(
            filtros,
            text="Hasta:"
            ).grid(
                row=1,
                column=2,
                padx=5,
                pady=5
        )

        ttk.Button(
            filtros,
            text="Buscar",
            command=self.buscar_reportes
        ).grid(
            row=1,
            column=5,
            padx=10
        )

        tabla_frame = tk.Frame(
            self.contenido
        )

        tabla_frame.pack(
            fill="both",
            expand=True,
            padx=20,
            pady=20
        )

        columnas = (
            "id",
            "codigo",
            "descripcion",
            "cantidad",
            "solicitante",
            "lider",
            "fabricacion",
            "fecha"
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

            self.tabla.heading(
                columna,
                text=encabezados[columna]
            )

            self.tabla.column(
                columna,
                width=110
            )

        scrollbar = ttk.Scrollbar(
            tabla_frame,
            orient="vertical",
            command=self.tabla.yview
        )

        self.tabla.configure(
            yscrollcommand=scrollbar.set
        )

        self.tabla.pack(
            side="left",
            fill="both",
            expand=True
        )

        scrollbar.pack(
            side="right",
            fill="y"
        )

        botones = tk.Frame(
            self.contenido
        )

        botones.pack(
            pady=10
        )

        ttk.Button(
            botones,
            text="Exportar resultados a Excel",
            command=self.exportar_resultados
        ).pack(
            side="left",
            padx=5
        )

        ttk.Button(
            botones,
            text="Volver",
            command=self.mostrar_inicio
        ).pack(
            side="left",
            padx=5
        )

    def autocompletar_codigo_reporte(self, event=None):

        texto = self.filtro_codigo.get().strip().lower()

        if not texto:
            resultados = self.materiales_filtro_codigo
        else:
            resultados = [
                material
                for material in self.materiales_filtro_codigo
                if (
                    texto in material.codigo.lower()
                    or texto in material.descripcion.lower()
                )
            ]

        self.filtro_codigo["values"] = [
            f"{material.codigo} - {material.descripcion}"
            for material in resultados
        ]


    def seleccionar_codigo_reporte(self, event=None):

        texto = self.filtro_codigo.get().strip()

        if " - " in texto:
            codigo = texto.split(" - ", 1)[0].strip()

            self.filtro_codigo.set(codigo)


    def buscar_reportes(self):

        try:

            codigo = (
                self.filtro_codigo.get()
                or None
            )

            solicitante = (
                self.filtro_solicitante.get()
                or None
            )

            lider = (
                self.filtro_lider.get()
                or None
            )

            fecha_inicio = None
            fecha_fin = None

            if self.filtro_fecha_inicio.get():

                fecha_inicio = datetime.strptime(
                    self.filtro_fecha_inicio.get(),
                    "%Y-%m-%d"
                ).date()

            if self.filtro_fecha_fin.get():

                fecha_fin = datetime.strptime(
                    self.filtro_fecha_fin.get(),
                    "%Y-%m-%d"
                ).date()

            registros = reportes.obtener_registros(
                codigo=codigo,
                solicitante=solicitante,
                lider=lider,
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
                        registro.fecha.strftime(
                            "%Y-%m-%d"
                        )
                    )
                )

            if not registros:

                messagebox.showinfo(
                    "Consulta",
                    "No se encontraron reportes."
                )

        except ValueError:

            messagebox.showerror(
                "Fecha inválida",
                "Utiliza el formato YYYY-MM-DD."
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