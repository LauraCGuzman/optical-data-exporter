# -*- coding: utf-8 -*-
"""
Módulo de interfaz de usuario
Maneja menús y diálogos de selección
"""

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, Toplevel, Listbox, Button, Checkbutton, IntVar, MULTIPLE, END, Frame, Label


class UI:
    """Gestiona la interfaz de usuario del programa"""
    
    def __init__(self):
        # Crear root window oculto para los diálogos
        self.root = tk.Tk()
        self.root.withdraw()
        # Asegurar que el root está listo para mostrar diálogos
        self.root.update()

    def mostrar_menu_principal(self):
        """
        Muestra el menú principal como diálogo gráfico y retorna la opción seleccionada

        Returns:
            'reflectancia', 'transmitancia_pv', 'transmitancia_csp', o None si cancela/sale
        """
        # Crear ventana de diálogo
        dialog = Toplevel(self.root)
        dialog.title("Exportador de Datos Ópticos")
        # Aumentamos un poco el alto (de 300 a 380) para la nueva opción
        dialog.geometry("450x380")
        dialog.resizable(False, False)

        # Traer ventana al frente
        dialog.lift()
        dialog.attributes('-topmost', True)

        # Frame principal
        main_frame = Frame(dialog, padx=30, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Título
        Label(
            main_frame,
            text="EXPORTADOR DE DATOS ÓPTICOS",
            font=("Arial", 14, "bold"),
            fg="#2E5090"
        ).pack(pady=(0, 20))

        Label(
            main_frame,
            text="Seleccione el tipo de medición a exportar:",
            font=("Arial", 11)
        ).pack(pady=(0, 15))

        # Variable para guardar selección
        seleccion = tk.StringVar(value="reflectancia")

        # Radio buttons
        rb_frame = Frame(main_frame)
        rb_frame.pack(pady=10)

        tk.Radiobutton(
            rb_frame,
            text="Reflectancia",
            variable=seleccion,
            value="reflectancia",
            font=("Arial", 11),
            padx=20,
            pady=5
        ).pack(anchor=tk.W)

        # Nueva Opción: Transmitancia PV
        tk.Radiobutton(
            rb_frame,
            text="Transmitancia PV",
            variable=seleccion,
            value="transmitancia_pv",
            font=("Arial", 11),
            padx=20,
            pady=5
        ).pack(anchor=tk.W)

        # Nueva Opción: Transmitancia CSP
        tk.Radiobutton(
            rb_frame,
            text="Transmitancia CSP",
            variable=seleccion,
            value="transmitancia_csp",
            font=("Arial", 11),
            padx=20,
            pady=5
        ).pack(anchor=tk.W)

        # Variable para resultado
        resultado = {'tipo': None}

        def aceptar():
            resultado['tipo'] = seleccion.get()
            dialog.destroy()

        def salir():
            resultado['tipo'] = None
            dialog.destroy()

        # Frame para botones
        button_frame = Frame(main_frame)
        button_frame.pack(pady=(20, 0))

        Button(
            button_frame,
            text="Aceptar",
            command=aceptar,
            bg="#4CAF50",
            fg="white",
            font=("Arial", 10, "bold"),
            padx=30,
            pady=8,
            width=12
        ).pack(side=tk.LEFT, padx=5)

        Button(
            button_frame,
            text="Salir",
            command=salir,
            bg="#f44336",
            fg="white",
            font=("Arial", 10, "bold"),
            padx=30,
            pady=8,
            width=12
        ).pack(side=tk.LEFT, padx=5)

        # Centrar ventana
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")

        # Hacer modal y forzar foco
        dialog.update()
        dialog.grab_set()
        dialog.focus_force()
        dialog.protocol("WM_DELETE_WINDOW", salir)
        dialog.wait_window()

        return resultado['tipo']

    def seleccionar_archivo_excel(self, tipo_medicion):
        """
        Muestra diálogo para seleccionar archivo Excel
        
        Args:
            tipo_medicion: 'reflectancia' o 'transmitancia'
        
        Returns:
            Path del archivo seleccionado o None si cancela
        """
        titulo = f"Seleccione archivo Excel con datos de {tipo_medicion.upper()}"
        
        filepath = filedialog.askopenfilename(
            title=titulo,
            filetypes=[
                ("Archivos Excel", "*.xlsx *.xlsm *.xls"),
                ("Todos los archivos", "*.*")
            ]
        )
        
        if not filepath:
            return None
        
        return filepath
    
    def seleccionar_hojas(self, sheet_names):
        """
        Muestra diálogo para seleccionar hojas del workbook
        
        Args:
            sheet_names: Lista de nombres de hojas
        
        Returns:
            Lista de nombres de hojas seleccionadas, o None si cancela
        """
        seleccionadas = []
        
        # Crear ventana de diálogo
        dialog = Toplevel(self.root)
        dialog.title("Seleccionar Hojas a Procesar")
        dialog.geometry("550x450")
        dialog.resizable(True, True)
        
        # Traer ventana al frente
        dialog.lift()
        dialog.attributes('-topmost', True)
        
        # Variable para "Todas las hojas"
        todas_var = IntVar()
        
        # Frame principal
        main_frame = Frame(dialog, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Instrucciones
        Label(
            main_frame,
            text="Seleccione las hojas que desea procesar:",
            font=("Arial", 11, "bold")
        ).pack(anchor=tk.W, pady=(0, 10))
        
        # Checkbox "Todas las hojas" - ARRIBA DEL TODO
        def toggle_todas():
            if todas_var.get() == 1:
                listbox.selection_set(0, END)
            else:
                listbox.selection_clear(0, END)
        
        check_frame = Frame(main_frame, bg="#E3F2FD", relief=tk.RAISED, bd=1)
        check_frame.pack(fill=tk.X, pady=(0, 10))
        
        Checkbutton(
            check_frame,
            text="✓ Seleccionar TODAS las hojas",
            variable=todas_var,
            command=toggle_todas,
            font=("Arial", 10, "bold"),
            bg="#E3F2FD",
            activebackground="#BBDEFB"
        ).pack(anchor=tk.W, padx=10, pady=8)
        
        # Frame para listbox y scrollbar
        list_frame = Frame(main_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # Scrollbar
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Listbox con hojas
        listbox = Listbox(
            list_frame,
            selectmode=MULTIPLE,
            yscrollcommand=scrollbar.set,
            font=("Arial", 10),
            height=15,
            selectbackground="#4CAF50",
            selectforeground="white"
        )
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=listbox.yview)
        
        # Agregar nombres de hojas
        for sheet_name in sheet_names:
            listbox.insert(END, sheet_name)
        
        # Info de hojas disponibles
        Label(
            main_frame,
            text=f"Total de hojas disponibles: {len(sheet_names)}",
            font=("Arial", 9),
            fg="#666"
        ).pack(anchor=tk.W)
        
        # Variable para guardar resultado
        resultado = {'seleccionadas': None}
        
        def aceptar():
            indices_seleccionados = listbox.curselection()
            if not indices_seleccionados:
                messagebox.showwarning(
                    "Sin selección",
                    "Por favor seleccione al menos una hoja.",
                    parent=dialog
                )
                return
            
            resultado['seleccionadas'] = [sheet_names[i] for i in indices_seleccionados]
            dialog.destroy()
        
        def cancelar():
            resultado['seleccionadas'] = None
            dialog.destroy()
        
        # Frame para botones
        button_frame = Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        Button(
            button_frame,
            text="Aceptar",
            command=aceptar,
            bg="#4CAF50",
            fg="white",
            font=("Arial", 10, "bold"),
            padx=20,
            pady=8,
            width=15
        ).pack(side=tk.LEFT, padx=(0, 10))
        
        Button(
            button_frame,
            text="Cancelar",
            command=cancelar,
            bg="#f44336",
            fg="white",
            font=("Arial", 10, "bold"),
            padx=20,
            pady=8,
            width=15
        ).pack(side=tk.LEFT)
        
        # Centrar ventana
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        # Hacer modal y forzar foco
        dialog.update()
        dialog.grab_set()
        dialog.focus_force()
        dialog.protocol("WM_DELETE_WINDOW", cancelar)
        dialog.wait_window()
        
        return resultado['seleccionadas']
    
    def confirmar_procesamiento(self, archivo, hojas, tipo_medicion):
        """
        Solicita confirmación antes de procesar usando diálogo gráfico
        
        Args:
            archivo: Path del archivo a procesar
            hojas: Lista de hojas a procesar
            tipo_medicion: 'reflectancia' o 'transmitancia'
        
        Returns:
            True si confirma, False si cancela
        """
        # Construir mensaje
        mensaje = f"Tipo de medición: {tipo_medicion.upper()}\n\n"
        mensaje += f"Archivo: {Path(archivo).name}\n\n"
        mensaje += f"Hojas a procesar: {len(hojas)}\n"
        
        # Si son muchas hojas, mostrar solo las primeras
        if len(hojas) <= 10:
            for hoja in hojas:
                mensaje += f"  • {hoja}\n"
        else:
            for hoja in hojas[:5]:
                mensaje += f"  • {hoja}\n"
            mensaje += f"  ... y {len(hojas) - 5} hojas más\n"
        
        mensaje += "\n¿Desea continuar con el procesamiento?"
        
        respuesta = messagebox.askyesno(
            "Confirmar Procesamiento",
            mensaje,
            icon='question'
        )
        
        return respuesta
    
    def mostrar_exito(self, num_hojas_procesadas, archivo_destino):
        """Muestra mensaje de éxito"""
        print("\n" + "=" * 60)
        print("  ✅ PROCESO COMPLETADO CON ÉXITO")
        print("=" * 60)
        print(f"\nHojas procesadas: {num_hojas_procesadas}")
        print(f"Datos exportados a: {archivo_destino}")
        print("\n" + "=" * 60)
        
        messagebox.showinfo(
            "Proceso Completado",
            f"Se procesaron {num_hojas_procesadas} hoja(s) correctamente.\n\n"
            f"Los datos se exportaron a:\n{archivo_destino}"
        )
    
    def mostrar_error(self, mensaje, log_path=None):
        """Muestra mensaje de error"""
        print("\n" + "=" * 60)
        print("  ❌ ERROR")
        print("=" * 60)
        print(f"\n{mensaje}")
        if log_path:
            print(f"\nSe generó un log de error en: {log_path}")
        print("\n" + "=" * 60)
        
        msg = mensaje
        if log_path:
            msg += f"\n\nSe generó un archivo de log con más detalles en:\n{log_path}"
        
        messagebox.showerror("Error", msg)
    
    def cerrar(self):
        """Cierra la interfaz"""
        if self.root:
            self.root.destroy()
