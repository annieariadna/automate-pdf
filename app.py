import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import threading
import os
from test_pdf import BalanceExtractorEnhanced  # Importamos tu algoritmo

class PDFToExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("🏦 Extractor PDF a Excel - Balance de Comprobación")
        self.root.geometry("600x500")
        self.root.resizable(True, True)
        
        # Variables
        self.selected_file = tk.StringVar()
        self.output_file = tk.StringVar()
        self.progress_var = tk.DoubleVar()
        self.status_var = tk.StringVar(value="Listo para procesar...")
        
        # Crear la interfaz
        self.create_widgets()
        
        # Centrar ventana
        self.center_window()
    
    def center_window(self):
        """Centrar la ventana en la pantalla"""
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (self.root.winfo_width() // 2)
        y = (self.root.winfo_screenheight() // 2) - (self.root.winfo_height() // 2)
        self.root.geometry(f"+{x}+{y}")
    
    def create_widgets(self):
        """Crear todos los widgets de la interfaz"""
        
        # Frame principal con padding
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configurar grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Título
        title_label = ttk.Label(
            main_frame, 
            text="Extractor PDF a Excel", 
            font=("Arial", 16, "bold")
        )
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 10))
        
        subtitle_label = ttk.Label(
            main_frame, 
            text="Balance de Comprobación - Banco de la Nación",
            font=("Arial", 10)
        )
        subtitle_label.grid(row=1, column=0, columnspan=3, pady=(0, 20))
        
        # Separador
        separator1 = ttk.Separator(main_frame, orient='horizontal')
        separator1.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        
        # Sección de selección de archivo
        file_frame = ttk.LabelFrame(main_frame, text="📄 Seleccionar Archivo PDF", padding="10")
        file_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        file_frame.columnconfigure(1, weight=1)
        
        ttk.Label(file_frame, text="Archivo:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        
        self.file_entry = ttk.Entry(file_frame, textvariable=self.selected_file, state="readonly")
        self.file_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        
        self.browse_button = ttk.Button(
            file_frame, 
            text="📁 Examinar", 
            command=self.browse_file
        )
        self.browse_button.grid(row=0, column=2)
        
        # Información del archivo
        self.file_info_label = ttk.Label(file_frame, text="", foreground="gray")
        self.file_info_label.grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=(5, 0))
        
        # Sección de archivo de salida
        output_frame = ttk.LabelFrame(main_frame, text="💾 Archivo de Salida", padding="10")
        output_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        output_frame.columnconfigure(1, weight=1)
        
        ttk.Label(output_frame, text="Guardar como:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        
        self.output_entry = ttk.Entry(output_frame, textvariable=self.output_file)
        self.output_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        
        self.output_browse_button = ttk.Button(
            output_frame, 
            text="📁 Guardar en", 
            command=self.browse_output_file
        )
        self.output_browse_button.grid(row=0, column=2)
        
        # Separador
        separator2 = ttk.Separator(main_frame, orient='horizontal')
        separator2.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        
        # Botones principales
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=6, column=0, columnspan=3, pady=(0, 15))
        
        self.process_button = ttk.Button(
            button_frame,
            text="🔄 Procesar PDF",
            command=self.process_file,
            style="Accent.TButton"
        )
        self.process_button.pack(side=tk.LEFT, padx=(0, 10))
        
        self.clear_button = ttk.Button(
            button_frame,
            text="🗑️ Limpiar",
            command=self.clear_form
        )
        self.clear_button.pack(side=tk.LEFT)
        
        # Barra de progreso
        progress_frame = ttk.LabelFrame(main_frame, text="📊 Progreso", padding="10")
        progress_frame.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        
        self.progress_bar = ttk.Progressbar(
            progress_frame, 
            variable=self.progress_var, 
            maximum=100, 
            mode='determinate'
        )
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        progress_frame.columnconfigure(0, weight=1)
        
        self.status_label = ttk.Label(progress_frame, textvariable=self.status_var)
        self.status_label.grid(row=1, column=0, sticky=tk.W)
        
        # Área de log
        log_frame = ttk.LabelFrame(main_frame, text="📋 Log de Proceso", padding="10")
        log_frame.grid(row=8, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(8, weight=1)
        
        # Crear el Text widget y scrollbar correctamente vinculados
        self.log_text = tk.Text(
            log_frame, 
            height=8, 
            wrap=tk.WORD, 
            state=tk.DISABLED,
            font=("Consolas", 9)
        )
        
        # Crear scrollbar vertical
        log_scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL)
        
        # Configurar la vinculación bidireccional entre Text y Scrollbar
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        log_scrollbar.configure(command=self.log_text.yview)
        
        # Posicionar los widgets usando grid
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        log_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Configurar estilos
        self.configure_styles()
    
    def configure_styles(self):
        """Configurar estilos personalizados"""
        style = ttk.Style()
        style.configure("Accent.TButton", font=("Arial", 10, "bold"))
    
    def browse_file(self):
        """Abrir diálogo para seleccionar archivo PDF"""
        filetypes = [
            ("Archivos PDF", "*.pdf"),
            ("Todos los archivos", "*.*")
        ]
        
        filename = filedialog.askopenfilename(
            title="Seleccionar archivo PDF",
            filetypes=filetypes
        )
        
        if filename:
            self.selected_file.set(filename)
            self.update_file_info(filename)
            self.auto_generate_output_name(filename)
    
    def browse_output_file(self):
        """Abrir diálogo para seleccionar ubicación de salida"""
        filename = filedialog.asksaveasfilename(
            title="Guardar archivo Excel como",
            defaultextension=".xlsx",
            filetypes=[
                ("Archivos Excel", "*.xlsx"),
                ("Todos los archivos", "*.*")
            ]
        )
        
        if filename:
            self.output_file.set(filename)
    
    def update_file_info(self, filepath):
        """Actualizar información del archivo seleccionado"""
        try:
            file_path = Path(filepath)
            file_size = file_path.stat().st_size / 1024 / 1024  # MB
            self.file_info_label.config(
                text=f"📏 Tamaño: {file_size:.2f} MB | 📁 {file_path.name}"
            )
        except Exception as e:
            self.file_info_label.config(text=f"⚠️ Error al obtener información: {e}")
    
    def auto_generate_output_name(self, input_file):
        """Generar automáticamente el nombre del archivo de salida"""
        input_path = Path(input_file)
        output_name = f"{input_path.stem}_balance_extraido.xlsx"
        output_path = input_path.parent / output_name
        self.output_file.set(str(output_path))
    
    def validate_inputs(self):
        """Validar las entradas del usuario"""
        if not self.selected_file.get():
            messagebox.showerror("Error", "Por favor selecciona un archivo PDF")
            return False
        
        if not Path(self.selected_file.get()).exists():
            messagebox.showerror("Error", "El archivo seleccionado no existe")
            return False
        
        if not self.selected_file.get().lower().endswith('.pdf'):
            messagebox.showerror("Error", "El archivo debe ser un PDF")
            return False
        
        if not self.output_file.get():
            messagebox.showerror("Error", "Por favor especifica un archivo de salida")
            return False
        
        # Verificar que el directorio de salida existe
        output_dir = Path(self.output_file.get()).parent
        if not output_dir.exists():
            messagebox.showerror("Error", f"El directorio de salida no existe: {output_dir}")
            return False
        
        return True
    
    def log_message(self, message):
        """Agregar mensaje al log"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.root.update_idletasks()
    
    def clear_log(self):
        """Limpiar el área de log"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
    
    def update_progress(self, value, status=""):
        """Actualizar barra de progreso y status"""
        self.progress_var.set(value)
        if status:
            self.status_var.set(status)
        self.root.update_idletasks()
    
    def process_file(self):
        """Procesar el archivo PDF en un hilo separado"""
        if not self.validate_inputs():
            return
        
        # Deshabilitar botón durante el procesamiento
        self.process_button.config(state="disabled")
        self.clear_log()
        
        # Ejecutar en hilo separado para no bloquear la UI
        thread = threading.Thread(target=self._process_file_thread)
        thread.daemon = True
        thread.start()
    
    def _process_file_thread(self):
        """Hilo para procesar el archivo"""
        try:
            self.log_message("🚀 Iniciando procesamiento...")
            self.update_progress(10, "Inicializando extractor...")
            
            # Crear instancia del extractor (tu algoritmo)
            extractor = BalanceExtractorEnhanced()
            
            self.update_progress(20, "Leyendo archivo PDF...")
            self.log_message(f"📖 Procesando: {Path(self.selected_file.get()).name}")
            
            # Extraer datos usando tu algoritmo
            data = extractor.extract_balance_data(self.selected_file.get())
            
            self.update_progress(70, "Datos extraídos, generando Excel...")
            
            if not data:
                raise Exception("No se encontraron datos válidos en el PDF")
            
            self.log_message(f"✅ Extraídos {len(data)} registros")
            
            # Guardar en Excel
            self.update_progress(90, "Guardando archivo Excel...")
            extractor.save_to_excel(data, self.output_file.get())
            
            self.update_progress(100, "✅ Proceso completado exitosamente")
            self.log_message(f"💾 Archivo guardado: {Path(self.output_file.get()).name}")
            self.log_message(f"📊 Total de cuentas procesadas: {len(data)}")
            
            # Mostrar mensaje de éxito
            self.root.after(0, self._show_success_message)
            
        except Exception as e:
            error_msg = f"❌ Error durante el procesamiento: {str(e)}"
            self.log_message(error_msg)
            self.update_progress(0, "Error en el procesamiento")
            
            # Mostrar error en el hilo principal
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
        
        finally:
            # Rehabilitar botón en el hilo principal
            self.root.after(0, lambda: self.process_button.config(state="normal"))
    
    def _show_success_message(self):
        """Mostrar mensaje de éxito y preguntar si abrir el archivo"""
        result = messagebox.askyesno(
            "✅ Proceso Completado",
            f"El archivo Excel se ha generado exitosamente.\n\n"
            f"📁 Ubicación: {self.output_file.get()}\n\n"
            f"¿Deseas abrir el archivo ahora?"
        )
        
        if result:
            self.open_output_file()
    
    def open_output_file(self):
        """Abrir el archivo de salida con la aplicación predeterminada"""
        try:
            import subprocess
            import sys
            
            if sys.platform == "win32":
                os.startfile(self.output_file.get())
            elif sys.platform == "darwin":  # macOS
                subprocess.run(["open", self.output_file.get()])
            else:  # Linux
                subprocess.run(["xdg-open", self.output_file.get()])
                
        except Exception as e:
            messagebox.showwarning(
                "Advertencia", 
                f"No se pudo abrir el archivo automáticamente: {e}"
            )
    
    def clear_form(self):
        """Limpiar el formulario"""
        self.selected_file.set("")
        self.output_file.set("")
        self.file_info_label.config(text="")
        self.update_progress(0, "Listo para procesar...")
        self.clear_log()

def main():
    """Función principal"""
    root = tk.Tk()
    app = PDFToExcelApp(root)
    
    # Configurar el icono si existe
    try:
        # Si tienes un archivo de icono, descomenta esta línea
        # root.iconbitmap("icon.ico")
        pass
    except:
        pass
    
    # Iniciar la aplicación
    root.mainloop()

if __name__ == "__main__":
    main()