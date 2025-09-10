import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import threading
import os
import pdfplumber
from test_pdf import BalanceExtractorEnhanced  # Importamos tu algoritmo

class PDFToExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("🏦 Extractor PDF a Excel - Balance de Comprobación")
        self.root.geometry("700x650")
        self.root.resizable(True, True)
        
        # Configurar tema oscuro
        self.setup_dark_theme()
        
        # Variables
        self.selected_file = tk.StringVar()
        self.output_file = tk.StringVar()
        self.progress_var = tk.DoubleVar()
        self.status_var = tk.StringVar(value="✨ Listo para procesar...")
        
        # Crear la interfaz
        self.create_widgets()
        
        # Centrar ventana
        self.center_window()
        
        # Aplicar efectos hover
        self.setup_hover_effects()
    
    def setup_dark_theme(self):
        """Configurar tema oscuro moderno"""
        # Colores del tema oscuro
        self.colors = {
            'bg_main': '#1e1e1e',           # Fondo principal
            'bg_secondary': '#2d2d2d',      # Fondo secundario
            'bg_accent': '#404040',         # Fondo de acentos
            'fg_primary': '#ffffff',        # Texto principal
            'fg_secondary': "#ffffff",      # Texto secundario
            'accent_blue': "#3779ab",       # Azul Microsoft
            'accent_light': "#5c89bb",      # Azul claro
            'success': '#107c10',           # Verde éxito
            'warning': '#ff8c00',           # Naranja advertencia
            'error': '#d13438',             # Rojo error
            'border': '#505050',            # Bordes
            'hover': '#383838'              # Hover
        }
        
        # Configurar colores de la ventana principal
        self.root.configure(bg=self.colors['bg_main'])
        
        # Configurar estilo ttk
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Configurar estilos personalizados
        self.configure_modern_styles()
    
    def configure_modern_styles(self):
        """Configurar estilos modernos para widgets ttk"""
        
        # Frame principal
        self.style.configure('Modern.TFrame',
                           background=self.colors['bg_main'],
                           relief='flat')
        
        # Labels principales
        self.style.configure('Title.TLabel',
                           background=self.colors['bg_main'],
                           foreground=self.colors['accent_blue'],
                           font=('Segoe UI', 18, 'bold'))
        
        self.style.configure('Subtitle.TLabel',
                           background=self.colors['bg_main'],
                           foreground=self.colors['fg_secondary'],
                           font=('Segoe UI', 11))
        
        self.style.configure('Modern.TLabel',
                           background=self.colors['bg_main'],
                           foreground=self.colors['fg_primary'],
                           font=('Segoe UI', 10))
        
        self.style.configure('Info.TLabel',
                           background=self.colors['bg_main'],
                           foreground=self.colors['fg_secondary'],
                           font=('Segoe UI', 9))
        
        # LabelFrames modernos
        self.style.configure('Modern.TLabelframe',
                           background=self.colors['bg_main'],
                           borderwidth=2,
                           relief='solid',
                           bordercolor=self.colors['border'])
        
        self.style.configure('Modern.TLabelframe.Label',
                           background=self.colors['bg_main'],
                           foreground=self.colors['accent_light'],
                           font=('Segoe UI', 11, 'bold'))
        
        # Entries modernos
        self.style.configure('Modern.TEntry',
                           fieldbackground=self.colors['bg_secondary'],
                           foreground=self.colors['fg_primary'],
                           bordercolor=self.colors['border'],
                           lightcolor=self.colors['accent_blue'],
                           darkcolor=self.colors['accent_blue'],
                           borderwidth=2,
                           insertcolor=self.colors['fg_primary'],
                           font=('Segoe UI', 10))
        
        # Botones principales
        self.style.configure('Primary.TButton',
                           background=self.colors['accent_blue'],
                           foreground='white',
                           borderwidth=0,
                           focuscolor='none',
                           padding=(20, 10),
                           font=('Segoe UI', 11, 'bold'))
        
        self.style.map('Primary.TButton',
                      background=[('active', self.colors['accent_light']),
                                ('pressed', '#005a9e')])
        
        # Botones secundarios
        self.style.configure('Secondary.TButton',
                           background=self.colors['bg_secondary'],
                           foreground=self.colors['fg_primary'],
                           borderwidth=2,
                           bordercolor=self.colors['border'],
                           focuscolor='none',
                           padding=(15, 8),
                           font=('Segoe UI', 10))
        
        self.style.map('Secondary.TButton',
                      background=[('active', self.colors['hover']),
                                ('pressed', self.colors['bg_accent'])])
        
        # Progress bar moderna
        self.style.configure('Modern.Horizontal.TProgressbar',
                           background=self.colors['accent_blue'],
                           troughcolor=self.colors['bg_secondary'],
                           borderwidth=0,
                           lightcolor=self.colors['accent_blue'],
                           darkcolor=self.colors['accent_blue'])
        
        # Separadores
        self.style.configure('Modern.TSeparator',
                           background=self.colors['border'])
    
    def center_window(self):
        """Centrar la ventana en la pantalla"""
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (self.root.winfo_width() // 2)
        y = (self.root.winfo_screenheight() // 2) - (self.root.winfo_height() // 2)
        self.root.geometry(f"+{x}+{y}")
    
    def create_widgets(self):
        """Crear todos los widgets de la interfaz"""
        
        # Frame principal con padding
        main_frame = ttk.Frame(self.root, padding="25", style='Modern.TFrame')
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configurar grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Header con título y subtítulo
        header_frame = ttk.Frame(main_frame, style='Modern.TFrame')
        header_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 20))
        header_frame.columnconfigure(0, weight=1)
        
        title_label = ttk.Label(
            header_frame, 
            text="🏦 Extractor PDF a Excel", 
            style='Title.TLabel'
        )
        title_label.grid(row=0, column=0)
        
        subtitle_label = ttk.Label(
            header_frame, 
            text="Balance de Comprobación - Banco de la Nación",
            style='Subtitle.TLabel'
        )
        subtitle_label.grid(row=1, column=0, pady=(5, 0))
        
        # Separador elegante
        separator1 = ttk.Separator(main_frame, orient='horizontal', style='Modern.TSeparator')
        separator1.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 20))
        
        # Sección de selección de archivo
        file_frame = ttk.LabelFrame(main_frame, text="📄 Seleccionar Archivo PDF", 
                                  padding="15", style='Modern.TLabelframe')
        file_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 20))
        file_frame.columnconfigure(1, weight=1)
        
        ttk.Label(file_frame, text="Archivo:", style='Modern.TLabel').grid(
            row=0, column=0, sticky=tk.W, padx=(0, 15))
        
        self.file_entry = ttk.Entry(file_frame, textvariable=self.selected_file, 
                                  state="readonly", style='Modern.TEntry')
        self.file_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 15))
        
        self.browse_button = ttk.Button(
            file_frame, 
            text="📁 Examinar", 
            command=self.browse_file,
            style='Secondary.TButton'
        )
        self.browse_button.grid(row=0, column=2)
        
        # Información del archivo
        self.file_info_label = ttk.Label(file_frame, text="", style='Info.TLabel')
        self.file_info_label.grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=(10, 0))
        
        # Sección de archivo de salida
        output_frame = ttk.LabelFrame(main_frame, text="💾 Archivo de Salida", 
                                    padding="15", style='Modern.TLabelframe')
        output_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 20))
        output_frame.columnconfigure(1, weight=1)
        
        ttk.Label(output_frame, text="Guardar como:", style='Modern.TLabel').grid(
            row=0, column=0, sticky=tk.W, padx=(0, 15))
        
        self.output_entry = ttk.Entry(output_frame, textvariable=self.output_file, 
                                    style='Modern.TEntry')
        self.output_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 15))
        
        self.output_browse_button = ttk.Button(
            output_frame, 
            text="📁 Guardar en", 
            command=self.browse_output_file,
            style='Secondary.TButton'
        )
        self.output_browse_button.grid(row=0, column=2)
        
        # Separador
        separator2 = ttk.Separator(main_frame, orient='horizontal', style='Modern.TSeparator')
        separator2.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 20))
        
        # Botones principales con mejor espaciado
        button_frame = ttk.Frame(main_frame, style='Modern.TFrame')
        button_frame.grid(row=5, column=0, columnspan=3, pady=(0, 20))
        
        self.process_button = ttk.Button(
            button_frame,
            text="Procesar PDF",
            command=self.process_file,
            style="Primary.TButton"
        )
        self.process_button.pack(side=tk.LEFT, padx=(0, 15))
        
        self.clear_button = ttk.Button(
            button_frame,
            text="🗑️ Limpiar",
            command=self.clear_form,
            style="Secondary.TButton"
        )
        self.clear_button.pack(side=tk.LEFT)
        
        # Barra de progreso moderna
        progress_frame = ttk.LabelFrame(main_frame, text="📊 Progreso", 
                                      padding="15", style='Modern.TLabelframe')
        progress_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 20))
        progress_frame.columnconfigure(0, weight=1)
        
        self.progress_bar = ttk.Progressbar(
            progress_frame, 
            variable=self.progress_var, 
            maximum=100, 
            mode='determinate',
            style='Modern.Horizontal.TProgressbar'
        )
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.status_label = ttk.Label(progress_frame, textvariable=self.status_var, 
                                    style='Modern.TLabel')
        self.status_label.grid(row=1, column=0, sticky=tk.W)
        
        # Área de log moderna
        log_frame = ttk.LabelFrame(main_frame, text="📋 Log de Proceso", 
                                 padding="15", style='Modern.TLabelframe')
        log_frame.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(7, weight=1)
        
        # Text widget con colores oscuros
        self.log_text = tk.Text(
            log_frame, 
            height=10, 
            wrap=tk.WORD, 
            state=tk.DISABLED,
            font=("Consolas", 9),
            bg=self.colors['bg_secondary'],
            fg=self.colors['fg_primary'],
            insertbackground=self.colors['accent_blue'],
            selectbackground=self.colors['accent_blue'],
            selectforeground='white',
            relief='flat',
            borderwidth=2,
            highlightbackground=self.colors['border'],
            highlightcolor=self.colors['accent_blue']
        )
        
        # Scrollbar con tema oscuro
        log_scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL)
        
        # Configurar la vinculación bidireccional entre Text y Scrollbar
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        log_scrollbar.configure(command=self.log_text.yview)
        
        # Posicionar los widgets usando grid
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        log_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
    
    def setup_hover_effects(self):
        """Configurar efectos hover para botones"""
        def on_enter(event, button, style_active):
            button.configure(style=style_active)
        
        def on_leave(event, button, style_normal):
            button.configure(style=style_normal)
        
        # Hover para botón principal
        self.process_button.bind("<Enter>", 
            lambda e: self.style.configure('Primary.TButton', 
                                         background=self.colors['accent_light']))
        self.process_button.bind("<Leave>", 
            lambda e: self.style.configure('Primary.TButton', 
                                         background=self.colors['accent_blue']))
        
        # Hover para botones secundarios
        buttons = [self.browse_button, self.output_browse_button, self.clear_button]
        for btn in buttons:
            btn.bind("<Enter>", 
                lambda e: self.style.configure('Secondary.TButton', 
                                             background=self.colors['hover']))
            btn.bind("<Leave>", 
                lambda e: self.style.configure('Secondary.TButton', 
                                             background=self.colors['bg_secondary']))
    
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
        
        try:
            from test_pdf import BalanceExtractorEnhanced
            temp_extractor = BalanceExtractorEnhanced()
            
            with pdfplumber.open(input_file) as pdf:
                temp_extractor.extracted_date = temp_extractor._extract_date_from_pdf(pdf)
            
            suggested_name = temp_extractor.get_excel_filename()
            output_path = input_path.parent / suggested_name
        
        except Exception as e:
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
        """Agregar mensaje al log con colores"""
        self.log_text.config(state=tk.NORMAL)
        
        # Configurar tags para diferentes tipos de mensajes
        self.log_text.tag_configure("success", foreground=self.colors['success'])
        self.log_text.tag_configure("warning", foreground=self.colors['warning'])
        self.log_text.tag_configure("error", foreground=self.colors['error'])
        self.log_text.tag_configure("info", foreground=self.colors['accent_light'])
        
        # Determinar el tipo de mensaje
        if "✅" in message or "🎉" in message or "💾" in message:
            tag = "success"
        elif "⚠️" in message or "advertencia" in message.lower():
            tag = "warning"
        elif "❌" in message or "error" in message.lower():
            tag = "error"
        elif "🚀" in message or "📖" in message or "📊" in message:
            tag = "info"
        else:
            tag = None
        
        if tag:
            self.log_text.insert(tk.END, f"{message}\n", tag)
        else:
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
            self.update_progress(10, "⚡ Inicializando extractor...")
            
            # Crear instancia del extractor (tu algoritmo)
            extractor = BalanceExtractorEnhanced()
            
            self.update_progress(20, "📖 Leyendo archivo PDF...")
            self.log_message(f"📖 Procesando: {Path(self.selected_file.get()).name}")
            
            # Extraer datos usando tu algoritmo
            data = extractor.extract_balance_data(self.selected_file.get())
            
            self.update_progress(70, "⚙️ Datos extraídos, generando Excel...")
            
            if not data:
                raise Exception("No se encontraron datos válidos en el PDF")
            
            self.log_message(f"✅ Extraídos {len(data)} registros")
            
            # Guardar en Excel
            self.update_progress(90, "💾 Guardando archivo Excel...")
            extractor.save_to_excel(data, self.output_file.get())
            
            self.update_progress(100, "🎉 Proceso completado exitosamente")
            self.log_message(f"💾 Archivo guardado: {Path(self.output_file.get()).name}")
            self.log_message(f"📊 Total de cuentas procesadas: {len(data)}")
            
            # Mostrar mensaje de éxito
            self.root.after(0, self._show_success_message)
            
        except Exception as e:
            error_msg = f"❌ Error durante el procesamiento: {str(e)}"
            self.log_message(error_msg)
            self.update_progress(0, "❌ Error en el procesamiento")
            
            # Mostrar error en el hilo principal
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
        
        finally:
            # Rehabilitar botón en el hilo principal
            self.root.after(0, lambda: self.process_button.config(state="normal"))
    
    def _show_success_message(self):
        """Mostrar mensaje de éxito y preguntar si abrir el archivo"""
        result = messagebox.askyesno(
            "🎉 Proceso Completado",
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
        self.update_progress(0, "✨ Listo para procesar...")
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