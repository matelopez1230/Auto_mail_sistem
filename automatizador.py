import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

def resource_path(relative_path):
    """Obtiene la ruta absoluta al recurso, funciona para desarrollo y para PyInstaller"""
    try:
        # PyInstaller crea una carpeta temporal y almacena la ruta en _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

class EmailSenderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Envío Masivo de Correos")
        self.root.geometry("900x750")
        
        # Configura el ícono de la aplicación
        try:
            self.root.iconbitmap(resource_path('icon.ico'))
        except:
            pass  # Si no hay ícono, continúa sin él
        
        # Mapeo de columnas (personalizable por el usuario)
        self.column_mapping = {
            'email': 'email',  # Columna que contiene los emails
            'nombre': 'nombre'  # Columna que contiene los nombres
        }
        
        # Variables de configuración SMTP
        self.smtp_config = {
            'servidor': '',
            'puerto': 587,
            'email': '',
            'password': ''
        }
        
        self.df_clientes = None
        self.current_excel_path = ""
        
        # Crear pestañas
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill='both', expand=True)
        
        # Crear las pestañas
        self.create_config_tab()
        self.create_data_tab()
        self.create_message_tab()
        self.create_send_tab()
        self.create_mapping_tab()  # Nueva pestaña para mapeo de columnas
    
    def create_config_tab(self):
        """Crea la pestaña de configuración SMTP"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Configuración SMTP")
        
        frame = ttk.LabelFrame(tab, text="Configuración del Servidor de Correo")
        frame.pack(pady=10, padx=10, fill='x')
        
        # Proveedores comunes
        ttk.Label(frame, text="Proveedor:").grid(row=0, column=0, padx=5, pady=5, sticky='e')
        self.provider_var = tk.StringVar()
        providers = ttk.Combobox(frame, textvariable=self.provider_var, 
                                values=["Personalizado", "Gmail", "Outlook", "Yahoo", "Office365"])
        providers.grid(row=0, column=1, padx=5, pady=5, sticky='we')
        providers.bind('<<ComboboxSelected>>', self.update_smtp_settings)
        
        # Servidor SMTP
        ttk.Label(frame, text="Servidor SMTP:").grid(row=1, column=0, padx=5, pady=5, sticky='e')
        self.server_entry = ttk.Entry(frame)
        self.server_entry.grid(row=1, column=1, padx=5, pady=5, sticky='we')
        
        # Puerto
        ttk.Label(frame, text="Puerto:").grid(row=2, column=0, padx=5, pady=5, sticky='e')
        self.port_entry = ttk.Entry(frame)
        self.port_entry.grid(row=2, column=1, padx=5, pady=5, sticky='we')
        self.port_entry.insert(0, "587")
        
        # Email
        ttk.Label(frame, text="Email:").grid(row=3, column=0, padx=5, pady=5, sticky='e')
        self.email_entry = ttk.Entry(frame)
        self.email_entry.grid(row=3, column=1, padx=5, pady=5, sticky='we')
        
        # Contraseña
        ttk.Label(frame, text="Contraseña:").grid(row=4, column=0, padx=5, pady=5, sticky='e')
        self.password_entry = ttk.Entry(frame, show="*")
        self.password_entry.grid(row=4, column=1, padx=5, pady=5, sticky='we')
        
        # Botón de prueba
        ttk.Button(frame, text="Probar Conexión", command=self.test_connection).grid(row=5, column=1, pady=10)
    
    def create_data_tab(self):
        """Crea la pestaña para importar datos de Excel"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Datos de Clientes")
        
        frame = ttk.LabelFrame(tab, text="Importar Datos desde Excel")
        frame.pack(pady=10, padx=10, fill='both', expand=True)
        
        # Botón para seleccionar archivo
        ttk.Button(frame, text="Seleccionar Archivo Excel", command=self.load_excel_file).pack(pady=10)
        
        # Vista previa de datos
        self.data_preview = ttk.Treeview(frame)
        self.data_preview.pack(fill='both', expand=True, pady=5)
        
        # Barra de desplazamiento
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self.data_preview.yview)
        scrollbar.pack(side='right', fill='y')
        self.data_preview.configure(yscrollcommand=scrollbar.set)
        
        # Etiqueta de información
        self.data_info = ttk.Label(frame, text="No se ha cargado ningún archivo")
        self.data_info.pack(pady=5)
    
    def create_message_tab(self):
        """Crea la pestaña para editar el mensaje"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Editor de Mensaje")
        
        frame = ttk.LabelFrame(tab, text="Componer Mensaje")
        frame.pack(pady=10, padx=10, fill='both', expand=True)
        
        # Instrucciones
        ttk.Label(frame, 
                 text="Escribe tu mensaje. Usa {nombre_columna} para incluir datos del Excel").pack(pady=5)
        
        # Editor de texto
        self.message_editor = scrolledtext.ScrolledText(frame, wrap=tk.WORD, height=15)
        self.message_editor.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Botón de previsualización
        ttk.Button(frame, text="Previsualizar Mensaje", command=self.preview_message).pack(pady=5)
        
        # Frame para previsualización
        preview_frame = ttk.LabelFrame(tab, text="Previsualización del Mensaje")
        preview_frame.pack(pady=10, padx=10, fill='both', expand=True)
        
        self.preview_text = scrolledtext.ScrolledText(preview_frame, wrap=tk.WORD, height=10, state='disabled')
        self.preview_text.pack(fill='both', expand=True, padx=5, pady=5)
    
    def create_send_tab(self):
        """Crea la pestaña para enviar los correos"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Enviar Correos")
        
        frame = ttk.LabelFrame(tab, text="Progreso del Envío")
        frame.pack(pady=10, padx=10, fill='both', expand=True)
        
        # Barra de progreso
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill='x', padx=10, pady=10)
        
        # Contador de envíos
        self.counter_label = ttk.Label(frame, text="0/0 correos enviados")
        self.counter_label.pack(pady=5)
        
        # Log de envío
        self.log_text = scrolledtext.ScrolledText(frame, height=15, state='disabled')
        self.log_text.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Botón de envío
        ttk.Button(frame, text="Iniciar Envío", command=self.start_sending).pack(pady=10)
    
    def create_mapping_tab(self):
        """Crea la pestaña para mapeo de columnas"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Mapeo de Columnas")
        
        frame = ttk.LabelFrame(tab, text="Configurar Mapeo de Columnas")
        frame.pack(pady=10, padx=10, fill='both', expand=True)
        
        # Información
        ttk.Label(frame, text="Asigna las columnas de tu Excel a los campos requeridos").pack(pady=5)
        
        # Frame para mapeos
        map_frame = ttk.Frame(frame)
        map_frame.pack(fill='x', pady=10)
        
        # Email
        ttk.Label(map_frame, text="Columna de Email:").grid(row=0, column=0, padx=5, pady=5, sticky='e')
        self.email_col_var = tk.StringVar()
        self.email_col_entry = ttk.Entry(map_frame, textvariable=self.email_col_var)
        self.email_col_entry.grid(row=0, column=1, padx=5, pady=5, sticky='we')
        self.email_col_var.set(self.column_mapping['email'])
        
        # Nombre
        ttk.Label(map_frame, text="Columna de Nombre:").grid(row=1, column=0, padx=5, pady=5, sticky='e')
        self.name_col_var = tk.StringVar()
        self.name_col_entry = ttk.Entry(map_frame, textvariable=self.name_col_var)
        self.name_col_entry.grid(row=1, column=1, padx=5, pady=5, sticky='we')
        self.name_col_var.set(self.column_mapping['nombre'])
        
        # Botón para guardar configuración
        ttk.Button(frame, text="Guardar Mapeo", command=self.save_mapping).pack(pady=10)
        
        # Información sobre columnas disponibles
        self.available_cols_label = ttk.Label(frame, text="Columnas disponibles: Ningún archivo cargado")
        self.available_cols_label.pack(pady=5)
    
    def save_mapping(self):
        """Guarda la configuración de mapeo de columnas"""
        self.column_mapping['email'] = self.email_col_var.get()
        self.column_mapping['nombre'] = self.name_col_var.get()
        messagebox.showinfo("Éxito", "Mapeo de columnas guardado correctamente")
    
    def update_smtp_settings(self, event):
        """Actualiza la configuración SMTP según el proveedor seleccionado"""
        provider = self.provider_var.get()
        
        # Borra los valores actuales
        self.server_entry.delete(0, tk.END)
        self.port_entry.delete(0, tk.END)
        
        # Configura los valores según el proveedor
        if provider == "Gmail":
            self.server_entry.insert(0, "smtp.gmail.com")
            self.port_entry.insert(0, "587")
        elif provider == "Outlook":
            self.server_entry.insert(0, "smtp-mail.outlook.com")
            self.port_entry.insert(0, "587")
        elif provider == "Yahoo":
            self.server_entry.insert(0, "smtp.mail.yahoo.com")
            self.port_entry.insert(0, "465")
        elif provider == "Office365":
            self.server_entry.insert(0, "smtp.office365.com")
            self.port_entry.insert(0, "587")
    
    def test_connection(self):
        """Prueba la conexión con el servidor SMTP"""
        try:
            # Obtiene la configuración desde la interfaz
            self.smtp_config = {
                'servidor': self.server_entry.get(),
                'puerto': int(self.port_entry.get()),
                'email': self.email_entry.get(),
                'password': self.password_entry.get()
            }
            
            # Validación básica
            if not all(self.smtp_config.values()):
                raise ValueError("Todos los campos son requeridos")
            
            # Intenta la conexión
            if self.smtp_config['puerto'] == 465:
                server = smtplib.SMTP_SSL(self.smtp_config['servidor'], self.smtp_config['puerto'])
            else:
                server = smtplib.SMTP(self.smtp_config['servidor'], self.smtp_config['puerto'])
                server.starttls()
            
            server.login(self.smtp_config['email'], self.smtp_config['password'])
            server.quit()
            
            messagebox.showinfo("Éxito", "Conexión exitosa con el servidor SMTP")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo conectar: {str(e)}")
    
    def load_excel_file(self):
        """Carga un archivo Excel con los datos de los clientes"""
        filepath = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        
        if filepath:
            try:
                # Lee el archivo Excel
                self.df_clientes = pd.read_excel(filepath)
                self.current_excel_path = filepath
                
                # Actualiza la vista previa
                self.update_data_preview()
                
                # Muestra las columnas disponibles para mapeo
                available_cols = ", ".join(self.df_clientes.columns)
                self.available_cols_label.config(
                    text=f"Columnas disponibles: {available_cols}"
                )
                
                messagebox.showinfo(
                    "Éxito", 
                    f"Archivo cargado: {os.path.basename(filepath)}\n"
                    f"{len(self.df_clientes)} registros encontrados"
                )
            
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo cargar el archivo: {str(e)}")
                self.df_clientes = None
                self.current_excel_path = ""
    
    def update_data_preview(self):
        """Actualiza la vista previa de los datos del Excel"""
        # Limpia la vista previa
        self.data_preview.delete(*self.data_preview.get_children())
        
        if self.df_clientes is not None:
            # Configura las columnas
            columns = list(self.df_clientes.columns)
            self.data_preview['columns'] = columns
            
            # Configura los encabezados
            for col in columns:
                self.data_preview.heading(col, text=col)
                self.data_preview.column(col, width=100, minwidth=50, anchor='w')
            
            # Agrega los datos (solo las primeras 10 filas)
            for _, row in self.df_clientes.head(10).iterrows():
                self.data_preview.insert("", tk.END, values=list(row))
            
            # Actualiza la información
            self.data_info.config(
                text=f"Archivo: {os.path.basename(self.current_excel_path)}\n"
                     f"Registros: {len(self.df_clientes)} | Columnas: {len(self.df_clientes.columns)}"
            )
        else:
            self.data_info.config(text="No se ha cargado ningún archivo")
    
    def validate_excel_structure(self):
        """Verifica que el Excel tenga las columnas necesarias según el mapeo"""
        if self.df_clientes is None:
            return False
        
        # Verifica que las columnas mapeadas existan en el DataFrame
        missing_columns = [
            col for col in self.column_mapping.values() 
            if col not in self.df_clientes.columns
        ]
        
        if missing_columns:
            messagebox.showerror(
                "Error", 
                f"El archivo Excel no tiene las columnas requeridas.\n"
                f"Columnas faltantes: {', '.join(missing_columns)}\n"
                f"Por favor, configura el mapeo correctamente."
            )
            return False
        
        return True
    
    def preview_message(self):
        """Muestra una previsualización del mensaje con datos reales"""
        if not self.validate_excel_structure():
            return
        
        message_text = self.message_editor.get("1.0", tk.END).strip()
        if not message_text:
            messagebox.showwarning("Advertencia", "Escribe un mensaje en el editor")
            return
        
        try:
            # Toma la primera fila como ejemplo
            sample_row = self.df_clientes.iloc[0]
            preview = message_text
            
            # Reemplaza todos los marcadores {columna} con valores reales
            for col in self.df_clientes.columns:
                if pd.notna(sample_row[col]):
                    preview = preview.replace(f"{{{col}}}", str(sample_row[col]))
            
            # Muestra la previsualización
            self.preview_text.config(state='normal')
            self.preview_text.delete('1.0', tk.END)
            self.preview_text.insert(tk.END, preview)
            self.preview_text.config(state='disabled')
        
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo generar la previsualización: {str(e)}")
    
    def start_sending(self):
        """Inicia el proceso de envío de correos"""
        # Validaciones previas
        if not all([self.smtp_config['servidor'], self.smtp_config['email'], self.smtp_config['password']]):
            messagebox.showerror("Error", "Configura primero la conexión SMTP")
            return
        
        if not self.validate_excel_structure():
            return
        
        message_text = self.message_editor.get("1.0", tk.END).strip()
        if not message_text:
            messagebox.showerror("Error", "Escribe un mensaje en el editor")
            return
        
        # Confirmación del usuario
        if not messagebox.askyesno(
            "Confirmar", 
            f"¿Estás seguro de enviar {len(self.df_clientes)} correos?\n"
            "Esta operación puede tardar varios minutos."
        ):
            return
        
        # Deshabilita las pestañas durante el envío
        for i in range(self.notebook.index("end")):
            self.notebook.tab(i, state='disabled')
        
        # Inicia el envío
        self.root.after(100, lambda: self.send_emails(message_text))
    
    def send_emails(self, message_text):
        """Realiza el envío masivo de correos"""
        total_emails = len(self.df_clientes)
        success_count = 0
        failure_count = 0
        
        self.log_message("=== INICIANDO ENVÍO DE CORREOS ===")
        
        try:
            # Conexión SMTP
            if self.smtp_config['puerto'] == 465:
                server = smtplib.SMTP_SSL(self.smtp_config['servidor'], self.smtp_config['puerto'])
            else:
                server = smtplib.SMTP(self.smtp_config['servidor'], self.smtp_config['puerto'])
                server.starttls()
            
            # Autenticación
            server.login(self.smtp_config['email'], self.smtp_config['password'])
            
            # Procesa cada cliente
            for i, (_, row) in enumerate(self.df_clientes.iterrows()):
                try:
                    # Personaliza el mensaje
                    email = row[self.column_mapping['email']]
                    personalized_msg = message_text
                    
                    # Reemplaza todos los marcadores
                    for col in self.df_clientes.columns:
                        if pd.notna(row[col]):
                            personalized_msg = personalized_msg.replace(f"{{{col}}}", str(row[col]))
                    
                    # Prepara el correo
                    msg = MIMEMultipart()
                    msg['From'] = self.smtp_config['email']
                    msg['To'] = email
                    msg['Subject'] = "Mensaje personalizado"
                    
                    msg.attach(MIMEText(personalized_msg, 'plain'))
                    
                    # Envía el correo
                    server.send_message(msg)
                    success_count += 1
                    self.log_message(f"✓ Enviado a {email}")
                
                except Exception as e:
                    failure_count += 1
                    self.log_message(f"✗ Error con {email}: {str(e)}")
                
                # Actualiza el progreso
                progress = (i + 1) / total_emails * 100
                self.progress_var.set(progress)
                self.counter_label.config(text=f"{i+1}/{total_emails} correos enviados")
                self.root.update()
            
            # Cierra la conexión
            server.quit()
            
            # Resultado final
            self.log_message(f"=== ENVÍO COMPLETADO ===")
            self.log_message(f"Correctos: {success_count}, Fallidos: {failure_count}")
            messagebox.showinfo(
                "Completado", 
                f"Proceso terminado:\n"
                f"- Correos enviados: {success_count}\n"
                f"- Errores: {failure_count}"
            )
        
        except Exception as e:
            self.log_message(f"ERROR GLOBAL: {str(e)}")
            messagebox.showerror("Error", f"Error en el envío: {str(e)}")
        
        finally:
            # Rehabilita las pestañas
            for i in range(self.notebook.index("end")):
                self.notebook.tab(i, state='normal')
            
            # Reinicia el progreso
            self.progress_var.set(0)
    
    def log_message(self, message):
        """Agrega un mensaje al registro de eventos"""
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')
        self.root.update()

if __name__ == "__main__":
    root = tk.Tk()
    app = EmailSenderApp(root)
    root.mainloop()