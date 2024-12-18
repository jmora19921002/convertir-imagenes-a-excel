import os  
import tkinter as tk  
from tkinter import filedialog, messagebox, Canvas, Scrollbar  
import customtkinter as ctk  
import pandas as pd  
from PIL import Image, ImageTk  
import pytesseract  

class ImageTextExtractorApp:  
    def __init__(self):  
        # Configuración de CustomTkinter  
        ctk.set_appearance_mode("System")  
        ctk.set_default_color_theme("blue")  

        # Ventana principal  
        self.root = ctk.CTk()  
        self.root.title("Extractor de Texto de Imágenes")  
        self.root.geometry("800x900")  

        # Lista para almacenar imágenes y sus textos  
        self.image_data = []  

        # Crear widgets  
        self.create_widgets()  

    def create_widgets(self):  
        # Frame principal  
        self.main_frame = ctk.CTkFrame(self.root)  
        self.main_frame.pack(padx=20, pady=20, fill="both", expand=True)  

        # Botón para seleccionar imágenes  
        self.select_button = ctk.CTkButton(  
            self.main_frame,   
            text="Seleccionar Imágenes",   
            command=self.select_images  
        )  
        self.select_button.pack(pady=10)  

        # Frame para listbox y preview de imagen  
        self.list_image_frame = ctk.CTkFrame(self.main_frame)  
        self.list_image_frame.pack(pady=10, fill="both", expand=True)  

        # Crear frame para listbox con scrollbar  
        self.listbox_frame = ctk.CTkFrame(self.list_image_frame)  
        self.listbox_frame.pack(side="left", fill="both", expand=True, padx=(0,10))  

        # Lista de imágenes seleccionadas con scrollbar  
        self.image_listbox = tk.Listbox(  
            self.listbox_frame,   
            width=50,   
            height=10  
        )  
        self.image_listbox.pack(side="left", fill="both", expand=True)  
        
        # Scrollbar para listbox  
        self.listbox_scrollbar = tk.Scrollbar(self.listbox_frame)  
        self.listbox_scrollbar.pack(side="right", fill="y")  
        
        self.image_listbox.config(yscrollcommand=self.listbox_scrollbar.set)  
        self.listbox_scrollbar.config(command=self.image_listbox.yview)  

        # Evento para mostrar imagen  
        self.image_listbox.bind('<<ListboxSelect>>', self.show_image_preview)  

        # Frame para preview de imagen  
        self.preview_frame = ctk.CTkFrame(self.list_image_frame)  
        self.preview_frame.pack(side="right", fill="both", expand=True)  

        # Canvas para preview de imagen  
        self.image_preview_canvas = Canvas(self.preview_frame, width=300, height=300)  
        self.image_preview_canvas.pack(fill="both", expand=True)  

        # Barra de progreso  
        self.progress_bar = ctk.CTkProgressBar(self.main_frame)  
        self.progress_bar.pack(pady=10, fill="x")  
        self.progress_bar.set(0)  # Iniciar en 0  

        # Área de configuración de OCR  
        self.ocr_frame = ctk.CTkFrame(self.main_frame)  
        self.ocr_frame.pack(pady=10, fill="x")  

        # Selector de idioma  
        self.language_label = ctk.CTkLabel(self.ocr_frame, text="Idioma:")  
        self.language_label.pack(side="left", padx=5)  
        
        self.language_var = ctk.StringVar(value="spa")  
        self.language_menu = ctk.CTkOptionMenu(  
            self.ocr_frame,   
            values=["spa", "eng", "fra", "deu"],  
            variable=self.language_var  
        )  
        self.language_menu.pack(side="left", padx=5)  

        # Botón para extraer texto  
        self.extract_button = ctk.CTkButton(  
            self.main_frame,   
            text="Extraer Texto",   
            command=self.extract_text,  
            state="disabled"  
        )  
        self.extract_button.pack(pady=10)  

        # Botón para generar Excel  
        self.excel_button = ctk.CTkButton(  
            self.main_frame,   
            text="Generar Excel",   
            command=self.generate_excel,  
            state="disabled"  
        )  
        self.excel_button.pack(pady=10)  

    def show_image_preview(self, event):  
        # Obtener índice seleccionado  
        if not self.image_listbox.curselection():  
            return  
        
        index = self.image_listbox.curselection()[0]  
        image_path = self.image_data[index]['ruta']  

        # Abrir y ajustar imagen  
        image = Image.open(image_path)  
        image.thumbnail((300, 300))  # Ajustar tamaño  
        photo = ImageTk.PhotoImage(image)  

        # Limpiar canvas y mostrar imagen  
        self.image_preview_canvas.delete("all")  
        self.image_preview_canvas.create_image(150, 150, image=photo)  
        self.image_preview_canvas.image = photo  # Mantener referencia  

    def select_images(self):  
        # Diálogo para seleccionar imágenes  
        filetypes = [  
            ("Archivos de imagen", "*.jpg *.jpeg *.png *.gif *.bmp")  
        ]  
        selected_images = filedialog.askopenfilenames(  
            title="Seleccionar Imágenes",   
            filetypes=filetypes  
        )  

        # Limpiar listas previas  
        self.image_data.clear()  
        self.image_listbox.delete(0, tk.END)  

        # Agregar imágenes  
        for image_path in selected_images:  
            self.image_data.append({  
                'ruta': image_path,  
                'nombre': os.path.basename(image_path),  
                'texto': ''  
            })  
            self.image_listbox.insert(tk.END, os.path.basename(image_path))  

        # Habilitar botones  
        if self.image_data:  
            self.extract_button.configure(state="normal")  
            self.excel_button.configure(state="disabled")  

    def extract_text(self):  
        # Configurar Tesseract (asegúrate de tener Tesseract instalado)  
        try:  
            # Resetear barra de progreso  
            self.progress_bar.set(0)  
            total_images = len(self.image_data)  

            # Iterar sobre las imágenes y extraer texto  
            for index, img_info in enumerate(self.image_data, 1):  
                # Abrir imagen  
                image = Image.open(img_info['ruta'])  
                
                # Extraer texto con Tesseract  
                texto = pytesseract.image_to_string(  
                    image,   
                    lang=self.language_var.get()  
                )  
                
                # Guardar texto  
                img_info['texto'] = texto.strip()  
                
                # Actualizar barra de progreso  
                progress = index / total_images  
                self.progress_bar.set(progress)  
                
                # Actualizar interfaz gráfica para mostrar progreso  
                self.root.update_idletasks()  

            # Asegurar que la barra de progreso llegue al 100%  
            self.progress_bar.set(1)  

            # Habilitar botón de Excel  
            self.excel_button.configure(state="normal")  
            
            # Mostrar mensaje de éxito  
            messagebox.showinfo("Éxito", "Texto extraído de todas las imágenes")  

        except Exception as e:  
            # Resetear barra de progreso en caso de error  
            self.progress_bar.set(0)  
            messagebox.showerror("Error", f"Error al extraer texto: {str(e)}")  

    def generate_excel(self):  
        # Crear DataFrame  
        df = pd.DataFrame(self.image_data)  
        
        # Diálogo para guardar archivo  
        excel_path = filedialog.asksaveasfilename(  
            defaultextension=".xlsx",  
            filetypes=[("Archivos Excel", "*.xlsx")]  
        )  
        
        if excel_path:  
            # Guardar en Excel  
            df.to_excel(excel_path, index=False)  
            messagebox.showinfo("Éxito", f"Archivo Excel guardado en {excel_path}")  

    def run(self):  
        self.root.mainloop()  

# Ejecutar la aplicación  
if __name__ == "__main__":  
    app = ImageTextExtractorApp()  
    app.run()