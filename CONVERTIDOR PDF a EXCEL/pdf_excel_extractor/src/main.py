import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    from pdf_parser import PDFParser
    from excel_writer import ExcelWriter
    print("Módulos importados correctamente")
except ImportError as e:
    print(f"Error al importar módulos: {e}")
    sys.exit(1)

class PDFExcelExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF a Excel  Cliente_DUER")
        
        # Establecer el tamaño de la ventana (reducir altura)
        self.root.geometry("400x100")
        
        # Agregar etiqueta de texto encima del botón
        self.label = tk.Label(root, text="SOLO PARA TALLAS XS, S, M, L, XL, XXL", 
                             font=("Arial", 10), fg="blue")
        self.label.pack(pady=(10, 5))
        
        self.button = ttk.Button(root, text="CLICK", command=self.extract_and_save)
        self.button.pack(padx=40, pady=10)

    def extract_and_save(self):
        file_paths = filedialog.askopenfilenames(
            title="Selecciona archivos PDF", filetypes=[("PDF Files", "*.pdf")]
        )
        if not file_paths:
            return

        parser = PDFParser()
        extracted_data = []
        for pdf_file in file_paths:
            data = parser.extract_data(pdf_file)
            if data:
                extracted_data.extend(data)

        if not extracted_data:
            messagebox.showwarning("Sin datos", "No se extrajo información de los PDF seleccionados.")
            return

        # Guardar en la carpeta del primer PDF seleccionado
        folder = os.path.dirname(file_paths[0])
        save_path = os.path.join(folder, "Excel_extraido.xlsx")

        excel_writer = ExcelWriter()
        excel_writer.write_to_excel(extracted_data, output_file=save_path)
        messagebox.showinfo("¡Listo!", f"Excel guardado en:\n{save_path}")
        os.startfile(save_path)

if __name__ == "__main__":
    try:
        print("Iniciando aplicación...")
        root = tk.Tk()
        print("Ventana Tkinter creada")
        app = PDFExcelExtractorApp(root)
        print("Aplicación inicializada, iniciando mainloop...")
        root.mainloop()
    except Exception as e:
        print(f"Error al ejecutar la aplicación: {e}")
        input("Presiona Enter para cerrar...")