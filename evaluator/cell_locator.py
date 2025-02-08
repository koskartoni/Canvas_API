import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

def build_activity_mapping(file_path, sheet_name="EVALUACIÓN",
                           merged_header_text="RESULTADO APRENDIZAJE-CRITERIO DE EVALUACIÓN PRÁCTICOS",
                           header_rows=(7, 8), activity_row=9):
    """
    Abre el archivo Excel y busca los rangos combinados en las filas de encabezado (header_rows)
    que contienen el texto especificado en merged_header_text.
    Para cada bloque encontrado, extrae de la fila activity_row los nombres de las actividades
    y construye un diccionario que asocia cada actividad a su referencia de celda (ej.: "AC9").
    """
    wb = load_workbook(file_path, data_only=True)
    ws = wb[sheet_name]
    mapping = {}

    # Recorremos cada rango de celdas combinadas de la hoja
    for merged_range in ws.merged_cells.ranges:
        min_row = merged_range.min_row
        max_row = merged_range.max_row
        min_col = merged_range.min_col
        max_col = merged_range.max_col

        # Verificar que el rango abarca las filas de encabezado definidas (7 y 8)
        if min_row == header_rows[0] and max_row == header_rows[1]:
            # El contenido se encuentra en la celda superior izquierda del rango combinado
            header_value = ws.cell(row=min_row, column=min_col).value
            if header_value and merged_header_text.lower() in str(header_value).lower():
                # Se ha identificado un bloque de actividades; se recorre cada columna del bloque
                for col in range(min_col, max_col + 1):
                    cell = ws.cell(row=activity_row, column=col)
                    if cell.value is not None:
                        activity_name = str(cell.value).strip()
                        cell_ref = f"{get_column_letter(col)}{activity_row}"
                        mapping[activity_name] = cell_ref
    return mapping

def select_file():
    """Abre un diálogo para seleccionar el archivo Excel y muestra la ruta en el entry."""
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        entry_file_path.delete(0, tk.END)
        entry_file_path.insert(0, file_path)

def map_activities():
    """Construye el mapeo de actividades y lo muestra en un mensaje."""
    file_path = entry_file_path.get()
    if not file_path:
        messagebox.showerror("Error", "Por favor, seleccione un archivo Excel.")
        return
    mapping = build_activity_mapping(file_path)
    if mapping:
        result_text = "Mapeo de Actividades:\n"
        for activity, cell_ref in mapping.items():
            result_text += f"  {activity}: {cell_ref}\n"
    else:
        result_text = "No se encontró el encabezado o no se pudieron mapear actividades."
    messagebox.showinfo("Resultado del Mapeo", result_text)

# Crear la interfaz gráfica con Tkinter
root = tk.Tk()
root.title("Mapeo de Actividades en Excel")

# Selección del archivo Excel
tk.Label(root, text="Ruta del archivo Excel:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
entry_file_path = tk.Entry(root, width=50)
entry_file_path.grid(row=0, column=1, padx=5, pady=5)
tk.Button(root, text="Seleccionar Archivo", command=select_file).grid(row=0, column=2, padx=5, pady=5)

# Botón para ejecutar el mapeo
tk.Button(root, text="Mapear Actividades", command=map_activities).grid(row=1, column=1, padx=5, pady=5)

root.mainloop()
