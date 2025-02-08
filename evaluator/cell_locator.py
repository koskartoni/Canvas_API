import pandas as pd
import tkinter as tk
from tkinter import filedialog


def load_evaluation_sheet(file_path, sheet_name="EVALUACIÓN"):
    """Carga la hoja de evaluación y devuelve el DataFrame"""
    xls = pd.ExcelFile(file_path, engine="openpyxl")
    return xls.parse(sheet_name)


def find_student_row(df, student_name):
    """Busca la fila donde se encuentra el estudiante."""
    for index, row in df.iterrows():
        if str(row.iloc[2]).strip().lower() == student_name.lower():
            return index
    return None


def find_activity_column(df, activity_name):
    """Busca la columna donde está la actividad."""
    activity_row = 7  # La fila 7 contiene los nombres de actividades
    for col_index, value in enumerate(df.iloc[activity_row]):
        if str(value).strip().lower() == activity_name.lower():
            return col_index
    return None


def locate_grade_cell(file_path, student_name, activity_name):
    """Devuelve la celda donde se debe insertar la nota."""
    df = load_evaluation_sheet(file_path)
    student_row = find_student_row(df, student_name)
    activity_column = find_activity_column(df, activity_name)

    if student_row is None:
        return f"Estudiante '{student_name}' no encontrado."
    if activity_column is None:
        return f"Actividad '{activity_name}' no encontrada."

    return f"Insertar nota en: Fila {student_row + 1}, Columna {activity_column + 1}"


def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    entry_file_path.delete(0, tk.END)
    entry_file_path.insert(0, file_path)


def search_cell():
    file_path = entry_file_path.get()
    student_name = entry_student.get()
    activity_name = entry_activity.get()
    result = locate_grade_cell(file_path, student_name, activity_name)
    label_result.config(text=result)


# Crear interfaz gráfica
root = tk.Tk()
root.title("Localizador de Celdas - Evaluación")

tk.Label(root, text="Ruta del archivo Excel:").grid(row=0, column=0)
entry_file_path = tk.Entry(root, width=50)
entry_file_path.grid(row=0, column=1)
tk.Button(root, text="Buscar", command=browse_file).grid(row=0, column=2)

tk.Label(root, text="Nombre del Estudiante:").grid(row=1, column=0)
entry_student = tk.Entry(root, width=30)
entry_student.grid(row=1, column=1)

tk.Label(root, text="Nombre de la Actividad:").grid(row=2, column=0)
entry_activity = tk.Entry(root, width=30)
entry_activity.grid(row=2, column=1)

tk.Button(root, text="Buscar Celda", command=search_cell).grid(row=3, column=1)

label_result = tk.Label(root, text="", fg="blue")
label_result.grid(row=4, column=0, columnspan=3)

root.mainloop()
