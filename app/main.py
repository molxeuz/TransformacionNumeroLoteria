import tkinter as tk
from tkinter import messagebox, filedialog
import pandas as pd

archivo_excel = None  # Variable para almacenar la ruta del archivo

def importar_excel():
    global archivo_excel
    archivo_excel = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if archivo_excel:
        try:
            xls = pd.ExcelFile(archivo_excel)
            messagebox.showinfo("Importación", f"Archivo Excel importado correctamente.\nHojas disponibles: {xls.sheet_names}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir el archivo: {e}")

def transformar_numero(numero):
    registros = []
    digitos = [int(d) for d in str(numero)]
    digitos = [(d + 5) % 10 for d in digitos]
    registros.append(digitos[:])

    nuevo_orden = [digitos[0], digitos[2], digitos[3], digitos[1]]
    nuevo_orden[0] = (nuevo_orden[0] + 5) % 10
    registros.append(nuevo_orden[:])

    nuevo_orden[0] = (nuevo_orden[0] + 5) % 10
    nuevo_orden[2], nuevo_orden[3] = nuevo_orden[3], nuevo_orden[2]
    registros.append(nuevo_orden[:])

    nuevo_orden = [(d + 5) % 10 for d in registros[1]]
    registros.append(nuevo_orden[:])

    nuevo_orden = [(d + 5) % 10 for d in registros[2]]
    registros.append(nuevo_orden[:])

    return registros

def procesar():
    numero = entrada.get()
    loteria = loteria_var.get()
    if not numero.isdigit() or len(numero) != 4:
        messagebox.showerror("Error", "Debes ingresar un número de 4 cifras.")
        return

    registros = transformar_numero(numero)
    resultado_texto.delete("1.0", tk.END)
    resultado_texto.insert(tk.END, f"Lotería: {loteria}\n\n")
    resultado_texto.insert(tk.END, "Registro de transformaciones:\n\n")
    for i, registro in enumerate(registros, 1):
        resultado_texto.insert(tk.END, f"Paso {i}: {''.join(map(str, registro))}\n")

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

def agregar_a_excel():
    if not archivo_excel:
        messagebox.showerror("Error", "No se ha importado un archivo Excel.")
        return

    numero = entrada.get()
    loteria = loteria_var.get()
    registros = transformar_numero(numero)

    try:
        # Cargar el archivo Excel
        wb = load_workbook(archivo_excel)

        # Si "Resultados" no existe, créalo
        if "Resultados" not in wb.sheetnames:
            ws = wb.create_sheet("Resultados")
        else:
            ws = wb["Resultados"]

        # Escribir encabezados en B2 y C2
        ws["B2"] = "Lotería"
        ws["C2"] = loteria

        ws["B3"] = "Número ingresado"
        ws["C3"] = numero

        # Escribir "Acción" en B4 - B8 y combinaciones en C4 - C8
        ws["B4"] = "Acción"
        for i, registro in enumerate(registros, start=4):
            ws[f"B{i}"] = "Acción"
            ws[f"C{i}"] = "".join(map(str, registro))

        # Ajustar el ancho de las columnas
        ws.column_dimensions["B"].width = 15
        ws.column_dimensions["C"].width = 20

        # Guardar el archivo
        wb.save(archivo_excel)
        messagebox.showinfo("Éxito", "Datos agregados al archivo Excel en formato tabla.")

    except Exception as e:
        messagebox.showerror("Error", f"No se pudo agregar al Excel: {e}")


# Configuración de la ventana
root = tk.Tk()
root.title("Transformador de Números")
root.geometry("400x400")

# Selección de lotería
tk.Label(root, text="Seleccione la lotería:").pack(pady=5)
loteria_var = tk.StringVar()
loteria_menu = tk.OptionMenu(root, loteria_var, "Lotería Nacional", "Lotería Popular", "Otra")
loteria_menu.pack()

# Entrada de número
tk.Label(root, text="Ingrese un número de 4 cifras:").pack(pady=5)
entrada = tk.Entry(root, font=("Arial", 14), justify="center")
entrada.pack()

# Botones
tk.Button(root, text="Transformar", command=procesar).pack(pady=5)
tk.Button(root, text="Importar Excel", command=importar_excel).pack(pady=5)
tk.Button(root, text="Agregar a Excel", command=agregar_a_excel).pack(pady=5)

# Área de resultados
resultado_texto = tk.Text(root, height=10, width=40)
resultado_texto.pack()

root.mainloop()
