import tkinter as tk
from tkinter import messagebox, filedialog
import pandas as pd
from openpyxl import load_workbook

# Configuración de la ventana principal
tk_root = tk.Tk()
tk_root.title("Transformador de Números")
tk_root.geometry("800x500")

tk_root.resizable(False, False)
archivo_excel = None  # Ruta del archivo Excel
registros_totales = []  # Historial de números procesados

# Función para importar un archivo Excel
def importar_excel():
    global archivo_excel
    archivo_excel = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if archivo_excel:
        try:
            xls = pd.ExcelFile(archivo_excel)
            messagebox.showinfo("Importación", f"Archivo Excel importado correctamente. Hojas disponibles: {xls.sheet_names}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir el archivo: {e}")

# Función para transformar un número
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

# Función para agregar un número al historial
def agregar_a_lista():
    numero = entrada_numero.get()
    loteria = entrada_loteria.get()
    if not numero.isdigit() or len(numero) != 4:
        messagebox.showerror("Error", "Debes ingresar un número de 4 cifras.")
        return
    registros = transformar_numero(numero)
    registros_totales.append((loteria, numero, registros))
    resultado_texto.insert(tk.END, f"Lotería: {loteria}, Número: {numero}\n")
    for i, registro in enumerate(registros, 1):
        resultado_texto.insert(tk.END, f"Paso {i}: {''.join(map(str, registro))}\n")
    resultado_texto.insert(tk.END, "--------------------\n")
    entrada_numero.delete(0, tk.END)
    entrada_loteria.delete(0, tk.END)

# Función para guardar el historial en un archivo Excel
def guardar_en_excel():
    if not archivo_excel:
        messagebox.showerror("Error", "No se ha importado un archivo Excel.")
        return
    if not registros_totales:
        messagebox.showerror("Error", "No hay registros para guardar.")
        return
    try:
        wb = load_workbook(archivo_excel)
        if "Resultados" not in wb.sheetnames:
            ws = wb.create_sheet("Resultados")
        else:
            ws = wb["Resultados"]
        col = ws.max_column + 2
        for loteria, numero, registros in registros_totales:
            ws.cell(row=2, column=col, value="Lotería")
            ws.cell(row=2, column=col+1, value=loteria)
            ws.cell(row=4, column=col, value="Número")
            ws.cell(row=4, column=col+1, value=numero)
            for i, registro in enumerate(registros, start=6):
                ws.cell(row=i, column=col, value="Datos")
                ws.cell(row=i, column=col+1, value="".join(map(str, registro)))
            col += 3
        wb.save(archivo_excel)
        messagebox.showinfo("Éxito", "Datos agregados al Excel.")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo guardar en Excel: {e}")

# Diseño de la interfaz
frame_izquierda = tk.Frame(tk_root, width=350, height=500)
frame_izquierda.pack(side=tk.LEFT, padx=10, pady=10)
frame_derecha = tk.Frame(tk_root, width=450, height=500)
frame_derecha.pack(side=tk.RIGHT, padx=10, pady=10)

# Sección de ingreso de datos
tk.Label(frame_izquierda, text="Ingrese la lotería:").pack(pady=5)
entrada_loteria = tk.Entry(frame_izquierda, font=("Arial", 14))
entrada_loteria.pack()
tk.Label(frame_izquierda, text="Ingrese un número de 4 cifras:").pack(pady=5)
entrada_numero = tk.Entry(frame_izquierda, font=("Arial", 14), justify="center")
entrada_numero.pack()
tk.Button(frame_izquierda, text="Transformar y Agregar", command=agregar_a_lista).pack(pady=5)
tk.Button(frame_izquierda, text="Importar Excel", command=importar_excel).pack(pady=5)
tk.Button(frame_izquierda, text="Guardar en Excel", command=guardar_en_excel).pack(pady=5)

# Sección del historial de números
tk.Label(frame_derecha, text="Historial de números procesados").pack()
resultado_texto = tk.Text(frame_derecha, height=25, width=50)
resultado_texto.pack()

# Iniciar la aplicación
tk_root.mainloop()
