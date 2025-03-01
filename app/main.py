import json
from tkinter import messagebox
import tkinter as tk
from tkinter import ttk

import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os
from dotenv import load_dotenv

load_dotenv()

CREDENTIALS_FILE = os.getenv("GOOGLE_SHEETS_CREDENTIALS_PATH")

# EN .ENV -> GOOGLE_SHEETS_CREDENTIALS_PATH=credentials.json

with open(CREDENTIALS_FILE, "r") as file: # Agregar (.env y credentials) dentro de app
    credentials_dict = json.load(file)

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
client = gspread.authorize(credentials)

print("✅ Conectado a Google Sheets")

SHEET_NAME = "TransformacionNumeroLoteria"
sheet = client.open(SHEET_NAME).sheet1

tk_root = tk.Tk()
tk_root.title("Transformador de Números")
tk_root.geometry("1000x600")
tk_root.resizable(False, False)

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

def registrar_y_guardar():
    numero = entrada_numero.get()
    loteria = entrada_loteria.get()
    if not numero.isdigit() or len(numero) != 4:
        messagebox.showerror("Error", "Debes ingresar un número de 4 cifras.")
        return
    registros = transformar_numero(numero)
    try:
        primera_fila = sheet.row_values(1)
        col_index = len(primera_fila) + 1
        datos = [
            ["Lotería", loteria],
            ["Número", numero],
        ]
        for r in registros:
            datos.append(["Datos", "".join(map(str, r))])
        start_col = gspread.utils.rowcol_to_a1(1, col_index)
        end_col = gspread.utils.rowcol_to_a1(len(datos), col_index + 1)
        sheet.update(f"{start_col}:{end_col}", datos)
        mensaje_estado.config(text="✅ Registro guardado en Google Sheets", fg="green")
        actualizar_tabla()
        entrada_numero.delete(0, tk.END)
        entrada_loteria.delete(0, tk.END)
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo guardar en Google Sheets: {e}")

def actualizar_tabla():
    try:
        data = sheet.get_all_values()
        for row in tree.get_children():
            tree.delete(row)
        loterias = data[0][1::2]
        numeros = data[1][1::2]

        for loteria, numero in zip(loterias, numeros):
            tree.insert("", tk.END, values=(loteria, numero))
    except Exception as e:
        print(f"Error al actualizar la tabla: {e}")

tk.Label(tk_root, text="Ingrese la lotería:").pack(pady=5)
entrada_loteria = tk.Entry(tk_root, font=("Arial", 14))
entrada_loteria.pack()

tk.Label(tk_root, text="Ingrese un número de 4 cifras:").pack(pady=5)
entrada_numero = tk.Entry(tk_root, font=("Arial", 14), justify="center")
entrada_numero.pack()

boton_registrar = tk.Button(tk_root, text="Registrar y Transformar", command=registrar_y_guardar, font=("Arial", 12), bg="#2ecc71", fg="white")
boton_registrar.pack(pady=10)

mensaje_estado = tk.Label(tk_root, text="", font=("Arial", 12))
mensaje_estado.pack(pady=5)

tk.Label(tk_root, text="Registros en Google Sheets").pack(pady=5)
tree = ttk.Treeview(tk_root, columns=("Lotería", "Número"), show="headings")

tree.heading("Lotería", text="Lotería")
tree.heading("Número", text="Número")
tree.pack(pady=10)

actualizar_tabla()
tk_root.mainloop()
