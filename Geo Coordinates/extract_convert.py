import re
import tkinter as tk
from tkinter import messagebox

def dms_to_dd(degrees, minutes, seconds, direction=None):
    dd = abs(degrees) + minutes / 60 + seconds / 3600
    if degrees < 0 or (direction in ['S', 'W']):
        dd *= -1
    return dd

def extract_and_convert(text):
    pattern = r'(-?\d+)°(\d+)[\'’](\d+(?:[.,]\d+)?)"'
    matches = re.findall(pattern, text)

    if len(matches) < 2:
        raise ValueError("Não foi possível encontrar Latitude e Longitude no texto")

    lon_deg, lon_min, lon_sec = matches[0]
    lon = dms_to_dd(int(lon_deg), int(lon_min), float(lon_sec.replace(',', '.')))

    lat_deg, lat_min, lat_sec = matches[1]
    lat = dms_to_dd(int(lat_deg), int(lat_min), float(lat_sec.replace(',', '.')))

    return f"{lat:.6f}, {lon:.6f}"

def converter():
    texto = entry.get()
    try:
        resultado = extract_and_convert(texto)
        result_var.set(resultado)
    except Exception as e:
        messagebox.showerror("Erro", str(e))

def copiar():
    resultado = result_var.get()
    if resultado:
        root.clipboard_clear()
        root.clipboard_append(resultado)

# Interface
root = tk.Tk()
root.title("Conversor DMS → DD")

tk.Label(root, text="Insira o texto com coordenadas:").pack(pady=5)

entry = tk.Entry(root, width=60)
entry.pack(pady=5)

tk.Button(root, text="Converter", command=converter).pack(pady=5)

result_var = tk.StringVar()
result_label = tk.Entry(root, textvariable=result_var, width=40, state="readonly", justify="center")
result_label.pack(pady=5)

tk.Button(root, text="Copiar", command=copiar).pack(pady=5)

root.mainloop()
