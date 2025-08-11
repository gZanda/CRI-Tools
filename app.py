import os
import tkinter as tk
import subprocess
from tkinter import filedialog, messagebox, ttk

def calc():
    subprocess.run( ["python", "C:/Users/gabriel.goncalves/Desktop/CRI-Tools/Caclcular Vencimento/calcular_vencimento.py"], check=True)


# Empty dir
def empty_dir():
    subprocess.run(["python", "C:/Users/gabriel.goncalves/Desktop/CRI-Tools/Empy Directories/empty_dir.py"], check=True)

# DOI
def doi():
    subprocess.run(["python", "C:/Users/gabriel.goncalves/Desktop/CRI-Tools/DOI Summary/doi_script.py"], check=True)


# Interface
janela = tk.Tk()
janela.title("Selecione sua Ferramenta")
janela.geometry("600x450")

btn_escolher1 = tk.Button(janela, text="Calcular Vencimento de Contrato", command=calc, font=("Arial", 12))
btn_escolher1.pack(pady=10)

btn_escolher2 = tk.Button(janela, text="Listar Diretórios Vazios", command=empty_dir, font=("Arial", 12))
btn_escolher2.pack(pady=10)

btn_escolher3 = tk.Button(janela, text="Relatório DOI", command=doi, font=("Arial", 12))
btn_escolher3.pack(pady=10)

janela.mainloop()