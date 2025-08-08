import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
from datetime import datetime

def encontrar_diretorios_vazios(diretorio_raiz, log_callback, progress_callback, fim_callback):
    diretorios_vazios = []
    contador = 0
    atualizacao_intervalo = 50  # Atualiza a barra/log a cada 50 pastas

    for raiz, dirs, arquivos in os.walk(diretorio_raiz):
        if not dirs and not arquivos:
            diretorios_vazios.append(raiz)
            log_callback(f"[VAZIO] {raiz}")
        contador += 1
        if contador % atualizacao_intervalo == 0:
            progress_callback(contador)

    fim_callback(diretorios_vazios, contador)

def salvar_em_txt(caminhos, nome_arquivo):
    with open(nome_arquivo, "w", encoding="utf-8") as arquivo:
        for caminho in caminhos:
            arquivo.write(caminho + "\n")

def escolher_pasta():
    pasta = filedialog.askdirectory(title="Escolha o diretório raiz")
    if not pasta:
        return

    log_text.delete("1.0", tk.END)
    progress_bar["value"] = 0
    progress_bar.update()

    def log_callback(texto):
        log_text.insert(tk.END, texto + "\n")
        if log_text.index('end-1c').split('.')[0] > '200':  # Limita para não travar
            log_text.delete("1.0", "2.0")
        log_text.see(tk.END)

    def progress_callback(atual):
        progress_bar["value"] = atual % 100  # Apenas incrementa visivelmente
        progress_bar.update()

    def fim_callback(resultados, total_scan):
        salvar_em_txt(resultados, "diretorios_vazios.txt")
        tempo = datetime.now().strftime("%H:%M:%S")
        msg = f"{len(resultados)} diretórios vazios encontrados entre {total_scan} analisados.\nSalvo em 'diretorios_vazios.txt' às {tempo}"
        messagebox.showinfo("Finalizado", msg)

    thread = threading.Thread(target=encontrar_diretorios_vazios, args=(pasta, log_callback, progress_callback, fim_callback))
    thread.start()

# Interface
janela = tk.Tk()
janela.title("Scanner de Diretórios Vazios")
janela.geometry("600x450")

btn_escolher = tk.Button(janela, text="Escolher pasta", command=escolher_pasta, font=("Arial", 12))
btn_escolher.pack(pady=10)

progress_bar = ttk.Progressbar(janela, orient="horizontal", length=500, mode="determinate", maximum=100)
progress_bar.pack(pady=10)

log_text = tk.Text(janela, height=20, width=70)
log_text.pack(pady=10)

janela.mainloop()
