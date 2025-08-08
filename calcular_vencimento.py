from datetime import datetime
from dateutil.relativedelta import relativedelta
import ttkbootstrap as tb
from ttkbootstrap.constants import *

def abrir_calculadora():
    janela = tb.Toplevel()
    janela.title("Calculadora de Vencimento Final")
    janela.geometry("350x250")
    
    # Widgets estilizados
    label1 = tb.Label(janela, text="Data da 1ª parcela (dd/mm/aaaa):")
    entrada_data = tb.Entry(janela, width=25, bootstyle="info")

    label2 = tb.Label(janela, text="Número de parcelas:")
    entrada_parcelas = tb.Entry(janela, width=25, bootstyle="info")

    label_resultado = tb.Label(janela, text="", font=("Arial", 12, "bold"), bootstyle="success")

    def calcular_vencimento():
        data_str = entrada_data.get()
        parcelas_str = entrada_parcelas.get()

        try:
            data_inicial = datetime.strptime(data_str, "%d/%m/%Y")
            parcelas = int(parcelas_str)

            if parcelas <= 0:
                raise ValueError("Parcelas deve ser maior que zero.")

            data_final = data_inicial + relativedelta(months=parcelas - 1)
            resultado = data_final.strftime("%d/%m/%Y")

            label_resultado.config(text=f"Vencimento final: {resultado}", bootstyle="success")
        except ValueError as e:
            label_resultado.config(text=f"Erro: {e}", bootstyle="danger")

    label1.pack(pady=5)
    entrada_data.pack()

    label2.pack(pady=5)
    entrada_parcelas.pack()

    botao = tb.Button(janela, text="Calcular", command=calcular_vencimento, bootstyle="primary")
    botao.pack(pady=10)

    label_resultado.pack(pady=10)

# ========== Janela principal com tema ==========

app = tb.Window(themename="flatly")  # Outros temas: "darkly", "cyborg", "journal", etc.
app.title("Painel de Ferramentas")
app.geometry("300x200")

label_menu = tb.Label(app, text="Escolha uma ferramenta:", font=("Arial", 14))
label_menu.pack(pady=20)

btn_calc = tb.Button(app, text="Calculadora de Vencimento", width=25, command=abrir_calculadora, bootstyle="info")
btn_calc.pack(pady=5)

app.mainloop()
