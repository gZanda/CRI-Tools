from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib import colors
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import time  # só para simulação

# Globals 
consulta_path = None
geral_path = None

# ================================================================================ CONVERSIONS and FORMATTING ================================================================================ #
def formatar_cpf_cnpj(valor):
    # Garante que valores vazios ou 'nan' sejam tratados como vazio
    if pd.isna(valor) or str(valor).strip().lower() in ('nan', ''):  # ALTERAÇÃO
        return ""
    valor_limpo = "".join(filter(str.isdigit, str(valor)))
    if len(valor_limpo) == 11:  # CPF
        return f"{valor_limpo[:3]}.{valor_limpo[3:6]}.{valor_limpo[6:9]}-{valor_limpo[9:]}"
    elif len(valor_limpo) == 14:  # CNPJ
        return f"{valor_limpo[:2]}.{valor_limpo[2:5]}.{valor_limpo[5:8]}/{valor_limpo[8:12]}-{valor_limpo[12:]}"
    else:
        return valor_limpo

# === Conversão numérica para formato brasileiro ===
def to_numeric_brazilian_series(col):
    s = col.astype(str).fillna('')
    # Se tiver vírgula assume formato BR (milhares ponto, decimal vírgula)
    has_comma = s.str.contains(',', regex=False)
    s2 = s.copy()
    s2[has_comma] = s[has_comma].str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
    s2[~has_comma] = s[~has_comma].str.replace(',', '', regex=False)  # remove possíveis vírgulas soltas
    return pd.to_numeric(s2, errors='coerce')

# ================================================================================ READING FROM EXCEL ================================================================================ #

def main(consulta_path, geral_path):

    time.sleep(3)

    # PLANILHA CONSULTA (df)
    df = pd.read_excel(consulta_path)

    # Renomeação de Colunas
    novos_nomes = [
        "Ato",
        "Matrícula",
        "Protocolo",
        "Registro",
        "Tipo Transação",
        "Descrição Transação",
        "Valor Alienação",
        "Situação Construção",
        "Forma Alienação/Aquisição",
        "Data Alienação",
        "Base de Cálculo",
        "Tipo Imóvel",
        "Descrição"
    ]
    df.columns = novos_nomes

    # Mudança de Ordem
    nova_ordem = [
        "Ato",
        "Matrícula",
        "Protocolo",
        "Registro",
        "Tipo Transação",
        "Descrição Transação",
        "Valor Alienação",
        "Base de Cálculo",
        "Situação Construção",
        "Forma Alienação/Aquisição",
        "Data Alienação",
        "Tipo Imóvel",
        "Descrição"
    ]
    df = df[nova_ordem]

    # PLANILHA GERAL (df2)

    # Forçar leitura como String (lê tudo com o formato que estiver, mas essas 3 colunas como string)
    df2 = pd.read_excel(
        geral_path,
        dtype={
            "CPF/CNPJ Alienante/Transmitente": str,
            "CPF/CNPJ Adquirente": str,
            "Participação na operação": str
        }
    )

    # Renomeação de Colunas
    df2.columns = [
        "Data Registro",
        "Matrícula",
        "Livro",
        "Folha",
        "Ato",
        "Situação",
        "Atribuição DOI",
        "Tipo Transação",
        "Descrição Transação",
        "Retificação Ato",
        "Data Alienação",
        "Forma Aquisição",
        "Valor Não Consta",
        "Valor Alienação",
        "Base de Cálculo",
        "Tipo Imóvel",
        "Tipo Imóvel Outros",
        "Situação Construção",
        "Localização",
        "Área Não Consta",
        "Área",
        "Endereço Imóvel",
        "Número",
        "Complemento",
        "Bairro",
        "CEP",
        "Município",
        "UF",
        "NIRF/Cadastro Imobiliário",
        "Cadastro Fiscal",
        "Valor ITBI",
        "Categoria",
        "CPF/CNPJ Transmitente",
        "Participação Transmitente",
        "CPF/CNPJ Representante",
        "CPF/CNPJ Adquirente",
        "Participação Adquirente",
        "CPF/CNPJ Representante"
    ]

    # Remoção de Colunas Desnecessárias
    colunas_para_remover = [
        "Livro", "Folha", "Situação", "Atribuição DOI",
        "Tipo Transação", "Descrição Transação", "Retificação Ato",
        "Data Alienação", "Valor Não Consta", "Valor Alienação", "Base de Cálculo",
        "Tipo Imóvel", "Tipo Imóvel Outros", "Situação Construção",
        "Área Não Consta", "CEP", "UF", "Valor ITBI",
        "CPF/CNPJ Representante", "CPF/CNPJ Representante"
    ]
    df2.drop(columns=colunas_para_remover, inplace=True)

    # ================================================================================ ADJUSTING DATA ================================================================================ #

    # 1) Normalizar Matrícula em df2 para o mesmo formato que você fez em df  # ALTERAÇÃO
    if "Matrícula" in df2.columns:
        df2["Matrícula"] = (
            df2["Matrícula"].astype(str)
            .str.extract(r"(\d[\d\.]*)")[0]
            .str.replace(".", "", regex=False)
        )
        df2["Matrícula"] = pd.to_numeric(df2["Matrícula"], errors="coerce").astype("Int64")

    # 2) Normalizar datas em df2 para dd/mm/YYYY (para usar como fallback)  # ALTERAÇÃO
    for _col in ["Data Registro", "Data Alienação"]:
        if _col in df2.columns:
            df2[_col] = pd.to_datetime(df2[_col], errors="coerce", format="%d/%m/%Y").dt.strftime("%d/%m/%Y")   # Formato da data editado aqui

    # 3) Extrair código do Ato (ex: "R.2") em ambas as tabelas para matching robusto  # ALTERAÇÃO
    df["Ato_cod"] = df["Ato"].astype(str).str.extract(r'([A-Za-z]\.\d+)')[0].fillna('').str.strip()
    df2["Ato_cod"] = df2["Ato"].astype(str).str.extract(r'([A-Za-z]\.\d+)')[0].fillna('').str.strip()

    # 4) Criar coluna numérica de 'Valor Alienação' em df2 para fallback por valor (tratamento BR/EN)  # ALTERAÇÃO
    if "Valor Alienação" in df2.columns:
        df2["Valor Alienação_num"] = to_numeric_brazilian_series(df2["Valor Alienação"])

    # ======================
    # AJUSTE MATRÍCULA E PROTOCOLO NA PLANILHA 1 (já existente)
    # ======================
    df["Matrícula"] = (
        df["Matrícula"].astype(str)
        .str.extract(r"(\d[\d\.]*)")[0]
        .str.replace(".", "", regex=False)
    )
    df["Matrícula"] = pd.to_numeric(df["Matrícula"], errors="coerce").astype("Int64")

    df["Protocolo"] = (
        df["Protocolo"].astype(str)
        .str.extract(r"(\d[\d\.]*)")[0]
        .str.replace(".", "", regex=False)
    )
    df["Protocolo"] = pd.to_numeric(df["Protocolo"], errors="coerce").astype("Int64")

    # ======================
    # FORMATAÇÃO DE DATAS (PLANILHA 1)
    # ======================
    for coluna_data in ["Registro", "Data Alienação"]:
        if coluna_data in df.columns:
            df[coluna_data] = pd.to_datetime(df[coluna_data], errors="coerce").dt.strftime("%d/%m/%Y")

    # Text Styles
    style = getSampleStyleSheet()

    normal_style = ParagraphStyle(
        'Normal',
        parent=style['Normal'],
        fontName='Helvetica',
        fontSize=12,       # ALTERAÇÃO - tamanho da fonte
        leading=18         # ALTERAÇÃO - espaçamento entre linhas
    )

    subtitle_style = ParagraphStyle(
        'Subtitle',
        parent=style['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=14,       # ALTERAÇÃO - tamanho da fonte
        leading=18
    )

    title_style = ParagraphStyle(
        'Title',
        parent=style['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=16,       # ALTERAÇÃO - tamanho da fonte
        leading=18
    )

    story = []

    # ================================================================================ MAIN PROCESSING ================================================================================ #
    for _, linha in df.iterrows():
        matricula = linha["Matrícula"]
        ato = linha["Ato"]
        ato_cod = linha.get("Ato_cod", "") if "Ato_cod" in linha.index else ""


        # ==== CABEÇALHO DO PDF ====
        matricula_str = f"{int(matricula):,}".replace(",", ".")
        story.append(Paragraph(f"MATRÍCULA: {matricula_str}", title_style))
        story.append(Spacer(1, 6))

        # ==== BLOCO ESQUERDO (Dados da Transação) ====
        bloco_esquerdo = []
        bloco_esquerdo.append(Paragraph("Dados da Transação:", subtitle_style))
        for campo, valor in linha.items():
            # Evita imprimir colunas auxiliares como Ato_cod (opcional)
            if campo == "Ato_cod":
                continue
            bloco_esquerdo.append(Paragraph(f"<b>{campo}:</b> {valor if pd.notna(valor) else ''}", normal_style))

        # ==== FILTRO EM DF2: primeiro por Matrícula + Ato_cod (se achar) ====
        dados_2 = pd.DataFrame()
        if (pd.notna(matricula)):
            if ato_cod:
                # tenta pelo código do ato (ex: "R.2")  # ALTERAÇÃO
                dados_2 = df2[(df2["Matrícula"] == matricula) & (df2["Ato_cod"] == ato_cod)].copy()

            # fallbacks caso não ache:
            if dados_2.empty:
                # 1) tentar por Data Registro (se coluna existir)
                reg = linha.get("Registro", "")
                if reg and "Data Registro" in df2.columns:
                    dados_2 = df2[(df2["Matrícula"] == matricula) & (df2["Data Registro"] == reg)].copy()

            if dados_2.empty:
                # 2) tentar por Data Alienação
                data_al = linha.get("Data Alienação", "")
                if data_al and "Data Alienação" in df2.columns:
                    dados_2 = df2[(df2["Matrícula"] == matricula) & (df2["Data Alienação"] == data_al)].copy()

            if dados_2.empty:
                # 3) tentar por Valor Alienação (numérico), se disponível  # ALTERAÇÃO
                if "Valor Alienação_num" in df2.columns and "Valor Alienação" in linha.index:
                    try:
                        v = linha["Valor Alienação"]
                        # normalizar o valor da linha (mesma lógica)
                        if pd.isna(v):
                            vnum = None
                        else:
                            vs = str(v)
                            if ',' in vs:
                                vs2 = vs.replace('.', '').replace(',', '.')
                            else:
                                vs2 = vs.replace(',', '')
                            vnum = float(vs2)
                        if vnum is not None:
                            # garantir que df2 tenha a coluna numérica criada (feito antes)
                            dados_2 = df2[(df2["Matrícula"] == matricula) & (df2["Valor Alienação_num"].notna()) & (df2["Valor Alienação_num"].round(2) == round(vnum, 2))].copy()
                    except Exception:
                        pass

            # última tentativa: apenas por matrícula (poderá trazer mais linhas, mas é fallback)  # ALTERAÇÃO
            if dados_2.empty:
                dados_2 = df2[df2["Matrícula"] == matricula].copy()

        # ==== BLOCO DIREITO (Informações Complementares) ====
        bloco_direito = []
        if not dados_2.empty:
            bloco_direito.append(Paragraph("Informações Complementares:", subtitle_style))
            # pega a primeira combinação relevante (reserve os duplicados)
            dados_fixos = dados_2[[
                "Localização", "Área", "Endereço Imóvel", "Número",
                "Complemento", "Bairro", "Município", "NIRF/Cadastro Imobiliário",
                "Cadastro Fiscal", "Categoria"
            ]].drop_duplicates().iloc[0]
            for campo, valor in dados_fixos.items():
                bloco_direito.append(Paragraph(f"<b>{campo}:</b> {valor if pd.notna(valor) else ''}", normal_style))
        else:
            bloco_direito.append(Paragraph("Nenhuma informação complementar encontrada.", normal_style))

        # ==== MONTAR DUAS COLUNAS LADO A LADO ====
        tabela_duas_colunas = Table(
            [[bloco_esquerdo, bloco_direito]],
            colWidths=[260, 260]
        )
        tabela_duas_colunas.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ]))

        story.append(tabela_duas_colunas)
        story.append(Spacer(1, 6))

        # ==== TABELA DE ENVOLVIDOS ====
        story.append(Paragraph("Envolvidos:", subtitle_style))

        # Seleciona apenas as colunas de envolvidos do dados_2 (pode haver várias linhas)
        envolvidos = pd.DataFrame()
        if not dados_2.empty:
            envolvidos = dados_2[[
                "CPF/CNPJ Transmitente", "Participação Transmitente",
                "CPF/CNPJ Adquirente", "Participação Adquirente"
            ]].copy()

        # garantir vazio em nulos  # ALTERAÇÃO
        if not envolvidos.empty:
            envolvidos = envolvidos.fillna('')  # ALTERAÇÃO

            envolvidos["CPF/CNPJ Transmitente"] = envolvidos["CPF/CNPJ Transmitente"].apply(formatar_cpf_cnpj)
            envolvidos["CPF/CNPJ Adquirente"] = envolvidos["CPF/CNPJ Adquirente"].apply(formatar_cpf_cnpj)

            data_table = [list(envolvidos.columns)] + envolvidos.values.tolist()

            t = Table(data_table, repeatRows=1)
            t.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#d3d3d3")),
                ('GRID', (0,0), (-1,-1), 0.5, colors.black),
                ('FONTSIZE', (0,0), (-1,-1), 9),
                ('VALIGN', (0,0), (-1,-1), 'TOP'),
                ('ALIGN', (1,1), (-1,-1), 'CENTER')
            ]))
            story.append(t)
        else:
            # Sem envolvidos encontrados — evita tabela somente com cabeçalho
            story.append(Paragraph("Nenhum envolvido listado.", normal_style))  # ALTERAÇÃO

        story.append(PageBreak())


    # Saving PDF
    pdf_path = filedialog.asksaveasfilename(
    title="Salvar relatório como",
    defaultextension=".pdf",
    filetypes=[("PDF files", "*.pdf")],
    initialfile="relatório_DOI.pdf"
)
    if not pdf_path:
        messagebox.showwarning("Aviso", "Nenhum caminho de arquivo selecionado. PDF não será salvo.")
        return
    else:
        doc = SimpleDocTemplate(pdf_path, pagesize=A4, rightMargin=20, leftMargin=20, topMargin=10, bottomMargin=10)
        doc.build(story)
        print(f"✅ PDF gerado: {pdf_path}")

# ================================================================================ GUI ================================================================================ #

# Funções para escolher arquivos
def escolher_consulta():
    global consulta_path
    caminho = filedialog.askopenfilename(title="Selecione o arquivo Consulta", filetypes=[("Excel", "*.xlsx *.xls")])
    if caminho:
        consulta_path = caminho
        lbl_consulta.config(text=f"Consulta: {caminho}")

def escolher_geral():
    global geral_path
    caminho = filedialog.askopenfilename(title="Selecione o arquivo Geral", filetypes=[("Excel", "*.xlsx *.xls")])
    if caminho:
        geral_path = caminho
        lbl_geral.config(text=f"Geral: {caminho}")

# Função para executar
def executar():
    if not consulta_path or not geral_path:
        messagebox.showerror("Erro", "Selecione os dois arquivos antes de executar.")
        return
    # Limpa área principal e mostra loading
    limpar_area()
    tk.Label(frame_principal, text="Processando, aguarde...", font=("Arial", 14)).pack(pady=20)
    progress = ttk.Progressbar(frame_principal, mode="indeterminate")
    progress.pack(pady=10, fill="x", padx=40)
    progress.start(10)

    def rodar():
        try:
            main(consulta_path, geral_path)
            mostrar_sucesso()
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro: {e}")
            mostrar_tela_inicial()

    threading.Thread(target=rodar).start()

def mostrar_sucesso():
    limpar_area()
    tk.Label(frame_principal, text="Processamento concluído!", font=("Arial", 14), fg="green").pack(pady=20)
    tk.Button(frame_principal, text="Repetir", font=("Arial", 12), command=mostrar_tela_inicial).pack(pady=5)
    tk.Button(frame_principal, text="Fechar", font=("Arial", 12), command=janela.quit).pack(pady=5)

def limpar_area():
    for widget in frame_principal.winfo_children():
        widget.destroy()

def mostrar_tela_inicial():
    limpar_area()
    global lbl_consulta, lbl_geral
    # Escolher arquivo de Consuulta
    btn_consulta = tk.Button(frame_principal, text="Planilha Consulta", command=escolher_consulta, font=("Arial", 12))
    btn_consulta.pack(pady=5)
    lbl_consulta = tk.Label(frame_principal, text="Consulta: (nenhum arquivo selecionado)", wraplength=500)
    lbl_consulta.pack(pady=5)

    # Escolher arquivo Geral
    btn_geral = tk.Button(frame_principal, text="Planilha Geral", command=escolher_geral, font=("Arial", 12))
    btn_geral.pack(pady=5)
    lbl_geral = tk.Label(frame_principal, text="Geral: (nenhum arquivo selecionado)", wraplength=500)
    lbl_geral.pack(pady=5)

    # Botão para executar
    btn_exec = tk.Button(frame_principal, text="Executar Função", command=executar, font=("Arial", 14), bg="green", fg="white")
    btn_exec.pack(pady=20)

# Tela Inicial
janela = tk.Tk()
janela.title("Relatório DOI")
janela.geometry("600x300")

frame_principal = tk.Frame(janela)
frame_principal.pack(expand=True, fill="both")

mostrar_tela_inicial()

janela.mainloop()