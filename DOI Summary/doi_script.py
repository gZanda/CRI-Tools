from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib import colors
import pandas as pd

# ==== Função de formatação ajustada ====  # ALTERAÇÃO
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

# ======================
# PLANILHA CONSULTA (df)
# ======================
df = pd.read_excel("Consulta.xlsx")

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

# ======================
# PLANILHA GERAL (df2) — força leitura de CPF/CNPJ como texto
# ======================
df2 = pd.read_excel(
    "Geral.xlsx",
    dtype={
        "CPF/CNPJ Alienante/Transmitente": str,
        "CPF/CNPJ Adquirente": str,
        "Participação na operação": str
    }
)

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
    "Inscrição NIRF",
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

colunas_para_remover = [
    "Livro", "Folha", "Situação", "Atribuição DOI",
    "Tipo Transação", "Descrição Transação", "Retificação Ato",
    "Data Alienação", "Valor Não Consta", "Valor Alienação", "Base de Cálculo",
    "Tipo Imóvel", "Tipo Imóvel Outros", "Situação Construção",
    "Área Não Consta", "CEP", "UF", "Valor ITBI",
    "CPF/CNPJ Representante", "CPF/CNPJ Representante"
]
df2.drop(columns=colunas_para_remover, inplace=True)

# ======================
# AJUSTE MATRÍCULA E PROTOCOLO NA PLANILHA 1
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












# ======================
# ESCREVER NO PDF
# ======================
# ==== ESTILOS ====
style = getSampleStyleSheet()

# Corpo do texto
normal_style = ParagraphStyle(
    'Normal',
    parent=style['Normal'],
    fontName='Helvetica',
    fontSize=12,       # ALTERAÇÃO - tamanho da fonte
    leading=18         # ALTERAÇÃO - espaçamento entre linhas
)

# Títulos
subtitle_style = ParagraphStyle(
    'Subtitle',
    parent=style['Heading2'],
    fontName='Helvetica-Bold',
    fontSize=14,       # ALTERAÇÃO - tamanho da fonte
    leading=18
)

# Títulos
title_style = ParagraphStyle(
    'Title',
    parent=style['Heading1'],
    fontName='Helvetica-Bold',
    fontSize=16,       # ALTERAÇÃO - tamanho da fonte
    leading=18
)


story = []

for matricula in df["Matrícula"].dropna().unique():
    story.append(Paragraph(f"MATRÍCULA: {matricula}", title_style))
    story.append(Spacer(1, 6))

    dados_1 = df[df["Matrícula"] == matricula].iloc[0]

    # ==== BLOCO ESQUERDO (Dados da Transação) ====
    bloco_esquerdo = []
    bloco_esquerdo.append(Paragraph("Dados da Transação:", subtitle_style))
    for campo, valor in dados_1.items():
        bloco_esquerdo.append(Paragraph(f"<b>{campo}:</b> {valor if pd.notna(valor) else ''}", normal_style))

    # ==== BLOCO DIREITO (Informações Complementares) ====
    bloco_direito = []
    dados_2 = df2[df2["Matrícula"] == matricula]
    if not dados_2.empty:
        bloco_direito.append(Paragraph("Informações Complementares:", subtitle_style))
        dados_fixos = dados_2[[
            "Localização", "Área", "Endereço Imóvel", "Número",
            "Complemento", "Bairro", "Município", "Inscrição NIRF",
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
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
    ]))

    story.append(tabela_duas_colunas)
    story.append(Spacer(1, 12))

    # ==== TABELA DE ENVOLVIDOS ====
    story.append(Paragraph("👥 Envolvidos:", subtitle_style))
    envolvidos = dados_2[[
        "CPF/CNPJ Transmitente", "Participação Transmitente",
        "CPF/CNPJ Adquirente", "Participação Adquirente"
    ]].copy()

    envolvidos = envolvidos.fillna('')  # 🔹 Remove 'nan' e deixa vazio  # ALTERAÇÃO

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
    story.append(PageBreak())

# ==== GERAR PDF ====
pdf_path = "relatorio_por_matricula_organizado.pdf"
doc = SimpleDocTemplate(pdf_path, pagesize=A4, rightMargin=20, leftMargin=20, topMargin=10, bottomMargin=10)
doc.build(story)
print(f"✅ PDF gerado: {pdf_path}")
