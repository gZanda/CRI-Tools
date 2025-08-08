from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib import colors
import pandas as pd
import numpy as np

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

# Helper: converte números no formato brasileiro/inglês para float  # ALTERAÇÃO
def to_numeric_brazilian_series(col):
    s = col.astype(str).fillna('')
    # Se tiver vírgula assume formato BR (milhares ponto, decimal vírgula)
    has_comma = s.str.contains(',', regex=False)
    s2 = s.copy()
    s2[has_comma] = s[has_comma].str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
    s2[~has_comma] = s[~has_comma].str.replace(',', '', regex=False)  # remove possíveis vírgulas soltas
    return pd.to_numeric(s2, errors='coerce')

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

# Ordenar por Registro (data) crescente
# df = df.sort_values(by="Registro", ascending=True)

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
# ALTERAÇÕES DE NORMALIZAÇÃO (IMPORTANTE)
# ======================
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
        df2[_col] = pd.to_datetime(df2[_col], errors="coerce").dt.strftime("%d/%m/%Y")

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


# ======================
# ESCREVER NO PDF (estilos)
# ======================
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

# ======================
# LAÇO PRINCIPAL: AGORA ITERA LINHA A LINHA (cada linha = 1 ato)  # ALTERAÇÃO
# ======================
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

# ==== GERAR PDF ====
pdf_path = "relatorio_por_matricula_organizado.pdf"
doc = SimpleDocTemplate(pdf_path, pagesize=A4, rightMargin=20, leftMargin=20, topMargin=10, bottomMargin=10)
doc.build(story)
print(f"✅ PDF gerado: {pdf_path}")
