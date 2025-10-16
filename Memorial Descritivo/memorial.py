import re

# Arquivo de entrada e saída
entrada = "entrada.txt"
saida = "saida.txt"

# Lê todo o texto do arquivo de entrada
with open(entrada, "r", encoding="utf-8") as f:
    texto = f.read()

# Regex para pegar vértices + coordenadas
padrao = r"[Vv]értice\s*'?(\w+)'?.*?E\s*=?\s*([\d\.]+)\s*m?\s*e\s*N\s*=?\s*([\d\.]+)"

vertices = re.findall(padrao, texto)

# Escreve a saída formatada em outro arquivo
with open(saida, "w", encoding="utf-8") as f:
    for nome, e, n in vertices:
        f.write(f"vértice {nome} E={e} N={n}\n")

print(f"Processo concluído ✅. Saída salva em {saida}")
