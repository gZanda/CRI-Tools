from docx2pdf import convert
import os

# Pasta de origem (onde estão os .docx)
pasta_origem = r"\\192.168.0.15\Documentos_2019\POP's - 2025"

# Pasta de destino (onde serão salvos os PDFs)
pasta_destino = r"C:\Users\gabriel.goncalves\Desktop\A"

# Cria a pasta de destino se não existir
os.makedirs(pasta_destino, exist_ok=True)

# Converte tudo da origem para o destino
convert(pasta_origem, pasta_destino)
