import re

def dms_to_dd(degrees, minutes, seconds, direction=None):
    """Converte coordenadas DMS para DD"""
    dd = abs(degrees) + minutes / 60 + seconds / 3600
    if degrees < 0 or (direction in ['S', 'W']):
        dd *= -1
    return dd

def extract_and_convert(text):
    # Expressão regular para capturar valores de coordenadas no formato DMS
    pattern = r'(-?\d+)°(\d+)[\'’](\d+(?:[.,]\d+)?)"'
    matches = re.findall(pattern, text)

    if len(matches) < 2:
        raise ValueError("Não foi possível encontrar Latitude e Longitude no texto")

    # Longitude
    lon_deg, lon_min, lon_sec = matches[0]
    lon = dms_to_dd(int(lon_deg), int(lon_min), float(lon_sec.replace(',', '.')))

    # Latitude
    lat_deg, lat_min, lat_sec = matches[1]
    lat = dms_to_dd(int(lat_deg), int(lat_min), float(lat_sec.replace(',', '.')))

    return f"{lat:.6f}, {lon:.6f}"

# Input do usuário
texto = input("Digite o texto com as coordenadas DMS: ")
print(extract_and_convert(texto))
