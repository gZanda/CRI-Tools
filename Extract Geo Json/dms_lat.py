import re
import json

# Função para converter DMS -> graus decimais (DD)
def dms_to_dd(dms_str):
    """
    Converte string no formato -45°03'36,015" para float em graus decimais.
    """
    regex = re.match(r'(-?\d+)°(\d+)\'([\d,\.]+)"', dms_str.strip())
    if not regex:
        raise ValueError(f"Formato inválido: {dms_str}")
    graus, minutos, segundos = regex.groups()
    graus = int(graus)
    minutos = int(minutos)
    segundos = float(segundos.replace(",", "."))
    
    dd = abs(graus) + minutos/60 + segundos/3600
    if graus < 0:
        dd = -dd
    return dd

# Texto de exemplo (use o memorial completo aqui)
texto = """
Inicia-se a descrição deste imóvel no vértice FOY-P-14743, Longitude: -45°03'36,015", Latitude: -21°30'37,566" e Altitude: 925,04 m, deste segue confrontando com CNS: 05.932-9 | Mat. 17322 | EDUARDO PERILLO no azimute 105°33' e distância de 80,27 m até o vértice FOY-P-14744, Longitude: -45°03'33,329", Latitude: -21°30'38,266" e Altitude: 931,96 m, deste segue no azimute 100°55' e distância de 76,19 m até o vértice FOY-P-14745, Longitude: -45°03'30,730", Latitude: -21°30'38,735" e Altitude: 937,71 m, deste segue no azimute 81°52' e distância de 69,18 m até o vértice FOY-P-14746, Longitude: -45°03'28,351", Latitude: -21°30'38,418" e Altitude: 940,62 m, deste segue confrontando com CNS: 05.932-9 | Mat. 13241 | MARCILENE APARECIDA DOS SANTOS no azimute 147°23' e distância de 29,34 m até o vértice FOY-P-14747, Longitude: -45°03'27,802", Latitude: -21°30'39,221" e Altitude: 938,22 m, deste segue no azimute 91°58' e distância de 442,01 m até o vértice FOY-P-14748, Longitude: -45°03'12,456", Latitude: -21°30'39,718" e Altitude: 969,62 m, deste segue confrontando com CNS: 05.932-9 | Mat. 13243 | MARCIO ANTONIO DO SANTOS no azimute 189°32' e distância de 706,04 m até o vértice FOY-P-14749, Longitude: -45°03'16,520", Latitude: -21°31'02,353" e Altitude: 976,60 m, deste segue no azimute 284°35' e distância de 312,00 m até o vértice FOY-P-14736, Longitude: -45°03'27,009", Latitude: -21°30'59,797" e Altitude: 958,07 m, deste segue confrontando com CNS: 05.932-9 | Mat. 13249 | MATHEUS VINICIUS SOUZA SANTOS no azimute 277°42' e distância de 286,57 m até o vértice FOY-P-14735, Longitude: -45°03'36,875", Latitude: -21°30'58,546" e Altitude: 936,95 m, deste segue no azimute 278°11' e distância de 176,98 m até o vértice FOY-P-14734, Longitude: -45°03'42,961", Latitude: -21°30'57,727" e Altitude: 912,41 m, deste segue no azimute 270°46' e distância de 59,12 m até o vértice FOY-P-14733, Longitude: -45°03'45,014", Latitude: -21°30'57,701" e Altitude: 907,73 m, deste segue no azimute 263°39' e distância de 129,91 m até o vértice RMBP-P-4292, Longitude: -45°03'49,500", Latitude: -21°30'58,168" e Altitude: 908,43 m, deste segue confrontando com CNS: 04.732-4 | Mat. 6048 | JOAO DA MATA no azimute 305°24' e distância de 21,40 m até o vértice RMBP-P-4291, Longitude: -45°03'50,106", Latitude: -21°30'57,765" e Altitude: 907,48 m, deste segue no azimute 0°20' e distância de 48,94 m até o vértice RMBP-P-4290, Longitude: -45°03'50,096", Latitude: -21°30'56,174" e Altitude: 905,86 m, deste segue no azimute 352°52' e distância de 19,25 m até o vértice RMBP-P-4289, Longitude: -45°03'50,179", Latitude: -21°30'55,553" e Altitude: 908,77 m, deste segue no azimute 49°21' e distância de 34,86 m até o vértice RMBP-P-4288, Longitude: -45°03'49,260", Latitude: -21°30'54,815" e Altitude: 909,38 m, deste segue no azimute 110°14' e distância de 20,28 m até o vértice RMBP-P-4287, Longitude: -45°03'48,599", Latitude: -21°30'55,043" e Altitude: 904,42 m, deste segue no azimute 21°45' e distância de 30,60 m até o vértice RMBP-P-4286, Longitude: -45°03'48,205", Latitude: -21°30'54,119" e Altitude: 902,99 m, deste segue no azimute 302°36' e distância de 8,51 m até o vértice RMBP-P-4285, Longitude: -45°03'48,454", Latitude: -21°30'53,970" e Altitude: 904,34 m, deste segue no azimute 213°41' e distância de 24,96 m até o vértice RMBP-P-4284, Longitude: -45°03'48,935", Latitude: -21°30'54,645" e Altitude: 905,10 m, deste segue no azimute 282°28' e distância de 24,20 m até o vértice RMBP-P-4283, Longitude: -45°03'49,756", Latitude: -21°30'54,475" e Altitude: 904,98 m, deste segue no azimute 357°36' e distância de 10,34 m até o vértice RMBP-P-4282, Longitude: -45°03'49,771", Latitude: -21°30'54,139" e Altitude: 909,13 m, deste segue no azimute 293°39' e distância de 41,48 m até o vértice RMBP-P-4281, Longitude: -45°03'51,091", Latitude: -21°30'53,598" e Altitude: 902,19 m, deste segue no azimute 32°26' e distância de 70,78 m até o vértice RMBP-P-4280, Longitude: -45°03'49,772", Latitude: -21°30'51,656" e Altitude: 905,61 m, deste segue no azimute 37°52' e distância de 38,73 m até o vértice RMBP-P-4279, Longitude: -45°03'48,946", Latitude: -21°30'50,662" e Altitude: 908,16 m, deste segue no azimute 255°00' e distância de 28,07 m até o vértice RMBP-P-4278, Longitude: -45°03'49,888", Latitude: -21°30'50,898" e Altitude: 911,75 m, deste segue no azimute 345°56' e distância de 43,95 m até o vértice RMBP-P-4277, Longitude: -45°03'50,259", Latitude: -21°30'49,512" e Altitude: 900,47 m, deste segue no azimute 32°58' e distância de 46,60 m até o vértice RMBP-P-4276, Longitude: -45°03'49,378", Latitude: -21°30'48,241" e Altitude: 908,44 m, deste segue no azimute 64°58' e distância de 21,60 m até o vértice RMBP-P-4275, Longitude: -45°03'48,698", Latitude: -21°30'47,944" e Altitude: 907,79 m, deste segue no azimute 1°27' e distância de 74,50 m até o vértice RMBP-P-4274, Longitude: -45°03'48,632", Latitude: -21°30'45,523" e Altitude: 906,21 m, deste segue no azimute 342°54' e distância de 25,07 m até o vértice RMBP-P-4273, Longitude: -45°03'48,888", Latitude: -21°30'44,744" e Altitude: 907,46 m, deste segue no azimute 57°33' e distância de 22,99 m até o vértice RMBP-P-4272, Longitude: -45°03'48,214", Latitude: -21°30'44,343" e Altitude: 904,29 m, deste segue no azimute 147°45' e distância de 32,81 m até o vértice RMBP-P-4271, Longitude: -45°03'47,606", Latitude: -21°30'45,245" e Altitude: 905,87 m, deste segue no azimute 67°55' e distância de 39,61 m até o vértice RMBP-P-4270, Longitude: -45°03'46,331", Latitude: -21°30'44,761" e Altitude: 907,49 m, deste segue no azimute 7°22' e distância de 34,55 m até o vértice RMBP-P-4269, Longitude: -45°03'46,177", Latitude: -21°30'43,647" e Altitude: 904,84 m, deste segue no azimute 50°41' e distância de 29,28 m até o vértice RMBP-P-4268, Longitude: -45°03'45,390", Latitude: -21°30'43,044" e Altitude: 903,93 m, deste segue no azimute 333°20' e distância de 16,94 m até o vértice RMBP-P-4267, Longitude: -45°03'45,654", Latitude: -21°30'42,552" e Altitude: 903,88 m, deste segue no azimute 41°33' e distância de 39,01 m até o vértice RMBP-P-4266, Longitude: -45°03'44,755", Latitude: -21°30'41,603" e Altitude: 904,94 m, deste segue no azimute 41°41' e distância de 86,21 m até o vértice RMBP-P-4265, Longitude: -45°03'42,763", Latitude: -21°30'39,510" e Altitude: 905,90 m, deste segue no azimute 135°19' e distância de 23,70 m até o vértice RMBP-P-4264, Longitude: -45°03'42,184", Latitude: -21°30'40,058" e Altitude: 904,07 m, deste segue no azimute 192°48' e distância de 35,96 m até o vértice RMBP-P-4263, Longitude: -45°03'42,461", Latitude: -21°30'41,198" e Altitude: 904,39 m, deste segue no azimute 161°47' e distância de 42,84 m até o vértice RMBP-P-4262, Longitude: -45°03'41,996", Latitude: -21°30'42,521" e Altitude: 905,39 m, deste segue no azimute 28°20' e distância de 28,62 m até o vértice RMBP-P-4261, Longitude: -45°03'41,524", Latitude: -21°30'41,702" e Altitude: 905,50 m, deste segue no azimute 17°20' e distância de 26,46 m até o vértice RMBP-P-4260, Longitude: -45°03'41,250", Latitude: -21°30'40,881" e Altitude: 904,34 m, deste segue no azimute 28°52' e distância de 60,25 m até o vértice RMBP-P-4259, Longitude: -45°03'40,239", Latitude: -21°30'39,166" e Altitude: 906,03 m, deste segue confrontando com CNS: 05.932-9 | Mat. 17322 | EDUARDO PERILLO no azimute 75°20' e distância de 14,02 m até o vértice FOY-P-14750, Longitude: -45°03'39,768", Latitude: -21°30'39,051" e Altitude: 907,99 m, deste segue no azimute 62°23' e distância de 19,74 m até o vértice FOY-P-14751, Longitude: -45°03'39,160", Latitude: -21°30'38,753" e Altitude: 912,48 m, deste segue no azimute 68°38' e distância de 51,22 m até o vértice FOY-P-14776, Longitude: -45°03'37,503", Latitude: -21°30'38,147" e Altitude: 917,99 m, deste segue no azimute 67°20' e distância de 46,40 m até o vértice FOY-P-14743, Longitude: -45°03'36,015", Latitude: -21°30'37,566" e Altitude: 925,04 m, ponto inicial desta descrição.
"""

# Regex para capturar vértices com Lat/Lon
pattern = re.compile(
    r"vértice\s*([A-Z0-9\-]+).*?Longitude:\s*([-0-9°\'\",\.]+).*?Latitude:\s*([-0-9°\'\",\.]+)", 
    re.IGNORECASE
)

vertices = []
for match in pattern.finditer(texto):
    vertice = match.group(1)
    lon_dms = match.group(2).strip()
    lat_dms = match.group(3).strip()
    
    lon = dms_to_dd(lon_dms)
    lat = dms_to_dd(lat_dms)
    
    vertices.append((vertice, lon, lat))

# Criar lista de coordenadas
coords_geo = [[lon, lat] for _, lon, lat in vertices]

# Fechar polígono
if coords_geo[0] != coords_geo[-1]:
    coords_geo.append(coords_geo[0])

# Montar GeoJSON
geojson = {
    "type": "FeatureCollection",
    "features": [
        {
            "type": "Feature",
            "properties": {"name": "Perímetro do imóvel"},
            "geometry": {
                "type": "Polygon",
                "coordinates": [coords_geo]
            }
        }
    ]
}

# Salvar em arquivo
with open("perimetro_dms.geojson", "w", encoding="utf-8") as f:
    json.dump(geojson, f, ensure_ascii=False, indent=2)

print("GeoJSON salvo em perimetro_dms.geojson")
