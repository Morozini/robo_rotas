import xml.etree.ElementTree as ET
import pandas as pd
from maps import mapsAuto

def kml_to_xlsx(kml_file, xlsx_file):

    tree = ET.parse(kml_file)
    root = tree.getroot()

    ns = {'kml': 'http://www.opengis.net/kml/2.2'}

    data = []

    for placemark in root.findall('.//kml:Placemark', ns):
        name = placemark.find('kml:name', ns).text if placemark.find('kml:name', ns) is not None else 'N/A'
        
        data.append([name])

    df = pd.DataFrame(data, columns=['Name'])

    df.to_excel(xlsx_file, index=False)

def executar(data):

    arquivo_rota_convertido = f'{data[1]}\\{data[2]}_{data[3]}.xlsx'

    kml_to_xlsx(rf'{data[0]}', f'{arquivo_rota_convertido}')
    mapsAuto(arquivo_rota_convertido)

