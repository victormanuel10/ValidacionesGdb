import xml.etree.ElementTree as ET
from shapely.geometry import Point
import fiona
from fiona.crs import from_epsg


xtf_file = 'xtf_sonson_urbano_20250224_1.xtf'

# Ruta de salida para el shapefile
shapefile_output = 'output.shp'

# Leer el archivo XTF (XML)
tree = ET.parse(xtf_file)
root = tree.getroot()

# Crear un esquema de fiona para el shapefile
schema = {
    'geometry': 'Point',
    'properties': {'id': 'int'},
}

# Crear el shapefile
with fiona.open(shapefile_output, 'w', driver='ESRI Shapefile', crs=from_epsg(4326), schema=schema) as output:
    # Iterar sobre las coordenadas en el archivo XTF
    for i, point in enumerate(root.findall('.//point')):  # Suponiendo que hay puntos en el XTF
        coords = point.find('coordinates').text
        lon, lat = map(float, coords.split(','))  # Asumiendo que las coordenadas están separadas por coma
        
        # Crear un punto de Shapely
        geometry = Point(lon, lat)
        
        # Escribir el punto en el shapefile
        output.write({
            'geometry': geometry.__geo_interface__,
            'properties': {'id': i}
        })

print(f"Shapefile guardado en {shapefile_output}")

