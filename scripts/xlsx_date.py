"""
Funciones necesarias para poder determinar la fecha de los archivos xlsx
"""

import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
import sys

def get_xlsx_dates(archivo):
    """
    Funcion que retorna la fecha de creacion y de modificacion de un archivo
    """

    file_path = Path(archivo)

    if not file_path.exists() or file_path.suffix.lower() != ".xlsx":
        return None
    
    try:
        with zipfile.ZipFile(file_path, 'r') as z:
            if "docProps/core.xml" in z.namelist():
                with z.open("docProps/core.xml") as core:
                    tree = ET.parse(core)
                    root = tree.getroot()

                    ns = {
                        'cp' : 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
                        'dc' : 'http://purl.org/dc/elements/1.1/',
                        'dcterms' : 'http://purl.org/dc/terms/'
                    }

                    created = root.find('dcterms:created', ns)
                    modified = root.find('dcterms:modified', ns)

                    return {
                        'file' : str(file_path),
                        'created' : created.text if created is not None else None,
                        'modified' : modified.text if modified is not None else None 
                    }
            else:
                return {
                    'file' : str(file_path),
                    'created' : None,
                    'modified' : None
                }
    except Exception as e:
        return {
                    'file' : str(file_path),
                    'error' : str(e)
                }

if __name__ == "__main__":
    op = sys.argv[1:]
    respuesta = get_xlsx_dates(op[0])

    print(f"el archivo {op[0]} tiene {respuesta}")