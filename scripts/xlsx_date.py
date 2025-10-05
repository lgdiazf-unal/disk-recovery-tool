"""
Funciones necesarias para poder determinar la fecha de los archivos xlsx
"""

import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
import sys
import shutil
from datetime import datetime

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

def get_year(metadata, dato="modified"):
    """
    Funcion que retorna el year
    """
    data_date = metadata.get(dato)

    if data_date:
        try:
            dt = datetime.fromisoformat(data_date.replace('Z', '+00:00'))
            return dt.year
        except ValueError:
            return None
        
        return None

def copy_to_year_folder(file_path, year, base_dir):
    """
    Funcion que copia el archivo xlsx a su respondiente carpeta
    """
    if year is None:
        year_folder = Path(base_dir) / "unknow_year"
    else:
        year_folder = Path(base_dir) / str(year)

    year_folder.mkdir(parents=True, exist_ok=True)

    file_path = Path(file_path)
    dest_path = year_folder / file_path.name
    shutil.copy2(file_path, dest_path)

    return dest_path


if __name__ == "__main__":
    op = sys.argv[1:]
    
    PATH_ARCHIVO = op[0]
    CARPETA_RAIZ = op[1]

    respuesta = get_xlsx_dates(PATH_ARCHIVO)
    year = get_year(respuesta)

    destino = copy_to_year_folder(
        PATH_ARCHIVO,
        year,
        CARPETA_RAIZ)

    print(f"el archivo {PATH_ARCHIVO} Se guarda en {year}")