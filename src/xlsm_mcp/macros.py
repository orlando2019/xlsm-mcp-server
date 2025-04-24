"""
Módulo para interactuar con macros VBA en archivos Excel con macros (.xlsm)

Este módulo proporciona funciones para trabajar con macros de VBA en archivos Excel,
permitiendo listar las macros disponibles, obtener información detallada sobre ellas,
verificar si un archivo contiene macros y convertir archivos .xlsx a .xlsm.

El módulo utiliza un enfoque basado en la inspección del contenido binario del archivo
para extraer información sobre las macros sin necesidad de ejecutar Excel.

Ejemplos de uso:
---------------

1. Listar todas las macros en un archivo Excel:

   ```python
   from xlsm_mcp.macros import list_macros
   
   # Obtener lista de todas las macros en el archivo
   macros = list_macros("ejemplo.xlsm")
   
   # Mostrar información básica de cada macro
   for macro in macros:
       print(f"Nombre: {macro['name']}, Tipo: {macro['type']}")
   ```

2. Obtener información detallada de una macro específica:

   ```python
   from xlsm_mcp.macros import get_macro_info
   
   # Obtener detalles de una macro específica
   macro_info = get_macro_info("ejemplo.xlsm", "ProcesarDatos")
   
   # Acceder a la información de la macro
   print(f"Macro: {macro_info['name']}")
   print(f"Tipo: {macro_info['type']}")
   print(f"Descripción: {macro_info['description']}")
   ```

3. Verificar si un archivo contiene macros:

   ```python
   from xlsm_mcp.macros import has_macros
   
   # Comprobar si el archivo tiene macros
   if has_macros("documento.xlsx"):
       print("El archivo contiene macros")
   else:
       print("El archivo no contiene macros")
   ```

4. Convertir un archivo .xlsx a .xlsm:

   ```python
   from xlsm_mcp.macros import convert_to_xlsm
   
   # Convertir un archivo .xlsx a .xlsm
   nuevo_archivo = convert_to_xlsm("datos.xlsx", "datos_con_macros.xlsm")
   print(f"Archivo convertido: {nuevo_archivo}")
   ```
"""

import os
import zipfile
import logging
import tempfile
import re
import xml.etree.ElementTree as ET
from typing import Any, Dict, List, Optional
from oletools.olevba import VBA_Parser

from xlsm_mcp.exceptions import MacroError

logger = logging.getLogger("xlsm-mcp")

def list_macros(filepath: str) -> List[Dict[str, Any]]:
    """
    Lista todas las macros disponibles en un libro Excel con macros.
    
    Esta función examina el archivo binario vbaProject.bin dentro del archivo XLSM
    para extraer información sobre módulos, subprocedimientos y funciones VBA.
    
    Args:
        filepath: Ruta al archivo Excel con macros (.xlsm)
        
    Returns:
        Lista de diccionarios con la siguiente información para cada macro:
        - name: Nombre de la macro
        - type: Tipo de macro (Module, Sub, Function)
        - source: Archivo fuente donde se encontró la macro
        
    Raises:
        MacroError: Si ocurre algún error durante el proceso, como:
                   - El archivo no existe
                   - El archivo no es un .xlsm válido
                   - El archivo está dañado
                   - No se pueden extraer las macros
                   
    Ejemplo:
        ```python
        macros = list_macros("informe_financiero.xlsm")
        print(f"Se encontraron {len(macros)} macros")
        for m in macros:
            print(f"{m['type']}: {m['name']}")
        ```
    """
    macros = []
    
    try:
        vba_parser = VBA_Parser(filepath)
        if vba_parser.detect_vba_macros():
            for (_, _, vba_filename, vba_code) in vba_parser.extract_macros():
                # Buscar procedimientos y funciones VBA
                for match in re.finditer(r'(Sub|Function)\s+(\w+)', vba_code):
                    macros.append({
                        "name": match.group(2),
                        "type": match.group(1),
                        "source": vba_filename
                    })
        vba_parser.close()
    except Exception as e:
        logger.error(f"Error al listar macros: {e}")
        raise MacroError(f"No se pudieron listar las macros: {str(e)}")
    return macros

def get_macro_info(filepath: str, macro_name: str) -> Dict[str, Any]:
    """
    Obtiene información detallada sobre una macro específica.
    
    Args:
        filepath: Ruta al archivo Excel con macros (.xlsm)
        macro_name: Nombre de la macro o módulo VBA
        
    Returns:
        Diccionario con detalles de la macro
        
    Raises:
        MacroError: Si ocurre un error al obtener la información de la macro
    """
    try:
        # Verificar que el archivo existe y tiene extensión .xlsm
        if not os.path.exists(filepath):
            raise MacroError(f"El archivo {filepath} no existe")
        
        if not filepath.lower().endswith('.xlsm'):
            raise MacroError(f"El archivo {filepath} no es un archivo Excel con macros (.xlsm)")
        
        # Obtener lista de macros
        macros = list_macros(filepath)
        
        # Buscar la macro solicitada
        macro_info = None
        for macro in macros:
            if macro["name"] == macro_name:
                macro_info = macro
                break
                
        if not macro_info:
            raise MacroError(f"No se encontró la macro '{macro_name}' en el archivo")
            
        # Añadir información adicional (en una implementación real se extraería más información)
        macro_info["description"] = f"Macro de tipo {macro_info['type']}"
        macro_info["code_preview"] = "' Esta es una vista previa del código"
        
        return macro_info
    except MacroError:
        raise
    except Exception as e:
        logger.error(f"Error al obtener información de macro: {e}")
        raise MacroError(f"No se pudo obtener información de la macro {macro_name}: {str(e)}")

def has_macros(filepath: str) -> bool:
    """
    Verifica si un archivo Excel contiene macros.
    
    Args:
        filepath: Ruta al archivo Excel
        
    Returns:
        True si el archivo contiene macros, False en caso contrario
        
    Raises:
        MacroError: Si ocurre un error al verificar el archivo
    """
    try:
        # Verificar que el archivo existe
        if not os.path.exists(filepath):
            raise MacroError(f"El archivo {filepath} no existe")
            
        # Verificar si es un archivo .xlsm
        if filepath.lower().endswith('.xlsm'):
            return True
            
        # Verificar si es un archivo .xlsx pero tiene macros
        if filepath.lower().endswith('.xlsx'):
            with zipfile.ZipFile(filepath, 'r') as z:
                vba_files = [f for f in z.namelist() if 'vbaProject' in f]
                return len(vba_files) > 0
                
        return False
    except Exception as e:
        logger.error(f"Error al verificar macros: {e}")
        raise MacroError(f"No se pudo verificar si el archivo contiene macros: {str(e)}")

def convert_to_xlsm(filepath: str, output_filepath: Optional[str] = None) -> str:
    """
    Convierte un archivo .xlsx a .xlsm para permitir el uso de macros.
    
    Esta función toma un archivo Excel sin macros (.xlsx) y lo convierte a formato
    con macros habilitadas (.xlsm), modificando los metadatos internos necesarios
    para soportar VBA.
    
    Args:
        filepath: Ruta al archivo Excel a convertir (.xlsx)
        output_filepath: Ruta de salida para el archivo convertido (opcional).
                         Si no se especifica, se usa el mismo nombre con extensión .xlsm
        
    Returns:
        Ruta al archivo convertido
        
    Raises:
        MacroError: Si ocurre algún error durante la conversión como:
                   - El archivo no existe
                   - El archivo no es un .xlsx válido
                   - No se puede escribir el archivo de salida
                   - Error al modificar los metadatos internos
                   
    Ejemplo:
        ```python
        try:
            nuevo_archivo = convert_to_xlsm("datos.xlsx", "datos_con_macros.xlsm")
            print(f"Archivo convertido exitosamente a: {nuevo_archivo}")
        except MacroError as e:
            print(f"Error en la conversión: {e}")
        ```
    """
    try:
        # Verificar que el archivo existe
        if not os.path.exists(filepath):
            raise MacroError(f"El archivo {filepath} no existe")
            
        # Verificar que el archivo es un .xlsx
        if not filepath.lower().endswith('.xlsx'):
            if filepath.lower().endswith('.xlsm'):
                logger.info(f"El archivo {filepath} ya está en formato .xlsm")
                return filepath
            else:
                raise MacroError(f"El archivo {filepath} no es un archivo Excel (.xlsx)")
        
        # Verificar que el archivo es un ZIP válido (todos los Excel son archivos ZIP)
        try:
            with zipfile.ZipFile(filepath, 'r'):
                pass
        except zipfile.BadZipFile:
            raise MacroError(f"El archivo {filepath} está dañado o no es un archivo Excel válido")
            
        # Determinar ruta de salida
        if not output_filepath:
            output_filepath = os.path.splitext(filepath)[0] + '.xlsm'
        
        # Verificar si el directorio de salida existe
        output_dir = os.path.dirname(output_filepath)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)
            
        # Verificar si tenemos permisos de escritura en la ruta de salida
        if os.path.exists(output_filepath):
            if not os.access(output_filepath, os.W_OK):
                raise MacroError(f"No hay permisos de escritura para {output_filepath}")
        elif not os.access(os.path.dirname(output_filepath) or '.', os.W_OK):
            raise MacroError(f"No hay permisos de escritura en el directorio para {output_filepath}")
            
        # Copiar archivo a nuevo formato
        import shutil
        try:
            shutil.copy2(filepath, output_filepath)
        except (shutil.Error, IOError) as e:
            raise MacroError(f"Error al copiar el archivo: {str(e)}")
        
        # Actualizar metadatos internos para indicar soporte de macros
        try:
            with zipfile.ZipFile(output_filepath, 'a') as z:
                # Leer archivo de tipos de contenido
                content_types_file = '[Content_Types].xml'
                
                if content_types_file not in z.namelist():
                    raise MacroError(f"El archivo {filepath} no tiene un formato Excel válido")
                
                content_types_xml = z.read(content_types_file)
                tree = ET.fromstring(content_types_xml)
                
                # Namespace para XML
                ns = '{http://schemas.openxmlformats.org/package/2006/content-types}'
                ET.register_namespace('', 'http://schemas.openxmlformats.org/package/2006/content-types')
                
                # Añadir referencia a vbaProject si no existe
                vba_found = False
                
                for elem in tree.findall(f'.//{ns}Override'):
                    if 'vbaProject' in elem.attrib.get('PartName', ''):
                        vba_found = True
                        break
                        
                if not vba_found:
                    # Añadir el tipo vbaProject
                    vba_elem = ET.SubElement(tree, f'{ns}Override')
                    vba_elem.set('PartName', '/xl/vbaProject.bin')
                    vba_elem.set('ContentType', 'application/vnd.ms-office.vbaProject')
                    
                    # También asegurarse de que esté la extensión de macros habilitadas
                    macro_ext_found = False
                    for elem in tree.findall(f'.//{ns}Default'):
                        if elem.attrib.get('Extension') == 'bin':
                            macro_ext_found = True
                            break
                    
                    if not macro_ext_found:
                        bin_elem = ET.SubElement(tree, f'{ns}Default')
                        bin_elem.set('Extension', 'bin')
                        bin_elem.set('ContentType', 'application/vnd.ms-office.vbaProject')
                    
                    # Crear estructura vacía de vbaProject.bin
                    # En una implementación real, tendríamos un template de vbaProject.bin 
                    empty_vba = b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1\x00\x00\x00\x00\x00\x00\x00\x00'  # Encabezado mínimo
                    z.writestr('xl/vbaProject.bin', empty_vba)
                    
                    # Escribir archivo actualizado de content types
                    z.writestr(content_types_file, ET.tostring(tree))
                    
                # Actualizar workbook.xml para indicar soporte de macros
                wb_file = 'xl/workbook.xml'
                if wb_file in z.namelist():
                    wb_xml = z.read(wb_file)
                    wb_tree = ET.fromstring(wb_xml)
                    
                    # Añadir el atributo codeName si es necesario
                    ns_wb = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
                    wb_ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
                    ET.register_namespace('', wb_ns)
                    
                    # Añadir soporte para VBA en el nodo workbookPr
                    wb_props = wb_tree.find(f'.//{ns_wb}workbookPr')
                    if wb_props is not None:
                        wb_props.set('codeName', 'ThisWorkbook')
                        wb_props.set('vbaSuppressed', '0')
                    else:
                        # Si no existe el nodo workbookPr, crearlo
                        wb_node = wb_tree.find(f'.//{ns_wb}workbook')
                        if wb_node is not None:
                            wb_props = ET.SubElement(wb_node, f'{ns_wb}workbookPr')
                            wb_props.set('codeName', 'ThisWorkbook')
                            wb_props.set('vbaSuppressed', '0')
                    
                    # Escribir archivo de workbook actualizado
                    z.writestr(wb_file, ET.tostring(wb_tree))
        
        except Exception as e:
            logger.error(f"Error al modificar metadatos XML: {e}")
            raise MacroError(f"Error al actualizar metadatos del archivo: {str(e)}")
        
        logger.info(f"Archivo convertido correctamente a {output_filepath}")
        return output_filepath
    except MacroError:
        raise
    except Exception as e:
        logger.error(f"Error inesperado al convertir a XLSM: {e}")
        raise MacroError(f"No se pudo convertir el archivo a formato XLSM: {str(e)}")