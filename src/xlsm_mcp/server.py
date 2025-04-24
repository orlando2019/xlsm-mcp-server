import sys
import json
import asyncio
import logging
import os
from typing import Any, List, Dict, Optional

from mcp.server.fastmcp import FastMCP

# Importar excepciones
from xlsm_mcp.exceptions import (
    ValidationError,
    WorkbookError,
    SheetError,
    DataError,
    MacroError,
    FormattingError
)

# Importar funcionalidades
from xlsm_mcp.workbook import get_workbook_info, create_workbook
from xlsm_mcp.sheet import create_worksheet, copy_sheet, delete_sheet, rename_sheet
from xlsm_mcp.data import read_excel_range, write_data
from xlsm_mcp.macros import list_macros, get_macro_info
from xlsm_mcp.formatting import format_range

# Configurar logger
logger = logging.getLogger("xlsm-mcp")

# Inicializar FastMCP
mcp = FastMCP("xlsm-mcp", description="Servidor MCP para manipular archivos Excel con macros (.xlsm)")

@mcp.tool()
async def read_data_from_excel(
    filepath: str,
    sheet_name: str,
    start_cell: str = "A1",
    end_cell: Optional[str] = None,
    include_formulas: bool = False
) -> Dict[str, Any]:
    """
    Lee datos de una hoja de Excel.
    
    Args:
        filepath: Ruta al archivo Excel
        sheet_name: Nombre de la hoja
        start_cell: Celda inicial (por defecto "A1")
        end_cell: Celda final (opcional)
        include_formulas: Si es True, incluye las fórmulas en lugar de sus valores
        
    Returns:
        Diccionario con los datos leídos
    """
    try:
        # Validar que el archivo existe
        if not os.path.exists(filepath):
            error_msg = f"El archivo no existe en la ruta: {filepath}"
            logger.error(error_msg)
            return {
                "success": False,
                "error": error_msg,
                "error_type": "FileNotFoundError"
            }
            
        data = read_excel_range(filepath, sheet_name, start_cell, end_cell, include_formulas)
        return {
            "success": True,
            "data": data,
            "message": f"Datos leídos correctamente de {sheet_name} ({start_cell} a {end_cell or 'final'})"
        }
    except ValidationError as e:
        error_msg = f"Error de validación: {str(e)}"
        logger.error(error_msg)
        return {
            "success": False,
            "error": error_msg,
            "error_type": "ValidationError"
        }
    except WorkbookError as e:
        error_msg = f"Error con el libro Excel: {str(e)}"
        logger.error(error_msg)
        return {
            "success": False,
            "error": error_msg,
            "error_type": "WorkbookError"
        }
    except SheetError as e:
        error_msg = f"Error con la hoja '{sheet_name}': {str(e)}"
        logger.error(error_msg)
        return {
            "success": False,
            "error": error_msg,
            "error_type": "SheetError"
        }
    except DataError as e:
        error_msg = f"Error con los datos: {str(e)}"
        logger.error(error_msg)
        return {
            "success": False,
            "error": error_msg,
            "error_type": "DataError"
        }
    except Exception as e:
        error_msg = f"Error inesperado al leer datos: {str(e)}"
        logger.error(error_msg)
        return {
            "success": False,
            "error": error_msg,
            "error_type": "UnexpectedError"
        }

@mcp.tool()
async def write_data_to_excel(
    filepath: str,
    sheet_name: str,
    data: List[Dict],
    start_cell: str = "A1"
) -> Dict[str, Any]:
    """
    Escribe datos en una hoja de Excel.
    
    Args:
        filepath: Ruta al archivo Excel
        sheet_name: Nombre de la hoja
        data: Lista de diccionarios con los datos a escribir
        start_cell: Celda inicial donde comenzar a escribir
        
    Returns:
        Diccionario con el resultado de la operación
    """
    try:
        result = write_data(filepath, sheet_name, data, start_cell)
        return {
            "success": True,
            "message": f"Datos escritos correctamente en {sheet_name} desde {start_cell}"
        }
    except Exception as e:
        logger.error(f"Error al escribir datos: {e}")
        return {
            "success": False,
            "error": str(e)
        }

@mcp.tool()
async def create_new_workbook(
    filepath: str,
    with_macros: bool = True
) -> Dict[str, Any]:
    """
    Crea un nuevo libro de Excel con opción de habilitar macros.
    
    Args:
        filepath: Ruta donde guardar el archivo
        with_macros: Si es True, crea un archivo .xlsm con macros habilitadas
        
    Returns:
        Diccionario con el resultado de la operación
    """
    try:
        create_workbook(filepath, with_macros)
        return {
            "success": True,
            "message": f"Libro creado correctamente en {filepath}"
        }
    except Exception as e:
        logger.error(f"Error al crear libro: {e}")
        return {
            "success": False,
            "error": str(e)
        }

@mcp.tool()
async def create_new_worksheet(
    filepath: str,
    sheet_name: str
) -> Dict[str, Any]:
    """
    Crea una nueva hoja en un libro de Excel existente.
    
    Args:
        filepath: Ruta al archivo Excel
        sheet_name: Nombre de la nueva hoja
        
    Returns:
        Diccionario con el resultado de la operación
    """
    try:
        create_worksheet(filepath, sheet_name)
        return {
            "success": True,
            "message": f"Hoja '{sheet_name}' creada correctamente"
        }
    except Exception as e:
        logger.error(f"Error al crear hoja: {e}")
        return {
            "success": False,
            "error": str(e)
        }

@mcp.tool()
async def get_workbook_metadata(
    filepath: str,
    include_macros: bool = True
) -> Dict[str, Any]:
    """
    Obtiene metadatos de un libro Excel incluyendo información sobre hojas y macros.
    
    Args:
        filepath: Ruta al archivo Excel
        include_macros: Si es True, incluye información sobre macros disponibles
        
    Returns:
        Diccionario con metadatos del libro
    """
    try:
        info = get_workbook_info(filepath, include_macros)
        return {
            "success": True,
            "data": info
        }
    except Exception as e:
        logger.error(f"Error al obtener metadatos: {e}")
        return {
            "success": False,
            "error": str(e)
        }

@mcp.tool()
async def list_macros_in_workbook(
    filepath: str
) -> Dict[str, Any]:
    """
    Lista todas las macros disponibles en un libro Excel.
    
    Args:
        filepath: Ruta al archivo Excel
        
    Returns:
        Diccionario con lista de macros
    """
    try:
        macros = list_macros(filepath)
        return {
            "success": True,
            "data": macros
        }
    except Exception as e:
        logger.error(f"Error al listar macros: {e}")
        return {
            "success": False,
            "error": str(e)
        }

@mcp.tool()
async def get_macro_details(
    filepath: str,
    macro_name: str
) -> Dict[str, Any]:
    """
    Obtiene detalles sobre una macro específica.
    
    Args:
        filepath: Ruta al archivo Excel
        macro_name: Nombre de la macro
        
    Returns:
        Diccionario con detalles de la macro
    """
    try:
        info = get_macro_info(filepath, macro_name)
        return {
            "success": True,
            "data": info
        }
    except Exception as e:
        logger.error(f"Error al obtener detalles de macro: {e}")
        return {
            "success": False,
            "error": str(e)
        }

@mcp.tool()
async def format_cell_range(
    filepath: str,
    sheet_name: str,
    start_cell: str,
    end_cell: Optional[str] = None,
    bold: bool = False,
    italic: bool = False,
    font_size: Optional[int] = None,
    font_color: Optional[str] = None,
    bg_color: Optional[str] = None
) -> Dict[str, Any]:
    """
    Aplica formato a un rango de celdas.
    
    Args:
        filepath: Ruta al archivo Excel
        sheet_name: Nombre de la hoja
        start_cell: Celda inicial
        end_cell: Celda final (opcional)
        bold: Si es True, aplica negrita
        italic: Si es True, aplica cursiva
        font_size: Tamaño de fuente (opcional)
        font_color: Color de texto en formato hex (#RRGGBB)
        bg_color: Color de fondo en formato hex (#RRGGBB)
        
    Returns:
        Diccionario con el resultado de la operación
    """
    try:
        format_range(
            filepath, 
            sheet_name, 
            start_cell, 
            end_cell, 
            bold=bold, 
            italic=italic, 
            font_size=font_size, 
            font_color=font_color, 
            bg_color=bg_color
        )
        
        return {
            "success": True,
            "message": f"Formato aplicado correctamente a rango {start_cell}:{end_cell or start_cell}"
        }
    except Exception as e:
        logger.error(f"Error al aplicar formato: {e}")
        return {
            "success": False,
            "error": str(e)
        }

def read_message():
    """Lee un mensaje del stdin."""
    line = sys.stdin.readline()
    return json.loads(line)

def write_message(message):
    """Escribe un mensaje en stdout."""
    sys.stdout.write(json.dumps(message) + "\n")
    sys.stdout.flush()

def run_server_stdio():
    """Inicia el servidor utilizando el protocolo stdio."""
    mcp.run(transport="stdio")

async def run_server_async():
    """Inicia el servidor de forma asíncrona."""
    await mcp.run_async(transport="stdio")

if __name__ == "__main__":
    run_server_stdio()