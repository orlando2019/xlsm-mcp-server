"""
Módulo para validación de datos de entrada en el servidor MCP.

Este módulo proporciona funciones de validación para diferentes tipos de datos
utilizados en el servidor, como rutas de archivo, referencias de celdas, etc.
"""

import os
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

from xlsm_mcp.exceptions import ValidationError

def validate_file_path(
    filepath: Union[str, Path], 
    must_exist: bool = True,
    file_extensions: Optional[List[str]] = None
) -> Path:
    """
    Valida que una ruta de archivo sea válida y cumpla con los requisitos.
    
    Args:
        filepath: Ruta a validar
        must_exist: Si es True, verifica que el archivo exista
        file_extensions: Lista de extensiones permitidas (ej. ['.xlsx', '.xlsm'])
    
    Returns:
        Objeto Path con la ruta validada
    
    Raises:
        ValidationError: Si la validación falla
    """
    try:
        # Convertir a objeto Path
        path = Path(filepath)
        
        # Verificar si debe existir
        if must_exist and not path.exists():
            raise ValidationError(f"El archivo {path} no existe")
        
        # Verificar extensión si es necesario
        if file_extensions:
            if not any(path.name.lower().endswith(ext.lower()) for ext in file_extensions):
                valid_exts = ", ".join(file_extensions)
                raise ValidationError(
                    f"El archivo {path.name} no tiene una extensión válida. "
                    f"Extensiones permitidas: {valid_exts}"
                )
        
        return path
    except Exception as e:
        if isinstance(e, ValidationError):
            raise
        raise ValidationError(f"Ruta de archivo inválida: {str(e)}")

def validate_sheet_name(sheet_name: str, max_length: int = 31) -> str:
    """
    Valida que un nombre de hoja sea válido según las restricciones de Excel.
    
    Args:
        sheet_name: Nombre de hoja a validar
        max_length: Longitud máxima permitida (31 en Excel)
    
    Returns:
        Nombre de hoja validado
    
    Raises:
        ValidationError: Si el nombre no es válido
    """
    if not sheet_name:
        raise ValidationError("El nombre de la hoja no puede estar vacío")
    
    if len(sheet_name) > max_length:
        raise ValidationError(f"El nombre de la hoja no puede tener más de {max_length} caracteres")
    
    # Caracteres prohibidos en nombres de hojas de Excel
    invalid_chars = ['\\', '/', '?', '*', '[', ']', ':', ' ']
    for char in invalid_chars:
        if char in sheet_name:
            raise ValidationError(f"El nombre de la hoja contiene caracteres inválidos: '{char}'")
    
    return sheet_name

def validate_cell_reference(cell_ref: str) -> str:
    """
    Valida que una referencia de celda tenga un formato válido (ej. A1, B2, etc.).
    
    Args:
        cell_ref: Referencia de celda a validar
    
    Returns:
        Referencia de celda validada
    
    Raises:
        ValidationError: Si la referencia no es válida
    """
    if not cell_ref:
        raise ValidationError("La referencia de celda no puede estar vacía")
    
    # Formato de celda: una o más letras seguidas de uno o más números
    pattern = re.compile(r'^[A-Za-z]+[1-9]\d*$')
    if not pattern.match(cell_ref):
        raise ValidationError(f"Formato de referencia de celda inválido: {cell_ref}")
    
    return cell_ref.upper()

def validate_cell_range(start_cell: str, end_cell: Optional[str] = None) -> Tuple[str, Optional[str]]:
    """
    Valida un rango de celdas.
    
    Args:
        start_cell: Celda inicial
        end_cell: Celda final (opcional)
    
    Returns:
        Tupla (celda_inicial, celda_final) validada
    
    Raises:
        ValidationError: Si el rango no es válido
    """
    # Validar celda inicial
    start = validate_cell_reference(start_cell)
    
    # Si hay celda final, validarla
    if end_cell:
        end = validate_cell_reference(end_cell)
        
        # Extraer componentes para comprobar orden
        start_col, start_row = split_cell_reference(start)
        end_col, end_row = split_cell_reference(end)
        
        # Verificar que la celda final está a la derecha y abajo de la inicial
        if end_col < start_col or end_row < start_row:
            raise ValidationError(
                f"Rango inválido: la celda final {end_cell} debe estar a la derecha "
                f"y abajo de la celda inicial {start_cell}"
            )
        
        return start, end
    
    return start, None

def split_cell_reference(cell_ref: str) -> Tuple[str, int]:
    """
    Divide una referencia de celda en sus componentes de columna y fila.
    
    Args:
        cell_ref: Referencia de celda (ej. A1, B2)
    
    Returns:
        Tupla (columna, fila)
    """
    # Asegurar que la celda es válida
    cell_ref = validate_cell_reference(cell_ref)
    
    # Encontrar el índice donde terminan las letras
    idx = 0
    while idx < len(cell_ref) and cell_ref[idx].isalpha():
        idx += 1
    
    # Dividir en componentes
    col = cell_ref[:idx]
    row = int(cell_ref[idx:])
    
    return col, row

def validate_excel_data(data: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Valida que los datos a escribir en Excel tengan un formato válido.
    
    Args:
        data: Lista de diccionarios con datos a escribir
    
    Returns:
        Datos validados
    
    Raises:
        ValidationError: Si los datos no son válidos
    """
    if not isinstance(data, list):
        raise ValidationError("Los datos deben ser una lista de diccionarios")
    
    if not data:
        raise ValidationError("La lista de datos no puede estar vacía")
    
    for i, row in enumerate(data):
        if not isinstance(row, dict):
            raise ValidationError(f"La fila {i+1} debe ser un diccionario")
        
        if not row:
            raise ValidationError(f"La fila {i+1} no puede estar vacía")
    
    return data

def validate_color(color: str) -> str:
    """
    Valida que un color tenga un formato hexadecimal válido.
    
    Args:
        color: Color en formato hexadecimal (ej. #FF0000)
    
    Returns:
        Color validado
    
    Raises:
        ValidationError: Si el color no es válido
    """
    if not color:
        raise ValidationError("El color no puede estar vacío")
    
    # Aceptar con o sin #
    if color.startswith('#'):
        hex_color = color[1:]
    else:
        hex_color = color
    
    # Verificar longitud (RGB o RGBA)
    if len(hex_color) not in [6, 8]:
        raise ValidationError(f"Formato de color inválido: {color}")
    
    # Verificar que sean caracteres hexadecimales
    try:
        int(hex_color, 16)
    except ValueError:
        raise ValidationError(f"Formato de color inválido: {color}")
    
    # Normalizar a formato con #
    return f"#{hex_color.upper()}" 