"""
Módulo para definir excepciones personalizadas del servidor MCP para archivos XLSM.
"""

class XLSMBaseError(Exception):
    """Clase base para todas las excepciones del servidor XLSM MCP."""
    pass

class ValidationError(XLSMBaseError):
    """Excepción lanzada cuando falla una validación."""
    pass

class WorkbookError(XLSMBaseError):
    """Excepción lanzada cuando hay un error relacionado con el libro de Excel."""
    pass

class SheetError(XLSMBaseError):
    """Excepción lanzada cuando hay un error relacionado con una hoja de Excel."""
    pass

class DataError(XLSMBaseError):
    """Excepción lanzada cuando hay un error al leer o escribir datos."""
    pass

class MacroError(XLSMBaseError):
    """Excepción lanzada cuando hay un error relacionado con macros."""
    pass

class FormattingError(XLSMBaseError):
    """Excepción lanzada cuando hay un error al aplicar formato."""
    pass

class RangeError(XLSMBaseError):
    """Excepción lanzada cuando hay un error con un rango de celdas."""
    pass