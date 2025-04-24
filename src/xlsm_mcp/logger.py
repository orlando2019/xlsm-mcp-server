"""
Módulo para configurar el sistema de logging del servidor MCP de Excel.

Este módulo configura un sistema de logging con rotación de archivos
y diferentes niveles de detalle según el entorno.
"""

import os
import sys
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Optional, Union

# Nombre del logger principal
LOGGER_NAME = "xlsm-mcp"

# Niveles de log disponibles
LOG_LEVELS = {
    "debug": logging.DEBUG,
    "info": logging.INFO,
    "warning": logging.WARNING,
    "error": logging.ERROR,
    "critical": logging.CRITICAL
}

def get_log_directory() -> Path:
    """
    Determina y crea (si no existe) el directorio para archivos de log.
    
    En sistemas Windows, usa %APPDATA%/xlsm-mcp/logs/
    En sistemas Unix/Linux, usa ~/.xlsm-mcp/logs/
    
    Returns:
        Path al directorio de logs
    """
    if sys.platform.startswith('win'):
        log_dir = Path(os.environ.get('APPDATA', '.')) / 'xlsm-mcp' / 'logs'
    else:
        log_dir = Path.home() / '.xlsm-mcp' / 'logs'
    
    # Crear directorio si no existe
    log_dir.mkdir(parents=True, exist_ok=True)
    
    return log_dir

def setup_logging(
    log_level: str = "info",
    log_file: Optional[Union[str, Path]] = None,
    max_size_mb: int = 10,
    backup_count: int = 5,
    console_output: bool = True
) -> logging.Logger:
    """
    Configura el sistema de logging para el servidor.
    
    Args:
        log_level: Nivel de log (debug, info, warning, error, critical)
        log_file: Ruta al archivo de log (si es None, se usa una ruta por defecto)
        max_size_mb: Tamaño máximo del archivo de log en MB antes de rotar
        backup_count: Número de archivos de respaldo a mantener
        console_output: Si es True, también muestra logs en la consola
        
    Returns:
        Logger configurado
    """
    # Obtener nivel de log
    level = LOG_LEVELS.get(log_level.lower(), logging.INFO)
    
    # Determinar archivo de log
    if log_file is None:
        log_dir = get_log_directory()
        log_file = log_dir / 'xlsm-mcp.log'
    else:
        log_file = Path(log_file)
        
        # Crear directorio si no existe
        log_file.parent.mkdir(parents=True, exist_ok=True)
    
    # Configurar logger principal
    logger = logging.getLogger(LOGGER_NAME)
    logger.setLevel(level)
    
    # Eliminar handlers existentes
    for handler in logger.handlers[:]:
        logger.removeHandler(handler)
    
    # Formato del log
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    # Configurar handler de archivo con rotación
    file_handler = RotatingFileHandler(
        log_file,
        maxBytes=max_size_mb * 1024 * 1024,  # Convertir MB a bytes
        backupCount=backup_count,
        encoding='utf-8'
    )
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    
    # Añadir handler para consola
    if console_output:
        console_handler = logging.StreamHandler(sys.stderr)
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)
    
    # Evitar propagar logs a los handlers raíz
    logger.propagate = False
    
    logger.debug(f"Sistema de logging inicializado. Nivel: {log_level}, Archivo: {log_file}")
    
    return logger

def get_logger() -> logging.Logger:
    """
    Obtiene el logger configurado o uno por defecto si no se ha configurado.
    
    Returns:
        Logger configurado
    """
    logger = logging.getLogger(LOGGER_NAME)
    
    # Si el logger no tiene handlers, configurarlo con valores por defecto
    if not logger.handlers:
        return setup_logging()
    
    return logger 