import sys
import json
import logging
import argparse
from .server import run_server_stdio
from .logger import setup_logging

def parse_arguments():
    """
    Analiza los argumentos de línea de comandos.
    
    Returns:
        Objeto con los argumentos parseados
    """
    parser = argparse.ArgumentParser(
        description="Servidor MCP para archivos Excel con macros (.xlsm)"
    )
    
    parser.add_argument(
        "--log-level",
        choices=["debug", "info", "warning", "error", "critical"],
        default="info",
        help="Nivel de detalle para el log (default: info)"
    )
    
    parser.add_argument(
        "--log-file",
        help="Ruta al archivo de log (por defecto se usa una ubicación estándar)"
    )
    
    parser.add_argument(
        "--no-console-log",
        action="store_true",
        help="Deshabilita la salida de logs a la consola"
    )
    
    return parser.parse_args()

def main():
    """
    Inicia el servidor MCP para archivos XLSM utilizando el protocolo stdio.
    Este servidor permite manipular archivos Excel con macros (.xlsm).
    """
    try:
        # Parsear argumentos
        args = parse_arguments()
        
        # Configurar logging
        logger = setup_logging(
            log_level=args.log_level,
            log_file=args.log_file,
            console_output=not args.no_console_log
        )
        
        # Mostrar mensaje solo en stderr para no interferir con el protocolo stdio
        print("Servidor MCP para XLSM iniciado", file=sys.stderr)
        logger.info("Servidor iniciado")
        
        # Iniciar servidor usando el protocolo stdio
        run_server_stdio()
    except Exception as e:
        logger = logging.getLogger("xlsm-mcp")
        logger.error(f"Error al iniciar el servidor: {e}")
        import traceback
        logger.error(traceback.format_exc())
        sys.exit(1)

if __name__ == "__main__":
    main()