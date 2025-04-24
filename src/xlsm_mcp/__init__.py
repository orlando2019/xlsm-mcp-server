"""
Servidor MCP para archivos Excel con macros (.xlsm)

Este paquete implementa el Model Context Protocol (MCP) para manipular archivos Excel
que contienen macros (.xlsm). Utiliza el protocolo stdio para la comunicación, lo que
permite integrarse fácilmente con clientes MCP como Claude.
"""

__version__ = "0.1.0"

# Exportar elementos de módulos
from .workbook import (
    create_workbook,
    get_workbook_info,
    open_workbook
)

from .sheet import (
    create_worksheet,
    copy_sheet,
    delete_sheet,
    rename_sheet
)

from .data import (
    read_excel_range,
    write_data,
    append_data
)

from .macros import (
    list_macros,
    get_macro_info,
    has_macros,
    convert_to_xlsm
)

from .formatting import (
    format_range,
    apply_conditional_formatting,
    clear_formatting,
    set_column_width,
    set_row_height,
    create_named_style,
    apply_named_style
)

# Exportar excepciones
from .exceptions import (
    XLSMBaseError,
    ValidationError,
    WorkbookError,
    SheetError,
    DataError,
    MacroError,
    FormattingError,
    RangeError
)

# Exportar funciones de servidor
from .server import (
    mcp,
    run_server_stdio,
    run_server_async
)
