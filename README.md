# 📊 Servidor MCP para archivos Excel con macros (.xlsm)

Este servidor implementa el Model Context Protocol (MCP) para manipular archivos Excel que contienen macros (.xlsm). Utiliza el protocolo stdio para la comunicación, lo que permite integrarse fácilmente con clientes MCP como Claude.

## 🌟 ¿Qué es MCP?

MCP (Model Context Protocol) es un protocolo que permite a los modelos de lenguaje interactuar con herramientas externas. Con este servidor, Claude y otros asistentes AI pueden manipular archivos Excel con macros de forma nativa, ampliando sus capacidades para ayudar en tareas de análisis de datos y automatización de oficina.

## ✨ Características

- Creación y manipulación de archivos Excel con macros (.xlsm)
- Lectura y escritura de datos en hojas de cálculo
- Gestión de hojas (crear, eliminar, renombrar)
- Listar y obtener información de macros VBA
- Aplicar formato a rangos de celdas

## 🔧 Instalación

```bash
pip install xlsm-mcp-server
```

## 📝 Uso

### Configuración para Claude

Agrega a tu configuración de Claude:

```json
"mcpServers": {
  "xlsm": {
    "command": "python",
    "args": ["-m", "xlsm_mcp"]
  }
}
```

### 🛠️ Herramientas disponibles

- `read_data_from_excel`: Lee datos de una hoja de Excel
- `write_data_to_excel`: Escribe datos en una hoja de Excel
- `create_new_workbook`: Crea un nuevo libro de Excel con opción de habilitar macros
- `create_new_worksheet`: Crea una nueva hoja en un libro de Excel existente
- `get_workbook_metadata`: Obtiene metadatos del libro, incluyendo información sobre macros
- `list_macros_in_workbook`: Lista todas las macros disponibles en un libro
- `get_macro_details`: Obtiene información detallada sobre una macro específica
- `format_cell_range`: Aplica formato a un rango de celdas

## 💡 Ejemplos

### Leer datos de un archivo Excel

```python
# Ejemplo de uso para leer datos
result = await read_data_from_excel("ejemplo.xlsm", "Hoja1", "A1", "C10")
print(result["data"])
```

### Listar macros en un archivo

```python
# Ejemplo para listar macros
macros = await list_macros_in_workbook("ejemplo.xlsm")
for macro in macros["data"]:
    print(f"Macro: {macro['name']}")
```

## 📋 Casos de uso

Este servidor es especialmente útil para:
- Analistas de datos que trabajan con modelos AI
- Automatización de tareas administrativas
- Generación y manipulación de informes financieros
- Integración de IA con flujos de trabajo basados en Excel

## 👥 Contribuir

Las contribuciones son bienvenidas. Por favor, abre un issue o pull request en el repositorio.

## 👨‍💻 Autor

Desarrollado por Orlando Ospino ([@OrlandoOspino](https://github.com/orlando2019))

## 📄 Licencia

MIT