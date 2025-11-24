# Macro-sir Python helpers

Herramientas en Python que replican la lógica de normalización y lectura de
reportes contenida en los archivos VBA originales.

## Funciones principales

- `analysis.normalization.normalize`: genera claves de encabezado sin tildes ni
  separadores.
- `analysis.normalization.date_only_ex2` y `time_to_sec_ex`: convierten fechas
  MDY/DMY y horas `hh:mm`/`hhmm` a objetos de Python.
- `analysis.normalization.servicio_id_from_components` y
  `normalize_vehiculo_key`: producen los mismos identificadores únicos de
  servicio que VBA.
- `analysis.normalization.read_division_excel` y
  `read_visitas_excel`: detectan encabezados flexibles, leen hojas de Excel con
  `pandas.read_excel` y devuelven `DataFrame` listos para construir la línea de
  tiempo.
- `analysis.normalization.build_timeline`: combina servicios y visitas en una
  tabla ordenada para Gantt.

## Ejecución por línea de comandos

```
python -m analysis.cli \
  --division-file /ruta/division.xlsx \
  --visitas-file /ruta/visitas.xlsx \
  --output-csv timeline.csv
```

También pueden usarse variables de entorno `DIVISION_FILE`, `VISITAS_FILE`,
`OUTPUT_CSV` y `OUTPUT_XLSX` para evitar argumentos explícitos.

## Instalación y dependencias

Este proyecto está probado con Python 3.10+ y requiere `pandas>=1.5` y
`openpyxl>=3.1` para leer archivos Excel.

1. (Opcional) Crea un entorno virtual:

   ```
   python -m venv .venv
   source .venv/bin/activate
   ```

2. Instala las dependencias mínimas:

   ```
   pip install -r requirements.txt
   ```

3. Para ejecutar las pruebas y herramientas de formateo:

   ```
   pip install -r requirements-dev.txt
   ```

Las dependencias de desarrollo incluyen `pytest` para las pruebas y `black`/`ruff`
para el formateo y linting.
