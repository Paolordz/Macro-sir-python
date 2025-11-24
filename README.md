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

## Convención de nombres

Todas las tablas generadas usan **nombres en inglés con `snake_case`** para
que cada columna sea autoexplicativa. Las principales son:

- `division_name`: etiqueta derivada del nombre del archivo de división
  (por ejemplo `Division Norte`).
- `vehicle_id`: identificador sanitizado del vehículo/unidad.
- `start_time` y `end_time`: límites temporales del servicio o visita.
- `distance_km` y `duration_minutes`: medidas numéricas coherentes entre
  servicios y visitas.
- `client_site`: cliente o sitio relacionado, cuando el dato existe.
- `service_id`: identificador estable generado a partir de vehículo, fecha y
  secuencia.
- `event_type`: valores `SERVICE` o `VISIT` según el origen de la fila.

## Dependencias

Se requieren `pandas` y `openpyxl` para leer archivos Excel. Instala los
paquetes con `pip install pandas openpyxl`.
