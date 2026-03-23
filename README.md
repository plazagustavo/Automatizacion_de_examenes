# Conversor de Excel

Script de línea de comandos que convierte archivos Excel con datos de pacientes en planillas formateadas con columnas por tipo de estudio.

## Instalación

```bash
pip install -r requirements.txt
```

`xlwings` es opcional (autowidth de columnas). `openpyxl` es opcional pero recomendado para formato profesional.

## Uso

Colocá los archivos `.xlsx` o `.xls` en la misma carpeta que `main.py` y ejecutá:

```bash
python main.py
```

Los archivos de salida se generan en la misma carpeta con el prefijo `output_sorted_`.

## Qué hace

- Lee todos los `.xlsx`/`.xls` de la carpeta (ignora los que empiezan con `output_` o `~$`)
- Extrae datos de empresa de posiciones fijas del Excel
- Procesa pacientes por CUIL, ordenados alfabéticamente A-Z
- Genera una planilla con columnas por tipo de estudio y una `X` donde corresponde
- Si `openpyxl` está disponible aplica formato profesional; si no, usa pandas como fallback
- Si `xlwings` está disponible aplica autowidth a las columnas

## Estructura

```
conversor_excel/
├── main.py
├── requirements.txt
├── core/
│   ├── patients.py        ← extracción y procesamiento de pacientes
│   ├── processor.py       ← orquestador del flujo completo
│   └── file_utils.py      ← búsqueda de archivos y autowidth
└── writers/
    ├── openpyxl_writer.py ← escritura con formato profesional
    └── pandas_writer.py   ← escritura básica (fallback)
```

---
Creado por [Gustavo Plaza](https://github.com/plazagustavo)
