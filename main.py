import os
import sys
import time

from datos.file_utils import find_input_files, apply_autowidth_excel
from datos.processor import process_excel_file

script_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(script_dir)


def main():
    carpeta = os.path.dirname(os.path.abspath(__file__))
    archivos = find_input_files(carpeta)

    if not archivos:
        print("No se encontraron archivos .xlsx o .xls para procesar en la carpeta.")
        input("\nPresiona Enter para cerrar el programa...")
        return

    print(f"Archivos encontrados: {len(archivos)}")
    for archivo in archivos:
        print(f"  • {archivo}")

    archivos_procesados = 0
    archivos_con_error = 0
    start_time_total = time.time()

    for archivo in archivos:
        print(f"\nProcesando: {archivo}")

        start_time = time.time()
        input_file = os.path.join(carpeta, archivo)
        nombre_sin_extension = os.path.splitext(archivo)[0]
        output_file = os.path.join(carpeta, f"output_sorted_{nombre_sin_extension}.xlsx")

        result = process_excel_file(input_file, output_file)
        if result and result[0]:
            apply_autowidth_excel(output_file)
            elapsed = time.time() - start_time
            print(f"Archivo procesado en {elapsed:.2f}s -> {os.path.basename(output_file)}")
            archivos_procesados += 1
        else:
            print(f"Error al procesar: {archivo}")
            archivos_con_error += 1

    elapsed_total = time.time() - start_time_total
    print(f"\nTotal: {len(archivos)} archivos | OK: {archivos_procesados} | Error: {archivos_con_error}")
    print(f"Tiempo total: {elapsed_total:.2f}s")

    input("\nPresiona Enter para cerrar el programa...")


if __name__ == "__main__":
    main()
