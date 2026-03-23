import os

try:
    import xlwings as xw
    XLWINGS_AVAILABLE = True
except ImportError:
    XLWINGS_AVAILABLE = False


def find_input_files(folder):
    archivos = []
    try:
        for archivo in os.listdir(folder):
            if (archivo.endswith(('.xlsx', '.xls'))
                    and not archivo.startswith('~$')
                    and not archivo.startswith('output_')):
                archivos.append(archivo)
    except Exception as e:
        print(f"Error listando archivos: {e}")
    return archivos


def apply_autowidth_excel(file_path):
    if not XLWINGS_AVAILABLE:
        return True

    try:
        app = xw.App(visible=False)
        libro = app.books.open(file_path)

        for hoja in libro.sheets:
            hoja.range('B:B').column_width = 38
            hoja.range('C:C').column_width = 11

            used_range = hoja.used_range
            if used_range:
                last_col = used_range.last_cell.column
                for col in range(1, last_col + 1):
                    if col in (2, 3):
                        continue
                    letra = xw.utils.col_name(col)
                    hoja.range(f'{letra}:{letra}').columns.autofit()

        libro.save()
        libro.close()
        app.quit()
        return True

    except Exception as e:
        print(f"Error aplicando autowidth: {e}")
        try:
            libro.close()
            app.quit()
        except Exception:
            pass
        return False
