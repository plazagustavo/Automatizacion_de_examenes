import traceback
import pandas as pd

from datos.patients import extract_company_data, process_all_patients, build_employees_and_exams
from output.openpyxl_writer import write_with_openpyxl, OPENPYXL_AVAILABLE
from output.pandas_writer import write_with_pandas


def process_excel_file(input_file, output_file):
    try:
        df_header = pd.read_excel(input_file)
        df_no_header = pd.read_excel(input_file, header=None)

        company_data = extract_company_data(df_header)
        patient_info, patient_numbers = process_all_patients(df_no_header)
        final_employees_data, exams_list, exam_count = build_employees_and_exams(patient_info)

        if OPENPYXL_AVAILABLE:
            result = write_with_openpyxl(output_file, company_data, final_employees_data, exams_list, exam_count, patient_numbers)
        else:
            result = write_with_pandas(output_file, company_data, final_employees_data, exams_list, exam_count, patient_numbers)

        return result

    except Exception as e:
        print(f"Error procesando {input_file}: {e}")
        traceback.print_exc()
        return False, None
