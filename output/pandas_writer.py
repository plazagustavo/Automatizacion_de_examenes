import pandas as pd

COMPANY_FIELDS = [
    ('Proceso', ''),
    ('Empresa', 'Empresa'),
    ('CUIT', 'CUIT'),
    ('Contrato', 'Contrato'),
    ('Domicilio', 'Domicilio'),
    ('Localidad', 'Localidad'),
    ('Provincia', 'Provincia'),
    ('Telefono', 'Telefono'),
    ('Contacto', ''),
    ('Email', 'Email'),
]


def write_with_pandas(output_file, company_data, final_employees_data, exams_list, exam_count, patient_numbers):
    try:
        num_cols = 3 + len(exams_list)

        company_rows = []
        for label, key in COMPANY_FIELDS:
            row = [''] * num_cols
            row[1] = label
            row[2] = company_data.get(key, '') if key else ''
            company_rows.append(row)

        company_rows.append([''] * num_cols)

        headers = ['Id', 'Empleado', 'CUIL'] + exams_list

        sorted_employees = sorted(final_employees_data.items(), key=lambda x: x[1]['name'].upper())
        employee_rows = []
        for cuil, data in sorted_employees:
            row = [''] * num_cols
            row[0] = patient_numbers[cuil]
            row[1] = data['name']
            row[2] = data['cuil']
            for i, exam in enumerate(exams_list):
                if any(e.strip() == exam.strip() for e in data['exams']):
                    row[3 + i] = 'X'
            employee_rows.append(row)

        employee_rows.append([''] * num_cols)

        exam_rows = []
        for exam, count in exam_count.items():
            row = [''] * num_cols
            row[0] = count
            row[1] = exam
            exam_rows.append(row)

        all_rows = company_rows + [headers] + employee_rows + exam_rows
        pd.DataFrame(all_rows).to_excel(output_file, index=False, header=False)
        return True, len(employee_rows)

    except Exception as e:
        print(f"Error en pandas writer: {e}")
        return False, None
