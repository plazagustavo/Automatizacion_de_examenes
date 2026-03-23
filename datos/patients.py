import pandas as pd

CUIL_COL = 2
NOMBRE_COL = 15
DESCRIPCION_COL = 4

COMPANY_COLUMN_MAP = {
    'Empresa':    (0, 7),
    'CUIT':       (0, 1),
    'Contrato':   (0, 6),
    'Domicilio':  (0, 9),
    'Localidad':  (0, 16),
    'Provincia':  (0, 11),
    'Telefono':   (0, 14),
    'Email':      (0, 8),
}

PREFERRED_EXAM_ORDER = [
    "EXAMEN CLINICO",
    "AUDIOMETRIA",
    "ESPIROMETRIA",
    "CUESTIONARIO OSTEOARTICULAR COLUMNA LUMBOSACRA",
    "CUESTIONARIO DE SEGMENTOS COMPROMETIDOS",
    "RX",
    "RX DE TORAX",
]


def extract_company_data(df):
    company_data = {key: '' for key in COMPANY_COLUMN_MAP}

    if len(df) == 0:
        return company_data

    for field, (row_idx, col_idx) in COMPANY_COLUMN_MAP.items():
        if len(df.columns) > col_idx and len(df) > row_idx:
            val = df.iloc[row_idx, col_idx]
            company_data[field] = str(val) if pd.notna(val) else ''

    return company_data


def process_all_patients(df):
    patient_numbers = {}
    patient_info = {}
    unique_cuils = []

    for i in range(1, len(df)):
        try:
            if not (CUIL_COL < len(df.columns)
                    and NOMBRE_COL < len(df.columns)
                    and DESCRIPCION_COL < len(df.columns)):
                continue

            cuil_value = df.iloc[i, CUIL_COL]
            cuil = str(cuil_value).strip() if pd.notna(cuil_value) else ""

            nombre_value = df.iloc[i, NOMBRE_COL]
            nombre = str(nombre_value).strip() if pd.notna(nombre_value) else ""

            desc_value = df.iloc[i, DESCRIPCION_COL]
            descripcion = str(desc_value).strip() if pd.notna(desc_value) else ""

            if cuil and cuil.lower() != "cuil" and len(cuil) > 5:
                if cuil not in unique_cuils:
                    unique_cuils.append(cuil)
                    patient_info[cuil] = {
                        "nombre": nombre,
                        "estudios": [descripcion] if descripcion else []
                    }
                else:
                    if descripcion and descripcion not in patient_info[cuil]["estudios"]:
                        patient_info[cuil]["estudios"].append(descripcion)
        except Exception:
            continue

    cuil_nombre_pairs = [(cuil, info["nombre"]) for cuil, info in patient_info.items()]
    cuil_nombre_pairs.sort(key=lambda x: x[1].upper())

    for i, (cuil, _) in enumerate(cuil_nombre_pairs):
        patient_numbers[cuil] = i + 1

    return patient_info, patient_numbers


def build_employees_and_exams(patient_info):
    final_employees_data = {}
    exams_set = set()
    exam_count = {}

    for cuil, info in patient_info.items():
        final_employees_data[cuil] = {
            'name': info['nombre'],
            'cuil': cuil,
            'exams': info['estudios']
        }
        for exam in info['estudios']:
            if exam:
                key = exam.strip()
                exams_set.add(key)
                exam_count[key] = exam_count.get(key, 0) + 1

    exams_list = []
    for preferred in PREFERRED_EXAM_ORDER:
        matches = [e for e in exams_set if preferred.upper() in e.upper()]
        for match in matches:
            if match in exams_set:
                exams_list.append(match)
                exams_set.remove(match)

    exams_list.extend(sorted(exams_set))

    return final_employees_data, exams_list, exam_count
