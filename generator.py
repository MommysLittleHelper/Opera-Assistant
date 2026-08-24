from openpyxl import load_workbook

FIELD_MAP = {
    "номер ктк": "номер ктк",
    "REF": "Букинг",
    "статус": "Клиент",
    "фактический объем, м3": "объем, м3",
    "вес, кг": "вес, кг",
    "кол-во мест": "кол-во мест",
}

def norm(value):
    return "" if value is None else str(value).strip().lower()

def get_headers(ws):
    return {norm(cell.value): cell.column for cell in ws[1] if cell.value is not None}

def require_column(headers, name):
    col = headers.get(norm(name))
    if col is None:
        raise ValueError(f"В исходном файле не найдена колонка: {name}")
    return col

def generate_proработка(source_path, template_path, output_path):
    source_wb = load_workbook(source_path, data_only=True)
    source_ws = source_wb.active
    template_wb = load_workbook(template_path)
    template_ws = template_wb.active

    source_headers = get_headers(source_ws)
    template_headers = get_headers(template_ws)
    source_cols = {src: require_column(source_headers, src) for src in FIELD_MAP}
    template_cols = {dst: require_column(template_headers, dst) for dst in FIELD_MAP.values()}

    target_row = 2
    for source_row in range(2, source_ws.max_row + 1):
        values = [source_ws.cell(source_row, col).value for col in source_cols.values()]
        if all(v in (None, "") for v in values):
            continue
        for src_name, dst_name in FIELD_MAP.items():
            template_ws.cell(target_row, template_cols[dst_name]).value = source_ws.cell(source_row, source_cols[src_name]).value
        target_row += 1

    template_wb.save(output_path)
