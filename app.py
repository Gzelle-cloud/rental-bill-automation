import os
if os.path.exists('.env'):
    from dotenv import load_dotenv
    load_dotenv()

import re
import json
import anthropic
from flask import Flask, request, jsonify, send_file, render_template
from pypdf import PdfReader
from openpyxl import load_workbook
import tempfile
import shutil

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

XLSX_DEFAULT = os.path.join(os.path.dirname(__file__), 'kvitanciya.xlsx')

# vol  = row for volume (объём услуг)
# corr = row for correction (корректировки)
# pay  = row for payment total (к оплате)
# payer: 'tenant' (да) / 'landlord' (нет)
ROW_MAP = {
    # Housing services — pay rows 47-55
    'взнос на капитальный ремонт':   {'vol': 5,  'corr': 36, 'pay': 47, 'payer': 'landlord'},
    'водоотведение одн':             {'vol': 6,  'corr': 37, 'pay': 48, 'payer': 'tenant'},
    'горячее в/с (носитель) одн':    {'vol': 7,  'corr': 38, 'pay': 49, 'payer': 'tenant'},
    'горячее в/с (энергия) одн':     {'vol': 8,  'corr': 39, 'pay': 50, 'payer': 'tenant'},
    'содержание жилого помещения':   {'vol': 9,  'corr': 40, 'pay': 51, 'payer': 'landlord'},
    'холодное в/с одн':              {'vol': 10, 'corr': 41, 'pay': 52, 'payer': 'tenant'},
    'электроснабжение день одн':     {'vol': 11, 'corr': 42, 'pay': 53, 'payer': 'tenant'},
    'электроснабжение ночь одн':     {'vol': 12, 'corr': 43, 'pay': 54, 'payer': 'tenant'},
    'электроснабжение одн':          {'vol': 13, 'corr': 44, 'pay': 55, 'payer': 'tenant'},
    # Utility services — pay rows 76-81
    'водоотведение':                 {'vol': 16, 'corr': 68, 'pay': 76, 'payer': 'tenant'},
    'горячее в/с (носитель)':        {'vol': 17, 'corr': 69, 'pay': 77, 'payer': 'tenant'},
    'горячее в/с (энергия)':         {'vol': 18, 'corr': 70, 'pay': 78, 'payer': 'tenant'},
    'обращение с тко':               {'vol': 19, 'corr': 71, 'pay': 79, 'payer': 'landlord'},
    'обращение c тко':               {'vol': 19, 'corr': 71, 'pay': 79, 'payer': 'landlord'},
    'отопление':                     {'vol': 20, 'corr': 72, 'pay': 80, 'payer': 'tenant'},
    'холодное в/с':                  {'vol': 21, 'corr': 73, 'pay': 81, 'payer': 'tenant'},
}

SYSTEM_PROMPT = """Ты — эксперт по обработке российских квитанций ЖКХ.
Из текста квитанции извлеки данные ТОЛЬКО из первой основной таблицы (расчёт размера платы за жилищные и коммунальные услуги).

Верни JSON строго в таком формате:
{
  "period": "март 2026",
  "columns_present": ["перерасчеты", "задолженность", "оплачено"],
  "services": [
    {
      "name": "точное название услуги из квитанции",
      "volume": число или null,
      "perechet": число,
      "zadolzhennost": число,
      "oplacheno": число
    }
  ]
}

ПРАВИЛА:
- "volume" = значение из колонки "Объем услуг"
- "perechet" = значение из колонки перерасчётов (0 если колонки нет)
- "zadolzhennost" = значение из колонки задолженности/переплаты, сохраняй знак (0 если нет)
- "oplacheno" = значение из колонки "Оплачено, руб." (0 если нет)
Строку "Добровольное страхование" игнорируй.
Возвращай ТОЛЬКО валидный JSON без markdown."""


def extract_text_from_pdf(pdf_bytes):
    with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as f:
        f.write(pdf_bytes)
        f.flush()
        reader = PdfReader(f.name)
        text = ""
        for page in reader.pages:
            text += page.extract_text() + "\n"
    os.unlink(f.name)
    return text


def normalize(name: str) -> str:
    s = name.strip().lower().replace('ё', 'е')
    return re.sub(r'\s+', ' ', s)


def to_float(v) -> float:
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    return float(str(v).replace('\xa0', '').replace(' ', '').replace(',', '.'))


def parse_with_claude(pdf_text: str) -> dict:
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise RuntimeError("ANTHROPIC_API_KEY is missing in environment")
    client = anthropic.Anthropic(api_key=api_key)
    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=3000,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": f"Текст квитанции:\n\n{pdf_text}"}]
    )
    raw = response.content[0].text.strip()
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    parsed = json.loads(raw.strip())
    print(f"\n=== AI извлёк: {parsed.get('period')} ===")
    for s in parsed.get('services', []):
        print(f"  {s['name']}: vol={s.get('volume')}, per={s.get('perechet',0)}, zadolzh={s.get('zadolzhennost',0)}, opl={s.get('oplacheno',0)}")
    return parsed


def find_row_info(name: str) -> dict | None:
    """Exact match first, then word-boundary fuzzy match."""
    key = normalize(name)
    if key in ROW_MAP:
        return ROW_MAP[key]
    for k, v in ROW_MAP.items():
        if re.search(r'\b' + re.escape(k) + r'\b', key):
            return v
    return None


def find_month_column(wb) -> int | None:
    """Find next empty column by checking row 5 (actual volume data)."""
    ws = wb['Рассчёты']
    for col in range(5, 18):
        if ws.cell(row=5, column=col).value in (None, ''):
            return col
    return None


def calc_correction(svc: dict) -> float:
    """correction = perechet + zadolzhennost - oplacheno"""
    return round(
        to_float(svc.get('perechet')) +
        to_float(svc.get('zadolzhennost')) -
        to_float(svc.get('oplacheno')),
        2
    )


def calc_totals(parsed: dict, ws, col: int, electricity_ipu: float) -> tuple[float, float]:
    """
    Compute tenant and landlord totals directly in Python.
    Formula per service: к_оплате = tariff * volume + correction
    Tariff is read from col D (column 4) of the Excel template.
    Matches Excel formula: =SUM($D{vol_row} * {C}{vol_row}, {C}{corr_row})
    """
    tenant = 0.0
    landlord = 0.0
    seen = set()

    columns_present = parsed.get('columns_present', [])
    for svc in parsed.get('services', []):
        row_info = find_row_info(svc['name'])
        if not row_info or row_info['pay'] in seen:
            continue
        seen.add(row_info['pay'])

        # Tariff from col D, volume from parsed PDF
        tariff = to_float(ws.cell(row=row_info['vol'], column=4).value)
        volume = to_float(svc.get('volume'))
        corr = calc_correction(svc)

        # к оплате = tariff * volume + correction (mirrors Excel formula)
        # For utility services (3.3): IF(result > 0, result, 0)
        amount = tariff * volume + corr
        if row_info['pay'] >= 76:  # utility services have IF > 0 guard
            amount = max(amount, 0.0)
        amount = round(amount, 2)

        print(f"  [{svc['name']}]: tariff={tariff} * vol={volume} + corr={corr} = {amount} ({row_info['payer']})")

        if row_info['payer'] == 'tenant':
            tenant += amount
        else:
            landlord += amount

    tenant += electricity_ipu
    return round(tenant, 2), round(landlord, 2)


def write_formulas(ws, col: int):
    from openpyxl.utils import get_column_letter
    C = get_column_letter(col)

    # 2.1 Начислено по тарифу (жилищные)
    for calc_row, vol_row in {25:5, 26:6, 27:7, 28:8, 29:9, 30:10, 31:11, 32:12}.items():
        ws.cell(row=calc_row, column=col).value = f'=$D{vol_row} * {C}{vol_row}'
    ws.cell(row=34, column=col).value = f'=SUM({C}25:{C}33)'

    # 2.2 Корректировки (жилищные) — данные из PDF
    ws.cell(row=45, column=col).value = f'=SUM({C}36:{C}44)'

    # 2.3 К оплате (жилищные)
    for pay_row, calc_row, corr_row in [(47,25,36),(48,26,37),(49,27,38),(50,28,39),
                                         (51,29,40),(52,30,41),(53,31,42),(54,32,43)]:
        ws.cell(row=pay_row, column=col).value = f'=SUM({C}{calc_row},{C}{corr_row})'
    ws.cell(row=55, column=col).value = f'=SUM({C}33,{C}44)'
    ws.cell(row=56, column=col).value = f'=SUM({C}47:{C}55)'

    # 3.1 Начислено по тарифу (коммунальные)
    for calc_row, vol_row in {60:16, 61:17, 62:18, 63:19, 64:20, 65:21}.items():
        ws.cell(row=calc_row, column=col).value = f'=$D{vol_row} * {C}{vol_row}'
    ws.cell(row=66, column=col).value = f'=SUM({C}60:{C}65)'

    # 3.2 Корректировки (коммунальные) — данные из PDF
    ws.cell(row=74, column=col).value = f'=SUM({C}68:{C}73)'

    # 3.3 К оплате (коммунальные) с IF > 0
    for pay_row, calc_row, corr_row in [(76,60,68),(77,61,69),(78,62,70),
                                         (79,63,71),(80,64,72),(81,65,73)]:
        ws.cell(row=pay_row, column=col).value = f'=IF(SUM({C}{calc_row},{C}{corr_row})>0,SUM({C}{calc_row},{C}{corr_row}),0)'
    ws.cell(row=82, column=col).value = f'=SUM({C}76:{C}81)'

    # 4. Итоги
    ws.cell(row=84, column=col).value = f'=SUM({C}56,{C}82)'
    ws.cell(row=86, column=col).value = f'=SUMIFS({C}$23:{C}$82,$B$23:$B$82,$B86)'
    ws.cell(row=87, column=col).value = f'=SUMIFS({C}$23:{C}$82,$B$23:$B$82,$B87)'
    ws.cell(row=88, column=col).value = f'=SUM({C}86:{C}87)-{C}84'
    ws.cell(row=92, column=col).value = f'=SUM({C}86,{C}90)'


def fill_service_data(ws, parsed: dict, col: int):
    """Write volumes and corrections from PDF into worksheet."""
    for svc in parsed.get('services', []):
        row_info = find_row_info(svc['name'])
        if row_info:
            vol = svc.get('volume')
            if vol is not None:
                ws.cell(row=row_info['vol'], column=col).value = vol
            ws.cell(row=row_info['corr'], column=col).value = calc_correction(svc)
        else:
            print(f"  [SKIP — не найдено]: {svc['name']}")


def write_to_excel(parsed: dict, electricity_ipu: float, output_path: str, source_xlsx: str = None):
    src = source_xlsx if source_xlsx else XLSX_DEFAULT
    shutil.copy(src, output_path)
    wb = load_workbook(output_path)
    ws = wb['Рассчёты']

    col = find_month_column(wb)
    if col is None:
        raise ValueError("Нет свободных колонок в таблице (все 12 месяцев заполнены)")

    write_formulas(ws, col)
    fill_service_data(ws, parsed, col)
    ws.cell(row=90, column=col).value = electricity_ipu

    # Calculate totals before saving (tariffs readable from template)
    tenant_total, landlord_total = calc_totals(parsed, ws, col, electricity_ipu)

    wb.save(output_path)
    return col, tenant_total, landlord_total


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/process', methods=['POST'])
def process():
    if 'pdf' not in request.files:
        return jsonify({'error': 'PDF файл не загружен'}), 400

    pdf_file = request.files['pdf']
    electricity = request.form.get('electricity', '0')
    xlsx_file = request.files.get('xlsx')

    try:
        electricity_val = float(electricity.replace(',', '.'))
    except ValueError:
        return jsonify({'error': 'Некорректная сумма электроэнергии'}), 400

    source_xlsx = None
    if xlsx_file and xlsx_file.filename:
        tmp_xlsx = os.path.join(tempfile.gettempdir(), 'source_kvitanciya.xlsx')
        xlsx_file.save(tmp_xlsx)
        source_xlsx = tmp_xlsx

    try:
        pdf_text = extract_text_from_pdf(pdf_file.read())
    except Exception as e:
        return jsonify({'error': f'Ошибка чтения PDF: {str(e)}'}), 400

    try:
        parsed = parse_with_claude(pdf_text)
    except Exception as e:
        return jsonify({'error': f'Ошибка AI-обработки: {str(e)}'}), 500

    period = parsed.get('period', 'updated').replace(' ', '_')
    out_name = f'Квитанции_{period}.xlsx'
    output_path = os.path.join(tempfile.gettempdir(), out_name)

    try:
        col, tenant_total, landlord_total = write_to_excel(
            parsed, electricity_val, output_path, source_xlsx=source_xlsx
        )
    except Exception as e:
        return jsonify({'error': f'Ошибка записи в Excel: {str(e)}'}), 500

    return jsonify({
        'success': True,
        'period': parsed.get('period'),
        'services_count': len(parsed.get('services', [])),
        'electricity_ipu': electricity_val,
        'filename': out_name,
        'tenant_total': tenant_total,
        'landlord_total': landlord_total,
    })


@app.route('/download')
def download():
    fname = request.args.get('file', 'Квитанции_updated.xlsx')
    path = os.path.join(tempfile.gettempdir(), fname)
    if not os.path.exists(path):
        return jsonify({'error': 'Файл не найден'}), 404
    return send_file(path, as_attachment=True, download_name=fname)


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5050))
    app.run(host='0.0.0.0', port=port)
