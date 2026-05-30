import json
import base64
import zipfile
import io
import re
from datetime import date, timedelta
from http.server import BaseHTTPRequestHandler

def get_week_number(d):
    return d.isocalendar()[1]

def patch_xlsm(file_bytes, params):
    vyx_num = params.get('vyxNum', '')
    fmt_date = params.get('fmtDate', '')
    time_val = params.get('time', '17:00')
    berth = params.get('berth', '')
    sign = params.get('sign', '')
    today_fmt = params.get('todayFmt', '')
    selected_sheets = params.get('selectedSheets', [])

    SIGNS = ['Олексій АЛЬОХІН', 'Олександр ЛИТВИНОВ',
             'Морський агент: Олексій АЛЬОХІН', 'Морський агент: Олександр ЛИТВИНОВ']

    orig_zip = zipfile.ZipFile(io.BytesIO(file_bytes))
    out_buf = io.BytesIO()

    # Получаем mapping sheet name -> file
    wb_xml = orig_zip.read('xl/workbook.xml').decode('utf-8')
    rels_xml = orig_zip.read('xl/_rels/workbook.xml.rels').decode('utf-8')

    rid_to_file = dict(re.findall(r'Id="([^"]+)"[^>]+Target="([^"]+)"', rels_xml))
    sheet_map = re.findall(r'name="([^"]+)"[^>]+r:id="([^"]+)"', wb_xml)
    name_to_file = {}
    for name, rid in sheet_map:
        f = rid_to_file.get(rid, '')
        if f:
            name_to_file[name] = ('xl/' + f) if not f.startswith('xl/') else f

    # Листы для включения в выходной файл
    keep_sheets = set(['DATA'] + selected_sheets)

    with zipfile.ZipFile(out_buf, 'w', allowZip64=True) as new_zip:
        for item in orig_zip.infolist():
            data = orig_zip.read(item.filename)

            # Патчим sharedStrings.xml
            if item.filename == 'xl/sharedStrings.xml':
                text = data.decode('utf-8')

                # Замена подписи
                if sign:
                    for orig_sign in SIGNS:
                        text = text.replace(orig_sign, sign)

                # Дата в шапке -> сегодня
                if vyx_num and today_fmt:
                    text = re.sub(
                        r'Вих\. № [^<]*від \d{2}\.\d{2}\.\d{4}',
                        'Вих. № ' + vyx_num + ' від ' + today_fmt,
                        text
                    )

                # Дата/время события
                if fmt_date:
                    text = re.sub(
                        r'\d{2}\.\d{2}\.\d{4} о \d{2}:\d{2}',
                        fmt_date + ' о ' + time_val,
                        text
                    )
                    # LINE BOAT формат
                    date_parts = fmt_date.split('.')
                    if len(date_parts) == 3:
                        day, mon, yr = date_parts[0], date_parts[1], date_parts[2]
                        text = re.sub(
                            r'«\d+» \d+\.\d+р\. о \d{2}:\d{2} год\.',
                            '«' + day + '» ' + mon + '.' + yr + 'р. о ' + time_val + ' год.',
                            text
                        )

                data = text.encode('utf-8')

            # Патчим DATA sheet
            elif item.filename == 'xl/worksheets/sheet1.xml':
                text = data.decode('utf-8')

                if fmt_date:
                    # B4 - дата захода
                    text = re.sub(
                        r'<c r="B4"[^>]*t="s"[^>]*><v>[^<]*</v></c>',
                        '<c r="B4" s="260" t="inlineStr"><is><t>' + fmt_date + '</t></is></c>',
                        text
                    )
                if berth:
                    # B37 - причал
                    berth_safe = berth.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
                    text = re.sub(
                        r'<c r="B37"[^>]*t="s"[^>]*><v>[^<]*</v></c>',
                        '<c r="B37" s="220" t="inlineStr"><is><t>' + berth_safe + '</t></is></c>',
                        text
                    )
                data = text.encode('utf-8')

            # Патчим workbook.xml — убираем лишние листы
            elif item.filename == 'xl/workbook.xml':
                text = data.decode('utf-8')
                # Убираем листы не из keep_sheets
                for name, rid in sheet_map:
                    if name not in keep_sheets:
                        text = re.sub(
                            r'<sheet[^>]+name="' + re.escape(name) + r'"[^/]*/\s*>',
                            '',
                            text
                        )
                data = text.encode('utf-8')

            new_item = zipfile.ZipInfo(item.filename)
            new_item.compress_type = item.compress_type
            new_item.date_time = item.date_time
            new_zip.writestr(new_item, data)

    orig_zip.close()
    return out_buf.getvalue()


class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        content_length = int(self.headers.get('Content-Length', 0))
        body = self.rfile.read(content_length)

        try:
            payload = json.loads(body)
            file_b64 = payload.get('file', '')
            params = payload.get('params', {})

            file_bytes = base64.b64decode(file_b64)
            result_bytes = patch_xlsm(file_bytes, params)
            result_b64 = base64.b64encode(result_bytes).decode('utf-8')

            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(json.dumps({'ok': True, 'file': result_b64}).encode())

        except Exception as e:
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(json.dumps({'ok': False, 'error': str(e)}).encode())

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()
