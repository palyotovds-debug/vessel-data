from http.server import BaseHTTPRequestHandler
import json, base64, zipfile, io, re


def patch_xlsm(file_bytes, params):
    vyx_num   = params.get('vyxNum', '')
    fmt_date  = params.get('fmtDate', '')
    time_val  = params.get('time', '17:00')
    berth     = params.get('berth', '')
    sign      = params.get('sign', '')
    today_fmt = params.get('todayFmt', '')
    selected  = params.get('selectedSheets', [])

    SIGNS = [
        'Олексій АЛЬОХІН', 'Олександр ЛИТВИНОВ',
        'Морський агент: Олексій АЛЬОХІН',
        'Морський агент: Олександр ЛИТВИНОВ',
    ]

    orig = zipfile.ZipFile(io.BytesIO(file_bytes))
    wb_xml   = orig.read('xl/workbook.xml').decode('utf-8')
    rels_xml = orig.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    rid2file = dict(re.findall(r'Id="([^"]+)"[^>]+Target="([^"]+)"', rels_xml))
    smap     = re.findall(r'name="([^"]+)"[^>]+r:id="([^"]+)"', wb_xml)
    keep     = set(['DATA'] + selected)

    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w', allowZip64=True) as zout:
        for item in orig.infolist():
            data = orig.read(item.filename)

            # ── sharedStrings ────────────────────────────────────────────
            if item.filename == 'xl/sharedStrings.xml':
                t = data.decode('utf-8')

                # Подпись
                if sign:
                    for s in SIGNS:
                        t = t.replace(s, sign)

                # Вих. № → сегодня
                if vyx_num and today_fmt:
                    t = re.sub(
                        r'Вих\. № [^<]*від \d{2}\.\d{2}\.\d{4}',
                        'Вих. № ' + vyx_num + ' від ' + today_fmt, t)

                # Дата+время события "дд.мм.рррр о чч:мм"
                if fmt_date:
                    t = re.sub(
                        r'\d{2}\.\d{2}\.\d{4} о \d{2}:\d{2}',
                        fmt_date + ' о ' + time_val, t)
                    p = fmt_date.split('.')
                    if len(p) == 3:
                        t = re.sub(
                            r'«\d+» \d+\.\d+р\. о \d{2}:\d{2} год\.',
                            '«' + p[0] + '» ' + p[1] + '.' + p[2] + 'р. о ' + time_val + ' год.', t)

                # Причал — заменяем оригинальное значение '14/15; 16; 17'
                if berth:
                    be = berth.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
                    # Заменяем любой вариант причала в тексте
                    t = re.sub(r'14/15(?:; 16)?(?:; 17)?', be, t)
                    # И отдельные номера если вдруг
                    t = re.sub(r'(?<!\d)14/15(?!\d)', be, t)

                data = t.encode('utf-8')

            # ── DATA sheet (sheet1.xml) ───────────────────────────────────
            elif item.filename == 'xl/worksheets/sheet1.xml':
                t = data.decode('utf-8')

                # B4 — дата захода
                if fmt_date:
                    t = re.sub(
                        r'<c r="B4"[^>]*t="s"[^>]*><v>[^<]*</v></c>',
                        '<c r="B4" s="260" t="inlineStr"><is><t>' + fmt_date + '</t></is></c>', t)

                data = t.encode('utf-8')

            # ── workbook.xml — оставляем только нужные листы ─────────────
            elif item.filename == 'xl/workbook.xml':
                t = data.decode('utf-8')
                for name, _ in smap:
                    if name not in keep:
                        t = re.sub(
                            r'<sheet[^>]+name="' + re.escape(name) + r'"[^/]*/\s*>', '', t)
                data = t.encode('utf-8')

            ni = zipfile.ZipInfo(item.filename)
            ni.compress_type = item.compress_type
            ni.date_time = item.date_time
            zout.writestr(ni, data)

    orig.close()
    return out.getvalue()


class handler(BaseHTTPRequestHandler):

    def do_OPTIONS(self):
        self.send_response(200)
        self._cors()
        self.end_headers()

    def do_POST(self):
        try:
            body = self.rfile.read(int(self.headers.get('Content-Length', 0)))
            pl   = json.loads(body)
            res  = patch_xlsm(base64.b64decode(pl['file']), pl.get('params', {}))
            resp = json.dumps({'ok': True, 'file': base64.b64encode(res).decode()}).encode()
            self.send_response(200)
            self._cors()
            self.send_header('Content-Type', 'application/json')
            self.send_header('Content-Length', str(len(resp)))
            self.end_headers()
            self.wfile.write(resp)
        except Exception as e:
            resp = json.dumps({'ok': False, 'error': str(e)}).encode()
            self.send_response(500)
            self._cors()
            self.send_header('Content-Type', 'application/json')
            self.send_header('Content-Length', str(len(resp)))
            self.end_headers()
            self.wfile.write(resp)

    def _cors(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
