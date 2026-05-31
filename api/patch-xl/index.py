from http.server import BaseHTTPRequestHandler
import json, base64, zipfile, io, re


def set_col_width(xml, col_num, width):
    """Устанавливает ширину колонки в XML листа."""
    w = str(width)
    n = str(col_num)
    # Обновляем существующую
    replaced = re.sub(
        r'(<col[^>]+min="'+n+r'"[^>]+max="'+n+r'"[^>]+)width="[^"]+"',
        r'\1width="'+w+'"', xml)
    if replaced != xml:
        return replaced
    replaced = re.sub(
        r'(<col[^>]+max="'+n+r'"[^>]+min="'+n+r'"[^>]+)width="[^"]+"',
        r'\1width="'+w+'"', xml)
    if replaced != xml:
        return replaced
    # Добавляем новую
    new_col = '<col min="'+n+'" max="'+n+'" width="'+w+'" customWidth="1"/>'
    if '<cols>' in xml:
        return xml.replace('</cols>', new_col+'</cols>')
    return xml.replace('<sheetData>', '<cols>'+new_col+'</cols><sheetData>')


def replace_cell(xml, cell_ref, value, style='260'):
    """Заменяет или вставляет ячейку с inlineStr значением."""
    safe = value.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
    new_cell = '<c r="'+cell_ref+'" s="'+style+'" t="inlineStr"><is><t>'+safe+'</t></is></c>'
    # Пустая самозакрывающаяся ячейка
    replaced = re.sub(r'<c r="'+cell_ref+r'"[^/]*/>', new_cell, xml)
    if replaced != xml:
        return replaced
    # Непустая ячейка
    return re.sub(r'<c r="'+cell_ref+r'"[^>]*>(?:.*?)</c>', new_cell, xml, flags=re.DOTALL)


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

    orig     = zipfile.ZipFile(io.BytesIO(file_bytes))
    wb_xml   = orig.read('xl/workbook.xml').decode('utf-8')
    rels_xml = orig.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    rid2file = dict(re.findall(r'Id="([^"]+)"[^>]+Target="([^"]+)"', rels_xml))
    smap     = re.findall(r'name="([^"]+)"[^>]+r:id="([^"]+)"', wb_xml)
    keep     = set(['DATA'] + selected)

    be = berth.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;') if berth else ''

    # Маппинг имя листа → файл
    name_to_file = {}
    for name, rid in smap:
        f = rid2file.get(rid, '')
        if f:
            name_to_file[name] = ('xl/'+f) if not f.startswith('xl/') else f

    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w', allowZip64=True) as zout:
        for item in orig.infolist():
            data = orig.read(item.filename)

            # ── sharedStrings ────────────────────────────────────────────
            if item.filename == 'xl/sharedStrings.xml':
                t = data.decode('utf-8')
                if sign:
                    for s in SIGNS:
                        t = t.replace(s, sign)
                if vyx_num and today_fmt:
                    t = re.sub(
                        r'Вих\. № [^<]*від \d{2}\.\d{2}\.\d{4}',
                        'Вих. № '+vyx_num+' від '+today_fmt, t)
                if fmt_date:
                    t = re.sub(
                        r'\d{2}\.\d{2}\.\d{4} о \d{2}:\d{2}',
                        fmt_date+' о '+time_val, t)
                    p = fmt_date.split('.')
                    if len(p) == 3:
                        t = re.sub(
                            r'«\d+» \d+\.\d+р\. о \d{2}:\d{2} год\.',
                            '«'+p[0]+'» '+p[1]+'.'+p[2]+'р. о '+time_val+' год.', t)
                if be:
                    t = re.sub(r'<si><t>14/15[^<]*</t></si>', '<si><t>'+be+'</t></si>', t)
                data = t.encode('utf-8')

            # ── DATA sheet ───────────────────────────────────────────────
            elif item.filename == 'xl/worksheets/sheet1.xml':
                t = data.decode('utf-8')
                if fmt_date:
                    t = replace_cell(t, 'B4', fmt_date, '260')
                    t = replace_cell(t, 'G4', fmt_date, '191')
                if time_val:
                    t = replace_cell(t, 'G5', time_val, '191')
                # Расширяем колонки в DATA
                for col_n, w in [(5, 22.0), (6, 20.0), (7, 20.0), (8, 20.0)]:
                    t = set_col_width(t, col_n, w)
                data = t.encode('utf-8')

            # ── Листы заявок — расширяем колонки ─────────────────────────
            elif any(item.filename == name_to_file.get(n,'') for n in selected):
                t = data.decode('utf-8')
                # Расширяем колонки G(7) H(8) F(6) I(9) для подписей
                for col_n, w in [(5, 22.0), (6, 20.0), (7, 20.0), (8, 20.0), (9, 18.0)]:
                    t = set_col_width(t, col_n, w)
                data = t.encode('utf-8')

            # ── workbook.xml ─────────────────────────────────────────────
            elif item.filename == 'xl/workbook.xml':
                t = data.decode('utf-8')
                for name, _ in smap:
                    if name not in keep:
                        t = re.sub(r'<sheet[^>]+name="'+re.escape(name)+r'"[^/]*/\s*>', '', t)
                t = re.sub(r'<calcPr[^/]*/>', '<calcPr calcId="0" fullCalcOnLoad="1"/>', t)
                if '<calcPr' not in t:
                    t = t.replace('</workbook>', '<calcPr calcId="0" fullCalcOnLoad="1"/></workbook>')
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
