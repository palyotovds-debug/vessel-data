from http.server import BaseHTTPRequestHandler
import json, base64, zipfile, io, re


def esc(s):
    return str(s).replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')


def replace_cell_inline(xml, cell_ref, value, style):
    """Заменяет ячейку на inlineStr. Если не найдена — вставляет в строку."""
    safe = esc(value)
    new_cell = '<c r="'+cell_ref+'" s="'+style+'" t="inlineStr"><is><t>'+safe+'</t></is></c>'
    replaced = re.sub(r'<c r="'+re.escape(cell_ref)+r'"[^/]*/>', new_cell, xml)
    if replaced != xml:
        return replaced
    replaced = re.sub(r'<c r="'+re.escape(cell_ref)+r'"[^>]*>.*?</c>', new_cell, xml, flags=re.DOTALL)
    if replaced != xml:
        return replaced
    row_num = re.search(r'\d+', cell_ref).group()
    row_pat = r'(<row r="'+row_num+r'"[^>]*>)(.*?)(</row>)'
    def insert_cell(m):
        return m.group(1) + m.group(2) + new_cell + m.group(3)
    return re.sub(row_pat, insert_cell, xml, flags=re.DOTALL)


def update_formula_v(xml, cell_ref, new_value):
    """Обновляет <v> в ячейке с формулой (любой тип), сохраняя формулу."""
    safe = esc(new_value)
    # Заменяем <v>...</v> внутри ячейки с формулой
    pattern = r'(<c r="'+re.escape(cell_ref)+r'"[^>]*>(?:(?!<v>).)*?<f[^>]*>[^<]*</f>[^<]*)<v>[^<]*</v>'
    replaced = re.sub(pattern, r'\g<1><v>'+safe+'</v>', xml, flags=re.DOTALL)
    if replaced != xml:
        return replaced
    # Альтернативный порядок f/v
    pattern2 = r'(<c r="'+re.escape(cell_ref)+r'"[^>]*>.*?)<v>[^<]*</v>(.*?</c>)'
    return re.sub(pattern2, r'<v>'+safe+'</v>'.replace('', '\1').replace('', '\2'), xml, flags=re.DOTALL)


def replace_formula_str_v(xml, cell_ref, new_value):
    """Обновляет <v> в ячейке t="str" с формулой."""
    safe = esc(new_value)
    def repl(m):
        return m.group(1) + safe + m.group(2)
    return re.sub(
        r'(<c r="'+re.escape(cell_ref)+r'"[^>]*>.*?<v>)[^<]*(</v>)',
        repl, xml, flags=re.DOTALL)


def replace_ss_string(ss_xml, ss_list, idx, new_value):
    """Заменяет текст в конкретном <si> sharedStrings по индексу."""
    if idx >= len(ss_list):
        return ss_xml
    old_si = ss_list[idx]
    safe = esc(new_value)
    new_si = '<si><t>'+safe+'</t></si>'
    return ss_xml.replace(old_si, new_si, 1)


def set_col_width(xml, col_num, width):
    """Устанавливает ширину конкретной колонки."""
    w = str(width)
    n = str(col_num)
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
    new_col = '<col min="'+n+'" max="'+n+'" width="'+w+'" customWidth="1"/>'
    if '<cols>' in xml:
        return xml.replace('</cols>', new_col+'</cols>')
    return xml.replace('<sheetData>', '<cols>'+new_col+'</cols><sheetData>')


def patch_xlsm(file_bytes, params):
    vyx_num   = params.get('vyxNum', '')
    fmt_date  = params.get('fmtDate', '')   # dd.mm.yyyy — дата заявки
    time_val  = params.get('time', '17:00')
    berth     = params.get('berth', '')
    sign      = params.get('sign', '')
    today_fmt = params.get('todayFmt', '')  # dd.mm.yyyy — сегодня
    selected  = params.get('selectedSheets', [])

    orig     = zipfile.ZipFile(io.BytesIO(file_bytes))
    wb_xml   = orig.read('xl/workbook.xml').decode('utf-8')
    rels_xml = orig.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    rid2file = dict(re.findall(r'Id="([^"]+)"[^>]+Target="([^"]+)"', rels_xml))
    smap     = re.findall(r'name="([^"]+)"[^>]+r:id="([^"]+)"', wb_xml)

    keep = set(['DATA'] + selected)
    be   = esc(berth) if berth else ''

    name_to_file = {}
    for name, rid in smap:
        f = rid2file.get(rid, '')
        if f:
            name_to_file[name] = ('xl/'+f) if not f.startswith('xl/') else f

    excluded_files = set()
    for name, rid in smap:
        if name not in keep:
            f = rid2file.get(rid, '')
            if f:
                excluded_files.add(('xl/'+f) if not f.startswith('xl/') else f)

    selected_files = set()
    for name in selected:
        f = name_to_file.get(name, '')
        if f:
            selected_files.add(f)

    # ── Патч sharedStrings ───────────────────────────────────────────────
    ss_xml  = orig.read('xl/sharedStrings.xml').decode('utf-8')
    ss_list = re.findall(r'<si>.*?</si>', ss_xml, re.DOTALL)

    # Подписи
    # ss[343] = "Олексій АЛЬОХІН" (Букс бласт G41, Шварт ТБТ H37, LINE BOAT I34, Шварт Адм F31)
    # ss[517] = "Олександр ЛИТВИНОВ" (Зняття B31)
    # ss[348] = "Морський агент: Олексій АЛЬОХІН" (Букс порт A28)
    if sign:
        sign_name = sign.replace('Морський агент: ', '').strip()
        sign_full = 'Морський агент: ' + sign_name
        for idx in [343, 517]:
            ss_xml = replace_ss_string(ss_xml, ss_list, idx, sign_name)
            ss_list = re.findall(r'<si>.*?</si>', ss_xml, re.DOTALL)
        for idx in [348]:
            ss_xml = replace_ss_string(ss_xml, ss_list, idx, sign_full)
            ss_list = re.findall(r'<si>.*?</si>', ss_xml, re.DOTALL)

    # Вих.№
    # ss[20]  = только номер "334/15"              → Шварт Адм C10
    # ss[344] = "Вих. № 334/17"                    → Букс порт J9
    # ss[520] = "Вих. № 211/3 від 18.05.2026"      → Букс бласт A8
    if vyx_num:
        ss_xml = replace_ss_string(ss_xml, ss_list, 20, vyx_num)
        ss_list = re.findall(r'<si>.*?</si>', ss_xml, re.DOTALL)
        ss_xml = replace_ss_string(ss_xml, ss_list, 344, 'Вих. № '+vyx_num)
        ss_list = re.findall(r'<si>.*?</si>', ss_xml, re.DOTALL)
        if today_fmt:
            ss_xml = replace_ss_string(ss_xml, ss_list, 520, 'Вих. № '+vyx_num+' від '+today_fmt)
            ss_list = re.findall(r'<si>.*?</si>', ss_xml, re.DOTALL)

    # Дата+время в текстах (Шварт ТБТ, LINE BOAT и др.)
    if fmt_date:
        ss_xml = re.sub(
            r'\d{2}\.\d{2}\.\d{4} о \d{2}:\d{2}',
            fmt_date+' о '+time_val, ss_xml)
        p = fmt_date.split('.')
        if len(p) == 3:
            ss_xml = re.sub(
                r'«\d+» \d+\.\d+р\. о \d{2}:\d{2} год\.',
                '«'+p[0]+'» '+p[1]+'.'+p[2]+'р. о '+time_val+' год.', ss_xml)

    # Причал в текстах
    if be:
        ss_xml = re.sub(r'<si><t>14/15[^<]*</t></si>', '<si><t>'+be+'</t></si>', ss_xml)

    ss_bytes = ss_xml.encode('utf-8')

    # ── Сборка нового zip ────────────────────────────────────────────────
    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w', allowZip64=True) as zout:
        for item in orig.infolist():
            fn = item.filename

            if fn in excluded_files:
                continue

            data = orig.read(fn)

            if fn == 'xl/sharedStrings.xml':
                data = ss_bytes

            elif fn == 'xl/worksheets/sheet1.xml':
                # DATA лист — вставляем все значения
                t = data.decode('utf-8')
                if fmt_date:
                    t = replace_cell_inline(t, 'B4', fmt_date, '260')
                    t = replace_cell_inline(t, 'G4', fmt_date, '191')
                    # G2 = дата today (тянут формулы листов: DATA!G2)
                    t = replace_cell_inline(t, 'G2', today_fmt if today_fmt else fmt_date, '188')
                if vyx_num:
                    # H2 = Вих.№ (тянут формулы: DATA!H2)
                    t = replace_cell_inline(t, 'H2', vyx_num, '203')
                if time_val:
                    t = replace_cell_inline(t, 'G5', time_val, '191')
                data = t.encode('utf-8')

            elif fn == 'xl/worksheets/sheet14.xml' and fn in selected_files:
                # Шварт Адм прих: C11 = DATA!G2 (дата today)
                t = data.decode('utf-8')
                if today_fmt:
                    t = update_formula_v(t, 'C11', today_fmt)
                data = t.encode('utf-8')

            elif fn == 'xl/worksheets/sheet15.xml' and fn in selected_files:
                # Шварт Адм отход: C9 = DATA!G2
                t = data.decode('utf-8')
                if today_fmt:
                    t = update_formula_v(t, 'C9', today_fmt)
                data = t.encode('utf-8')

            elif fn == 'xl/worksheets/sheet16.xml' and fn in selected_files:
                # Букс порт прих: B9 = DATA!G2 (дата today), I31 = DATA!H2
                t = data.decode('utf-8')
                if today_fmt:
                    t = update_formula_v(t, 'B9', today_fmt)
                    t = update_formula_v(t, 'I31', today_fmt)
                data = t.encode('utf-8')

            elif fn == 'xl/worksheets/sheet17.xml' and fn in selected_files:
                # Букс порт отх: C9 = DATA!G2
                t = data.decode('utf-8')
                if today_fmt:
                    t = update_formula_v(t, 'C9', today_fmt)
                data = t.encode('utf-8')

            elif fn in ('xl/worksheets/sheet18.xml', 'xl/worksheets/sheet19.xml') and fn in selected_files:
                # Букс бласт: расширяем G(7) для подписи
                t = data.decode('utf-8')
                t = set_col_width(t, 7, 28.0)
                data = t.encode('utf-8')

            elif fn in ('xl/worksheets/sheet28.xml', 'xl/worksheets/sheet29.xml') and fn in selected_files:
                # Зняття порт/митн: A6 = формула WEEKNUM, A7 = формула TODAY()
                t = data.decode('utf-8')
                if vyx_num:
                    t = replace_formula_str_v(t, 'A6', 'Вих. №'+vyx_num)
                if today_fmt:
                    t = replace_formula_str_v(t, 'A7', 'Дата: '+today_fmt)
                data = t.encode('utf-8')

            elif fn == 'xl/workbook.xml':
                t = data.decode('utf-8')
                for name, _ in smap:
                    if name not in keep:
                        t = re.sub(
                            r'<sheet\b[^>]*\bname="'+re.escape(name)+r'"[^/]*/\s*>',
                            '', t)
                data = t.encode('utf-8')

            elif fn == 'xl/_rels/workbook.xml.rels':
                t = data.decode('utf-8')
                for name, rid in smap:
                    if name not in keep:
                        t = re.sub(
                            r'<Relationship\b[^>]*\bId="'+re.escape(rid)+r'"[^/]*/\s*>',
                            '', t)
                data = t.encode('utf-8')

            ni = zipfile.ZipInfo(fn)
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
            import traceback
            resp = json.dumps({'ok': False, 'error': str(e), 'trace': traceback.format_exc()}).encode()
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
