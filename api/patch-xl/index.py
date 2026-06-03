from http.server import BaseHTTPRequestHandler
import json, base64, zipfile, io, re


def esc(s):
    return str(s).replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')


def replace_cell_inline(xml, cell_ref, value, style):
    safe = esc(value)
    new_cell = '<c r="'+cell_ref+'" s="'+style+'" t="inlineStr"><is><t>'+safe+'</t></is></c>'
    r = re.sub(r'<c r="'+re.escape(cell_ref)+r'"[^/]*/>', new_cell, xml)
    if r != xml: return r
    r = re.sub(r'<c r="'+re.escape(cell_ref)+r'"[^>]*>.*?</c>', new_cell, xml, flags=re.DOTALL)
    if r != xml: return r
    row_num = re.search(r'\d+', cell_ref).group()
    def ins(m): return m.group(1)+m.group(2)+new_cell+m.group(3)
    return re.sub(r'(<row r="'+row_num+r'"[^>]*>)(.*?)(</row>)', ins, xml, flags=re.DOTALL)


def update_formula_v(xml, cell_ref, new_value):
    safe = esc(new_value)
    def repl(m): return m.group(1)+safe+m.group(2)
    return re.sub(
        r'(<c r="'+re.escape(cell_ref)+r'"[^>]*>.*?<v>)[^<]*(</v>)',
        repl, xml, flags=re.DOTALL)


def replace_ss_by_idx(ss_xml, ss_list, idx, new_value):
    """Заменяет конкретный <si> по индексу, используя позицию в файле."""
    if idx >= len(ss_list): return ss_xml
    old_si = ss_list[idx]
    safe = esc(new_value)
    new_si = '<si><t>'+safe+'</t></si>'
    # Заменяем только первое вхождение (они уникальны по позиции)
    return ss_xml.replace(old_si, new_si, 1)


def find_ss_indices(ss_list, xml_sheets):
    """
    Динамически находит ss индексы нужных значений сканируя листы.
    Возвращает словарь: тип → {idx: int, cell_refs: [...]}
    """
    def get_val(idx):
        try: return ''.join(re.findall(r'<t[^>]*>([^<]*)</t>', ss_list[idx]))
        except: return ''

    result = {
        'berth_idx': None,           # ss idx причала (из DATA строки Причал)
        'vyx_num_idxs': [],          # ss idx только-номер (список: разные колонки)
        'vyx_full_idxs': [],         # ss idx "Вих. № X/Y"
        'vyx_with_date_idxs': [],    # ss idx "Вих. № X/Y від ДД.ММ.ГГГГ"
        'eta_dt1_idxs': [],          # ss idx "ДД.ММ.ГГГГ о ЧЧ:ММ"
        'eta_dt2_idxs': [],          # ss idx "«ДД» ММ.ГГГГр. о ЧЧ:ММ год."
        'sign_name_idxs': [],        # ss idx только-имя подписи
        'sign_full_idxs': [],        # ss idx "Морський агент: Имя"
    }

    # Паттерны для определения типа
    def is_vyx_num(v):    return bool(re.match(r'^\d{2,3}/\d{1,2}$', v))
    def is_vyx_full(v):   return bool(re.match(r'^Вих\. № \d+/\d+$', v))
    def is_vyx_date(v):   return bool(re.match(r'^Вих\. № \d+/\d+ від \d', v))
    def is_eta_dt1(v):    return bool(re.match(r'^\d{2}\.\d{2}\.\d{4} о \d{2}:\d{2}$', v))
    def is_eta_dt2(v):    return bool(re.match(r'^«\d+» \d+\.\d{4}р\. о', v))
    def is_sign_name(v):  return bool(re.match(r'^[А-ЯІЇЄҐ][а-яіїєґ\']+ [А-ЯІЇЄҐ]{3,}$', v))
    def is_sign_full(v):  return bool(re.match(r'^Морський агент: [А-ЯІЇЄҐ]', v))

    # Сканируем все ss
    for i, si in enumerate(ss_list):
        v = ''.join(re.findall(r'<t[^>]*>([^<]*)</t>', si))
        if is_vyx_date(v)  and i not in result['vyx_with_date_idxs']:
            result['vyx_with_date_idxs'].append(i)
        elif is_vyx_full(v) and i not in result['vyx_full_idxs']:
            result['vyx_full_idxs'].append(i)
        elif is_vyx_num(v) and i not in result['vyx_num_idxs']:
            result['vyx_num_idxs'].append(i)
        if is_eta_dt1(v) and i not in result['eta_dt1_idxs']:
            result['eta_dt1_idxs'].append(i)
        if is_eta_dt2(v) and i not in result['eta_dt2_idxs']:
            result['eta_dt2_idxs'].append(i)
        if is_sign_name(v) and i not in result['sign_name_idxs']:
            result['sign_name_idxs'].append(i)
        if is_sign_full(v) and i not in result['sign_full_idxs']:
            result['sign_full_idxs'].append(i)

    # Причал: ищем в DATA строку с "Причал" в col A
    data_xml = xml_sheets.get('sheet1.xml', '')
    rows = re.findall(r'<row r="\d+"[^>]*>.*?</row>', data_xml, re.DOTALL)
    for row in rows:
        cells = re.findall(r'<c r="([^"]+)"[^>]*t="s"[^>]*><v>(\d+)</v></c>', row)
        for ref, idx_s in cells:
            col = re.match(r'[A-Z]+', ref).group()
            if col == 'A' and get_val(int(idx_s)) in ('Причал', 'причал', 'ПРИЧАЛ'):
                # Берём ss индексы из col B этой же строки
                for ref2, idx2_s in cells:
                    col2 = re.match(r'[A-Z]+', ref2).group()
                    if col2 == 'B':
                        result['berth_idx'] = int(idx2_s)
                break

    # Вых.№ только-номер: уточняем — нужен тот индекс, что стоит в C10 Шварт Адм
    # Или H10 DATA. Ищем в листах заявок ячейку рядом с "Вих. №"
    for sheet_xml in xml_sheets.values():
        cells_s = re.findall(r'<c r="([^"]+)"[^>]*t="s"[^>]*><v>(\d+)</v></c>', sheet_xml)
        for ref, idx_s in cells_s:
            v = get_val(int(idx_s))
            if v == 'Вих. №':
                # Следующая ячейка в той же строке = номер
                row_n = re.search(r'\d+', ref).group()
                row_match = re.search(r'<row r="'+row_n+r'"[^>]*>.*?</row>', sheet_xml, re.DOTALL)
                if row_match:
                    row_cells = re.findall(r'<c r="([^"]+)"[^>]*t="s"[^>]*><v>(\d+)</v></c>', row_match.group())
                    for ref2, idx2_s in row_cells:
                        v2 = get_val(int(idx2_s))
                        if is_vyx_num(v2):
                            idx2 = int(idx2_s)
                            if idx2 not in result['vyx_num_idxs']:
                                result['vyx_num_idxs'].insert(0, idx2)

    return result


def patch_xlsm(file_bytes, params):
    vyx_num   = params.get('vyxNum', '')
    fmt_date  = params.get('fmtDate', '')    # dd.mm.yyyy — дата заявки (ETA)
    time_val  = params.get('time', '17:00')
    berth     = params.get('berth', '').strip()
    sign      = params.get('sign', '').strip()
    today_fmt = params.get('todayFmt', '')   # dd.mm.yyyy — сегодня

    selected  = params.get('selectedSheets', [])

    orig     = zipfile.ZipFile(io.BytesIO(file_bytes))
    wb_xml   = orig.read('xl/workbook.xml').decode('utf-8')
    rels_xml = orig.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    rid2file = dict(re.findall(r'Id="([^"]+)"[^>]+Target="([^"]+)"', rels_xml))
    smap     = re.findall(r'name="([^"]+)"[^>]+r:id="([^"]+)"', wb_xml)

    keep = set(['DATA'] + selected)

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

    selected_files = set(
        name_to_file[n] for n in selected if n in name_to_file
    )

    # Загружаем все листы-заявки в память для сканирования
    sheet_xmls = {}
    for item in orig.infolist():
        if item.filename.startswith('xl/worksheets/sheet') and item.filename.endswith('.xml'):
            bname = item.filename.split('/')[-1]
            sheet_xmls[bname] = orig.read(item.filename).decode('utf-8')

    # ── Патч sharedStrings ───────────────────────────────────────────────
    ss_xml  = orig.read('xl/sharedStrings.xml').decode('utf-8')
    ss_list = re.findall(r'<si>.*?</si>', ss_xml, re.DOTALL)

    # Динамически находим индексы
    indices = find_ss_indices(ss_list, sheet_xmls)

    def get_val(idx):
        try: return ''.join(re.findall(r'<t[^>]*>([^<]*)</t>', ss_list[idx]))
        except: return ''

    # Подписи
    if sign:
        sign_name = sign.replace('Морський агент: ', '').strip()
        sign_full = 'Морський агент: ' + sign_name
        for idx in indices['sign_name_idxs']:
            ss_xml = replace_ss_by_idx(ss_xml, ss_list, idx, sign_name)
            ss_list = re.findall(r'<si>.*?</si>', ss_xml, re.DOTALL)
        for idx in indices['sign_full_idxs']:
            ss_xml = replace_ss_by_idx(ss_xml, ss_list, idx, sign_full)
            ss_list = re.findall(r'<si>.*?</si>', ss_xml, re.DOTALL)

    # Вих.№
    if vyx_num:
        # Только номер (Шварт Адм C10 и др.)
        for idx in indices['vyx_num_idxs']:
            ss_xml = replace_ss_by_idx(ss_xml, ss_list, idx, vyx_num)
            ss_list = re.findall(r'<si>.*?</si>', ss_xml, re.DOTALL)
        # "Вих. № X/Y"
        for idx in indices['vyx_full_idxs']:
            ss_xml = replace_ss_by_idx(ss_xml, ss_list, idx, 'Вих. № '+vyx_num)
            ss_list = re.findall(r'<si>.*?</si>', ss_xml, re.DOTALL)
        # "Вих. № X/Y від ДД.ММ.ГГГГ"
        if today_fmt:
            for idx in indices['vyx_with_date_idxs']:
                ss_xml = replace_ss_by_idx(ss_xml, ss_list, idx, 'Вих. № '+vyx_num+' від '+today_fmt)
                ss_list = re.findall(r'<si>.*?</si>', ss_xml, re.DOTALL)

    # ETA дата + время
    if fmt_date:
        # "ДД.ММ.ГГГГ о ЧЧ:ММ"
        for idx in indices['eta_dt1_idxs']:
            ss_xml = replace_ss_by_idx(ss_xml, ss_list, idx, fmt_date+' о '+time_val)
            ss_list = re.findall(r'<si>.*?</si>', ss_xml, re.DOTALL)
        # "«ДД» ММ.ГГГГр. о ЧЧ:ММ год."
        p = fmt_date.split('.')
        if len(p) == 3:
            lbl = '«'+p[0]+'» '+p[1]+'.'+p[2]+'р. о '+time_val+' год.'
            for idx in indices['eta_dt2_idxs']:
                ss_xml = replace_ss_by_idx(ss_xml, ss_list, idx, lbl)
                ss_list = re.findall(r'<si>.*?</si>', ss_xml, re.DOTALL)

    # Причал — только ss индекс из строки "Причал" DATA
    if berth and indices['berth_idx'] is not None:
        ss_xml = replace_ss_by_idx(ss_xml, ss_list, indices['berth_idx'], berth)
        ss_list = re.findall(r'<si>.*?</si>', ss_xml, re.DOTALL)

    ss_bytes = ss_xml.encode('utf-8')

    # ── Сборка zip ───────────────────────────────────────────────────────
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
                # DATA лист
                t = data.decode('utf-8')
                if fmt_date:
                    t = replace_cell_inline(t, 'B4', fmt_date, '260')
                    t = replace_cell_inline(t, 'G4', fmt_date, '191')
                if time_val:
                    t = replace_cell_inline(t, 'G5', time_val, '191')
                if today_fmt:
                    t = replace_cell_inline(t, 'G2', today_fmt, '188')
                if vyx_num:
                    t = replace_cell_inline(t, 'H2', vyx_num, '203')
                data = t.encode('utf-8')

            elif fn in selected_files:
                # Листы заявок — обновляем кешированные <v> формульных ячеек
                t = data.decode('utf-8')

                def patch_formula_v(xml, ref, transform):
                    """Читает <v> ячейки ref, применяет transform, записывает обратно."""
                    # Ищем ячейку с формулой (не self-closing)
                    m = re.search(
                        r'<c r="'+re.escape(ref)+r'"[^>]*>\s*<f[^>]*>.*?</f>\s*<v>([^<]*)</v>',
                        xml, re.DOTALL)
                    if not m:
                        return xml
                    old_v = m.group(1)
                    new_v = transform(old_v)
                    if new_v == old_v:
                        return xml
                    return update_formula_v(xml, ref, new_v)

                def repl_berth(v):
                    return re.sub(r'\d+/\d+(?:\s*;\s*\d+)*', berth, v) if berth else v

                def repl_eta(v):
                    return re.sub(
                        r'орієнтовний час прибуття (?!\d{2}\.\d{2})[^.]*\.',
                        'орієнтовний час прибуття '+fmt_date+' о '+time_val+'.',
                        v, count=1) if (fmt_date and time_val) else v

                def repl_b21_date(v):
                    if not (fmt_date and time_val): return v
                    return re.sub(
                        r'постановки:\s+\S*р\., о \S*',
                        'постановки:   '+fmt_date+'р., о '+time_val, v)

                # ── Дата today (шапка листа) — DATA!G2 ──────────────────
                if today_fmt:
                    # C11 (Шварт Адм прих), C9 (Шварт Адм отход, Букс порт отх), B9 (Букс порт прих)
                    for ref in ['C11', 'C9', 'B9']:
                        t = update_formula_v(t, ref, today_fmt)
                    # I31 (Букс порт прих низ), H26 (Букс порт отх низ)
                    for ref in ['I31', 'H26']:
                        t = update_formula_v(t, ref, today_fmt)

                # ── Причал + дата/время в формульных текстах ────────────
                # B17 = Шварт Адм прих: текст с причалом
                t = patch_formula_v(t, 'B17', repl_berth)
                # B14 = Шварт Адм отход: текст с причалом
                t = patch_formula_v(t, 'B14', repl_berth)
                # B21 = Шварт Адм прих (дата+время) и Шварт ТБТ прих/отход (причал+ETA)
                t = patch_formula_v(t, 'B21', lambda v: repl_eta(repl_berth(repl_b21_date(v))))
                # D27 = Букс порт прих: причал (DATA!E37)
                t = patch_formula_v(t, 'D27', repl_berth)
                # C23 = Букс порт отх: причал (DATA!E37)
                t = patch_formula_v(t, 'C23', repl_berth)
                # A19 = Букс порт прих: длинный текст с причалом
                t = patch_formula_v(t, 'A19', repl_berth)
                # A15 = Букс порт отх: длинный текст с причалом
                t = patch_formula_v(t, 'A15', repl_berth)
                # A18 = Букс бласт прих/отход: текст с причалом
                t = patch_formula_v(t, 'A18', repl_berth)
                # A20 = LINE BOAT: текст с причалом
                t = patch_formula_v(t, 'A20', repl_berth)
                # A26/A15 = Зняття: текст с причалом
                t = patch_formula_v(t, 'A26', repl_berth)
                t = patch_formula_v(t, 'A15', repl_berth)

                # ── Зняття: A6 / A7 ──────────────────────────────────────
                if vyx_num:
                    t = update_formula_v(t, 'A6', 'Вих. №'+vyx_num)
                if today_fmt:
                    t = update_formula_v(t, 'A7', 'Дата: '+today_fmt)

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
