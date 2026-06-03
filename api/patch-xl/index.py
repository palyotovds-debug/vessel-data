from http.server import BaseHTTPRequestHandler
import json, base64, zipfile, io, re


def esc(s):
    return str(s).replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')


def replace_cell_inline(xml, cell_ref, value, style):
    """Заменяет ячейку на inlineStr. Если нет — вставляет в строку."""
    safe = esc(value)
    new_cell = '<c r="'+cell_ref+'" s="'+style+'" t="inlineStr"><is><t>'+safe+'</t></is></c>'
    replaced = re.sub(r'<c r="'+re.escape(cell_ref)+r'"[^/]*/>', new_cell, xml)
    if replaced != xml:
        return replaced
    replaced = re.sub(r'<c r="'+re.escape(cell_ref)+r'"[^>]*>.*?</c>', new_cell, xml, flags=re.DOTALL)
    if replaced != xml:
        return replaced
    row_num = re.search(r'\d+', cell_ref).group()
    def insert_cell(m):
        return m.group(1) + m.group(2) + new_cell + m.group(3)
    return re.sub(r'(<row r="'+row_num+r'"[^>]*>)(.*?)(</row>)', insert_cell, xml, flags=re.DOTALL)


def update_formula_v(xml, cell_ref, new_value):
    """Обновляет кешированное <v> в ячейке с формулой."""
    safe = esc(new_value)
    def repl(m):
        return m.group(1) + safe + m.group(2)
    return re.sub(
        r'(<c r="'+re.escape(cell_ref)+r'"[^>]*>.*?<v>)[^<]*(</v>)',
        repl, xml, flags=re.DOTALL)


def replace_ss_string(ss_xml, ss_list, idx, new_value):
    """Заменяет текст конкретного <si> в sharedStrings."""
    if idx >= len(ss_list):
        return ss_xml
    old_si = ss_list[idx]
    safe = esc(new_value)
    return ss_xml.replace(old_si, '<si><t>'+safe+'</t></si>', 1)


def set_col_width(xml, col_num, width):
    w, n = str(width), str(col_num)
    for pat in [
        r'(<col[^>]+min="'+n+r'"[^>]+max="'+n+r'"[^>]+)width="[^"]+"',
        r'(<col[^>]+max="'+n+r'"[^>]+min="'+n+r'"[^>]+)width="[^"]+"',
    ]:
        r = re.sub(pat, r'\1width="'+w+'"', xml)
        if r != xml:
            return r
    new_col = '<col min="'+n+'" max="'+n+'" width="'+w+'" customWidth="1"/>'
    if '<cols>' in xml:
        return xml.replace('</cols>', new_col+'</cols>')
    return xml.replace('<sheetData>', '<cols>'+new_col+'</cols><sheetData>')


def patch_xlsm(file_bytes, params):
    vyx_num   = params.get('vyxNum', '')
    fmt_date  = params.get('fmtDate', '')    # dd.mm.yyyy — дата заявки (ETA/прибытие)
    time_val  = params.get('time', '17:00')  # HH:MM
    berth     = params.get('berth', '')      # номер причала
    sign      = params.get('sign', '')       # подпись
    today_fmt = params.get('todayFmt', '')   # dd.mm.yyyy — сегодня (дата документа)
    selected  = params.get('selectedSheets', [])

    orig     = zipfile.ZipFile(io.BytesIO(file_bytes))
    wb_xml   = orig.read('xl/workbook.xml').decode('utf-8')
    rels_xml = orig.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    rid2file = dict(re.findall(r'Id="([^"]+)"[^>]+Target="([^"]+)"', rels_xml))
    smap     = re.findall(r'name="([^"]+)"[^>]+r:id="([^"]+)"', wb_xml)

    keep = set(['DATA'] + selected)
    be   = berth.strip() if berth else ''

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

    # ── Подготовка sharedStrings ─────────────────────────────────────────
    ss_xml  = orig.read('xl/sharedStrings.xml').decode('utf-8')
    ss_list = re.findall(r'<si>.*?</si>', ss_xml, re.DOTALL)

    # Подписи
    # ss[343] = "Олексій АЛЬОХІН"  — используется в Букс бласт G41, Шварт ТБТ H37, LINE BOAT I34, Шварт Адм F31
    # ss[517] = "Олександр ЛИТВИНОВ" — Зняття B31
    # ss[348] = "Морський агент: Олексій АЛЬОХІН" — Букс порт A28
    if sign:
        sign_name = sign.replace('Морський агент: ', '').strip()
        sign_full = 'Морський агент: ' + sign_name
        for idx in [343, 517]:
            ss_xml = replace_ss_string(ss_xml, ss_list, idx, sign_name)
            ss_list = re.findall(r'<si>.*?</si>', ss_xml, re.DOTALL)
        ss_xml = replace_ss_string(ss_xml, ss_list, 348, sign_full)
        ss_list = re.findall(r'<si>.*?</si>', ss_xml, re.DOTALL)

    # Вих.№
    # ss[20]  = только номер "334/15"             — Шварт Адм C10
    # ss[344] = "Вих. № 334/17"                   — Букс порт J9
    # ss[520] = "Вих. № 211/3 від 18.05.2026"     — Букс бласт A8
    if vyx_num:
        ss_xml = replace_ss_string(ss_xml, ss_list, 20, vyx_num)
        ss_list = re.findall(r'<si>.*?</si>', ss_xml, re.DOTALL)
        ss_xml = replace_ss_string(ss_xml, ss_list, 344, 'Вих. № '+vyx_num)
        ss_list = re.findall(r'<si>.*?</si>', ss_xml, re.DOTALL)
        if today_fmt:
            ss_xml = replace_ss_string(ss_xml, ss_list, 520, 'Вих. № '+vyx_num+' від '+today_fmt)
            ss_list = re.findall(r'<si>.*?</si>', ss_xml, re.DOTALL)

    # Дата+время заявки (ETA)
    # ss[521] = "19.05.2026 о 17:00"            — Букс бласт F34
    # ss[522] = "«23» 05.2026р. о 17:00 год."   — LINE BOAT A26
    if fmt_date:
        ss_xml = replace_ss_string(ss_xml, ss_list, 521, fmt_date+' о '+time_val)
        ss_list = re.findall(r'<si>.*?</si>', ss_xml, re.DOTALL)
        p = fmt_date.split('.')
        if len(p) == 3:
            lbl = '«'+p[0]+'» '+p[1]+'.'+p[2]+'р. о '+time_val+' год.'
            ss_xml = replace_ss_string(ss_xml, ss_list, 522, lbl)
            ss_list = re.findall(r'<si>.*?</si>', ss_xml, re.DOTALL)

    # Причал
    # ss[519] = "14/15; 16; 17" — используется в DATA B37/C37/D37/E37
    # Это основной источник для всех формул кроме LINE BOAT
    if be:
        ss_xml = replace_ss_string(ss_xml, ss_list, 519, be)
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

            # sharedStrings
            if fn == 'xl/sharedStrings.xml':
                data = ss_bytes

            # ── DATA sheet ───────────────────────────────────────────────
            elif fn == 'xl/worksheets/sheet1.xml':
                t = data.decode('utf-8')
                # B4 = дата заявки (отображение в DATA)
                if fmt_date:
                    t = replace_cell_inline(t, 'B4', fmt_date, '260')
                    # G4 = дата заявки (тянут формулы: Шварт Адм B21, Шварт ТБТ E16)
                    t = replace_cell_inline(t, 'G4', fmt_date, '191')
                # G5 = время (тянут формулы: Шварт Адм B21, Шварт ТБТ E16)
                if time_val:
                    t = replace_cell_inline(t, 'G5', time_val, '191')
                # G2 = дата today — ВАЖНО: НЕ вставляем как inlineStr, т.к. формулы
                # DATA!G2 (Шварт Адм C11, Букс порт B9) ожидают текст.
                # Вставляем как inlineStr с today_fmt
                if today_fmt:
                    t = replace_cell_inline(t, 'G2', today_fmt, '188')
                # H2 = Вих.№ (на случай если где-то ссылается)
                if vyx_num:
                    t = replace_cell_inline(t, 'H2', vyx_num, '203')
                data = t.encode('utf-8')

            # ── Шварт Адм прих (sheet14) ────────────────────────────────
            elif fn == 'xl/worksheets/sheet14.xml' and fn in selected_files:
                t = data.decode('utf-8')
                # C11 = DATA!G2 = дата today
                if today_fmt:
                    t = update_formula_v(t, 'C11', today_fmt)
                # B21 = "Очікувальний час постановки: {date}р., о {time}"
                # формула использует DATA!G4 и DATA!G5
                if fmt_date and time_val:
                    new_v21 = ('Очікувальний час постановки:   '
                               +fmt_date+'р., о '+time_val
                               +'                                                                        Характеристика судна:')
                    t = update_formula_v(t, 'B21', new_v21)
                # G12 = текст с причалом (формула DATA!E37)
                if be:
                    # обновляем <v> — формула подтянет при открытии
                    old_v = re.search(r'<c r="G12"[^>]*>.*?<v>([^<]*)</v>', t, re.DOTALL)
                    if old_v:
                        new_text = old_v.group(1).replace('14/15; 16; 17', be).replace('14/15', be)
                        t = update_formula_v(t, 'G12', new_text)
                data = t.encode('utf-8')

            # ── Шварт Адм отход (sheet15) ───────────────────────────────
            elif fn == 'xl/worksheets/sheet15.xml' and fn in selected_files:
                t = data.decode('utf-8')
                if today_fmt:
                    t = update_formula_v(t, 'C9', today_fmt)
                data = t.encode('utf-8')

            # ── Букс порт прих (sheet16) ─────────────────────────────────
            elif fn == 'xl/worksheets/sheet16.xml' and fn in selected_files:
                t = data.decode('utf-8')
                # B9 = DATA!G2 = today
                if today_fmt:
                    t = update_formula_v(t, 'B9', today_fmt)
                # D27 = DATA!E37 = причал (t="str")
                if be:
                    t = update_formula_v(t, 'D27', be)
                # B27 = DATA!E37 (ss ячейка, ссылается на DATA)
                # I31 = DATA!H2 = today дата внизу
                if today_fmt:
                    t = update_formula_v(t, 'I31', today_fmt)
                # B13 — длинный текст с причалом
                if be:
                    old_v = re.search(r'<c r="B13"[^>]*>.*?<v>([^<]*)</v>', t, re.DOTALL)
                    if old_v:
                        new_text = old_v.group(1).replace('14/15; 16; 17', be).replace('14/15', be)
                        t = update_formula_v(t, 'B13', new_text)
                data = t.encode('utf-8')

            # ── Букс порт отх (sheet17) ──────────────────────────────────
            elif fn == 'xl/worksheets/sheet17.xml' and fn in selected_files:
                t = data.decode('utf-8')
                if today_fmt:
                    t = update_formula_v(t, 'C9', today_fmt)
                if be:
                    t = update_formula_v(t, 'B23', be)
                    old_v = re.search(r'<c r="B13"[^>]*>.*?<v>([^<]*)</v>', t, re.DOTALL)
                    if old_v:
                        new_text = old_v.group(1).replace('14/15; 16; 17', be).replace('14/15', be)
                        t = update_formula_v(t, 'B13', new_text)
                data = t.encode('utf-8')

            # ── Букс бласт прих (sheet18) ───────────────────────────────
            elif fn == 'xl/worksheets/sheet18.xml' and fn in selected_files:
                t = data.decode('utf-8')
                # Причал в тексте D14/A18
                if be:
                    for ref in ['D14', 'A18']:
                        old_v = re.search(r'<c r="'+ref+r'"[^>]*>.*?<v>([^<]*)</v>', t, re.DOTALL)
                        if old_v:
                            new_text = old_v.group(1).replace('14/15; 16; 17', be).replace('14/15', be)
                            t = update_formula_v(t, ref, new_text)
                # Подпись G41 — расширяем мёрж вправо добавив colspan в ячейки
                # Вместо расширения колонки — добавим ещё ячеек к мёржу через изменение mergeCell
                t = re.sub(
                    r'(<mergeCell ref="G41:)([^"]+)(")',
                    r'\g<1>I41\3', t)
                data = t.encode('utf-8')

            # ── Букс бласт отход (sheet19) ──────────────────────────────
            elif fn == 'xl/worksheets/sheet19.xml' and fn in selected_files:
                t = data.decode('utf-8')
                if be:
                    for ref in ['D14', 'A18']:
                        old_v = re.search(r'<c r="'+ref+r'"[^>]*>.*?<v>([^<]*)</v>', t, re.DOTALL)
                        if old_v:
                            new_text = old_v.group(1).replace('14/15; 16; 17', be).replace('14/15', be)
                            t = update_formula_v(t, ref, new_text)
                t = re.sub(r'(<mergeCell ref="G41:)([^"]+)(")', r'\g<1>I41\3', t)
                data = t.encode('utf-8')

            # ── Шварт ТБТ прих (sheet20) ─────────────────────────────────
            elif fn == 'xl/worksheets/sheet20.xml' and fn in selected_files:
                t = data.decode('utf-8')
                # E16 = большой текст с причалом и датой
                if be or fmt_date or time_val:
                    old_v = re.search(r'<c r="E16"[^>]*>.*?<v>([^<]*)</v>', t, re.DOTALL)
                    if old_v:
                        new_text = old_v.group(1)
                        if be:
                            new_text = new_text.replace('14/15; 16; 17', be).replace('14/15', be)
                        if fmt_date:
                            # В тексте есть "орієнтовний час прибуття ." — вставляем дату и время
                            new_text = re.sub(
                                r'орієнтовний час прибуття [^.]*\.',
                                'орієнтовний час прибуття '+fmt_date+' о '+time_val+'.',
                                new_text)
                        t = update_formula_v(t, 'E16', new_text)
                # B21 = текст с причалом
                if be:
                    old_v = re.search(r'<c r="B21"[^>]*>.*?<v>([^<]*)</v>', t, re.DOTALL)
                    if old_v:
                        new_text = old_v.group(1).replace('14/15; 16; 17', be).replace('14/15', be)
                        t = update_formula_v(t, 'B21', new_text)
                data = t.encode('utf-8')

            # ── Шварт ТБТ отход (sheet21) ────────────────────────────────
            elif fn == 'xl/worksheets/sheet21.xml' and fn in selected_files:
                t = data.decode('utf-8')
                if be or fmt_date or time_val:
                    for ref in ['E16', 'B21']:
                        old_v = re.search(r'<c r="'+ref+r'"[^>]*>.*?<v>([^<]*)</v>', t, re.DOTALL)
                        if old_v:
                            new_text = old_v.group(1)
                            if be:
                                new_text = new_text.replace('14/15; 16; 17', be).replace('14/15', be)
                            if fmt_date and ref == 'E16':
                                new_text = re.sub(
                                    r'орієнтовний час прибуття [^.]*\.',
                                    'орієнтовний час прибуття '+fmt_date+' о '+time_val+'.',
                                    new_text)
                            t = update_formula_v(t, ref, new_text)
                data = t.encode('utf-8')

            # ── LINE BOAT (sheet25) ──────────────────────────────────────
            elif fn == 'xl/worksheets/sheet25.xml' and fn in selected_files:
                t = data.decode('utf-8')
                # H15 — большой текст с причалом (жёстко "14/15" в формуле)
                if be:
                    old_v = re.search(r'<c r="H15"[^>]*>.*?<v>([^<]*)</v>', t, re.DOTALL)
                    if old_v:
                        new_text = old_v.group(1).replace('14/15', be)
                        t = update_formula_v(t, 'H15', new_text)
                    # Также A20 (основной текст)
                    old_v2 = re.search(r'<c r="A20"[^>]*>.*?<v>([^<]*)</v>', t, re.DOTALL)
                    if old_v2:
                        new_text2 = old_v2.group(1).replace('14/15', be)
                        t = update_formula_v(t, 'A20', new_text2)
                data = t.encode('utf-8')

            # ── Зняття порт (sheet28) ────────────────────────────────────
            elif fn == 'xl/worksheets/sheet28.xml' and fn in selected_files:
                t = data.decode('utf-8')
                # A6: формула WEEKNUM → <v>
                if vyx_num:
                    t = update_formula_v(t, 'A6', 'Вих. №'+vyx_num)
                # A7: формула TODAY() → <v>
                if today_fmt:
                    t = update_formula_v(t, 'A7', 'Дата: '+today_fmt)
                # A26: CONCATENATE с причалом
                if be:
                    old_v = re.search(r'<c r="A26"[^>]*>.*?<v>([^<]*)</v>', t, re.DOTALL)
                    if old_v:
                        new_text = old_v.group(1).replace('14/15; 16; 17', be).replace('14/15', be)
                        t = update_formula_v(t, 'A26', new_text)
                data = t.encode('utf-8')

            # ── Зняття митн (sheet29) ────────────────────────────────────
            elif fn == 'xl/worksheets/sheet29.xml' and fn in selected_files:
                t = data.decode('utf-8')
                if vyx_num:
                    t = update_formula_v(t, 'A6', 'Вих. №'+vyx_num)
                if today_fmt:
                    t = update_formula_v(t, 'A7', 'Дата: '+today_fmt)
                if be:
                    old_v = re.search(r'<c r="A26"[^>]*>.*?<v>([^<]*)</v>', t, re.DOTALL)
                    if old_v:
                        new_text = old_v.group(1).replace('14/15; 16; 17', be).replace('14/15', be)
                        t = update_formula_v(t, 'A26', new_text)
                data = t.encode('utf-8')

            # ── workbook.xml ─────────────────────────────────────────────
            elif fn == 'xl/workbook.xml':
                t = data.decode('utf-8')
                for name, _ in smap:
                    if name not in keep:
                        t = re.sub(
                            r'<sheet\b[^>]*\bname="'+re.escape(name)+r'"[^/]*/\s*>',
                            '', t)
                data = t.encode('utf-8')

            # ── workbook.xml.rels ────────────────────────────────────────
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
