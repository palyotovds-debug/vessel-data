import json
import base64
import zipfile
import io
import re

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

    SIGNS = [
        'Олексій АЛЬОХІН', 'Олександр ЛИТВИНОВ',
        'Морський агент: Олексій АЛЬОХІН',
        'Морський агент: Олександр ЛИТВИНОВ'
    ]

    orig_zip = zipfile.ZipFile(io.BytesIO(file_bytes))

    # Mapping sheet name -> file
    wb_xml = orig_zip.read('xl/workbook.xml').decode('utf-8')
    rels_xml = orig_zip.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    rid_to_file = dict(re.findall(r'Id="([^"]+)"[^>]+Target="([^"]+)"', rels_xml))
    sheet_map = re.findall(r'name="([^"]+)"[^>]+r:id="([^"]+)"', wb_xml)

    keep_sheets = set(['DATA'] + selected_sheets)

    out_buf = io.BytesIO()
    with zipfile.ZipFile(out_buf, 'w', allowZip64=True) as new_zip:
        for item in orig_zip.infolist():
            data = orig_zip.read(item.filename)

            if item.filename == 'xl/sharedStrings.xml':
                text = data.decode('utf-8')
                if sign:
                    for s in SIGNS:
                        text = text.replace(s, sign)
                if vyx_num and today_fmt:
                    text = re.sub(
                        r'Вих\. № [^<]*від \d{2}\.\d{2}\.\d{4}',
                        'Вих. № ' + vyx_num + ' від ' + today_fmt,
                        text
                    )
                if fmt_date:
                    text = re.sub(
                        r'\d{2}\.\d{2}\.\d{4} о \d{2}:\d{2}',
                        fmt_date + ' о ' + time_val,
                        text
                    )
                    parts = fmt_date.split('.')
                    if len(parts) == 3:
                        day, mon, yr = parts[0], parts[1], parts[2]
                        text = re.sub(
                            r'«\d+» \d+\.\d+р\. о \d{2}:\d{2} год\.',
                            '«' + day + '» ' + mon + '.' + yr + 'р. о ' + time_val + ' год.',
                            text
                        )
                data = text.encode('utf-8')

            elif item.filename == 'xl/worksheets/sheet1.xml':
                text = data.decode('utf-8')
                if fmt_date:
                    text = re.sub(
                        r'<c r="B4"[^>]*t="s"[^>]*><v>[^<]*</v></c>',
                        '<c r="B4" s="260" t="inlineStr"><is><t>' + fmt_date + '</t></is></c>',
                        text
                    )
                if berth:
                    b = berth.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
                    text = re.sub(
                        r'<c r="B37"[^>]*t="s"[^>]*><v>[^<]*</v></c>',
                        '<c r="B37" s="220" t="inlineStr"><is><t>' + b + '</t></is></c>',
                        text
                    )
                data = text.encode('utf-8')

            elif item.filename == 'xl/workbook.xml':
                text = data.decode('utf-8')
                for name, rid in sheet_map:
                    if name not in keep_sheets:
                        text = re.sub(
                            r'<sheet[^>]+name="' + re.escape(name) + r'"[^/]*/\s*>',
                            '', text
                        )
                data = text.encode('utf-8')

            new_item = zipfile.ZipInfo(item.filename)
            new_item.compress_type = item.compress_type
            new_item.date_time = item.date_time
            new_zip.writestr(new_item, data)

    orig_zip.close()
    return out_buf.getvalue()


def handler(request):
    if request.method == 'OPTIONS':
        return Response('', status=200, headers={
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'POST, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type'
        })

    if request.method != 'POST':
        return Response(json.dumps({'ok': False, 'error': 'Method not allowed'}),
                       status=405, headers={'Content-Type': 'application/json'})

    try:
        payload = request.json
        file_b64 = payload.get('file', '')
        params = payload.get('params', {})
        file_bytes = base64.b64decode(file_b64)
        result_bytes = patch_xlsm(file_bytes, params)
        result_b64 = base64.b64encode(result_bytes).decode('utf-8')

        return Response(
            json.dumps({'ok': True, 'file': result_b64}),
            status=200,
            headers={
                'Content-Type': 'application/json',
                'Access-Control-Allow-Origin': '*'
            }
        )
    except Exception as e:
        return Response(
            json.dumps({'ok': False, 'error': str(e)}),
            status=500,
            headers={
                'Content-Type': 'application/json',
                'Access-Control-Allow-Origin': '*'
            }
        )
