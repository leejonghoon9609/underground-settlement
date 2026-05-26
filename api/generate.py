import json
import base64
from io import BytesIO
import openpyxl

RATES = {
    'realtime': {'labor': 6209, 'machine': 9437},
    'probe':    {'labor': 3104, 'machine': 4719},
}
RATIOS = {
    'indirectLabor': 0.10, 'accident': 0.0356, 'employment': 0.0101,
    'pension': 0.0475, 'health': 0.03595, 'safety': 0.0178,
    'elderly': 0.1314, 'genMgmt': 0.03, 'profit': 0.135,
    'contract': 0.749, 'vat': 0.10,
}

def calc_cost(exposed_km, probe_km, method):
    probe_m = probe_km * 1000
    r = RATES[method]
    dl = int(probe_m * r['labor'])
    me = int(probe_m * r['machine'])
    il = int(dl * RATIOS['indirectLabor'])
    lt = dl + il
    ac = int(lt * RATIOS['accident'])
    em = int(lt * RATIOS['employment'])
    pe = int(dl * RATIOS['pension'])
    he = int(dl * RATIOS['health'])
    sa = int(dl * RATIOS['safety'])
    el = int(he * RATIOS['elderly'])
    et = ac + em + pe + he + sa + el + me
    gm = int((lt + et) * RATIOS['genMgmt'])
    pr = int((lt + et + gm) * RATIOS['profit'])
    wc = lt + et + gm + pr
    ct = int((wc - sa) * RATIOS['contract']) + sa
    fi = (ct // 1000) * 1000
    vt = round(fi * RATIOS['vat'])
    return {
        'directLabor': dl, 'indirectLabor': il, 'laborTotal': lt,
        'machineExp': me, 'accident': ac, 'employment': em,
        'pension': pe, 'health': he, 'safety': sa, 'elderly': el,
        'expTotal': et, 'generalMgmt': gm, 'profit': pr,
        'workCost': wc, 'finalCost': fi, 'vat': vt, 'totalWithVat': fi + vt,
    }

COST_ROW_MAP = {
    10: 'directLabor', 15: 'indirectLabor', 16: 'laborTotal',
    19: 'accident', 20: 'employment', 21: 'pension',
    22: 'health', 23: 'safety', 24: 'elderly',
    27: 'machineExp', 28: 'expTotal', 29: 'generalMgmt',
    30: 'profit', 31: 'workCost', 32: 'finalCost',
    40: 'finalCost', 41: 'vat', 42: 'totalWithVat',
}

def handler(request):
    if request.method == 'OPTIONS':
        return Response('', 200, {
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'POST, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type',
        })

    if request.method != 'POST':
        return Response(json.dumps({'error': 'Method not allowed'}), 405, {
            'Access-Control-Allow-Origin': '*',
        })

    try:
        body = json.loads(request.body)
        projects = body['projects']
        template_b64 = body['template']

        template_bytes = base64.b64decode(template_b64)
        wb = openpyxl.load_workbook(BytesIO(template_bytes))

        ws_detail = wb['세부내역']
        COL = {
            'gubun': 3, 'region': 4, 'surveyName': 5, 'workCode': 6,
            'workName': 7, 'tangoType': 8, 'tangoKm': 10,
            'exposedKm': 11, 'probeKm': 12, 'finalCost': 15, 'remark': 16
        }
        for i, p in enumerate(projects):
            cost = calc_cost(p['exposedKm'], p['probeKm'], p['method'])
            row = 5 + i
            ws_detail.cell(row, COL['gubun']).value      = p.get('gubun', '')
            ws_detail.cell(row, COL['region']).value     = p.get('region', '')
            ws_detail.cell(row, COL['surveyName']).value = p.get('surveyName', '')
            ws_detail.cell(row, COL['workCode']).value   = p.get('workCode', '')
            ws_detail.cell(row, COL['workName']).value   = p.get('workName', '')
            ws_detail.cell(row, COL['tangoType']).value  = p.get('tangoType', '')
            ws_detail.cell(row, COL['tangoKm']).value    = p.get('tangoKm', 0)
            ws_detail.cell(row, COL['exposedKm']).value  = p.get('exposedKm', 0)
            ws_detail.cell(row, COL['probeKm']).value    = p.get('probeKm', 0)
            ws_detail.cell(row, COL['finalCost']).value  = cost['finalCost']
            ws_detail.cell(row, COL['remark']).value     = p.get('remark', '')

        ws_cost = wb['원가계산서']
        for i, p in enumerate(projects):
            cost = calc_cost(p['exposedKm'], p['probeKm'], p['method'])
            start_col = 13 + i * 3
            ws_cost.cell(5, start_col).value = f"{i+1}. {p.get('workCode', '')}"
            for row, key in COST_ROW_MAP.items():
                val = cost.get(key, 0)
                ws_cost.cell(row, start_col).value     = val
                ws_cost.cell(row, start_col + 1).value = 0
                ws_cost.cell(row, start_col + 2).value = val

        output = BytesIO()
        wb.save(output)
        output.seek(0)
        result_b64 = base64.b64encode(output.read()).decode('utf-8')

        return Response(json.dumps({'file': result_b64}), 200, {
            'Access-Control-Allow-Origin': '*',
            'Content-Type': 'application/json',
        })

    except Exception as e:
        return Response(json.dumps({'error': str(e)}), 500, {
            'Access-Control-Allow-Origin': '*',
            'Content-Type': 'application/json',
        })
