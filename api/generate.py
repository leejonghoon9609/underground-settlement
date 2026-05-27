from http.server import BaseHTTPRequestHandler
import json
import base64
import os
from io import BytesIO
import openpyxl

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'template.xlsx')

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

class handler(BaseHTTPRequestHandler):

    def send_cors_headers(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type, Accept')

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_cors_headers()
        self.send_header('Content-Length', '0')
        self.end_headers()

    def do_POST(self):
        try:
            length = int(self.headers.get('Content-Length', 0))
            body = json.loads(self.rfile.read(length))
            projects = body['projects']

            # 로컬 template.xlsx 직접 읽기
            with open(TEMPLATE_PATH, 'rb') as f:
                template_bytes = f.read()

            wb = openpyxl.load_workbook(BytesIO(template_bytes))

            # 세부내역 값 주입
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

            # 원가계산서 값 주입
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

            resp = json.dumps({'file': result_b64}).encode('utf-8')
            self.send_response(200)
            self.send_cors_headers()
            self.send_header('Content-Type', 'application/json')
            self.send_header('Content-Length', str(len(resp)))
            self.end_headers()
            self.wfile.write(resp)

        except Exception as e:
            resp = json.dumps({'error': str(e)}).encode('utf-8')
            self.send_response(500)
            self.send_cors_headers()
            self.send_header('Content-Type', 'application/json')
            self.send_header('Content-Length', str(len(resp)))
            self.end_headers()
            self.wfile.write(resp)
