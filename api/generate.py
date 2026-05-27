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
    wc_floor = (wc // 1000) * 1000  # 공사비 천원 절사
    ct = int((wc_floor - sa) * RATIOS['contract']) + sa
    fi = ct  # 낙찰가 절사 없음
    vt = round(fi * RATIOS['vat'])
    return {
        'directLabor': dl, 'indirectLabor': il, 'laborTotal': lt,
        'machineExp': me, 'accident': ac, 'employment': em,
        'pension': pe, 'health': he, 'safety': sa, 'elderly': el,
        'expTotal': et, 'generalMgmt': gm, 'profit': pr,
        'workCost': wc_floor, 'finalCost': fi, 'vat': vt, 'totalWithVat': fi + vt,
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

            wb = openpyxl.load_workbook(TEMPLATE_PATH)

            # 세부내역 값 주입 (5행부터 순서대로)
            ws_detail = wb['세부내역']
            for i, p in enumerate(projects):
                cost = calc_cost(p['exposedKm'], p['probeKm'], p['method'])
                row = 5 + i
                ws_detail.cell(row, 3).value  = p.get('gubun', '')
                ws_detail.cell(row, 4).value  = p.get('region', '')
                ws_detail.cell(row, 5).value  = p.get('surveyName', '')
                ws_detail.cell(row, 6).value  = p.get('workCode', '')
                ws_detail.cell(row, 7).value  = p.get('workName', '')
                ws_detail.cell(row, 8).value  = p.get('tangoType', '')
                ws_detail.cell(row, 10).value = p.get('tangoKm', 0)
                ws_detail.cell(row, 11).value = p.get('exposedKm', 0)
                ws_detail.cell(row, 12).value = p.get('probeKm', 0)
                ws_detail.cell(row, 15).value = cost['finalCost']
                ws_detail.cell(row, 16).value = p.get('remark', '')

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
            import traceback
            err_msg = traceback.format_exc()
            resp = json.dumps({'error': str(e), 'traceback': err_msg}).encode('utf-8')
            self.send_response(500)
            self.send_cors_headers()
            self.send_header('Content-Type', 'application/json')
            self.send_header('Content-Length', str(len(resp)))
            self.end_headers()
            self.wfile.write(resp)
