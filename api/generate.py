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

# 세부내역 시트: tangoType 텍스트 → 고정 행 번호
TANGO_ROW_MAP = {
    '[A-2. 인입관로] 인입 관로 공급': 4,
    '[B-3. 기간선로] 신설/증설/보강': 5,
    '[B-3. 기간선로] 신축국사 연계 간선선로': 6,
    '[C-2. 프론트홀 선로(5G)] 용량증설(5G)': 7,
    '[E-4. 지장이설] 원인자 공사': 8,
    '[E-4. 지장이설] 지중 인프라 확보': 9,
    '[E-4. 지장이설] 순수 지장 이설': 10,
    '[G-3. 프론트홀 선로(4G)] 용량증설(4G)': 11,
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

            wb = openpyxl.load_workbook(TEMPLATE_PATH)

            # 세부내역 값 주입 (tangoType 기준 고정 행)
            ws_detail = wb['세부내역']
            for p in projects:
                tango = p.get('tangoType', '')
                row = TANGO_ROW_MAP.get(tango)
                if row is None:
                    continue
                cost = calc_cost(p['exposedKm'], p['probeKm'], p['method'])
                survey_km = round(p.get('exposedKm', 0) + p.get('probeKm', 0), 3)
                ws_detail.cell(row, 2).value = survey_km
                ws_detail.cell(row, 3).value = p.get('tangoKm', 0)
                ws_detail.cell(row, 4).value = p.get('exposedKm', 0)
                ws_detail.cell(row, 5).value = p.get('probeKm', 0)
                ws_detail.cell(row, 7).value = cost['finalCost']
                ws_detail.cell(row, 9).value = p.get('remark', '')

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
