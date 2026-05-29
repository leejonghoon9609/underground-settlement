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
    'pension': 0.0475, 'health': 0.03595,
    'elderly': 0.1314, 'genMgmt': 0.03, 'profit': 0.135,
    'contract': 0.749, 'vat': 0.10,
}

# 탱고 사업구분 텍스트 매핑 (세부내역 H열에 입력될 값 → 갑지 SUMIF 키와 일치해야 함)
TANGO_LABEL_MAP = {
    'A-2':  '[A-2. 인입관로] 인입 관로 공급',
    'B-3a': '[B-3. 기간선로] 신설/증설/보강',
    'B-3b': '[B-3. 기간선로] 신축국사 연계 간선선로',
    'C-2':  '[C-2. 프론트홀 선로(5G)] 용량증설(5G)',
    'E-4a': '[E-4. 지장이설] 원인자 공사',
    'E-4b': '[E-4. 지장이설] 지중 인프라 확보',
    'E-4c': '[E-4. 지장이설] 순수 지장 이설',
    'G-3':  '[G-3. 프론트홀 선로(4G)] 용량증설(4G)',
}

def calc_cost(exposed_km, probe_km, method, survey_name=''):
    probe_m = probe_km * 1000
    safety_rate = 0.0178 if '수도권지사' in (survey_name or '') else 0.0164
    r = RATES[method]
    dl = int(probe_m * r['labor'])
    me = int(probe_m * r['machine'])
    il = int(dl * RATIOS['indirectLabor'])
    lt = dl + il
    ac = int(lt * RATIOS['accident'])
    em = int(lt * RATIOS['employment'])
    pe = int(dl * RATIOS['pension'])
    he = int(dl * RATIOS['health'])
    sa = int(dl * safety_rate)
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

            # ── 세부내역 시트 값 주입 ────────────────────────────────
            # 데이터 시작 행: 5행, 최대 16건 (5~20행)
            # I열(측량거리)은 템플릿에 =K+L 수식이 있으므로 덮어쓰지 않음
            # H열(사업구분)에 갑지 SUMIF 키와 일치하는 텍스트를 입력해야 갑지에 값이 표시됨
            ws_detail = wb['세부내역']
            for i, p in enumerate(projects):
                if i >= 16:
                    break  # 최대 16건
                row = 5 + i
                tango = p.get('tangoType', '')
                tango_label = TANGO_LABEL_MAP.get(tango, tango)
                cost = calc_cost(
                    p.get('exposedKm', 0),
                    p.get('probeKm', 0),
                    p.get('method', 'probe'),
                    p.get('surveyName', '')
                )

                ws_detail.cell(row, 3).value  = p.get('division', '')      # C(3): 구분
                ws_detail.cell(row, 4).value  = p.get('region', '')        # D(4): 지역
                ws_detail.cell(row, 5).value  = p.get('surveyName', '')    # E(5): 공공측량명칭
                ws_detail.cell(row, 6).value  = p.get('workCode', '')      # F(6): 공사코드
                ws_detail.cell(row, 7).value  = p.get('workName', '')      # G(7): 공사명
                ws_detail.cell(row, 8).value  = tango_label                # H(8): 사업구분(Tango) ← SUMIF 키
                # I(9): 수식(=K+L) 유지 — 덮어쓰지 않음
                ws_detail.cell(row, 10).value = p.get('tangoKm', 0)       # J(10): Tango 굴착거리
                ws_detail.cell(row, 11).value = p.get('exposedKm', 0)     # K(11): 노출측량
                ws_detail.cell(row, 12).value = p.get('probeKm', 0)       # L(12): 탐사측량
                ws_detail.cell(row, 15).value = cost['finalCost']          # O(15): 원가계산서 공사비금액(낙찰가)
                ws_detail.cell(row, 16).value = p.get('remark', '')        # P(16): 비고

            # ── 공공측량 갑지는 세부내역 SUMIF 수식으로 자동 계산됨 ──
            # 별도 값 주입 불필요

            # ── 원가계산서 값 주입 ──────────────────────────────────
            ws_cost = wb['원가계산서']
            for i, p in enumerate(projects):
                cost = calc_cost(
                    p.get('exposedKm', 0),
                    p.get('probeKm', 0),
                    p.get('method', 'probe'),
                    p.get('surveyName', '')
                )
                start_col = 13 + i * 3  # M열(13)부터 3열씩
                ws_cost.cell(5, start_col).value = p.get('workCode', '')
                for row_num, key in COST_ROW_MAP.items():
                    val = cost.get(key, 0)
                    ws_cost.cell(row_num, start_col).value     = val   # Capex
                    ws_cost.cell(row_num, start_col + 1).value = 0     # Opex
                    ws_cost.cell(row_num, start_col + 2).value = val   # 계

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
