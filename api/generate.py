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

def calc_cost(exposed_km, probe_km, method, survey_name=''):
    probe_m = probe_km * 1000
    safety_rate = 0.0178 if '수도권지사' in (survey_name or '') else 0.0164
    r = RATES[method]
    dl = int(probe_m * r['labor'])
    me = int(probe_m * r['machine'])
    il = int(dl * RATIOS['indirectLabor'])
    lt = dl + il                                      # 노무비소계(16행)
    ac = int(lt * RATIOS['accident'])
    em = int(lt * RATIOS['employment'])
    pe = int(dl * RATIOS['pension'])
    he = int(dl * RATIOS['health'])
    sa = int(dl * safety_rate)
    el = int(he * RATIOS['elderly'])
    et = ac + em + pe + he + sa + el + me             # 경비소계(28행)
    gm = int((lt + et) * RATIOS['genMgmt'])
    pr = int((lt + et + gm) * RATIOS['profit'])
    wc = lt + et + gm + pr                            # 공사비(31행)
    ct = int((wc - sa) * RATIOS['contract']) + sa     # 낙찰가(32행): 절사 없음
    fi = (ct // 1000) * 1000                          # 공사비합계(40행): 천원 절사
    vt = round(fi * RATIOS['vat'])
    return {
        'directLabor': dl,      # 10행
        'indirectLabor': il,    # 15행
        'laborTotal': lt,       # 16행 노무비소계
        'accident': ac,         # 19행
        'employment': em,       # 20행
        'pension': pe,          # 21행
        'health': he,           # 22행
        'safety': sa,           # 23행
        'elderly': el,          # 24행
        'machineExp': me,       # 27행
        'expTotal': et,         # 28행 경비소계
        'generalMgmt': gm,      # 29행
        'profit': pr,           # 30행
        'workCost': wc,         # 31행 공사비
        'contractCost': ct,     # 32행 낙찰가 (절사 없음)
        'finalCost': fi,        # 40행 공사비합계 (천원 절사)
        'vat': vt,              # 41행 부가세
        'totalWithVat': fi + vt, # 42행 총계
        'safetyRate': safety_rate,
    }

# 샘플 기준 행번호 → cost 키 매핑 (수식 행 41,42 제외)
COST_ROW_MAP = {
    10: 'directLabor',
    15: 'indirectLabor',
    16: 'laborTotal',
    19: 'accident',
    20: 'employment',
    21: 'pension',
    22: 'health',
    23: 'safety',
    24: 'elderly',
    27: 'machineExp',
    28: 'expTotal',
    29: 'generalMgmt',
    30: 'profit',
    31: 'workCost',
    32: 'contractCost',
    40: 'finalCost',
    # 41(부가세), 42(총계): 템플릿 수식 유지 → 입력 안 함
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

            # ─── 공공측량 갑지 A1셀 제목 동적 생성 ───
            branch = body.get('branchName', '수도권지사')
            year   = body.get('year', '2026')
            month  = body.get('month', '')
            ws_gab = wb['공공측량 갑지']
            ws_gab.cell(1, 1).value = f"{year}년 기성준공내역서_도급_SKTNS_{branch}_측량({month}월)"

            # ─── 세부내역: 5행부터 순서대로 입력 ───
            ws_detail = wb['세부내역']
            for i, p in enumerate(projects):
                row = 5 + i
                if row > 20:
                    break
                cost = calc_cost(
                    p.get('exposedKm', 0),
                    p.get('probeKm', 0),
                    p.get('method', 'probe'),
                    p.get('surveyName', '')
                )
                ws_detail.cell(row, 3).value  = p.get('gubun', p.get('division', ''))
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

            # ─── 원가계산서: 건별로 M(13)열부터 3열씩 입력 ───
            ws_cost = wb['원가계산서']
            for i, p in enumerate(projects):
                cost = calc_cost(
                    p.get('exposedKm', 0),
                    p.get('probeKm', 0),
                    p.get('method', 'probe'),
                    p.get('surveyName', '')
                )
                col = 13 + i * 3  # M=13, P=16, S=19, V=22 ...

                # 5행: 공사코드
                ws_cost.cell(5, col).value = f"{i+1}. {p.get('workCode','')}"

                # E23셀: 안전관리비 실제 적용 비율
                ws_cost.cell(23, 5).value = cost['safetyRate']

                # 각 항목 Capex, Opex=0, 계 입력
                for row_num, key in COST_ROW_MAP.items():
                    val = cost.get(key, 0)
                    ws_cost.cell(row_num, col).value     = val  # Capex
                    ws_cost.cell(row_num, col + 1).value = 0    # Opex
                    ws_cost.cell(row_num, col + 2).value = val  # 계

                # 41행(부가세), 42행(총계) 직접 값으로 입력
                ws_cost.cell(41, col).value     = cost['vat']           # 부가세 Capex
                ws_cost.cell(41, col + 1).value = 0                     # 부가세 Opex
                ws_cost.cell(41, col + 2).value = cost['vat']           # 부가세 계
                ws_cost.cell(42, col).value     = cost['totalWithVat']  # 총계 Capex
                ws_cost.cell(42, col + 1).value = 0                     # 총계 Opex
                ws_cost.cell(42, col + 2).value = cost['totalWithVat']  # 총계 계

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
