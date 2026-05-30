from http.server import BaseHTTPRequestHandler
import json
import os
import urllib.request
import urllib.error

SUPABASE_URL = os.environ.get('SUPABASE_URL', '')
SUPABASE_KEY = os.environ.get('SUPABASE_KEY', '')

def supabase_request(method, path, data=None, params=None):
    url = f"{SUPABASE_URL}/rest/v1/{path}"
    if params:
        url += '?' + '&'.join(f"{k}={v}" for k, v in params.items())
    headers = {
        'apikey': SUPABASE_KEY,
        'Authorization': f'Bearer {SUPABASE_KEY}',
        'Content-Type': 'application/json',
        'Prefer': 'return=representation',
    }
    body = json.dumps(data).encode('utf-8') if data else None
    req = urllib.request.Request(url, data=body, headers=headers, method=method)
    try:
        with urllib.request.urlopen(req) as res:
            return json.loads(res.read().decode('utf-8'))
    except urllib.error.HTTPError as e:
        return {'error': e.read().decode('utf-8')}

class handler(BaseHTTPRequestHandler):

    def send_cors_headers(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, DELETE, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type, Accept')

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_cors_headers()
        self.send_header('Content-Length', '0')
        self.end_headers()

    def send_json(self, data, status=200):
        resp = json.dumps(data, ensure_ascii=False).encode('utf-8')
        self.send_response(status)
        self.send_cors_headers()
        self.send_header('Content-Type', 'application/json')
        self.send_header('Content-Length', str(len(resp)))
        self.end_headers()
        self.wfile.write(resp)

    def do_GET(self):
        # 전체 조회 또는 연도별 조회
        from urllib.parse import urlparse, parse_qs
        parsed = urlparse(self.path)
        qs = parse_qs(parsed.query)
        year = qs.get('year', [None])[0]

        params = {'order': 'created_at.desc'}
        if year:
            params['year'] = f'eq.{year}'

        result = supabase_request('GET', 'projects', params=params)
        self.send_json(result)

    def do_POST(self):
        length = int(self.headers.get('Content-Length', 0))
        body = json.loads(self.rfile.read(length))
        action = body.get('action', 'save')

        if action == 'save':
            # 사업 저장
            projects = body.get('projects', [])
            year = body.get('year', '2026')
            month = body.get('month', '')
            rows = []
            for p in projects:
                rows.append({
                    'year': year,
                    'month': month,
                    'gubun': p.get('gubun', ''),
                    'region': p.get('region', ''),
                    'survey_name': p.get('surveyName', ''),
                    'work_code': p.get('workCode', ''),
                    'work_name': p.get('workName', ''),
                    'tango_type': p.get('tangoType', ''),
                    'tango_km': p.get('tangoKm', 0),
                    'exposed_km': p.get('exposedKm', 0),
                    'probe_km': p.get('probeKm', 0),
                    'method': p.get('method', 'probe'),
                    'final_cost': p.get('finalCost', 0),
                    'vat': p.get('vat', 0),
                    'total_with_vat': p.get('totalWithVat', 0),
                    'remark': p.get('remark', ''),
                    'is_carried_over': p.get('isCarriedOver', False),
                    'carried_over_from': p.get('carriedOverFrom', ''),
                })
            result = supabase_request('POST', 'projects', data=rows)
            self.send_json({'ok': True, 'data': result})

        elif action == 'delete':
            # 사업 삭제
            project_id = body.get('id')
            result = supabase_request('DELETE', 'projects',
                                      params={'id': f'eq.{project_id}'})
            self.send_json({'ok': True})

        elif action == 'carry_over':
            # 이월 처리
            project_id = body.get('id')
            to_month = body.get('toMonth', '')
            result = supabase_request('PATCH', 'projects',
                                      data={'is_carried_over': True,
                                            'carried_over_from': to_month},
                                      params={'id': f'eq.{project_id}'})
            self.send_json({'ok': True})
