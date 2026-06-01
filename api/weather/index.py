from http.server import BaseHTTPRequestHandler
import json, urllib.request, urllib.error

class handler(BaseHTTPRequestHandler):

    def do_OPTIONS(self):
        self.send_response(200)
        self._cors()
        self.end_headers()

    def do_GET(self):
        try:
            url = 'https://api.windy.com/api/point-forecast/v2'
            body = json.dumps({
                'lat': 46.301,
                'lon': 30.655,
                'model': 'gfs',
                'parameters': ['temp', 'wind', 'windDir', 'precip', 'clouds'],
                'levels': ['surface'],
                'key': 'Mc6FW3CVOfwMpGafNoHBZPO6YXzC3DVk'
            }).encode('utf-8')

            req = urllib.request.Request(
                url,
                data=body,
                headers={'Content-Type': 'application/json'},
                method='POST'
            )
            with urllib.request.urlopen(req, timeout=10) as r:
                data = r.read()

            self.send_response(200)
            self._cors()
            self.send_header('Content-Type', 'application/json')
            self.send_header('Content-Length', str(len(data)))
            self.end_headers()
            self.wfile.write(data)

        except Exception as e:
            resp = json.dumps({'error': str(e)}).encode()
            self.send_response(500)
            self._cors()
            self.send_header('Content-Type', 'application/json')
            self.send_header('Content-Length', str(len(resp)))
            self.end_headers()
            self.wfile.write(resp)

    def _cors(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
