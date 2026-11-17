import http.server
import mimetypes
import socketserver
import sys

mimetypes.add_type('application/javascript', '.mjs')
mimetypes.add_type('application/manifest+json', '.webmanifest')

port = int(sys.argv[1]) if len(sys.argv) > 1 else 8080

class Handler(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header('Cache-Control', 'no-cache')
        super().end_headers()

with socketserver.ThreadingTCPServer(('0.0.0.0', port), Handler) as httpd:
    print(f'Serving YunzyBook PWA at http://0.0.0.0:{port}/')
    httpd.serve_forever()
