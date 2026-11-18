"""שרת מקומי לבדיקת ה-PWA (סוגי MIME נכונים גם ב-Windows). הרצה: python test/serve.py"""
import functools, http.server, mimetypes, os, sys

mimetypes.add_type('text/javascript', '.js')
mimetypes.add_type('application/manifest+json', '.webmanifest')
mimetypes.add_type('image/svg+xml', '.svg')

root = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'docs')
port = int(sys.argv[1]) if len(sys.argv) > 1 else 8765
handler = functools.partial(http.server.SimpleHTTPRequestHandler, directory=root)
handler.func.extensions_map.update({'.js': 'text/javascript', '.webmanifest': 'application/manifest+json'})
http.server.ThreadingHTTPServer(('127.0.0.1', port), handler).serve_forever()
