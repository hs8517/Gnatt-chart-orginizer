import http.server
import socketserver
import json
import os
import webbrowser
import threading
import time

PORT = 8000
PORT = 8000
DATA_DIR = 'data'

class Handler(http.server.SimpleHTTPRequestHandler):
    def get_data_file_path(self, source):
        # source can be 'local' or 'cloud'
        if source not in ['local', 'cloud']:
            source = 'local'
        return os.path.join(DATA_DIR, source, 'data.json')

    def do_GET(self):
        if self.path.startswith('/api/data'):
            # Parse query params
            from urllib.parse import urlparse, parse_qs
            parsed_path = urlparse(self.path)
            query_params = parse_qs(parsed_path.query)
            source = query_params.get('source', ['local'])[0]
            
            data_file = self.get_data_file_path(source)

            self.send_response(200)
            self.send_header('Content-type', 'application/json')
            self.end_headers()
            
            if os.path.exists(data_file):
                with open(data_file, 'r', encoding='utf-8') as f:
                    self.wfile.write(f.read().encode('utf-8'))
            else:
                self.wfile.write(b'{}')
        else:
            return http.server.SimpleHTTPRequestHandler.do_GET(self)

    def do_POST(self):
        if self.path.startswith('/api/data'):
            # Parse query params
            from urllib.parse import urlparse, parse_qs
            parsed_path = urlparse(self.path)
            query_params = parse_qs(parsed_path.query)
            source = query_params.get('source', ['local'])[0]
            
            data_file = self.get_data_file_path(source)
            
            content_length = int(self.headers['Content-Length'])
            post_data = self.rfile.read(content_length)
            
            try:
                # Validate JSON
                json_data = json.loads(post_data.decode('utf-8'))
                
                # Ensure directory exists
                os.makedirs(os.path.dirname(data_file), exist_ok=True)
                
                with open(data_file, 'w', encoding='utf-8') as f:
                    json.dump(json_data, f, indent=4)
                    
                self.send_response(200)
                self.send_header('Content-type', 'application/json')
                self.end_headers()
                self.wfile.write(b'{"status": "success"}')
            except Exception as e:
                self.send_response(500)
                self.end_headers()
                self.wfile.write(f'{{"error": "{str(e)}"}}'.encode('utf-8'))
        else:
            self.send_response(404)
            self.end_headers()

def open_browser():
    time.sleep(1)
    webbrowser.open(f'http://localhost:{PORT}')

if __name__ == '__main__':
    print(f"Serving at http://localhost:{PORT}")
    
    # Initialize separate data directories
    os.makedirs(os.path.join(DATA_DIR, 'local'), exist_ok=True)
    os.makedirs(os.path.join(DATA_DIR, 'cloud'), exist_ok=True)

    threading.Thread(target=open_browser).start()
    
    with socketserver.TCPServer(("", PORT), Handler) as httpd:
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nShutting down server...")
            httpd.shutdown()
