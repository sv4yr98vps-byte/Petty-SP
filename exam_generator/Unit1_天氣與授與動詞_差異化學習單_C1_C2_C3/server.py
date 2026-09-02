import http.server
import socketserver
import socket
import webbrowser
import os
import sys

def get_local_ip():
    """取得本機在區域網路 (Wi-Fi/LAN) 中的 IP 位址"""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(('8.8.8.8', 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return '127.0.0.1'

PORT = 8000
DIRECTORY = os.path.dirname(os.path.abspath(__file__))

class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=DIRECTORY, **kwargs)

    def log_message(self, format, *args):
        print(f"[平板連線日誌] {self.address_string()} - {args[0]}")

if __name__ == '__main__':
    local_ip = get_local_ip()
    url = f"http://{local_ip}:{PORT}/index.html"

    print("=" * 60)
    print("🎓 國中英語 Unit 1 季節天氣與授與動詞｜平板教學伺服器已啟動")
    print("=" * 60)
    print(f"📡 電腦與平板請連接至【相同 Wi-Fi 網路】")
    print(f"🔗 本機瀏覽器網址：http://localhost:{PORT}/index.html")
    print(f"📱 平板掃描/連線網址：{url}")
    print("=" * 60)
    print("正在自動為您開啟瀏覽器首頁...")
    print("按 Ctrl + C 可隨時停止伺服器")
    print("=" * 60)

    try:
        webbrowser.open(url)
    except Exception:
        pass

    socketserver.TCPServer.allow_reuse_address = True
    try:
        with socketserver.TCPServer(("", PORT), Handler) as httpd:
            httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n伺服器已安全停止。")
        sys.exit(0)
