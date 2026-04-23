import subprocess
import sys
import os

def main():
    base_dir    = os.path.dirname(os.path.abspath(__file__))
    backend_dir = os.path.join(base_dir, 'backend')
    req_file    = os.path.join(backend_dir, 'requirements.txt')

    print("패키지 설치 중...")
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', '-r', req_file])

    print("=" * 50)
    print("  Maintenance Data Dashboard (SQLite)")
    print("  http://localhost:5000")
    print("  브라우저에서 위 주소로 접속하세요")
    print("=" * 50)

    subprocess.check_call([
        sys.executable, '-m', 'uvicorn',
        'app:app',
        '--app-dir', backend_dir,
        '--host', '0.0.0.0',
        '--port', '5000',
    ])

if __name__ == '__main__':
    main()
