"""Local development runner for PO Splitter backend.

Usage: python run_local.py [port]
Default port: 5000
"""
import sys
import os

# Add api folder to path so we can import index.py
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'api'))

from index import app

if __name__ == '__main__':
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 5000
    print(f'PO Splitter Backend starting on http://localhost:{port}')
    print(f'Using: openpyxl (pure Python) - no MS Excel/WPS required')
    print(f'Endpoints:')
    print(f'  GET  /api/health   - health check')
    print(f'  POST /api/preview  - preview Excel data')
    print(f'  POST /api/process  - process & download Excel')
    print()
    app.run(host='0.0.0.0', port=port, debug=True)
