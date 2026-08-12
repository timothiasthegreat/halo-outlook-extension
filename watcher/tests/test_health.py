"""Verify health endpoint: TestClient + live uvicorn smoke test."""
import json, asyncio, urllib.request, random
from fastapi.testclient import TestClient
from watcher.watcher import _build_health_app, _utcnow_iso, HEALTH_PORT, _run_health_server

print('=== TestClient ===')
for label, s in [('populated',{'conversations':5,'last_sync_at':'2026-08-12T15:22:00Z'}),('null',{'conversations':0,'last_sync_at':None})]:
    app = _build_health_app(s)
    cl = TestClient(app)
    r = cl.get('/health')
    d = r.json()
    assert r.status_code == 200, f"{label} status={r.status_code}"
    assert d['status'] == 'ok', f"{label} status={d['status']}"
    assert d['conversations'] == s['conversations'], f"{label} conv={d['conversations']}"
    assert d['last_sync_at'] == s['last_sync_at'], f"{label} ts={d['last_sync_at']}"
    print(f'  OK {label}')
print('  TestClient: PASS')

print('=== Live server ===')
port = random.randint(9000,9999)
async def srv():
    sh = {'conversations':3,'last_sync_at':'2026-08-12T16:00:00Z'}
    task = asyncio.create_task(_run_health_server(sh,port))
    await asyncio.sleep(1)
    r = urllib.request.urlopen(f'http://127.0.0.1:{port}/health',timeout=5)
    d = json.loads(r.read())
    assert d['status'] == 'ok', d
    assert d['conversations'] == 3, d
    assert d['last_sync_at'] == '2026-08-12T16:00:00Z', d
    task.cancel()
    try: await task
    except asyncio.CancelledError: pass
    print('  Live server: PASS')
asyncio.run(srv())
print('\nALL PASSED')