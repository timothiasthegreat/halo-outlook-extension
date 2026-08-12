import sys, json, asyncio, urllib.request, random
from fastapi.testclient import TestClient
from watcher.watcher import _build_health_app, _utcnow_iso, HEALTH_PORT, _run_health_server

ok = 0
def t(n, c, d=''):
    global ok
    if c: ok+=1; print(f'  PASS {n}')
    else: print(f'  FAIL {n}  {d}')

print('=== TestClient ===')
for label, s in [('populated',{'conversations':5,'last_sync_at':'2026-08-12T15:22:00Z'}),('null',{'conversations':0,'last_sync_at':None})]:
    app = _build_health_app(s)
    cl = TestClient(app)
    r = cl.get('/health')
    d = r.json()
    t(f'200 ({label})', r.status_code==200, str(r.status_code))
    t(f'ok ({label})', d['status']=='ok', d.get('status'))

print('=== Live server ===')
port = random.randint(9000,9999)
async def srv():
    global ok
    sh = {'conversations':3,'last_sync_at':'2026-08-12T16:00:00Z'}
    task = asyncio.create_task(_run_health_server(sh,port))
    await asyncio.sleep(2)
    try:
        r = urllib.request.urlopen(f'http://127.0.0.1:{port}/health',timeout=5)
        d = json.loads(r.read())
        t('live status=ok', d['status']=='ok', d.get('status'))
        t('live conv=3', d['conversations']==3, str(d.get('conversations')))
        t('live ts', d['last_sync_at']=='2026-08-12T16:00:00Z', d.get('last_sync_at'))
    except Exception as e:
        t('live', False, str(e))
    finally:
        task.cancel()
        try: await task
        except asyncio.CancelledError: pass
asyncio.run(srv())
print(f'\n{ok}/7 passed')
sys.exit(0 if ok==7 else 1)