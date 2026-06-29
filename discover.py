"""
Reads state.json for last scraped ID, scans forward to find today's upper bound,
then outputs a dynamic parallel job matrix for GitHub Actions.
"""
import json, os, sys, time
from curl_cffi import requests as cffi
import warnings
warnings.filterwarnings('ignore')

BASE = 'https://mdcat.pmdc.pk'

with open('state.json') as f:
    last_id = json.load(f)['last_id']

print(f'Last scraped ID: {last_id:,}')

s = cffi.Session(impersonate='chrome124')

def valid(cid):
    try:
        r = s.get(f'{BASE}/Payment/DownloadChallan/{cid}',
                  timeout=4, verify=False, allow_redirects=True)
        return r.status_code == 200 and len(r.content) > 30000
    except:
        return False

# Scan forward from last_id to find new upper bound
# Use large step to jump quickly, then binary-refine the boundary
new_upper = last_id
step = 200
cid = last_id + step

for _ in range(100):
    if valid(cid):
        new_upper = cid
        step = min(step * 2, 2000)  # grow step once we know there's data
        cid = new_upper + step
    else:
        if step <= 10:
            break
        step = step // 2
        cid = new_upper + step
    time.sleep(0.02)

print(f'New upper bound: {new_upper:,}')
print(f'New IDs today:   {new_upper - last_id:,}')

gf = open(os.environ['GITHUB_OUTPUT'], 'a')

if new_upper <= last_id:
    print('No new challans — skipping scrape.')
    gf.write('has_new=false\n')
    gf.write(f'new_upper={last_id}\n')
    gf.write('matrix={"include":[]}\n')
    gf.close()
    sys.exit(0)

# Split [last_id+1 .. new_upper+10000] into 20 parallel chunks
# Small buffer catches challans issued while discover was running
N_JOBS = 20
start  = last_id + 1
end    = new_upper + 10000
span   = end - start
chunk  = max(span // N_JOBS, 1)

chunks = []
for i in range(N_JOBS):
    cs = start + i * chunk
    ce = start + (i + 1) * chunk - 1 if i < N_JOBS - 1 else end
    chunks.append({'start': cs, 'end': ce})

matrix = json.dumps({'include': [{'chunk': c} for c in chunks]})
print(f'Chunks: {chunks}')

gf.write('has_new=true\n')
gf.write(f'new_upper={new_upper}\n')   # true last valid ID — state.json stays accurate
gf.write(f'scan_end={end}\n')          # last_valid+10000 buffer used only for scrape range
gf.write(f'matrix={matrix}\n')
gf.close()
