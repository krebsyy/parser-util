"""Find the first ever challan ID and today's upper bound."""
from curl_cffi import requests as cffi
import warnings, time
warnings.filterwarnings('ignore')

BASE = 'https://mdcat.pmdc.pk'
s = cffi.Session(impersonate='chrome124')

def valid(cid):
    try:
        r = s.get(f'{BASE}/Payment/DownloadChallan/{cid}',
                  timeout=10, verify=False, allow_redirects=True)
        return r.status_code == 200 and len(r.content) > 30000
    except:
        return False

# ── 1. Binary search for first valid challan ─────────────────────────────────
KNOWN_GOOD = 2220968283
lo = KNOWN_GOOD - 1000000
hi = KNOWN_GOOD
print(f'Binary searching for first challan between {lo:,} and {hi:,}...')

while hi - lo > 1:
    mid = (lo + hi) // 2
    v = valid(mid)
    tag = 'VALID' if v else 'empty'
    print(f'  {mid:,} -> {tag}')
    if v:
        hi = mid
    else:
        lo = mid
    time.sleep(0.05)

print(f'\nFirst challan ID: {hi:,}')

# ── 2. Find current upper bound ───────────────────────────────────────────────
print('\nScanning forward from 2221354800 to find latest ID...')
last = 2221354800
step = 200
cid = last + step
while True:
    if valid(cid):
        last = cid
        cid += step
    else:
        if step <= 10:
            break
        step = step // 2
        cid = last + step
    time.sleep(0.03)

print(f'Latest challan ID today: {last:,}')
print(f'New IDs since last test (a few hours ago): {last - 2221354800:,}')
print(f'\nSAFE SCRAPE RANGE FOR LAST DAY:')
print(f'  --start {hi} --end [check again on July 13]')
