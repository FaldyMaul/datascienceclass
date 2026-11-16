"""Builds the Arunika Fiber teaching dataset (fictional company, invented numbers).

Usage:  python tools/build_case_data.py
Writes: data/Arunika_Fiber_BI_Case.xlsx   (multi-sheet workbook for Google Sheets / Data Studio)
        data/case_meta.js                 (data dictionary + events for the case page)
        tools/findings.json               (check numbers used in the lecturer's guide)
Needs:  numpy, pandas, xlsxwriter
"""
import json, os
import numpy as np, pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
rng = np.random.default_rng(20261116)
N = 1500
MONTHS = pd.date_range('2025-11-01', periods=12, freq='MS')
M = len(MONTHS)
SNAP = pd.Timestamp('2026-11-15')          # day the data was extracted
mi = {m.strftime('%Y-%m'): i for i, m in enumerate(MONTHS)}
FEB, MAR, APR, JUN, JUL, AUG, SEP, OCT, JAN = mi['2026-02'], mi['2026-03'], mi['2026-04'], mi['2026-06'], mi['2026-07'], mi['2026-08'], mi['2026-09'], mi['2026-10'], mi['2026-01']

# ---------------------------------------------------------------- dimensions
CITIES = [  # name, province, island, lat, lon, weight, segment tilt [Family, WFH, Gamer, Student, HomeBiz]
    ('Jakarta', 'DKI Jakarta', 'Jawa', -6.2088, 106.8456, 1.6, [1.0, 1.5, 1.0, 0.7, 0.9]),
    ('Bandung', 'Jawa Barat', 'Jawa', -6.9175, 107.6191, 1.2, [1.0, 1.1, 1.2, 1.2, 0.9]),
    ('Surabaya', 'Jawa Timur', 'Jawa', -7.2575, 112.7521, 1.3, [1.1, 1.1, 1.0, 0.8, 1.1]),
    ('Semarang', 'Jawa Tengah', 'Jawa', -6.9667, 110.4167, 0.9, [1.2, 0.9, 0.9, 0.9, 1.0]),
    ('Yogyakarta', 'DI Yogyakarta', 'Jawa', -7.7956, 110.3695, 0.9, [0.8, 0.8, 1.1, 2.0, 0.9]),
    ('Malang', 'Jawa Timur', 'Jawa', -7.9666, 112.6326, 0.8, [0.9, 0.7, 1.2, 1.9, 0.9]),
    ('Medan', 'Sumatera Utara', 'Sumatera', 3.5952, 98.6722, 1.0, [1.2, 0.8, 1.0, 0.8, 1.3]),
    ('Palembang', 'Sumatera Selatan', 'Sumatera', -2.9761, 104.7754, 0.8, [1.3, 0.8, 0.9, 0.7, 1.2]),
    ('Pekanbaru', 'Riau', 'Sumatera', 0.5071, 101.4478, 0.7, [1.2, 0.9, 0.9, 0.7, 1.2]),
    ('Makassar', 'Sulawesi Selatan', 'Sulawesi', -5.1477, 119.4327, 0.9, [1.2, 0.8, 1.0, 0.9, 1.1]),
    ('Balikpapan', 'Kalimantan Timur', 'Kalimantan', -1.2379, 116.8529, 0.7, [1.1, 1.2, 0.9, 0.5, 1.1]),
    ('Denpasar', 'Bali', 'Bali', -8.6705, 115.2126, 0.8, [0.9, 1.2, 0.9, 0.6, 1.7]),
]
SUB = [('Pusat', 0.0, 0.0), ('Utara', 0.045, 0.01), ('Selatan', -0.045, -0.01), ('Timur', 0.0, 0.05)]
BUILD = {'Balikpapan Utara', 'Balikpapan Timur', 'Makassar Timur', 'Pekanbaru Selatan', 'Malang Utara', 'Denpasar Timur', 'Medan Selatan', 'Palembang Utara'}
SELL = {'Jakarta Pusat', 'Jakarta Selatan', 'Surabaya Pusat', 'Bandung Pusat', 'Semarang Pusat', 'Medan Pusat', 'Yogyakarta Pusat'}
TRAP = {'Pekanbaru Timur', 'Palembang Timur', 'Makassar Utara', 'Semarang Timur'}

SEG = ['Family with kids', 'Work-from-home professional', 'Gamer & streamer', 'Student / shared house', 'Home business']
SEG_P = np.array([.30, .22, .15, .18, .15])
PKG = [('P030', 'Start 30', 30, 275000), ('P050', 'Family 50', 50, 345000), ('P100', 'Stream 100', 100, 445000),
       ('P200', 'Pro 200', 200, 615000), ('P300', 'Ultra 300', 300, 875000)]
PKG_OLD = {0: 260000, 1: 325000}            # price before 1 Aug 2026
SPEED = np.array([p[2] for p in PKG])
PKG_BY_SEG = np.array([[.15, .40, .30, .12, .03], [.05, .25, .40, .22, .08], [.22, .30, .23, .15, .10],
                       [.45, .35, .15, .04, .01], [.10, .30, .30, .20, .10]])
TV_BASIC = 45000
PACKS = [('Sports Pack', 60000), ('Kids Pack', 30000), ('Movies Pack', 50000), ('Asian Drama Pack', 35000)]
ADDONS = [('Speed Booster 7 days', 35000, 'One-off'), ('Wi-Fi Extender', 350000, 'One-off'), ('Static IP', 100000, 'Monthly')]
CHANNELS = [  # name, genre, tier, base hours/month, zero share, licence cost per entitled subscriber per month
    ('Berita Kini', 'News', 'Basic TV', 4.0, .35, 900), ('Nusantara News', 'News', 'Basic TV', 2.2, .50, 800),
    ('Layar Keluarga', 'Entertainment', 'Basic TV', 7.5, .15, 1800), ('Kids Corner Lite', 'Kids', 'Basic TV', 3.0, .45, 900),
    ('Drama Asia Lite', 'Drama', 'Basic TV', 6.5, .25, 1200), ('Arena Lite', 'Sports', 'Basic TV', 2.6, .40, 1500),
    ('Musik Kita', 'Music', 'Basic TV', 2.0, .50, 600), ('Dunia Dok', 'Documentary', 'Basic TV', 1.8, .55, 1000),
    ('Dapur Rasa', 'Lifestyle', 'Basic TV', 2.4, .50, 700), ('Sinema Lite', 'Movies', 'Basic TV', 3.2, .40, 1400),
    ('Arena Bola', 'Sports', 'Sports Pack', 7.0, .15, 18000), ('Arena Racket', 'Sports', 'Sports Pack', 2.0, .50, 6000),
    ('Arena Golf', 'Sports', 'Sports Pack', 0.4, .85, 9000),
    ('Kids Corner', 'Kids', 'Kids Pack', 10.0, .08, 5000), ('Toon Planet', 'Kids', 'Kids Pack', 8.0, .12, 4000),
    ('Belajar Seru', 'Education', 'Kids Pack', 2.5, .45, 1500),
    ('Sinema Box', 'Movies', 'Movies Pack', 4.5, .20, 12000), ('Sinema Action', 'Movies', 'Movies Pack', 2.8, .30, 8000),
    ('Sinema Classic', 'Movies', 'Movies Pack', 0.5, .75, 7000),
    ('Drama Asia', 'Drama', 'Asian Drama Pack', 13.0, .04, 6000), ('K-Wave', 'Drama', 'Asian Drama Pack', 10.5, .06, 7000),
    ('Serial Mandarin', 'Drama', 'Asian Drama Pack', 4.0, .35, 3000),
]
AFF = {'News': [1.0, 1.3, .4, .5, 1.5], 'Entertainment': [1.3, .8, .6, 1.0, 1.0], 'Kids': [2.2, .3, .1, .1, .5],
       'Drama': [1.4, .8, .3, 1.5, .9], 'Sports': [1.0, 1.1, 1.3, .9, 1.0], 'Music': [.8, .8, .9, 1.4, 1.0],
       'Documentary': [.9, 1.3, .7, .6, .9], 'Lifestyle': [1.2, .9, .3, .7, 1.4], 'Movies': [1, 1, 1, 1, 1],
       'Education': [1.8, .3, .1, .6, .3]}
CATS = ['Video streaming', 'Gaming', 'Social media', 'Work & video calls', 'Education', 'Shopping & other']
CAT_ALPHA = np.array([[45, 8, 18, 7, 14, 8], [30, 5, 12, 38, 5, 10], [28, 45, 12, 4, 3, 8], [35, 18, 24, 5, 12, 6], [22, 3, 20, 25, 3, 27]], float)
METHODS = ['Auto-debit card', 'E-wallet', 'Bank transfer', 'Minimarket counter']
LATE_BASE = np.array([.02, .09, .15, .27])
CHANNELS_SALES = ['MyArunika app', 'Website', 'Call centre', 'Door-to-door sales', 'Partner store']
SALES_COST = dict(zip(CHANNELS_SALES, [8000, 15000, 60000, 180000, 90000]))

# ---------------------------------------------------------------- areas
rows = []
for c, prov, isl, lat, lon, w, tilt in CITIES:
    for s, dlat, dlon in SUB:
        name = f'{c} {s}'
        hh = int(rng.uniform(22000, 60000) * (1.25 if c in ('Jakarta', 'Surabaya', 'Bandung', 'Medan') else .8) / 100) * 100
        if name in BUILD:
            cov, pen, comp, cost, wl_rate, inc, wait = rng.uniform(.22, .38), rng.uniform(.36, .46), int(rng.integers(0, 2)), rng.uniform(1.4, 1.9), rng.uniform(.085, .13), rng.choice(['Mid', 'High'], p=[.6, .4]), rng.uniform(9, 14)
        elif name in SELL:
            cov, pen, comp, cost, wl_rate, inc, wait = rng.uniform(.86, .96), rng.uniform(.10, .15), int(rng.integers(3, 5)), rng.uniform(2.2, 3.0), rng.uniform(.02, .04), 'High', rng.uniform(2, 4)
        elif name in TRAP:
            hh = int(hh * 1.9 / 100) * 100
            cov, pen, comp, cost, wl_rate, inc, wait = rng.uniform(.18, .30), rng.uniform(.14, .19), int(rng.integers(2, 4)), rng.uniform(3.4, 4.2), rng.uniform(.048, .062), 'Low', rng.uniform(6, 9)
        else:
            cov, pen, comp, cost, wl_rate, inc, wait = rng.uniform(.55, .80), rng.uniform(.20, .30), int(rng.integers(1, 4)), rng.uniform(1.9, 2.8), rng.uniform(.02, .05), rng.choice(['Low', 'Mid', 'High'], p=[.25, .55, .2]), rng.uniform(3, 6)
        hp = int(hh * cov / 10) * 10
        subs = int(hp * pen)
        unc = hh - hp
        rows.append(dict(area_name=name, city=c, province=prov, island=isl, latitude=round(lat + dlat, 4), longitude=round(lon + dlon, 4),
                         income_tier=inc, households=hh, homes_passed=hp, uncovered_households=unc, subscribers=subs,
                         coverage_pct=round(hp / hh * 100, 1), penetration_pct=round(subs / hp * 100, 1),
                         waitlist_requests_12m=int(unc * wl_rate), competitor_count=comp,
                         build_cost_per_home_idr=int(cost * 1e6 / 1e4) * 10000, avg_install_wait_days=round(wait, 1)))
areas = pd.DataFrame(rows)
areas.insert(0, 'area_id', [f'A{i + 1:02d}' for i in range(len(areas))])
areas['geo_point'] = areas.latitude.astype(str) + ',' + areas.longitude.astype(str)
areas['area_type'] = np.where(areas.area_name.isin(BUILD), 'build', np.where(areas.area_name.isin(SELL), 'sell', np.where(areas.area_name.isin(TRAP), 'trap', 'mid')))

# area × month: coverage checks, waitlist requests, new installs
am = []
for _, a in areas.iterrows():
    trend = np.linspace(.7, 1.3, M) if a.area_type == 'build' else np.linspace(1.05, .95, M)
    wts = trend * rng.uniform(.85, 1.15, M); wts /= wts.sum()
    wl = rng.multinomial(a.waitlist_requests_12m, wts)
    inst_rate = {'build': .020, 'sell': .004, 'trap': .006, 'mid': .009}[a.area_type]
    promo = np.ones(M); promo[[FEB, MAR]] = 1.5
    inst = rng.poisson(a.subscribers * inst_rate * promo * rng.uniform(.85, 1.15, M))
    covered_checks = rng.poisson(inst * rng.uniform(2.2, 3.2, M) + 5)
    for m in range(M):
        am.append(dict(month=MONTHS[m], area_id=a.area_id, area_name=a.area_name, city=a.city, island=a.island,
                       coverage_checks=int(wl[m] + covered_checks[m]), waitlist_requests=int(wl[m]), new_installs=int(inst[m])))
area_monthly = pd.DataFrame(am)

# ---------------------------------------------------------------- customers
aw = areas.subscribers.values ** .55
aw = aw * np.where(areas.area_type == 'build', 1.35, 1.0); aw = aw / aw.sum()
a_idx = rng.choice(len(areas), N, p=aw)
A = areas.iloc[a_idx].reset_index(drop=True)
tilt = {c[0]: np.array(c[6]) for c in CITIES}
seg = np.array([rng.choice(5, p=(SEG_P * tilt[c]) / (SEG_P * tilt[c]).sum()) for c in A.city])
pkg0 = np.array([rng.choice(5, p=PKG_BY_SEG[s]) for s in seg])
new_p = np.where(A.area_type == 'build', .42, .16)
is_new = rng.random(N) < new_p
jw = np.ones(M); jw[[FEB, MAR]] = [1.5, 1.8]; jw /= jw.sum()
join_m = np.where(is_new, rng.choice(M, N, p=jw), -1)
old_ten = np.clip(rng.gamma(2.2, 13, N) + 1, 1, 84).astype(int)
join_date = []
for i in range(N):
    if join_m[i] >= 0:
        d = MONTHS[join_m[i]] + pd.Timedelta(days=int(rng.integers(0, 27)))
    else:
        d = MONTHS[0] - pd.DateOffset(months=int(old_ten[i])) + pd.Timedelta(days=int(rng.integers(0, 27)))
    join_date.append(pd.Timestamp(d).normalize())
join_date = pd.DatetimeIndex(join_date)
tenure = ((pd.Timestamp('2026-10-31') - join_date).days / 30.44).astype(int)
AGE = ['18-24', '25-34', '35-44', '45-54', '55+']
AGE_P = [[.03, .32, .42, .18, .05], [.08, .47, .30, .12, .03], [.38, .42, .14, .05, .01], [.72, .22, .04, .01, .01], [.05, .28, .34, .23, .10]]
age = np.array([rng.choice(5, p=AGE_P[s]) for s in seg])
HOME = ['Landed house', 'Apartment', 'Boarding house (kos)', 'Shop-house (ruko)']
HOME_P = [[.82, .15, .0, .03], [.50, .45, .03, .02], [.55, .30, .13, .02], [.20, .12, .66, .02], [.45, .05, .0, .50]]
home = np.array([rng.choice(4, p=HOME_P[s]) for s in seg])
hsize = np.clip(np.round(rng.normal([4.2, 2.4, 2.6, 5.0, 3.4], 1.0, (N, 5))[np.arange(N), seg]), 1, 9).astype(int)
eng = np.clip(rng.beta(2.2, 2.2, N) + np.array([.0, .10, .15, .08, -.05])[seg] - (age - 1) * .06, .03, .97)
METH_P = np.array([[.20, .28, .30, .22], [.32, .38, .24, .06], [.16, .50, .22, .12], [.04, .40, .16, .40], [.18, .22, .40, .20]])
method = np.array([rng.choice(4, p=METH_P[s]) for s in seg])
ACQ_P = {'build': [.10, .10, .08, .62, .10], 'sell': [.34, .30, .14, .08, .14], 'trap': [.15, .15, .12, .43, .15], 'mid': [.25, .22, .15, .25, .13]}
acq = np.array([rng.choice(5, p=ACQ_P[t]) for t in A.area_type])

mgrid = np.arange(M)[None, :]
active = mgrid >= join_m[:, None]
ten_m = np.maximum(0, ((MONTHS.values[None, :] - join_date.values[:, None]) / np.timedelta64(1, 'D') / 30.44)).astype(int)

# demand and upgrades
demand = np.array([31, 34, 60, 38, 25])[seg] * rng.lognormal(0, .30, N) * (.82 + .05 * hsize)
util0 = demand / SPEED[pkg0] * 100
p_up = np.where(pkg0 < 4, .03 + .13 * (util0 > 85), 0)
upw = np.ones(M); upw[SEP] = 5; upw[OCT] = 1.6; upw[0] = 0
up_m = np.full(N, 99)
for i in np.where(rng.random(N) < p_up)[0]:
    ok = np.where(active[i] & (mgrid[0] > max(join_m[i], 0)))[0]
    if len(ok):
        up_m[i] = rng.choice(ok, p=upw[ok] / upw[ok].sum())
up_step = np.where(rng.random(N) < .15, 2, 1)
pkg_m = np.where(mgrid >= up_m[:, None], np.minimum(4, pkg0 + up_step)[:, None], pkg0[:, None])
pkg_cur = pkg_m[:, -1]
fee_tab = np.array([[PKG_OLD.get(p, PKG[p][3]) if m < AUG else PKG[p][3] for m in range(M)] for p in range(5)])
fee_m = fee_tab[pkg_m, mgrid]

# TV
tv = rng.random(N) < np.array([.66, .40, .30, .26, .46])[seg]
PACK_P = np.array([[.20, .46, .30, .12], [.24, .03, .30, .07], [.30, .02, .28, .03], [.13, .02, .26, .15], [.24, .05, .32, .08]])
pack = (rng.random((N, 4)) < PACK_P[seg]) & tv[:, None]
pack_start = np.where(rng.random((N, 4)) < .35, rng.integers(1, M, (N, 4)), -1)
pack_start = np.maximum(pack_start, np.maximum(join_m, -1)[:, None])
wc = tv & ~pack[:, 0] & (rng.random(N) < .17) & (join_m < JUN)       # bought the Sports Pack for the World Cup
pack[:, 0] |= wc; pack_start[wc, 0] = JUN
movie_dead = rng.random(N) < .46                                     # owns Movies Pack, never watches
pack_on = pack[:, :, None] & (mgrid[None] >= pack_start[:, :, None]) & active[:, None, :]
pack_fee = (pack_on * np.array([p[1] for p in PACKS])[None, :, None]).sum(1)
static_ip = (seg == 4) & (rng.random(N) < .14)
static_start = np.where(rng.random(N) < .4, rng.integers(1, M, N), -1); static_start = np.maximum(static_start, join_m)
static_on = static_ip[:, None] & (mgrid >= static_start[:, None]) & active

# ---------------------------------------------------------------- internet usage
speed_m = SPEED[pkg_m]
util = np.clip(demand[:, None] * (1 + .010 * mgrid) * rng.lognormal(0, .08, (N, M)) / speed_m * 100, 3, 100)
season = np.ones((N, M)); season[:, [JUN, JUL]] = np.where(np.isin(seg, [0, 3]), 1.12, 1.03)[:, None]
gb = np.array([430, 360, 560, 390, 250])[seg][:, None] * rng.lognormal(0, .33, N)[:, None] * (1 + .012 * mgrid) * season * rng.lognormal(0, .10, (N, M))
gb = gb * np.where(util > 92, .90, 1)
up_ratio = np.array([.08, .22, .16, .10, .25])[seg][:, None] * rng.lognormal(0, .15, (N, M))
devices = np.clip(np.round(hsize[:, None] * np.array([2.3, 3.0, 3.2, 1.8, 2.6])[seg][:, None] + rng.normal(0, 1, (N, M))), 1, 30)
aq = np.where(A.area_type == 'build', 1.35, np.where(A.area_type == 'sell', .8, 1.0))
outage = rng.gamma(1.2, 25, (N, M)) * aq[:, None]
mk = (A.city == 'Makassar').values; md = (A.city == 'Medan').values
outage[mk, FEB] += rng.uniform(600, 1400, mk.sum())
outage[md, JUL] += rng.uniform(180, 520, md.sum())
measured = speed_m * np.clip(rng.normal(.93, .03, (N, M)) - .22 * (util > 90) - .0002 * outage, .35, 1.0)

# ---------------------------------------------------------------- payments
p_late = LATE_BASE[method][:, None] + np.array([0, -.02, .02, .08, .03])[seg][:, None] - .07 * (eng[:, None] - .5)
madj = np.zeros(M); madj[MAR] = .06; madj[APR] = .09; madj[JUL] = .04
p_late = p_late + madj[None, :]
p_late = p_late + np.where((mgrid >= AUG) & (pkg_m <= 1), .05, 0)
p_late = np.where(method[:, None] == 0, np.minimum(p_late, .04), p_late)
p_late = np.clip(p_late, .005, .75)
late = (rng.random((N, M)) < p_late) & active
days_late = np.where(late, 1 + rng.geometric(.11, (N, M)), 0)
due = np.array([m + pd.Timedelta(days=19) for m in MONTHS])
overdue_now = np.array([(SNAP - d).days for d in due])
unpaid = late & ((days_late > 45) | (days_late > overdue_now[None, :]))
days_late = np.where(unpaid, np.minimum(overdue_now[None, :], 90), days_late)
billed = fee_m + pack_fee + np.where(tv[:, None], TV_BASIC, 0) + np.where(static_on, 100000, 0)
on_time = active & ~late

# ---------------------------------------------------------------- app and loyalty
e2 = np.clip(eng[:, None] * 1.1 + np.where(mgrid >= MAR, .08, 0), .03, .97)
app_on = (rng.random((N, M)) < e2) & active
sessions = np.where(app_on, 1 + rng.poisson(eng[:, None] * 12 * np.where(mgrid >= MAR, 1.3, 1.0)), 0)
bill_views = rng.binomial(sessions, .30)
speed_tests = np.where(app_on, rng.poisson(.3 + 1.6 * (util > 85) + outage / 300), 0)
help_chats = np.where(app_on, rng.poisson(.08 + outage / 450), 0)
addon_views = np.where(app_on, rng.poisson(eng[:, None] * 1.6), 0)
earned = np.where((mgrid >= JAN) & active, 100 * on_time + 10 * np.minimum(sessions, 15) + 50 * (ten_m > 24), 0)
redeemer = rng.random(N) < np.clip(eng * 1.2 - .42, 0, .8)
redeemed = np.zeros((N, M), int); balance = np.zeros((N, M), int); bal = np.zeros(N, int)
for m in range(M):
    bal = bal + earned[:, m]
    go = redeemer & app_on[:, m] & (m > JAN) & (bal >= 300) & (rng.random(N) < .30)
    r = np.where(go, (bal * rng.uniform(.4, .9, N) // 100 * 100).astype(int), 0)
    redeemed[:, m] = r; bal = bal - r; balance[:, m] = bal
ever_redeemed = redeemed.sum(1) > 0

# ---------------------------------------------------------------- flatten helpers
cid = np.array([f'C{i + 1:05d}' for i in range(N)])
ii, mm = np.where(active)
def base_cols(i, m=None):
    d = dict(customer_id=cid[i], city=A.city.values[i], area_name=A.area_name.values[i], segment=np.array(SEG)[seg[i]])
    return d

def fact(extra):
    d = dict(month=MONTHS[mm], customer_id=cid[ii], city=A.city.values[ii], area_name=A.area_name.values[ii],
             segment=np.array(SEG)[seg[ii]], package_name=np.array([p[1] for p in PKG])[pkg_m[ii, mm]])
    d.update(extra)
    return pd.DataFrame(d)

internet_usage = fact(dict(speed_mbps=speed_m[ii, mm], download_gb=np.round(gb[ii, mm], 1), upload_gb=np.round((gb * up_ratio)[ii, mm], 1),
                           avg_devices=devices[ii, mm].astype(int), speed_utilisation_pct=np.round(util[ii, mm], 1),
                           measured_speed_mbps=np.round(measured[ii, mm], 1), outage_minutes=np.round(outage[ii, mm]).astype(int)))
status = np.where(unpaid, 'Unpaid', np.where(late, 'Late', 'On time'))
paid_date = [None if unpaid[i, m] else due[m] + pd.Timedelta(days=int(days_late[i, m]) - int(rng.integers(0, 12) if not late[i, m] else 0)) for i, m in zip(ii, mm)]
payments = fact(dict(payment_method=np.array(METHODS)[method[ii]], amount_billed_idr=billed[ii, mm].astype(int), due_date=due[mm],
                     paid_date=paid_date, days_late=days_late[ii, mm].astype(int), payment_status=status[ii, mm]))
payments.insert(0, 'payment_id', [f'PAY{k + 1:06d}' for k in range(len(payments))])
payments = payments.rename(columns={'month': 'billing_month'})
app_usage = fact(dict(app_active=np.where(app_on[ii, mm], 'Yes', 'No'), app_sessions=sessions[ii, mm], bill_views=bill_views[ii, mm],
                      speed_tests=speed_tests[ii, mm], help_chats=help_chats[ii, mm], addon_page_views=addon_views[ii, mm],
                      points_earned=earned[ii, mm], points_redeemed=redeemed[ii, mm], points_balance=balance[ii, mm]))

mix = np.vstack([rng.dirichlet(CAT_ALPHA[s] * .6) for s in seg])
beh = []
for k, cat in enumerate(CATS):
    g = gb * mix[:, k][:, None] * rng.lognormal(0, .12, (N, M))
    if cat == 'Gaming': g = g * (1 + .015 * mgrid)
    beh.append(g)
beh = np.stack(beh); beh = beh / beh.sum(0, keepdims=True) * gb
bf = []
for k, cat in enumerate(CATS):
    f = fact(dict(usage_category=cat, usage_gb=np.round(beh[k][ii, mm], 1)))
    bf.append(f.drop(columns=['package_name']))
internet_behavior = pd.concat(bf, ignore_index=True).sort_values(['month', 'customer_id', 'usage_category']).reset_index(drop=True)

# ---------------------------------------------------------------- TV
tv_channels = pd.DataFrame([dict(channel_id=f'CH{k + 1:02d}', channel_name=c[0], genre=c[1], tier=c[2],
                                 pack_price_idr=dict(PACKS).get(c[2], TV_BASIC), licence_cost_per_subscriber_idr=c[5]) for k, c in enumerate(CHANNELS)])
tvfac = rng.lognormal(0, .45, N)
pack_names = [p[0] for p in PACKS]
tvf = []; tv_hours_total = np.zeros((N, M)); genre_hours = {}
for k, (name, genre, tier, base, zero, lic) in enumerate(CHANNELS):
    if tier == 'Basic TV':
        access = tv[:, None] & active
    else:
        access = pack_on[:, pack_names.index(tier), :]
    mean = base * np.array(AFF[genre])[seg][:, None] * tvfac[:, None] * np.ones((1, M))
    if name == 'Drama Asia Lite': mean = mean * np.where(pack[:, 3], .3, 1.0)[:, None]
    if name == 'Kids Corner Lite': mean = mean * np.where(pack[:, 1], .3, 1.0)[:, None]
    if genre == 'Sports':
        boost = np.ones(M); boost[[JUN, JUL]] = 3.4 if name in ('Arena Bola', 'Arena Lite') else 1.0
        mean = mean * boost[None, :]
        if tier == 'Sports Pack': mean = mean * np.where(wc[:, None] & (mgrid > JUL), .06, 1.0)
    if tier == 'Movies Pack': mean = mean * np.where(movie_dead, 0.0, 1.0)[:, None]
    h = rng.gamma(1.4, np.maximum(mean, 1e-9) / 1.4)
    never = rng.random(N) < zero * .6
    h = np.where(never[:, None] | (rng.random((N, M)) < zero * .4), 0, h)
    h = np.round(np.where(access, h, 0), 1)
    tv_hours_total += h; genre_hours[genre] = genre_hours.get(genre, 0) + h
    a_i, a_m = np.where(access)
    tvf.append(pd.DataFrame(dict(month=MONTHS[a_m], customer_id=cid[a_i], city=A.city.values[a_i], segment=np.array(SEG)[seg[a_i]],
                                 channel_name=name, genre=genre, tier=tier, watch_hours=h[a_i, a_m])))
tv_viewing = pd.concat(tvf, ignore_index=True).sort_values(['month', 'customer_id', 'channel_name']).reset_index(drop=True)

# ---------------------------------------------------------------- sales
S = []
def sale(date, i, cat, prod, ch, promo, amount, billing):
    S.append(dict(transaction_date=pd.Timestamp(date).normalize(), customer_id=cid[i], city=A.city.values[i], area_name=A.area_name.values[i],
                  segment=SEG[seg[i]], product_category=cat, product_name=prod, sales_channel=ch, promo_code=promo,
                  amount_idr=int(amount), billing_type=billing, sales_cost_idr=SALES_COST[ch]))
def rday(m, lo=0): return MONTHS[m] + pd.Timedelta(days=int(rng.integers(lo, 27)))
SELF_P = [.48, .22, .14, .06, .10]
booster_n = np.zeros(N, int); addon_rev = np.zeros(N)
for i in range(N):
    if join_m[i] >= 0:
        promo = 'RAMADAN26' if (join_date[i] >= pd.Timestamp('2026-02-18') and join_date[i] <= pd.Timestamp('2026-03-20')) else 'NONE'
        sale(join_date[i], i, 'New installation', PKG[pkg0[i]][1], CHANNELS_SALES[acq[i]], promo, fee_m[i, join_m[i]] + (0 if promo != 'NONE' else 150000), 'Monthly')
    if up_m[i] < 99:
        m = up_m[i]; promo = 'UPGRADE50' if m == SEP else 'NONE'
        sale(rday(m), i, 'Speed upgrade', PKG[pkg_cur[i]][1], rng.choice(CHANNELS_SALES, p=[.50, .18, .24, .03, .05]), promo, fee_m[i, m] * (.5 if promo != 'NONE' else 1), 'Monthly')
    for k in range(4):
        if pack[i, k] and pack_start[i, k] >= 1 and pack_start[i, k] > join_m[i]:
            m = pack_start[i, k]; iswc = (k == 0 and wc[i])
            ch = rng.choice(CHANNELS_SALES, p=[.20, .10, .15, .45, .10]) if k == 2 else rng.choice(CHANNELS_SALES, p=SELF_P)
            sale(rday(m, 0) if not iswc else MONTHS[m] + pd.Timedelta(days=int(rng.integers(0, 12))), i, 'TV pack', PACKS[k][0], ch, 'WORLDCUP26' if iswc else 'NONE', PACKS[k][1] * (.5 if iswc else 1), 'Monthly')
            addon_rev[i] += PACKS[k][1] * (M - m)
    if static_ip[i] and static_start[i] >= 1 and static_start[i] > join_m[i]:
        sale(rday(static_start[i]), i, 'Add-on', 'Static IP', rng.choice(CHANNELS_SALES, p=[.3, .2, .4, .05, .05]), 'NONE', 100000, 'Monthly')
        addon_rev[i] += 100000 * (M - static_start[i])
    lam = np.where(util[i] > 85, .30 * (1.5 if seg[i] == 2 else 1.0), .015) * active[i]
    for m in np.where(rng.poisson(lam) > 0)[0]:
        for _ in range(1 + int(rng.random() < .25)):
            sale(rday(m), i, 'Add-on', 'Speed Booster 7 days', rng.choice(CHANNELS_SALES, p=[.82, .12, .06, 0, 0]), 'NONE', 35000, 'One-off')
            booster_n[i] += 1; addon_rev[i] += 35000
    if rng.random() < (.10 if (home[i] in (0, 3) and hsize[i] >= 4) else .03):
        m = rng.choice(np.where(active[i])[0])
        sale(rday(m), i, 'Add-on', 'Wi-Fi Extender', rng.choice(CHANNELS_SALES, p=[.35, .20, .15, .10, .20]), 'NONE', 350000, 'One-off')
        addon_rev[i] += 350000
sales = pd.DataFrame(S).sort_values('transaction_date').reset_index(drop=True)
sales.insert(0, 'transaction_id', [f'TRX{k + 1:06d}' for k in range(len(sales))])
sales.insert(2, 'month', sales.transaction_date.values.astype('datetime64[M]').astype('datetime64[ns]'))

# ---------------------------------------------------------------- feedback
TOPICS = ['Speed', 'Outage', 'Billing & payment', 'Installation', 'TV content', 'Customer service', 'Price', 'App']
SOURCES = ['App review', 'Call centre', 'Social media', 'Email survey']
TXT = {
    ('Speed', 'Negative'): ['Slow every evening, video calls keep freezing.', 'Game lag is terrible at night.'],
    ('Speed', 'Positive'): ['Fast and stable, even with everyone online.', 'Speed is exactly as promised.'],
    ('Outage', 'Negative'): ['No internet for hours and no clear update.', 'Connection dropped again today.'],
    ('Outage', 'Positive'): ['The outage was fixed faster than I expected.', 'Thanks for the quick repair.'],
    ('Billing & payment', 'Negative'): ['Paying at the minimarket is a hassle, I always forget.', 'My bill is confusing.'],
    ('Billing & payment', 'Positive'): ['Auto-debit makes paying effortless.', 'Easy to pay from the app.'],
    ('Installation', 'Negative'): ['Waited two weeks for the technician.', 'Installation was rescheduled twice.'],
    ('Installation', 'Positive'): ['Installed in two days, neat cabling.', 'Technician was friendly and quick.'],
    ('TV content', 'Negative'): ['I pay for the movie pack but there is nothing I want to watch.', 'Too many repeats.'],
    ('TV content', 'Positive'): ['The drama channels are the best part.', 'My kids love the cartoon channels.'],
    ('Customer service', 'Negative'): ['Forty minutes on hold at the call centre.', 'Agent could not solve my problem.'],
    ('Customer service', 'Positive'): ['Chat support solved it in minutes.', 'Helpful and polite agent.'],
    ('Price', 'Negative'): ['The price went up and nothing got better.', 'Too expensive for this speed.'],
    ('Price', 'Positive'): ['Good value for a family.', 'Fair price for the speed.'],
    ('App', 'Negative'): ['The app is confusing and slow.', 'I cannot find where to redeem my points.'],
    ('App', 'Positive'): ['The new app is much easier to use.', 'Love checking my usage in the app.'],
}
NEU = ['It is fine, nothing special.', 'Works most of the time.', 'Average experience.']
TP = np.array([.20, .09, .13, .07, .11, .16, .09, .15])
F = []
def fb(i, m, topic, pneg, src=None):
    src = src or rng.choice(SOURCES, p=[.38, .25, .2, .17])
    if topic == 'Customer service' and src == 'Call centre': pneg = min(.95, pneg + .2)
    u = rng.random()
    sent = 'Negative' if u < pneg else ('Neutral' if u < pneg + (1 - pneg) * .28 else 'Positive')
    score = {'Negative': rng.uniform(-1, -.3), 'Neutral': rng.uniform(-.3, .3), 'Positive': rng.uniform(.3, 1)}[sent]
    rating = int(np.clip(round(3 + 2 * score + rng.normal(0, .3)), 1, 5))
    text = rng.choice(NEU) if sent == 'Neutral' else rng.choice(TXT[(topic, sent)])
    F.append(dict(feedback_date=rday(m), customer_id=cid[i], city=A.city.values[i], area_name=A.area_name.values[i], segment=SEG[seg[i]],
                  package_name=PKG[pkg_m[i, m]][1], source=src, topic=topic, sentiment=sent, sentiment_score=round(score, 2), rating=rating, comment=text))
atype = A.area_type.values
nfb = rng.poisson(.20, (N, M)) * active
for i, m in zip(*np.where(nfb > 0)):
    for _ in range(nfb[i, m]):
        p = TP.copy()
        if util[i, m] > 85: p[0] *= 2.5
        if tv[i]: p[4] *= 2.2
        else: p[4] = 0
        p[3] *= .2
        if m < JAN + 0: pass
        t = rng.choice(8, p=p / p.sum()); topic = TOPICS[t]
        pneg = [.25 + .55 * (util[i, m] > 85), .85, .30 + .30 * (method[i] == 3) - .2 * (method[i] == 0), .22 + .50 * (atype[i] == 'build'),
                .18 + .40 * (pack[i, 2] and movie_dead[i]) - .10 * pack[i, 3], .45, .50, .62 if m < MAR else .18][t]
        fb(i, m, topic, float(np.clip(pneg, .03, .95)))
for i in np.where(join_m >= 0)[0]:
    if rng.random() < .75: fb(i, join_m[i], 'Installation', .72 if atype[i] == 'build' else .18)
for i in np.where(mk & active[:, FEB])[0]:
    for _ in range(rng.poisson(1.1)): fb(i, FEB, 'Outage', .93)
for i in np.where(md & active[:, JUL])[0]:
    for _ in range(rng.poisson(.5)): fb(i, JUL, 'Outage', .88)
for m in (AUG, SEP):
    for i in np.where(active[:, m] & (pkg_m[:, m] <= 1))[0]:
        if rng.random() < (.30 if m == AUG else .16): fb(i, m, 'Price', .92)
for m in range(JAN + 1, M):
    for i in np.where(active[:, m] & (balance[:, m] > 900) & ~redeemer)[0]:
        if rng.random() < .035: fb(i, m, 'App', .85, 'App review'); F[-1]['comment'] = 'I cannot find where to redeem my points.'
feedback = pd.DataFrame(F).sort_values('feedback_date').reset_index(drop=True)
feedback.insert(0, 'feedback_id', [f'FB{k + 1:05d}' for k in range(len(feedback))])
feedback.insert(2, 'month', feedback.feedback_date.values.astype('datetime64[M]').astype('datetime64[ns]'))

# ---------------------------------------------------------------- customers + customer_360
nact = active.sum(1)
def avg(x): return (x * active).sum(1) / nact
neg_fb = feedback[feedback.sentiment == 'Negative'].groupby('customer_id').size().reindex(cid).fillna(0).values
fb_n = feedback.groupby('customer_id').size().reindex(cid).fillna(0).values
fb_sc = feedback.groupby('customer_id').sentiment_score.mean().reindex(cid).fillna(0).values
late_n = (late & active).sum(1)
out_tot = (outage * active).sum(1)
sat = (8.75 - 1.7 * (avg(util) > 85) - np.minimum(2.2, out_tot / 800) - 1.2 * (late_n / nact) - .4 * ((pkg_cur <= 1) & (nact > 3))
       + .9 * ever_redeemed - .8 * ((tenure > 36) & ~ever_redeemed) - .5 * (pack[:, 2] & movie_dead) + .5 * pack[:, 3] - .25 * neg_fb + rng.normal(0, 1.15, N))
nps = np.clip(np.round(sat), 0, 10).astype(int)
nps_group = np.where(nps >= 9, 'Promoter', np.where(nps >= 7, 'Passive', 'Detractor'))
tb = np.where(tenure < 12, '0-11 months', np.where(tenure < 24, '12-23 months', np.where(tenure < 36, '24-35 months', '36+ months')))
customers = pd.DataFrame(dict(
    customer_id=cid, join_date=join_date, tenure_months=tenure, tenure_band=tb, area_id=A.area_id.values, area_name=A.area_name.values,
    city=A.city.values, province=A.province.values, island=A.island.values, segment=np.array(SEG)[seg], age_group=np.array(AGE)[age],
    home_type=np.array(HOME)[home], household_size=hsize, package_name=np.array([p[1] for p in PKG])[pkg_cur], speed_mbps=SPEED[pkg_cur],
    monthly_fee_idr=np.array([p[3] for p in PKG])[pkg_cur], acquisition_channel=np.array(CHANNELS_SALES)[acq], payment_method=np.array(METHODS)[method],
    tv_subscriber=np.where(tv, 'Yes', 'No'), sports_pack=np.where(pack[:, 0], 'Yes', 'No'), kids_pack=np.where(pack[:, 1], 'Yes', 'No'),
    movies_pack=np.where(pack[:, 2], 'Yes', 'No'), drama_pack=np.where(pack[:, 3], 'Yes', 'No')))
gnames = list(genre_hours); gstack = np.stack([(genre_hours[g] * active).sum(1) for g in gnames])
top_genre = np.where(tv & (gstack.sum(0) > 0), np.array(gnames)[gstack.argmax(0)], 'No TV')
top_cat = np.array(CATS)[(beh * active[None]).sum(2).argmax(0)]
c360 = customers[['customer_id', 'city', 'area_name', 'island', 'segment', 'age_group', 'home_type', 'tenure_months', 'tenure_band', 'package_name',
                  'speed_mbps', 'monthly_fee_idr', 'payment_method', 'tv_subscriber']].copy()
c360 = c360.assign(
    months_in_data=nact, avg_download_gb=np.round(avg(gb), 1), avg_speed_utilisation_pct=np.round(avg(util), 1), top_usage_category=top_cat,
    total_outage_minutes=np.round(out_tot).astype(int), bills=nact, late_or_unpaid_bills=late_n, on_time_rate_pct=np.round((1 - late_n / nact) * 100, 1),
    total_billed_idr=(billed * active).sum(1).astype(int), avg_app_sessions=np.round(avg(sessions), 1), app_active_months=app_on.sum(1),
    points_earned=earned.sum(1), points_redeemed=redeemed.sum(1), points_balance=balance[:, -1], booster_purchases=booster_n,
    addon_revenue_idr=addon_rev.astype(int), tv_packs_owned=pack.sum(1), avg_tv_hours=np.round(np.where(tv, avg(tv_hours_total), 0), 1),
    top_tv_genre=top_genre, feedback_count=fb_n.astype(int), negative_feedback_count=neg_fb.astype(int), avg_sentiment_score=np.round(fb_sc, 2),
    nps_score=nps, nps_group=nps_group)

products = pd.DataFrame(
    [dict(product_name=p[1], product_category='Internet package', speed_mbps=p[2], price_idr=p[3], billing_type='Monthly', note='Price before 1 Aug 2026: Rp %s' % format(PKG_OLD[k], ',') if k in PKG_OLD else '-') for k, p in enumerate(PKG)]
    + [dict(product_name='Basic TV', product_category='TV', speed_mbps=0, price_idr=TV_BASIC, billing_type='Monthly', note='10 basic channels; needed before any TV pack')]
    + [dict(product_name=p[0], product_category='TV pack', speed_mbps=0, price_idr=p[1], billing_type='Monthly', note='3 channels') for p in PACKS]
    + [dict(product_name=a[0], product_category='Add-on', speed_mbps=0, price_idr=a[1], billing_type=a[2], note='-') for a in ADDONS])
products['note'] = products.note.fillna('-')

events = pd.DataFrame([
    ('2026-01-05', '2026-01-05', 'Loyalty', 'Arunika Poin launched', 'Customers earn points for paying on time, using the app and staying longer than two years. Points can be swapped for rewards in the app.', 'app_usage, customer_360'),
    ('2026-02-10', '2026-02-14', 'Network', 'Undersea cable fault, Makassar', 'Customers in Makassar had long outages for several days.', 'internet_usage, feedback'),
    ('2026-02-18', '2026-03-20', 'Promo', 'RAMADAN26: free installation', 'New customers who joined in this window paid no installation fee.', 'sales, area_monthly'),
    ('2026-03-02', '2026-03-02', 'Product', 'MyArunika app redesign', 'New home screen, usage view and simpler payment flow.', 'app_usage, feedback'),
    ('2026-03-20', '2026-04-05', 'Season', 'Lebaran holiday', 'Many customers travel; household spending is high.', 'payments'),
    ('2026-06-01', '2026-07-19', 'Promo', 'WORLDCUP26: Sports Pack half price for the first month', 'Promotion around the football World Cup (11 June to 19 July).', 'sales, tv_viewing'),
    ('2026-07-08', '2026-07-21', 'Network', 'Storm damage, Medan', 'Repeated short outages in Medan.', 'internet_usage, feedback'),
    ('2026-08-01', '2026-08-01', 'Pricing', 'Price adjustment', 'Start 30 rose from Rp 260,000 to Rp 275,000 and Family 50 from Rp 325,000 to Rp 345,000. Other packages unchanged.', 'payments, feedback'),
    ('2026-09-01', '2026-09-30', 'Promo', 'UPGRADE50: half price on the first month of a faster package', 'Offered to every customer in the app and by the call centre.', 'sales, internet_usage'),
    ('2025-11-01', '2026-10-31', 'Roll-out', 'Fiber roll-out in newer areas', 'Arunika is still building its network. Coverage is uneven and next year\'s build budget covers only a few areas.', 'areas, area_monthly'),
], columns=['start_date', 'end_date', 'type', 'event', 'what_happened', 'tables_to_check'])
events['start_date'] = pd.to_datetime(events.start_date); events['end_date'] = pd.to_datetime(events.end_date)
events = events.sort_values('start_date').reset_index(drop=True)
areas_out = areas.drop(columns=['area_type'])

# ---------------------------------------------------------------- dictionary
T, N_, D, P = 'Text', 'Number', 'Date', 'Number (percent 0-100)'
C = 'Currency (IDR)'
COMMON = {'customer_id': (T, 'Customer key. Joins every customer-level sheet.'), 'city': (T, 'City of the customer or area. Set type to Geo > City for maps.'),
          'area_name': (T, 'Area inside the city (city + Pusat/Utara/Selatan/Timur).'), 'segment': (T, 'Customer segment: one of five household types.'),
          'package_name': (T, 'Internet package in that month (or today, in customers).'), 'month': (D, 'First day of the month the row describes.'),
          'island': (T, 'Island group.'), 'speed_mbps': (N_, 'Package speed in megabits per second (Mbps).')}
DICT = {
 'customers': ('One row per customer. Who they are and what they buy today.', 'customer', {
    'join_date': (D, 'Day the service was installed.'), 'tenure_months': (N_, 'Months as a customer on 31 Oct 2026.'), 'tenure_band': (T, 'Tenure grouped into four bands.'),
    'area_id': (T, 'Area key. Joins areas and area_monthly.'), 'province': (T, 'Province.'), 'age_group': (T, 'Age band of the account holder.'),
    'home_type': (T, 'Type of home.'), 'household_size': (N_, 'People living in the home.'), 'monthly_fee_idr': (C, 'Current monthly price of the internet package.'),
    'acquisition_channel': (T, 'Channel that sold the first installation.'), 'payment_method': (T, 'How the customer pays the monthly bill.'),
    'tv_subscriber': (T, 'Yes if the customer has Arunika TV (Basic TV).'), 'sports_pack': (T, 'Owns the Sports Pack (Yes/No).'), 'kids_pack': (T, 'Owns the Kids Pack (Yes/No).'),
    'movies_pack': (T, 'Owns the Movies Pack (Yes/No).'), 'drama_pack': (T, 'Owns the Asian Drama Pack (Yes/No).')}),
 'customer_360': ('One row per customer with twelve-month totals from every other sheet. The easiest sheet to start with.', 'customer', {
    'age_group': (T, 'Age band of the account holder.'), 'home_type': (T, 'Type of home.'), 'tenure_months': (N_, 'Months as a customer on 31 Oct 2026.'), 'tenure_band': (T, 'Tenure grouped into four bands.'),
    'monthly_fee_idr': (C, 'Current monthly price of the internet package.'), 'payment_method': (T, 'How the customer pays the monthly bill.'), 'tv_subscriber': (T, 'Yes if the customer has Arunika TV.'),
    'months_in_data': (N_, 'Months the customer was active in the 12-month window.'), 'avg_download_gb': (N_, 'Average gigabytes downloaded per month.'),
    'avg_speed_utilisation_pct': (P, 'Average share of the package speed used at the busiest hour. Above 85 means the line is saturated.'),
    'top_usage_category': (T, 'Category with the most data used.'), 'total_outage_minutes': (N_, 'Minutes without service over the window.'), 'bills': (N_, 'Bills issued.'),
    'late_or_unpaid_bills': (N_, 'Bills paid after the due date or not paid.'), 'on_time_rate_pct': (P, 'Share of bills paid on time.'), 'total_billed_idr': (C, 'Total billed over the window.'),
    'avg_app_sessions': (N_, 'Average MyArunika app sessions per month.'), 'app_active_months': (N_, 'Months with at least one app session.'), 'points_earned': (N_, 'Arunika Poin earned since January 2026.'),
    'points_redeemed': (N_, 'Points swapped for rewards.'), 'points_balance': (N_, 'Unused points on 31 Oct 2026.'), 'booster_purchases': (N_, 'Number of Speed Booster purchases.'),
    'addon_revenue_idr': (C, 'Revenue from TV packs and add-ons bought inside the window.'), 'tv_packs_owned': (N_, 'Number of TV packs owned (0 to 4).'), 'avg_tv_hours': (N_, 'Average TV hours watched per month (0 if no TV).'),
    'top_tv_genre': (T, 'Most-watched genre, or "No TV".'), 'feedback_count': (N_, 'Feedback records from this customer.'), 'negative_feedback_count': (N_, 'Feedback records with negative sentiment.'),
    'avg_sentiment_score': (N_, 'Average sentiment score, -1 to 1 (0 if no feedback).'), 'nps_score': (N_, 'Answer to "How likely are you to recommend us?" (0 to 10), October 2026 survey.'),
    'nps_group': (T, 'Promoter (9-10), Passive (7-8) or Detractor (0-6).')}),
 'areas': ('One row per area (48 areas in 12 cities). Network coverage and demand, for the whole market, not only the 1,500 sampled customers.', 'area', {
    'area_id': (T, 'Area key.'), 'province': (T, 'Province.'), 'latitude': (N_, 'Latitude of the area centre.'), 'longitude': (N_, 'Longitude of the area centre.'),
    'income_tier': (T, 'Typical household income: Low, Mid or High.'), 'households': (N_, 'Households living in the area.'), 'homes_passed': (N_, 'Households the fiber network already reaches.'),
    'uncovered_households': (N_, 'Households not yet reachable (households - homes_passed).'), 'subscribers': (N_, 'Paying customers in the area.'),
    'coverage_pct': (P, 'homes_passed / households.'), 'penetration_pct': (P, 'subscribers / homes_passed: how well we sell where we already have network.'),
    'waitlist_requests_12m': (N_, 'People who asked for service at an address the network does not reach, last 12 months.'), 'competitor_count': (N_, 'Other fiber providers in the area.'),
    'build_cost_per_home_idr': (C, 'Cost to extend the network to one more household.'), 'avg_install_wait_days': (N_, 'Average days from order to installation.'),
    'geo_point': (T, 'Latitude,longitude in one field. Set type to Geo > Latitude, Longitude for bubble maps.')}),
 'area_monthly': ('One row per area per month. Demand signals over time.', 'area x month', {
    'area_id': (T, 'Area key.'), 'coverage_checks': (N_, 'People who checked on the website whether their address can be served.'),
    'waitlist_requests': (N_, 'Checks at addresses the network does not reach.'), 'new_installs': (N_, 'New installations in the area.')}),
 'products': ('Price list: internet packages, TV and add-ons.', 'product', {
    'product_name': (T, 'Product name, as used in sales.'), 'product_category': (T, 'Internet package, TV, TV pack or Add-on.'), 'price_idr': (C, 'Current price.'),
    'billing_type': (T, 'Monthly or One-off.'), 'note': (T, 'Extra information.')}),
 'sales': ('One row per sale inside the window: new installations, speed upgrades, TV packs and add-ons.', 'transaction', {
    'transaction_id': (T, 'Sale key.'), 'transaction_date': (D, 'Day of the sale.'), 'product_category': (T, 'New installation, Speed upgrade, TV pack or Add-on.'),
    'product_name': (T, 'Product sold.'), 'sales_channel': (T, 'Channel that closed the sale.'), 'promo_code': (T, 'Promotion used, or NONE.'),
    'amount_idr': (C, 'Amount charged at the sale (first month or one-off price, after promotion).'), 'billing_type': (T, 'Monthly or One-off.'),
    'sales_cost_idr': (C, 'What the sale cost the company in that channel (commission, agent time).')}),
 'payments': ('One row per monthly bill.', 'customer x month', {
    'payment_id': (T, 'Bill key.'), 'billing_month': (D, 'Month the bill covers.'), 'payment_method': (T, 'How the customer pays.'), 'amount_billed_idr': (C, 'Internet package + TV + monthly add-ons.'),
    'due_date': (D, 'Bills are due on the 20th.'), 'paid_date': (D, 'Day the bill was paid. Empty only when payment_status is Unpaid.'), 'days_late': (N_, 'Days after the due date (0 = on time).'),
    'payment_status': (T, 'On time, Late or Unpaid.')}),
 'app_usage': ('One row per customer per month in the MyArunika app, including loyalty points.', 'customer x month', {
    'app_active': (T, 'Yes if the customer opened the app that month.'), 'app_sessions': (N_, 'Times the app was opened.'), 'bill_views': (N_, 'Times the bill screen was opened.'),
    'speed_tests': (N_, 'Speed tests run from the app.'), 'help_chats': (N_, 'Chats started with support.'), 'addon_page_views': (N_, 'Visits to the add-on shop.'),
    'points_earned': (N_, 'Arunika Poin earned that month (0 before January 2026).'), 'points_redeemed': (N_, 'Points swapped for rewards that month.'), 'points_balance': (N_, 'Unused points at month end.')}),
 'internet_usage': ('One row per customer per month: how much the line is used and how well it works.', 'customer x month', {
    'download_gb': (N_, 'Gigabytes downloaded.'), 'upload_gb': (N_, 'Gigabytes uploaded.'), 'avg_devices': (N_, 'Devices connected on a typical day.'),
    'speed_utilisation_pct': (P, 'Share of the package speed used at the busiest hour. Above 85 means the line is saturated.'),
    'measured_speed_mbps': (N_, 'Average speed measured at the modem.'), 'outage_minutes': (N_, 'Minutes without service.')}),
 'internet_behavior': ('One row per customer per month per usage category. The six categories add up to download_gb.', 'customer x month x category', {
    'usage_category': (T, 'Video streaming, Gaming, Social media, Work & video calls, Education, Shopping & other.'), 'usage_gb': (N_, 'Gigabytes used in that category.')}),
 'tv_channels': ('One row per TV channel (22 channels).', 'channel', {
    'channel_id': (T, 'Channel key.'), 'channel_name': (T, 'Channel name. Joins tv_viewing.'), 'genre': (T, 'Genre.'), 'tier': (T, 'Basic TV (every TV subscriber) or the pack that unlocks the channel.'),
    'pack_price_idr': (C, 'Monthly price of the tier.'), 'licence_cost_per_subscriber_idr': (C, 'What the company pays the channel owner each month for every customer who can watch it.')}),
 'tv_viewing': ('One row per TV customer, month and channel they can watch. Zero hours are kept, so "pays but does not watch" is visible.', 'customer x month x channel', {
    'channel_name': (T, 'Channel name.'), 'genre': (T, 'Genre.'), 'tier': (T, 'Basic TV or the pack that unlocks the channel.'), 'watch_hours': (N_, 'Hours watched that month.')}),
 'feedback': ('One row per piece of customer feedback, already labelled by a sentiment model.', 'feedback record', {
    'feedback_id': (T, 'Feedback key.'), 'feedback_date': (D, 'Day received.'), 'source': (T, 'App review, Call centre, Social media or Email survey.'), 'topic': (T, 'What the feedback is about (eight topics).'),
    'sentiment': (T, 'Negative, Neutral or Positive.'), 'sentiment_score': (N_, 'Model score from -1 (very negative) to 1 (very positive).'), 'rating': (N_, 'Star rating, 1 to 5.'), 'comment': (T, 'Short text of the feedback.')}),
 'events': ('What happened at the company during the twelve months. Use it to explain jumps in the charts.', 'event', {
    'start_date': (D, 'First day.'), 'end_date': (D, 'Last day.'), 'type': (T, 'Kind of event.'), 'event': (T, 'Short name.'), 'what_happened': (T, 'Description.'), 'tables_to_check': (T, 'Sheets where the effect may show.')}),
}
TABLES = dict(customers=customers, customer_360=c360, areas=areas_out, area_monthly=area_monthly, products=products, sales=sales, payments=payments,
              app_usage=app_usage, internet_usage=internet_usage, internet_behavior=internet_behavior, tv_channels=tv_channels, tv_viewing=tv_viewing,
              feedback=feedback, events=events)
dict_rows = []
for t, df in TABLES.items():
    desc, grain, cols = DICT[t]
    for c in df.columns:
        typ, d = cols.get(c) or COMMON[c]
        ex = df[c].dropna().iloc[0]
        ex = ex.strftime('%Y-%m-%d') if hasattr(ex, 'strftime') else str(ex)
        dict_rows.append(dict(sheet=t, column=c, type=typ, description=d, example=ex))
    assert not df.drop(columns=['paid_date'], errors='ignore').isna().any().any(), t
dictionary = pd.DataFrame(dict_rows)

# ---------------------------------------------------------------- write workbook
out = os.path.join(ROOT, 'data', 'Arunika_Fiber_BI_Case.xlsx')
readme = [
    ['Arunika Fiber — BI case dataset'], [''],
    ['Arunika Fiber is a FICTIONAL home-internet company. Every number in this workbook was generated for teaching. No real customer or company data.'],
    ['Window: November 2025 to October 2026 (12 months). 1,500 sampled customers, 48 areas in 12 Indonesian cities.'],
    ['The data is already clean: no duplicates, no missing values (except paid_date for unpaid bills), consistent names. Spend your time on questions, not on cleaning.'], [''],
    ['How to use it'],
    ['1. Read the sheet "events" and "data_dictionary" first.'],
    ['2. Start with "customer_360" (one row per customer). Move to the detailed sheets when you need time trends.'],
    ['3. In Google Data Studio, add each sheet you need as its own data source (Google Sheets connector, "Use first row as headers").'],
    ['4. Most detailed sheets already carry city, segment and package, so you rarely need to blend data.'], [''],
    ['Sheet', 'Rows', 'One row is', 'What it holds'],
] + [[t, len(df), DICT[t][1], DICT[t][0]] for t, df in TABLES.items()]
with pd.ExcelWriter(out, engine='xlsxwriter', datetime_format='yyyy-mm-dd', date_format='yyyy-mm-dd') as xw:
    wb = xw.book
    head = wb.add_format({'bold': True, 'bg_color': '#141B2A', 'font_color': '#FFFFFF', 'border': 0})
    bold = wb.add_format({'bold': True}); big = wb.add_format({'bold': True, 'font_size': 16, 'font_color': '#EB1D0A'})
    ws = wb.add_worksheet('README'); xw.sheets['README'] = ws
    for r, row in enumerate(readme):
        for c, v in enumerate(row):
            ws.write(r, c, v, big if r == 0 else (bold if (r in (6, 12)) else None))
    ws.set_column(0, 0, 22); ws.set_column(1, 1, 10); ws.set_column(2, 2, 30); ws.set_column(3, 3, 110)
    def put(name, df):
        df.to_excel(xw, sheet_name=name, index=False)
        w = xw.sheets[name]
        for c, col in enumerate(df.columns):
            w.write(0, c, col, head)
            width = max(len(col), int(df[col].astype(str).str.len().quantile(.95)) if len(df) else 0) + 2
            w.set_column(c, c, min(width, 60))
        w.freeze_panes(1, 0)
    put('data_dictionary', dictionary); put('events', TABLES['events'])
    for t, df in TABLES.items():
        if t != 'events': put(t, df)

meta = dict(tables=[dict(name=t, rows=len(df), grain=DICT[t][1], desc=DICT[t][0],
                         cols=[dict(c=r['column'], t=r['type'], d=r['description'], e=r['example']) for r in dict_rows if r['sheet'] == t]) for t, df in TABLES.items()],
            events=[dict(start=r.start_date.strftime('%Y-%m-%d'), end=r.end_date.strftime('%Y-%m-%d'), type=r.type, event=r.event, what=r.what_happened) for r in events.itertuples()])
with open(os.path.join(ROOT, 'data', 'case_meta.js'), 'w', encoding='utf-8', newline='\n') as f:
    f.write('window.CASE_META = ' + json.dumps(meta, ensure_ascii=False) + ';\n')

print({t: len(df) for t, df in TABLES.items()})
print('cells', sum(df.size for df in TABLES.values()), 'xlsx MB', round(os.path.getsize(out) / 1e6, 2))
