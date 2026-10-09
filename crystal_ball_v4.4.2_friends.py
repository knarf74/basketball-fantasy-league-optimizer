# NBA Fantasy Optimizer v3.18.3 — Risk Calibration
# Start with: python nba_fantasy_v2.4.py
from functools import lru_cache
import sys, subprocess, importlib.util, re, sqlite3, json, math, unicodedata, time
from html.parser import HTMLParser
from html import escape
import csv
from io import StringIO
from pathlib import Path
from datetime import datetime, timezone


def ensure(pkg, import_name=None):
    name = import_name or pkg
    if importlib.util.find_spec(name) is None:
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', '--user', pkg])

for pkg, mod in [('streamlit','streamlit'),('requests','requests'),('pulp','pulp'),('openpyxl','openpyxl'),]:
    ensure(pkg, mod)

import requests, pulp, streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

BUILD_ID='4.4.2-FRIENDS-EDITION'
API='https://nbafantasy.nba.com/api'
BASE=Path(__file__).resolve().parent
# Friends Edition: shared historical data may be bundled, but user writes are ephemeral.
# Cloud filesystem is not durable; profiles live in each visitor's Streamlit session.
import tempfile, shutil, gzip
CLOUD_FRIENDS_EDITION=True
DB=Path(tempfile.gettempdir())/'crystal_ball_friends_history_v442.sqlite'
_bundled_db=BASE/'nba_fantasy_v2.4.sqlite'
_bundled_gzip=BASE/'nba_fantasy_v2.4.sqlite.gz'
if not DB.exists():
    if _bundled_gzip.exists():
        with gzip.open(_bundled_gzip, 'rb') as source, open(DB, 'wb') as target:
            shutil.copyfileobj(source, target)
    elif _bundled_db.exists():
        shutil.copy2(_bundled_db, DB)
TIMEOUT=30
HORIZON_GWS=4
GW_DECAY=.85
FREE_TRANSFERS=2
EXTRA_TRANSFER_COST=100
CURRENT_SEASON_STARTED=False
CURRENT_FINISHED_GAMEDAYS=0
NO_HISTORY_MARKET_VALUE=2.20
RISK_STRATEGY='Balanced'


def num(v,d=0.0):
    try:return float(v)
    except:return d

def get_json(path):
    r=requests.get(f"{API}/{path.lstrip('/')}",timeout=TIMEOUT,headers={'User-Agent':'Mozilla/5.0'})
    r.raise_for_status(); return r.json()

def parse_gw(name):
    m=re.search(r'Gameweek\s+(\d+)',name or '',re.I); return int(m.group(1)) if m else None

def day_no(name):
    m=re.search(r'Day\s+(\d+)',name or '',re.I); return int(m.group(1)) if m else None

def db_connect():
    con=sqlite3.connect(DB)
    con.executescript('''
    CREATE TABLE IF NOT EXISTS snapshots(
      captured_at TEXT NOT NULL, player_id INTEGER NOT NULL, name TEXT, team_id INTEGER,
      price REAL, ppg REAL, form REAL, minutes REAL, total_points REAL, status TEXT,
      chance REAL, selected REAL, PRIMARY KEY(captured_at,player_id));
    CREATE TABLE IF NOT EXISTS runs(
      run_id INTEGER PRIMARY KEY AUTOINCREMENT, created_at TEXT, mode TEXT, budget_cap REAL,
      objective REAL, roster_points REAL, captain_bonus REAL, spent REAL, roster_json TEXT);
    CREATE TABLE IF NOT EXISTS settings(key TEXT PRIMARY KEY,value TEXT);
    CREATE TABLE IF NOT EXISTS durability_seasons(
      season_end INTEGER NOT NULL, player_key TEXT NOT NULL, player_name TEXT,
      games INTEGER, source TEXT, captured_at TEXT,
      PRIMARY KEY(season_end,player_key));
    CREATE TABLE IF NOT EXISTS fantasy_seasons(
      season_end INTEGER NOT NULL, player_key TEXT NOT NULL, player_name TEXT,
      games INTEGER, fantasy_ppg REAL, source TEXT, captured_at TEXT,
      PRIMARY KEY(season_end,player_key));
    CREATE TABLE IF NOT EXISTS team_profiles(
      profile_name TEXT PRIMARY KEY, roster_json TEXT NOT NULL, excluded_json TEXT NOT NULL DEFAULT '[]',
      budget REAL, free_transfers INTEGER, allow_paid INTEGER, max_test_transfers INTEGER, updated_at TEXT);
    ''')
    return con


def _session_profiles():
    """Private to this browser session; intentionally not durable across restarts."""
    if '_cb_friends_profiles' not in st.session_state:
        st.session_state['_cb_friends_profiles'] = {}
    return st.session_state['_cb_friends_profiles']


def list_team_profiles():
    return sorted(_session_profiles(), key=str.casefold)


def load_team_profile(name):
    item = _session_profiles().get(name)
    return dict(item) if item is not None else None


def save_team_profile(name, roster, excluded, budget, free_transfers, allow_paid, max_test_transfers):
    name = (name or '').strip()
    if not name:
        raise ValueError('Geef het teamprofiel een naam.')
    _session_profiles()[name] = {
        'roster': list(roster), 'excluded': list(excluded), 'budget': float(budget),
        'free_transfers': int(free_transfers), 'allow_paid': bool(allow_paid),
        'max_test_transfers': int(max_test_transfers),
    }


def delete_team_profile(name):
    _session_profiles().pop(name, None)


def save_snapshot(players,currency):
    stamp=datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
    con=db_connect()
    rows=[]
    for p in players:
        rows.append((stamp,p['id'],p.get('web_name'),p.get('team'),num(p.get('now_cost'))/currency,
                     num(p.get('points_per_game'))/currency,num(p.get('form'))/currency,num(p.get('minutes')),
                     num(p.get('total_points'))/currency,p.get('status'),p.get('chance_of_playing_next_round'),
                     num(p.get('selected_by_percent'))))
    con.executemany('INSERT OR IGNORE INTO snapshots VALUES(?,?,?,?,?,?,?,?,?,?,?,?)',rows)
    con.commit(); con.close(); _snapshot_statistics_index.cache_clear(); return stamp

@lru_cache(maxsize=1)
def _snapshot_statistics_index():
    """Read snapshot availability and counts in one SQLite pass, not 550 queries."""
    con=db_connect()
    rows=con.execute('SELECT player_id,status,chance FROM snapshots').fetchall()
    con.close()
    sums={}; counts={}
    for pid,status,chance in rows:
        counts[pid]=counts.get(pid,0)+1
        if chance is not None: value=max(0,min(1,num(chance)/100))
        elif status in ('i','u','s'): value=.15
        elif status=='d': value=.75
        else: value=1.0
        sums[pid]=sums.get(pid,0.0)+value
    return {pid:(n,sums[pid]/n) for pid,n in counts.items()}

def snapshot_count(pid):
    return _snapshot_statistics_index().get(pid,(0,None))[0]

def player_key(name):
    x=unicodedata.normalize('NFKD',name or '').encode('ascii','ignore').decode('ascii').lower()
    return re.sub(r'[^a-z0-9]+','',x)

def full_name(p):
    name=(str(p.get('first_name') or '')+' '+str(p.get('second_name') or '')).strip()
    return name or str(p.get('web_name') or '')

class _BRefTotalsParser(HTMLParser):
    def __init__(self):
        super().__init__(); self.in_tr=False; self.stat=None; self.buf=[]; self.row={}; self.rows=[]
    def handle_starttag(self,tag,attrs):
        a=dict(attrs)
        if tag=="tr": self.in_tr=True; self.row={}
        elif self.in_tr and tag in ("th","td"): self.stat=a.get("data-stat"); self.buf=[]
    def handle_data(self,data):
        if self.in_tr and self.stat is not None: self.buf.append(data)
    def handle_endtag(self,tag):
        if self.in_tr and tag in ("th","td") and self.stat is not None:
            self.row[self.stat]="".join(self.buf).strip(); self.stat=None; self.buf=[]
        elif tag=="tr" and self.in_tr:
            if self.row:self.rows.append(self.row)
            self.in_tr=False; self.row={}

def _store_durability_rows(season_end, rows, source):
    stamp=datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
    best={}
    for name,games in rows:
        key=player_key(name)
        if not key: continue
        try:g=int(float(games))
        except (ValueError,TypeError):continue
        if key not in best or g>best[key][1]: best[key]=(name,g)
    if not best:return 0
    con=db_connect()
    con.executemany('INSERT OR REPLACE INTO durability_seasons VALUES(?,?,?,?,?,?)',
        [(season_end,k,n,g,source,stamp) for k,(n,g) in best.items()])
    con.commit();con.close()
    return len(best)


def _stat_num(row,*names):
    for name in names:
        v=row.get(name)
        if v not in (None,''):
            try:return float(str(v).replace(',','').strip())
            except (ValueError,TypeError):pass
    return None

def _fantasy_from_row(row):
    name=(row.get('player') or row.get('name_display') or row.get('Name') or
          row.get('Player') or row.get('player_name') or row.get('PLAYER_NAME') or '').replace('*','').strip()
    gp=_stat_num(row,'games','g','GP','G','gp')
    pts=_stat_num(row,'pts','PTS','points','Points','ppg','PTS_per_game')
    reb=_stat_num(row,'trb','TRB','reb','REB','total_rebounds','Total Rebounds')
    ast=_stat_num(row,'ast','AST','assists','Assists')
    blk=_stat_num(row,'blk','BLK','blocks','Blocks')
    stl=_stat_num(row,'stl','STL','steals','Steals')
    if not name or not gp or gp<=0 or None in (pts,reb,ast,blk,stl):
        return None
    # Detect totals vs per-game rows. BRef totals are totals; common CSVs may be per-game.
    # If points is implausibly large for per-game, treat all five fields as totals.
    totals = pts > 100
    if totals:
        fp=(pts+reb+2*ast+3*blk+3*stl)/gp
    else:
        fp=pts+reb+2*ast+3*blk+3*stl
    if not (0 <= fp <= 100): return None
    return name,int(round(gp)),float(fp)

def _store_fantasy_rows(season_end, rows, source):
    stamp=datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
    best={}
    for row in rows:
        rec=_fantasy_from_row(row)
        if not rec: continue
        name,gp,fp=rec; key=player_key(name)
        if key and (key not in best or gp>best[key][1]): best[key]=(name,gp,fp)
    if not best:return 0
    con=db_connect()
    con.executemany('INSERT OR REPLACE INTO fantasy_seasons VALUES(?,?,?,?,?,?,?)',
        [(season_end,k,n,g,fp,source,stamp) for k,(n,g,fp) in best.items()])
    con.commit();con.close()
    return len(best)

def _load_csv_dicts(url):
    r=requests.get(url,timeout=TIMEOUT,headers={'User-Agent':'Mozilla/5.0'});r.raise_for_status()
    text=r.text.lstrip('\ufeff')
    try:dialect=csv.Sniffer().sniff(text[:4096],delimiters=',;\t')
    except csv.Error:dialect=csv.excel
    return list(csv.DictReader(StringIO(text),dialect=dialect))

def _load_bref_stat_rows(year):
    url=f'https://www.basketball-reference.com/leagues/NBA_{year}_totals.html'
    r=requests.get(url,timeout=TIMEOUT,headers={'User-Agent':'Mozilla/5.0'});r.raise_for_status()
    parser=_BRefTotalsParser();parser.feed(r.text)
    return parser.rows,url

@lru_cache(maxsize=1)
def _fantasy_history_index():
    con=db_connect()
    rows=con.execute('SELECT player_key,season_end,fantasy_ppg FROM fantasy_seasons ORDER BY player_key,season_end').fetchall()
    con.close()
    index={}
    for key,year,ppg in rows:
        index.setdefault(key,[]).append((int(year),float(ppg)))
    return {key:tuple(values) for key,values in index.items()}

@lru_cache(maxsize=4096)
def fantasy_history_by_key(key):
    return _fantasy_history_index().get(key,())

def three_year_fantasy(p):
    return dict(fantasy_history_by_key(player_key(full_name(p))))

def multiseason_baseline(p, api_ppg):
    hist=three_year_fantasy(p)
    weights={2024:.20,2025:.30,2026:.50}
    vals=[(y,hist[y],weights[y]) for y in (2024,2025,2026) if y in hist and hist[y]>0]
    if not vals:
        return {'baseline':api_ppg,'history':hist,'coverage':0,'source':'Fantasy API fallback'}
    denom=sum(w for _,_,w in vals)
    weighted=sum(v*w for _,v,w in vals)/denom
    # Blend the independent historical reconstruction with the official Fantasy API anchor.
    # Full 3Y history gets 70% reconstructed / 30% API; partial history is more conservative.
    hist_share={1:.35,2:.55,3:.70}[len(vals)]
    baseline=hist_share*weighted+(1-hist_share)*api_ppg if api_ppg>0 else weighted
    return {'baseline':baseline,'history':hist,'coverage':len(vals),'source':f'{len(vals)}Y reconstructed fantasy'}


def local_durability_status():
    """Offline-only startup: never wait for external historical websites."""
    con=db_connect()
    counts=dict(con.execute('SELECT season_end,COUNT(*) FROM durability_seasons GROUP BY season_end').fetchall())
    con.close()
    return {year:{'ok':counts.get(year,0)>=100,'count':counts.get(year,0),
                  'source':'SQLite lokaal' if counts.get(year,0) else 'ontbreekt (handmatig verversen)'}
            for year in (2024,2025,2026)}

def local_fantasy_status():
    con=db_connect()
    counts=dict(con.execute('SELECT season_end,COUNT(*) FROM fantasy_seasons GROUP BY season_end').fetchall())
    con.close()
    return {year:{'ok':counts.get(year,0)>=100,'count':counts.get(year,0),
                  'source':'SQLite lokaal' if counts.get(year,0) else 'ontbreekt (handmatig verversen)'}
            for year in (2024,2025,2026)}

def load_three_year_fantasy():
    csv_sources={
      2024:'https://raw.githubusercontent.com/EPKeep32/STA551/main/2023-2024%20NBA%20Player%20Stats%20-%20Regular.csv',
      2025:'https://raw.githubusercontent.com/Data-Dunkers/data/main/NBA/player/nba_player_stats_2024-2025.csv',
      2026:'https://raw.githubusercontent.com/Data-Dunkers/data/main/NBA/player/nba_player_stats_2025-2026.csv',
    }
    status={}
    for yr in (2024,2025,2026):
        con=db_connect(); n=con.execute('SELECT COUNT(*) FROM fantasy_seasons WHERE season_end=?',(yr,)).fetchone()[0]; con.close()
        if n>=100:
            status[yr]={'ok':True,'count':n,'source':'SQLite cache'}; continue
        loaded=0; errors=[]; used=''
        try:
            rows,url=_load_bref_stat_rows(yr)
            loaded=_store_fantasy_rows(yr,rows,url)
            if loaded:used='Basketball-Reference totals'
        except Exception as exc: errors.append('BRef: '+str(exc)[:100])
        if loaded<100:
            try:
                rows=_load_csv_dicts(csv_sources[yr])
                n2=_store_fantasy_rows(yr,rows,csv_sources[yr])
                if n2>loaded: loaded=n2; used='CSV fallback'
            except Exception as exc: errors.append('CSV: '+str(exc)[:100])
        status[yr]={'ok':loaded>=100,'count':loaded,'source':used or 'niet geladen','error':' | '.join(errors)}
    fantasy_history_by_key.cache_clear()
    _fantasy_history_index.cache_clear()
    return status

def _load_csv_url(url):
    r=requests.get(url,timeout=TIMEOUT,headers={'User-Agent':'Mozilla/5.0'});r.raise_for_status()
    text=r.text.lstrip('\ufeff')
    sample=text[:4096]
    try:
        dialect=csv.Sniffer().sniff(sample,delimiters=',;\t')
    except csv.Error:
        dialect=csv.excel
    reader=csv.DictReader(StringIO(text),dialect=dialect)
    rows=[]
    for row in reader:
        # Support common schemas: Name/GP, Player/G, player/gp.
        name=row.get('Name') or row.get('Player') or row.get('player') or row.get('PLAYER_NAME')
        games=row.get('GP') or row.get('G') or row.get('gp') or row.get('games')
        if name and games not in (None,''):rows.append((name,games))
    return rows

def _load_bref_html(year):
    url=f'https://www.basketball-reference.com/leagues/NBA_{year}_totals.html'
    r=requests.get(url,timeout=TIMEOUT,headers={'User-Agent':'Mozilla/5.0'});r.raise_for_status()
    parser=_BRefTotalsParser();parser.feed(r.text)
    rows=[]
    for row in parser.rows:
        name=(row.get('player') or row.get('name_display') or '').replace('*','').strip()
        games=row.get('games') or row.get('g')
        if name and games not in (None,''):rows.append((name,games))
    return rows,url

GP_2024_VERIFIED = {
    'nikola jokic':79,'shai gilgeous alexander':75,'luka doncic':70,
    'jalen brunson':77,'tyrese maxey':70,'lebron james':71,
    'devin booker':68,'paolo banchero':80,'victor wembanyama':71,
    'joel embiid':39,'anthony davis':76,'jayson tatum':74,
    'stephen curry':74,'zion williamson':70,'karl anthony towns':62
}

def validate_2024_rows(rows):
    """Reject a 2023-24 source if known players do not match regular-season GP."""
    got={player_key(n):int(float(g)) for n,g in rows if n and str(g).strip()}
    checks={player_key(k):v for k,v in GP_2024_VERIFIED.items()}
    matched=0
    for k,expected in checks.items():
        if k in got:
            matched+=1
            if got[k] != expected:
                return False,f'{k}: bron={got[k]}, verwacht={expected}'
    return matched>=10,f'{matched} sanity checks'

def load_three_year_durability(players, force=False):
    """Load real regular-season games played. Returns a status dict; never silently labels
    missing history as 100% durability.
    """
    seasons=(2024,2025,2026)
    status={}
    # Public CSV mirrors are fallbacks for environments where BRef blocks scripted requests.
    csv_sources={
      2024:'https://raw.githubusercontent.com/EPKeep32/STA551/main/2023-2024%20NBA%20Player%20Stats%20-%20Regular.csv',
      2025:'https://raw.githubusercontent.com/Data-Dunkers/data/main/NBA/player/nba_player_stats_2024-2025.csv',
      2026:'https://raw.githubusercontent.com/Data-Dunkers/data/main/NBA/player/nba_player_stats_2025-2026.csv',
    }
    for yr in seasons:
        con=db_connect()
        count=con.execute('SELECT COUNT(*) FROM durability_seasons WHERE season_end=?',(yr,)).fetchone()[0]
        con.close()
        if count>=100 and not force:
            if yr==2024:
                con=db_connect()
                cached=con.execute('SELECT player_name,games FROM durability_seasons WHERE season_end=2024').fetchall()
                con.close()
                ok,why=validate_2024_rows(cached)
                if ok:
                    status[yr]={'ok':True,'count':count,'source':'SQLite cache (sanity checked)'}
                    continue
                con=db_connect(); con.execute('DELETE FROM durability_seasons WHERE season_end=2024'); con.commit(); con.close()
            else:
                status[yr]={'ok':True,'count':count,'source':'SQLite cache'}
                continue

        loaded=0; used=''; errors=[]
        # Try BRef first because it is the direct season totals source.
        try:
            rows,url=_load_bref_html(yr)
            if yr==2024:
                ok,why=validate_2024_rows(rows)
                if not ok: raise ValueError('2024 sanity check mislukt: '+why)
            loaded=_store_durability_rows(yr,rows,url)
            if loaded:used='Basketball-Reference'
        except Exception as exc:
            errors.append('BRef: '+str(exc)[:90])

        if not loaded:
            try:
                rows=_load_csv_url(csv_sources[yr])
                if yr==2024:
                    ok,why=validate_2024_rows(rows)
                    if not ok: raise ValueError('2024 sanity check mislukt: '+why)
                loaded=_store_durability_rows(yr,rows,csv_sources[yr])
                if loaded:used='CSV fallback'
            except Exception as exc:
                errors.append('CSV: '+str(exc)[:90])


        if yr==2024 and not loaded:
            verified_rows=[(name,gp) for name,gp in GP_2024_VERIFIED.items()]
            loaded=_store_durability_rows(yr,verified_rows,'verified 2023-24 sanity subset')
            used='verified subset'
        status[yr]={'ok':bool(loaded),'count':loaded,'source':used or 'niet geladen',
                    'error':' | '.join(errors)}
    return status

@lru_cache(maxsize=1)
def _durability_history_index():
    con=db_connect()
    rows=con.execute('SELECT player_key,season_end,games FROM durability_seasons ORDER BY player_key,season_end').fetchall()
    con.close()
    index={}
    for key,year,games in rows:
        index.setdefault(key,[]).append((int(year),int(games)))
    return {key:tuple(values) for key,values in index.items()}

@lru_cache(maxsize=4096)
def history_by_key(key):
    return _durability_history_index().get(key,())

def three_year_games(p):
    return dict(history_by_key(player_key(full_name(p))))

def historical_snapshot_availability(pid):
    n,average=_snapshot_statistics_index().get(pid,(0,None))
    return average if n>=5 else None

def current_availability(p):
    chance=p.get('chance_of_playing_next_round')
    if chance is not None:return max(0,min(1,num(chance)/100))
    if p.get('status') in ('i','u','s'):return .15
    if p.get('status')=='d':return .75
    return 1.0

def durability_details(p):
    hist=three_year_games(p)
    years=[y for y in (2024,2025,2026) if y in hist]
    if years:
        weights={2024:.20,2025:.30,2026:.50}
        denom=sum(weights[y] for y in years)
        weighted=sum(weights[y]*min(1.0,hist[y]/82.0) for y in years)/denom
        low=sum(hist[y] < 60 for y in years)
        severe=sum(hist[y] < 45 for y in years)
        # Repeated low seasons matter more than one isolated injury season.
        years_count=len(hist)
        recurrence=1.0 if years_count<2 else max(.70,1.0-.07*max(0,low-1)-.05*max(0,severe-1))
        raw_score=max(.35,min(1.0,weighted*recurrence))
        confidence={1:.35,2:.70,3:1.0}.get(years_count,1.0)
        score=1.0-confidence*(1.0-raw_score)
        source=f'{years_count}Y games played'
    else:
        score=None; low=0; severe=0; recurrence=1.0; years_count=0; source='geen historie'
    observed=historical_snapshot_availability(p['id'])
    if observed is not None and score is not None:
        n=snapshot_count(p['id']); w=min(.20,n/40*.20)
        score=score*(1-w)+observed*w
        source+=' + local snapshots'
    return {'score':score,'games':hist,'years_count':years_count,
            'low_seasons':int(low),'severe_seasons':int(severe),
            'recurrence':recurrence,'source':source}

def durability_rate(p):
    score=durability_details(p)['score']
    return 1.0 if score is None else score

def current_season_games_sample(p):
    # Before the first completed Gameday, bootstrap player totals can still contain
    # preseason/previous-season values. Never interpret those as current-season games.
    if not CURRENT_SEASON_STARTED:
        return 0.0
    ppg_raw=num(p.get('points_per_game'))
    total_raw=num(p.get('total_points'))
    if ppg_raw<=0 or total_raw<=0:
        return 0.0
    games=total_raw/ppg_raw
    # A player cannot have played more games than completed league Gamedays. The cap
    # also protects us against stale/mixed API fields around the season rollover.
    cap=max(0.0,float(CURRENT_FINISHED_GAMEDAYS))
    return max(0.0,min(games,cap))

def configure_no_history_market_prior(players, currency):
    """Estimate a preseason FP/G-per-$m prior from established NBA players.

    This gives rookies / 0Y players a non-zero, market-informed starting point instead
    of treating missing NBA history as a projection of zero. It uses only players with
    at least two reconstructed seasons and winsorises the ratio before taking a median.
    """
    global NO_HISTORY_MARKET_VALUE
    vals=[]
    for q in players:
        hist=three_year_fantasy(q)
        if len([v for v in hist.values() if v and v>0]) < 2:
            continue
        price=max(.1,num(q.get('now_cost'))/currency)
        api=num(q.get('points_per_game'))/currency
        ms=multiseason_baseline(q,api)
        ratio=ms.get('baseline',0)/price if price else 0
        if 1.0 <= ratio <= 4.0:
            vals.append(ratio)
    if vals:
        vals.sort(); n=len(vals)
        med=vals[n//2] if n%2 else (vals[n//2-1]+vals[n//2])/2
        NO_HISTORY_MARKET_VALUE=max(1.6,min(2.8,med))
    return NO_HISTORY_MARKET_VALUE

def no_history_prior(p,currency,api_ppg):
    """Preseason prior for players with zero reconstructed NBA seasons.

    Price is used as a market expectation proxy. If the Fantasy API already supplies
    a positive PPG anchor, blend it with the market prior; otherwise use a discounted
    market prior. This is uncertainty-aware evidence, not a claim that a rookie will
    actually score this amount.
    """
    price=max(.1,num(p.get('now_cost'))/currency)
    market=price*NO_HISTORY_MARKET_VALUE
    if api_ppg>0:
        baseline=.70*api_ppg+.30*market
        source='0Y API + market prior'
    else:
        baseline=.78*market
        source='0Y market prior (discounted)'
    return {'baseline':max(0.0,baseline),'market':market,'source':source}

def empirical_baseline_calibration(p, baseline):
    """Backtest-calibrated preseason correction.

    Uses only information known before the target season: prior fantasy level and
    prior games played. Coefficients come from the v3.10.1 walk-forward diagnostics.
    This is deliberately bounded and is applied BEFORE role/durability/availability.
    Players with no reconstructed history are not boosted by this calibration.
    """
    if CURRENT_SEASON_STARTED or baseline <= 0:
        return {'raw':baseline,'adjusted':baseline,'delta':0.0,'level_delta':0.0,'gp_delta':0.0,'prev_gp':None}
    seasons=preseason_context(p).get('history_seasons',0)
    if seasons <= 0:
        return {'raw':baseline,'adjusted':baseline,'delta':0.0,'level_delta':0.0,'gp_delta':0.0,'prev_gp':None}
    # v3.10.1 bias was prediction - actual, so correction is -bias.
    if baseline < 20: level_delta=2.85
    elif baseline < 35: level_delta=-0.51
    elif baseline < 50: level_delta=-0.85
    else: level_delta=-2.60
    games=durability_details(p).get('games',{})
    prev_gp=games.get(2026)
    if prev_gp in (None,''):
        gp_delta=0.0
    elif float(prev_gp) < 30: gp_delta=3.59
    elif float(prev_gp) < 60: gp_delta=0.97
    else: gp_delta=0.41
    # Level is the stronger signal; GP adds a smaller availability/role-regime signal.
    delta=.65*level_delta + .35*gp_delta
    delta=max(-2.5,min(2.5,delta))
    adjusted=max(0.0,baseline+delta)
    return {'raw':baseline,'adjusted':adjusted,'delta':delta,'level_delta':level_delta,'gp_delta':gp_delta,'prev_gp':prev_gp}

def adaptive_projection_components(p,currency):
    # Season PPG is the stable anchor; current Fantasy `form` is the recent signal.
    # The clock is current-season games, never historical minutes.
    api_ppg=num(p.get('points_per_game'))/currency
    ms=multiseason_baseline(p,api_ppg)
    nohist=None
    if not CURRENT_SEASON_STARTED and ms.get('coverage',0)==0:
        nohist=no_history_prior(p,currency,api_ppg)
        ms=dict(ms)
        ms['baseline']=nohist['baseline']
        ms['source']=nohist['source']
    ppg=api_ppg if CURRENT_SEASON_STARTED else ms['baseline']
    form=num(p.get('form'))/currency
    mins=num(p.get('minutes'))
    games_est=current_season_games_sample(p)
    live_sample=min(1.0,games_est/10.0)
    recent_weight=(0.35*live_sample) if form>0 and ppg>0 else 0.0
    historical_weight=1.0-recent_weight
    blended=(historical_weight*ppg + recent_weight*form) if ppg>0 else max(0,form)
    # Reliability remains separate from the adaptive clock. Once the season is live,
    # current minutes may still guard tiny samples; preseason old minutes cannot age form.
    reliability=1.0 if not CURRENT_SEASON_STARTED else (min(1,max(.55,mins/600)) if blended>0 else 0)
    uncalibrated_base=max(0,blended*reliability)
    cal=empirical_baseline_calibration(p,uncalibrated_base)
    base=cal['adjusted']
    return {'ppg':ppg,'api_ppg':api_ppg,'multi':ms,'form':form,'minutes':mins,'games_est':games_est,
            'historical_weight':historical_weight,'recent_weight':recent_weight,
            'blended':blended,'reliability':reliability,'base_uncalibrated':uncalibrated_base,
            'calibration':cal,'no_history_prior':nohist,'base':base}

def base_projection(p,currency):
    return adaptive_projection_components(p,currency)['base']

def durability_factor(p):
    # Strong enough to matter, but not a player-specific hardcoded penalty.
    return .55+.45*durability_rate(p)

def b2b_factor(p): return .40+.60*durability_rate(p)

def risk_upside_components(p):
    """Generic uncertainty envelope; no player-specific overrides.
    Confidence combines availability, durability, role stability and historical coverage.
    Floor/upside are scenario bands around the model projection, not probability guarantees.
    """
    avail=current_availability(p)
    dd=durability_details(p)
    dur=dd.get('score')
    dur=0.72 if dur is None else max(0.0,min(1.0,float(dur)))
    years=int(dd.get('years_count') or 0)
    hist={0:.62,1:.76,2:.89,3:1.0}.get(min(3,years),1.0)
    role=max(.90,min(1.05,num(p.get('_role_factor',1.0),1.0)))
    role_conf=max(.65,1.0-abs(role-1.0)*3.0)
    confidence=max(.45,min(1.0,.30*avail+.30*dur+.20*hist+.20*role_conf))
    uncertainty=1.0-confidence
    # Conservative but useful display bands. Low-confidence players get wider ranges.
    floor_factor=max(.62,.86-.32*uncertainty)
    upside_factor=min(1.32,1.10+.30*uncertainty)
    return {'confidence':confidence,'uncertainty':uncertainty,
            'floor_factor':floor_factor,'upside_factor':upside_factor}

def strategy_factor(p):
    # v3.18.3: strategy changes player RANKING, not merely every projection by a tiny amount.
    # Safe leans toward the downside band; Aggressive leans toward the upside band.
    # The blend is deliberately bounded so uncertainty can matter without overwhelming talent.
    r=risk_upside_components(p)
    if RISK_STRATEGY=='Safe':
        return max(.78, min(1.0, .35 + .65*r['floor_factor']))
    if RISK_STRATEGY=='Aggressive':
        return min(1.18, max(1.0, .45 + .55*r['upside_factor']))
    return 1.0

def central_projection_from_strategy_value(p, value):
    # _proj is prepared under the active strategy. Undo that multiplier for transparent display.
    f=strategy_factor(p)
    return num(value)/(f if f else 1.0)

def risk_scenario_values(p, central):
    r=risk_upside_components(p)
    return {'floor':central*r['floor_factor'], 'expected':central,
            'ceiling':central*r['upside_factor'], 'confidence':r['confidence']}

def team_fixture_map(events):
    out={}
    for e in events:
        teams=set()
        for f in get_json(f"fixtures/?event={e['id']}"):
            if f.get('team_h') is not None: teams.add(int(f['team_h']))
            if f.get('team_a') is not None: teams.add(int(f['team_a']))
        out[e['id']]=teams
    return out

def build_b2b(events,active):
    result=set(); byg={}
    for e in events:byg.setdefault(parse_gw(e.get('name')),[]).append(e)
    for evs in byg.values():
        evs=sorted(evs,key=lambda e:(day_no(e.get('name')) or 999,e['id']))
        for a,b in zip(evs,evs[1:]):
            if day_no(a.get('name')) is not None and day_no(b.get('name'))==day_no(a.get('name'))+1:
                for tid in active[a['id']] & active[b['id']]:result.add((int(tid),b['id']))
    return result

def event_projection(p,e,first_gw,currency,b2b):
    # Hot path: optimizer calls this thousands of times, so reuse per-player scalars.
    avail=p.get('_availability')
    if avail is None: avail=current_availability(p)
    base=p.get('_base_projection')
    if base is None: base=base_projection(p,currency)
    risk=p.get('_durability_factor')
    if risk is None: risk=durability_factor(p)
    b2bf=p.get('_b2b_factor')
    if b2bf is None: b2bf=b2b_factor(p)
    egw=parse_gw(e.get('name')) or first_gw
    cur=avail if egw<=first_gw else avail+(1-avail)*min(1,.35*(egw-first_gw))
    if (int(p['team']),e['id']) in b2b: risk*=b2bf
    return base*cur*risk*num(p.get('_role_factor',1.0),1.0)*strategy_factor(p)


def role_adjustment_map(players,currency):
    """Same role formula, but evaluate each player's baseline only once.

    The old nested rivals loop repeatedly recalculated expensive historical
    projections for every teammate (quadratic work per NBA team).
    """
    by_team={}
    for p in players:
        if p.get('can_select',True) and not p.get('removed',False):
            by_team.setdefault(int(p['team']),[]).append(p)
    out={}
    for team,grp in by_team.items():
        eligible=[]
        for q in grp:
            if num(q.get('minutes'))>=900:
                v=base_projection(q,currency)
                if v>=24:
                    eligible.append((q['id'],v))
        eligible.sort(key=lambda item:item[1],reverse=True)
        # The top three rivals excluding oneself are always in the top four.
        top_four=eligible[:4]
        for p in grp:
            rival_load=sum(v for qid,v in [x for x in top_four if x[0]!=p['id']][:3])
            factor=max(.90,min(1.05,1.0-(rival_load-75.0)*.0018))
            out[p['id']]={'factor':factor,'rival_load':rival_load}
    return out

def shortlist_candidates(players, events, active, currency, first_gw, b2b, keep_ids=None, per_pos=24, role_map=None):
    """Keep a diverse, high-value candidate pool before building the expensive MILP."""
    keep_ids=set(keep_ids or [])
    pool=[dict(p) for p in players if p.get('can_select',True) and not p.get('removed',False)]
    for p in pool:
        role=(role_map or {}).get(p['id'],{}).get('factor',1.0)
        p['_role']=role
        p['_role_factor']=role
        p['_base_projection']=base_projection(p,currency)
        p['_availability']=current_availability(p)
        dr=durability_rate(p)
        p['_durability_rate']=dr
        p['_durability_factor']=.55+.45*dr
        p['_b2b_factor']=.40+.60*dr
        p['_proj']=uncertainty_regression(p,p['_base_projection']*preseason_context(p)['confidence']*p['_availability']*p['_durability_factor']*role,currency)['adjusted']*strategy_factor(p)
        price=max(.1,num(p.get('now_cost'))/currency)
        p['_value']=p['_proj']/price
        p['_sched']=sum(event_projection(p,e,first_gw,currency,b2b)
                        for e in events if int(p['team']) in active[e['id']])
    chosen=set(keep_ids)
    for tid in sorted({int(p['element_type']) for p in pool}):
        grp=[p for p in pool if int(p['element_type'])==tid]
        for key in ('_proj','_value','_sched'):
            chosen.update(p['id'] for p in sorted(grp,key=lambda q:q[key],reverse=True)[:per_pos])
        # Preserve useful cheap options across price bands.
        for lo,hi in ((0,6),(6,9),(9,12),(12,16),(16,999)):
            band=[p for p in grp if lo <= num(p.get('now_cost'))/currency < hi]
            chosen.update(p['id'] for p in sorted(band,key=lambda q:q['_proj'],reverse=True)[:5])
    return [p for p in pool if p['id'] in chosen]

def optimize(players,teams,types,rules,events,active,currency,weights,first_gw,b2b,budget_cap,
             locked=None,excluded=None,current_roster=None,free_transfers=2,allow_paid=False,
             fixed_candidate_ids=None, min_keep_ids=None, min_keep_count=None, max_transfers=None, exact_transfers=None):
    t_all=time.perf_counter()
    locked=set(locked or []); excluded=set(excluded or []); current_roster=set(current_roster or [])
    min_keep_ids=set(min_keep_ids or [])
    keep=locked|current_roster|min_keep_ids
    t0=time.perf_counter()
    role_map=role_adjustment_map(players,currency)
    if fixed_candidate_ids is None:
        candidates=[p for p in shortlist_candidates(players,events,active,currency,first_gw,b2b,keep_ids=keep,role_map=role_map)
                    if p['id'] not in excluded]
    else:
        fixed_candidate_ids=set(fixed_candidate_ids)
        candidates=[dict(p) for p in players
                    if p.get('can_select',True) and not p.get('removed',False)
                    and p['id'] in fixed_candidate_ids and p['id'] not in excluded]
        # Rebuild the same cached scalar features for the fixed pool.
        for p in candidates:
            role=role_map.get(p['id'],{}).get('factor',1.0)
            p['_role']=role; p['_role_factor']=role
            p['_base_projection']=base_projection(p,currency)
            p['_availability']=current_availability(p)
            dr=durability_rate(p)
            p['_durability_rate']=dr
            p['_durability_factor']=.55+.45*dr
            p['_b2b_factor']=.40+.60*dr
            p['_proj']=uncertainty_regression(p,p['_base_projection']*preseason_context(p)['confidence']*p['_availability']*p['_durability_factor']*role,currency)['adjusted']*strategy_factor(p)
    t_short=time.perf_counter()-t0
    ids={p['id'] for p in candidates}
    t0=time.perf_counter()
    if not locked<=ids: raise ValueError('Een locked speler is niet selecteerbaar of ook uitgesloten.')
    for p in candidates:
        p['_role_factor']=role_map.get(p['id'],{}).get('factor',1.0)
        p['_proj']=uncertainty_regression(p,p['_base_projection']*preseason_context(p)['confidence']*p['_availability']*p['_durability_factor']*p['_role_factor'],currency)['adjusted']*strategy_factor(p)
    ep={(p['id'],e['id']):event_projection(p,e,first_gw,currency,b2b)
        for p in candidates for e in events}
    prob=pulp.LpProblem('NBAFantasy24',pulp.LpMaximize)
    x={p['id']:pulp.LpVariable(f'x_{p["id"]}',0,1,'Binary') for p in candidates}
    y={(p['id'],e['id']):pulp.LpVariable(f'y_{p["id"]}_{e["id"]}',0,1,'Binary') for p in candidates for e in events}
    z={(p['id'],e['id']):pulp.LpVariable(f'z_{p["id"]}_{e["id"]}',0,1,'Binary') for p in candidates for e in events}
    expected=pulp.lpSum(ep[p['id'],e['id']]*(int(p['team']) in active[e['id']])*weights[e['id']]*y[p['id'],e['id']] for p in candidates for e in events)
    capbonus=pulp.lpSum(ep[p['id'],e['id']]*(int(p['team']) in active[e['id']])*weights[e['id']]*z[p['id'],e['id']] for p in candidates for e in events)
    transfer_penalty=0
    if current_roster:
        incoming=pulp.lpSum(x[p['id']] for p in candidates if p['id'] not in current_roster)
        if exact_transfers is not None:
            prob += incoming == int(exact_transfers)
        elif max_transfers is not None:
            prob += incoming <= int(max_transfers)
        if allow_paid:
            paid=pulp.LpVariable('paid_transfers',lowBound=0,cat='Integer')
            prob += paid >= incoming-int(free_transfers)
            transfer_penalty=100*paid
        else:
            prob += incoming <= int(free_transfers)
    prob += expected+capbonus-transfer_penalty
    prob += pulp.lpSum(x.values())==int(rules.get('squad_squadsize',10))
    prob += pulp.lpSum(int(p.get('now_cost',0))*x[p['id']] for p in candidates) <= int(round(budget_cap*currency))
    for pid in locked: prob += x[pid]==1
    if min_keep_ids and min_keep_count is not None:
        usable_keep=min_keep_ids & ids
        if len(usable_keep) < int(min_keep_count):
            raise ValueError('Te weinig hand-picked spelers beschikbaar voor deze repair-run.')
        prob += pulp.lpSum(x[pid] for pid in usable_keep) >= int(min_keep_count)
    for t in types:
        tid=int(t['id'])
        if t.get('squad_select') is not None: prob += pulp.lpSum(x[p['id']] for p in candidates if int(p['element_type'])==tid)==int(t['squad_select'])
    for tid in teams: prob += pulp.lpSum(x[p['id']] for p in candidates if int(p['team'])==int(tid))<=int(rules.get('squad_team_limit',2))
    for e in events:
        eid=e['id']; prob += pulp.lpSum(y[p['id'],eid] for p in candidates)==int(rules.get('squad_squadplay',5))
        for p in candidates:
            prob += y[p['id'],eid] <= x[p['id']]; prob += z[p['id'],eid] <= y[p['id'],eid]
            if int(p['team']) not in active[eid]:prob += z[p['id'],eid]==0
        for t in types:
            tid=int(t['id']); ys=pulp.lpSum(y[p['id'],eid] for p in candidates if int(p['element_type'])==tid)
            if t.get('squad_min_play') is not None:prob += ys>=int(t['squad_min_play'])
            if t.get('squad_max_play') is not None:prob += ys<=int(t['squad_max_play'])
    for gw in sorted({parse_gw(e.get('name')) for e in events}):
        prob += pulp.lpSum(z[p['id'],e['id']] for p in candidates for e in events if parse_gw(e.get('name'))==gw)==1
    t_model=time.perf_counter()-t0
    t0=time.perf_counter()
    scenario_exact = fixed_candidate_ids is not None
    solver = pulp.PULP_CBC_CMD(msg=False,timeLimit=30,gapRel=0.0 if scenario_exact else 0.03)
    status=prob.solve(solver)
    t_solve=time.perf_counter()-t0
    t0=time.perf_counter()
    if pulp.LpStatus[status] not in ('Optimal','Not Solved'):
        raise RuntimeError(f'Geen bruikbare selectie gevonden bij ${budget_cap:.1f}m.')
    selected=[p for p in candidates if pulp.value(x[p['id']])>.5]
    lineups={e['id']:[p for p in candidates if pulp.value(y[p['id'],e['id']])>.5] for e in events}
    caps={e['id']:next((p for p in candidates if pulp.value(z[p['id'],e['id']])>.5),None) for e in events}
    caps={k:v for k,v in caps.items() if v}
    return {'selected':selected,'lineups':lineups,'captains':caps,'objective':float(pulp.value(prob.objective) or 0),
            'roster_points':float(pulp.value(expected) or 0),'captain_bonus':float(pulp.value(capbonus) or 0),
            'spent':sum(num(p.get('now_cost'))/currency for p in selected),'status':pulp.LpStatus[status],
            'candidate_count':len(candidates),'candidate_ids':[p['id'] for p in candidates],
            'transfers_in':[p for p in selected if current_roster and p['id'] not in current_roster],
            'transfers_out':[pid for pid in current_roster if pid not in {p['id'] for p in selected}],
            'transfer_penalty':float(pulp.value(transfer_penalty) or 0) if allow_paid and current_roster else 0.0,
            'gross_objective':float((pulp.value(expected) or 0)+(pulp.value(capbonus) or 0)),
            'timing':{'shortlist':t_short,'model':t_model,'solve':t_solve,
                      'extract':time.perf_counter()-t0,'total':time.perf_counter()-t_all}}

def explain_result(result, events, active, currency, weights, first_gw, b2b):
    """Break an optimized result into player-level starting/captain contributions."""
    selected={p['id']:p for p in result.get('selected',[])}
    out={}
    for pid,p in selected.items():
        out[pid]={'name':p.get('web_name') or full_name(p),'starts':0,'active_starts':0,
                  'start_points':0.0,'captain_uses':0,'captain_bonus':0.0,'total':0.0}
    for e in events:
        eid=e['id']; w=weights.get(eid,1.0)
        for p in result.get('lineups',{}).get(eid,[]):
            pid=p['id']; row=out.setdefault(pid,{'name':p.get('web_name') or full_name(p),'starts':0,'active_starts':0,'start_points':0.0,'captain_uses':0,'captain_bonus':0.0,'total':0.0})
            row['starts']+=1
            if int(p['team']) in active[eid]:
                row['active_starts']+=1
                row['start_points']+=event_projection(p,e,first_gw,currency,b2b)*w
        cap=result.get('captains',{}).get(eid)
        if cap is not None and int(cap['team']) in active[eid]:
            pid=cap['id']; row=out.setdefault(pid,{'name':cap.get('web_name') or full_name(cap),'starts':0,'active_starts':0,'start_points':0.0,'captain_uses':0,'captain_bonus':0.0,'total':0.0})
            row['captain_uses']+=1
            row['captain_bonus']+=event_projection(cap,e,first_gw,currency,b2b)*w
    for row in out.values(): row['total']=row['start_points']+row['captain_bonus']
    return out

def gap_explanation(base, alt, events, active, currency, weights, first_gw, b2b):
    b=explain_result(base,events,active,currency,weights,first_gw,b2b)
    a=explain_result(alt,events,active,currency,weights,first_gw,b2b)
    ids=set(b)|set(a); rows=[]
    for pid in ids:
        br=b.get(pid,{'name':a.get(pid,{}).get('name',str(pid)),'starts':0,'active_starts':0,'start_points':0,'captain_uses':0,'captain_bonus':0,'total':0})
        ar=a.get(pid,{'name':br.get('name',str(pid)),'starts':0,'active_starts':0,'start_points':0,'captain_uses':0,'captain_bonus':0,'total':0})
        delta=ar['total']-br['total']
        if abs(delta)>.05 or ar['starts']!=br['starts'] or ar['captain_uses']!=br['captain_uses']:
            rows.append({'Speler':ar.get('name') or br.get('name'),
                         'Δ totaal':round(delta,1),'Δ startpunten':round(ar['start_points']-br['start_points'],1),
                         'Δ captain':round(ar['captain_bonus']-br['captain_bonus'],1),
                         'Starts':f"{br['starts']} → {ar['starts']}",
                         'Met fixture':f"{br['active_starts']} → {ar['active_starts']}",
                         'Captain':f"{br['captain_uses']} → {ar['captain_uses']}"})
    rows.sort(key=lambda r:abs(r['Δ totaal']),reverse=True)
    return rows

def net_swap_summary(base, alt, out_ids, in_ids, players, events, active, currency, weights, first_gw, b2b):
    """Pair OUT/IN by roster position and show intuitive net production change.
    The residual captures lineup/fixture interactions on unchanged players.
    """
    b=explain_result(base,events,active,currency,weights,first_gw,b2b)
    a=explain_result(alt,events,active,currency,weights,first_gw,b2b)
    pmap={p['id']:p for p in players}
    rows=[]; paired_effect=0.0
    for pos in sorted(set(int(pmap[x]['element_type']) for x in list(out_ids)+list(in_ids) if x in pmap)):
        outs=[x for x in out_ids if x in pmap and int(pmap[x]['element_type'])==pos]
        ins=[x for x in in_ids if x in pmap and int(pmap[x]['element_type'])==pos]
        # Largest production replacement first. Exact pairing is explanatory only;
        # roster optimization itself is solved jointly.
        outs=sorted(outs,key=lambda x:b.get(x,{}).get('total',0),reverse=True)
        ins=sorted(ins,key=lambda x:a.get(x,{}).get('total',0),reverse=True)
        for oid,iid in zip(outs,ins):
            old=b.get(oid,{'total':0,'start_points':0,'captain_bonus':0})
            new=a.get(iid,{'total':0,'start_points':0,'captain_bonus':0})
            effect=new['total']-old['total']; paired_effect+=effect
            rows.append({'Swap':f"{pmap[oid].get('web_name',oid)} → {pmap[iid].get('web_name',iid)}",
                         'OUT productie':round(old['total'],1),'IN productie':round(new['total'],1),
                         'Netto':round(effect,1)})
    total=alt['objective']-base['objective']
    residual=total-paired_effect
    rows.append({'Swap':'Rotatie / interactie overige spelers','OUT productie':'—','IN productie':'—','Netto':round(residual,1)})
    rows.append({'Swap':'TOTAAL','OUT productie':'','IN productie':'','Netto':round(total,1)})
    return rows

def save_run(result,mode,budget):
    con=db_connect(); con.execute('INSERT INTO runs(created_at,mode,budget_cap,objective,roster_points,captain_bonus,spent,roster_json) VALUES(?,?,?,?,?,?,?,?)',
      (datetime.now(timezone.utc).isoformat(),mode,budget,result['objective'],result['roster_points'],result['captain_bonus'],result['spent'],json.dumps([p['id'] for p in result['selected']])))
    con.commit();con.close()

def excel_bytes(result,teams,type_names,events,active,currency,weights,first_gw,b2b):
    from io import BytesIO
    wb=Workbook(); ws=wb.active; ws.title='Advies'
    ws.append(['NBA Fantasy Optimizer v3.18.1']); ws.append(['Budget gebruikt',result['spent'],'Objective',result['objective'],'Captainbonus',result['captain_bonus']])
    ws.append(['Speler','Team','Pos','Prijs ($m)','Projectie','Durability %','Actueel %','GP 23-24','GP 24-25','GP 25-26'])
    for p in sorted(result['selected'],key=lambda p:-p['_proj']):
        dd=durability_details(p); gp=dd['games']
        ws.append([p['web_name'],teams[int(p['team'])]['short_name'],type_names[int(p['element_type'])],num(p['now_cost'])/currency,round(p['_proj'],2),('ONBEKEND' if dd['score'] is None else round(dd['score']*100)),round(current_availability(p)*100),gp.get(2024,''),gp.get(2025,''),gp.get(2026,'')])
    wl=wb.create_sheet('Starting 5'); wl.append(['Gameday','Speler','Team','Projectie'])
    for e in events:
        for p in result['lineups'][e['id']]:
            pts=event_projection(p,e,first_gw,currency,b2b) if int(p['team']) in active[e['id']] else 0
            wl.append([e['name'],p['web_name'],teams[int(p['team'])]['short_name'],round(pts,2)])
    wc=wb.create_sheet('Captain'); wc.append(['Gameweek','Gameday','Captain','Bonus'])
    for e in events:
        if e['id'] in result['captains']:
            p=result['captains'][e['id']]; wc.append([parse_gw(e['name']),e['name'],p['web_name'],round(event_projection(p,e,first_gw,currency,b2b),2)])
    for sh in wb.worksheets:
        for c in sh[1]:c.font=Font(bold=True,color='FFFFFF');c.fill=PatternFill('solid',fgColor='1F4E78')
        for i in range(1,sh.max_column+1):sh.column_dimensions[get_column_letter(i)].width=22
    bio=BytesIO();wb.save(bio);return bio.getvalue()

@st.cache_data(ttl=900,show_spinner=False)
def load_live(): return get_json('bootstrap-static/')

@st.cache_data(ttl=900,show_spinner=False)
def load_fixtures(event_ids):
    return {eid:get_json(f'fixtures/?event={eid}') for eid in event_ids}

def html_table(rows, height=520):
    """Render list[dict] without st.dataframe, so Streamlit never imports pandas."""
    if not rows:
        st.info("Geen gegevens om te tonen.")
        return
    cols=list(rows[0].keys())
    css=f"""
    <style>
    .nba-wrap{{max-height:{int(height)}px;overflow:auto;border:1px solid #354b78;border-radius:10px}}
    .nba-table{{width:100%;border-collapse:collapse;font-size:14px}}
    .nba-table th{{position:sticky;top:0;background:#17365d;color:white;text-align:left;padding:8px;z-index:1}}
    .nba-table td{{padding:7px 8px;border-bottom:1px solid #263b60;white-space:nowrap;color:#eaf1ff}}
    .nba-table tr:nth-child(odd){{background:#101c36}}
    .nba-table tr:nth-child(even){{background:#172747}}
    .nba-table tr:hover{{background:#254071}}
    </style>"""
    head="".join(f"<th>{escape(str(c))}</th>" for c in cols)
    body=[]
    for row in rows:
        cells=[]
        for c in cols:
            v=row.get(c,"")
            if isinstance(v,float):
                txt=f"{v:.2f}".rstrip("0").rstrip(".")
            else:
                txt=str(v)
            cells.append(f"<td>{escape(txt)}</td>")
        body.append("<tr>"+"".join(cells)+"</tr>")
    html=css+f'<div class="nba-wrap"><table class="nba-table"><thead><tr>{head}</tr></thead><tbody>{"".join(body)}</tbody></table></div>'
    st.markdown(html,unsafe_allow_html=True)

def durability_coverage(players):
    """Coverage measured through the same player->history matching path used by the model."""
    out={yr:{'matched':0,'total':len(players),'pct':0.0} for yr in (2024,2025,2026)}
    for p in players:
        hist=three_year_games(p)
        for yr in (2024,2025,2026):
            if yr in hist and hist[yr] is not None:
                out[yr]['matched']+=1
    for yr in out:
        out[yr]['pct']=100*out[yr]['matched']/out[yr]['total'] if out[yr]['total'] else 0.0
    return out



def preseason_context(p, role_map=None):
    # Generic preseason uncertainty adjustment.
    # More historical seasons -> more confidence. Sparse history is shrunk mildly,
    # never boosted. Team/role context remains the separate role factor.
    dd=durability_details(p)
    games=dd.get('games',{}) if isinstance(dd,dict) else {}
    seasons=sum(1 for y in (2024,2025,2026) if games.get(y) not in (None,'',0))
    confidence={0:0.90,1:0.94,2:0.97,3:1.00}.get(seasons,1.00)
    return {'history_seasons':seasons,'confidence':confidence}

def uncertainty_regression(p, raw_projection, currency):
    """Safety guard for 0Y/API-fallback players only.

    Backtest calibration now handles players with historical evidence. For players
    with zero reconstructed seasons the backtest cannot estimate a correction, so
    retain a mild continuous anti-outlier guard instead of inventing a boost.
    """
    price=max(.1,num(p.get('now_cost'))/currency)
    seasons=preseason_context(p).get('history_seasons',0)
    value=raw_projection/price if price else 0.0
    if seasons>0 or raw_projection<=0:
        return {'raw':raw_projection,'adjusted':raw_projection,'penalty':0.0,
                'shrink':0.0,'pressure':0.0,'value_raw':value,'ceiling':None}
    max_shrink=.12
    pressure=1.0/(1.0+math.exp(-2.0*(value-2.5)))
    penalty=raw_projection*max_shrink*pressure
    adjusted=max(0.0,raw_projection-penalty)
    return {'raw':raw_projection,'adjusted':adjusted,'penalty':penalty,
            'shrink':max_shrink,'pressure':pressure,'value_raw':value,'ceiling':None}

def final_projection(p,currency,role_factor=1.0):
    raw=base_projection(p,currency)*preseason_context(p)['confidence']*current_availability(p)*durability_factor(p)*role_factor
    # Keep explorer/team projection aligned with the selected Risk & Upside strategy.
    p['_role_factor']=role_factor
    raw*=strategy_factor(p)
    return uncertainty_regression(p,raw,currency)['adjusted']

def projection_breakdown(p,currency,role_map):
    """Expose every multiplicative factor used by the current projection."""
    ap=adaptive_projection_components(p,currency)
    base=ap['base']
    availability=current_availability(p)
    dr=durability_rate(p)
    durability=.55+.45*dr
    role=role_map.get(p['id'],{}).get('factor',1.0)
    pc=preseason_context(p,role_map)
    confidence=pc['confidence']
    raw_final=base*confidence*availability*durability*role
    ur=uncertainty_regression(p,raw_final,currency)
    final=ur['adjusted']
    return {
        'PPG bron':round(num(p.get('points_per_game'))/currency,2),
        'FP 23-24':round(ap['multi']['history'].get(2024),2) if ap['multi']['history'].get(2024) is not None else '—',
        'FP 24-25':round(ap['multi']['history'].get(2025),2) if ap['multi']['history'].get(2025) is not None else '—',
        'FP 25-26':round(ap['multi']['history'].get(2026),2) if ap['multi']['history'].get(2026) is not None else '—',
        'Multi-season baseline':round(ap['multi']['baseline'],2),
        'Multi-season bron':ap['multi']['source'],
        'Form bron':round(num(p.get('form'))/currency,2),
        'Minuten bron':round(num(p.get('minutes')),1),
        'Totaalpunten bron':round(num(p.get('total_points'))/currency,1),
        'Games sample*':round(ap['games_est'],1),
        'Historisch gewicht%':round(ap['historical_weight']*100,1),
        'Actueel gewicht%':round(ap['recent_weight']*100,1),
        'Blended baseline':round(ap['blended'],2),
        'Minuten reliability%':round(ap['reliability']*100,1),
        'Basis vóór backtest-calibratie':round(ap.get('base_uncalibrated',base),2),
        'Backtest calibration':round(ap.get('calibration',{}).get('delta',0.0),2),
        'Vorige GP (calibratie)':ap.get('calibration',{}).get('prev_gp') if ap.get('calibration',{}).get('prev_gp') is not None else '—',
        'Basis':round(base,2),
        'Historische seizoenen':pc['history_seasons'],
        'Preseason confidence%':round(confidence*100,1),
        'Availability%':round(availability*100,1),
        'Durability%':round(dr*100,1),
        'Durability factor%':round(durability*100,1),
        'Role%':round(role*100,1),
        'Voor uncertainty-regressie':round(raw_final,2),
        'Uncertainty penalty':round(ur['penalty'],2),
        'Uncertainty shrink%':round(ur['shrink']*100,1),
        'Value-ceiling anchor':round(ur['ceiling'],2) if ur.get('ceiling') is not None else '—',
        'Projectie':round(final,2),
    }

def projection_audit(p,currency,role_map):
    """Transparent data-quality audit of the projection used by the optimizer."""
    bd=projection_breakdown(p,currency,role_map)
    price=max(.1,num(p.get('now_cost'))/currency)
    score=100
    flags=[]
    seasons=int(bd['Historische seizoenen'])
    if seasons==0:
        score-=35; flags.append('geen 3Y fantasy/GP-historie')
    elif seasons==1:
        score-=22; flags.append('slechts 1 historisch seizoen')
    elif seasons==2:
        score-=10; flags.append('2 van 3 historische seizoenen')
    if 'fallback' in str(bd['Multi-season bron']).lower():
        score-=18; flags.append('Fantasy API fallback')
    if bd['Durability%']>=99.9 and seasons==0:
        score-=10; flags.append('durability onbekend')
    role=float(bd['Role%'])
    if role<90 or role>110:
        score-=10; flags.append('sterke role-aanpassing')
    elif role<94 or role>106:
        score-=5; flags.append('merkbare role-aanpassing')
    api=float(bd['PPG bron']); base=float(bd['Multi-season baseline'])
    if api>0 and base>0:
        div=abs(base-api)/api
        if div>=.25:
            score-=12; flags.append('historie/API wijken ≥25% af')
        elif div>=.15:
            score-=6; flags.append('historie/API wijken ≥15% af')
    proj=float(bd['Projectie'])
    if proj<=0:
        score-=25; flags.append('nulprojectie')
    value=proj/price
    if value>=3.5:
        score-=8; flags.append('extreme value/prijs')
    elif value>=3.0:
        score-=4; flags.append('hoge value/prijs')
    if CURRENT_SEASON_STARTED and float(bd['Games sample*'])<3:
        score-=8; flags.append('kleine actuele sample')
    score=max(0,min(100,score))
    label='HOOG' if score>=85 else ('MIDDEL' if score>=65 else 'LAAG')
    return {'Confidence':score,'Niveau':label,'Flags':'; '.join(flags) if flags else 'geen rode vlaggen',
            'Waarde/prijs':round(value,2),**bd}


def projection_lab_rows(players,currency,role_map):
    rows=[]
    for p in players:
        a=projection_audit(p,currency,role_map)
        rows.append({'Speler':p.get('web_name') or full_name(p),
                     'Prijs':round(num(p.get('now_cost'))/currency,1),
                     'Projectie':a['Projectie'],'Conf.%':a['Confidence'],'Niveau':a['Niveau'],
                     'Hist. seizoenen':a['Historische seizoenen'],'Bron':a['Multi-season bron'],
                     'API PPG':a['PPG bron'],'3Y baseline':a['Multi-season baseline'],
                     'Role%':a['Role%'],'Durability%':a['Durability%'],
                     'Value/prijs':a['Waarde/prijs'],'Waarschuwingen':a['Flags']})
    return rows


def historical_projection_backtest(min_games=15):
    """Walk-forward validation of the reconstructed multi-season fantasy baseline.

    Uses only seasons that were already known before each target season:
    2024 -> 2025 and 2024+2025 -> 2026. This deliberately does NOT claim to
    backtest future prices, roles, injuries, fixtures or captaincy.
    """
    con=db_connect()
    rows=con.execute(
        'SELECT season_end,player_key,player_name,games,fantasy_ppg FROM fantasy_seasons '
        'WHERE season_end IN (2024,2025,2026)'
    ).fetchall()
    con.close()
    by_key={}
    for yr,key,name,games,fp in rows:
        if games is None or fp is None: continue
        by_key.setdefault(key,{})[int(yr)]={'name':name,'games':int(games),'fp':float(fp)}
    details=[]
    for key,h in by_key.items():
        for target in (2025,2026):
            actual=h.get(target)
            if not actual or actual['games'] < int(min_games): continue
            priors=[]
            if target==2025 and 2024 in h and h[2024]['games']>=int(min_games):
                priors=[(2024,h[2024]['fp'],1.0)]
            elif target==2026:
                if 2024 in h and h[2024]['games']>=int(min_games): priors.append((2024,h[2024]['fp'],.40))
                if 2025 in h and h[2025]['games']>=int(min_games): priors.append((2025,h[2025]['fp'],.60))
            if not priors: continue
            denom=sum(w for _,_,w in priors)
            pred=sum(fp*w for _,fp,w in priors)/denom
            err=pred-actual['fp']
            prior_years=[yr for yr,_,_ in priors]
            prior_games=[h[yr]['games'] for yr in prior_years if yr in h]
            prior_fps=[h[yr]['fp'] for yr in prior_years if yr in h]
            avg_prior_games=sum(prior_games)/len(prior_games) if prior_games else 0.0
            avg_prior_fp=sum(prior_fps)/len(prior_fps) if prior_fps else 0.0
            details.append({'Speler':actual['name'],'Doelseizoen':f'{target-1}-{str(target)[-2:]}',
                            'Historie':len(priors),'Voorspeld':round(pred,2),'Werkelijk':round(actual['fp'],2),
                            'Fout':round(err,2),'Abs. fout':round(abs(err),2),'Games':actual['games'],
                            'Vorige GP':round(avg_prior_games,1),'Vorige FP/G':round(avg_prior_fp,1)})
    def metrics(sub):
        if not sub:return {'N':0,'MAE':0.0,'RMSE':0.0,'Bias':0.0}
        es=[float(r['Fout']) for r in sub]
        return {'N':len(es),'MAE':sum(abs(x) for x in es)/len(es),
                'RMSE':math.sqrt(sum(x*x for x in es)/len(es)),
                'Bias':sum(es)/len(es)}
    groups=[]
    for label,sub in [('Alles',details),('2024-25',[r for r in details if r['Doelseizoen']=='2024-25']),
                      ('2025-26',[r for r in details if r['Doelseizoen']=='2025-26']),
                      ('1 seizoen historie',[r for r in details if r['Historie']==1]),
                      ('2 seizoenen historie',[r for r in details if r['Historie']==2])]:
        m=metrics(sub)
        groups.append({'Groep':label,'N':m['N'],'MAE':round(m['MAE'],2),'RMSE':round(m['RMSE'],2),'Bias':round(m['Bias'],2)})
    details.sort(key=lambda r:r['Abs. fout'],reverse=True)
    return groups,details


def backtest_diagnostic_groups(details):
    def metrics(label, sub):
        if not sub: return {'Groep':label,'N':0,'MAE':0.0,'RMSE':0.0,'Bias':0.0}
        es=[float(r['Fout']) for r in sub]
        return {'Groep':label,'N':len(es),'MAE':round(sum(abs(x) for x in es)/len(es),2),
                'RMSE':round(math.sqrt(sum(x*x for x in es)/len(es)),2),'Bias':round(sum(es)/len(es),2)}
    gp_bands=[('Vorige GP <30',lambda x:x<30),('Vorige GP 30-59',lambda x:30<=x<60),('Vorige GP 60+',lambda x:x>=60)]
    fp_bands=[('Vorige FP/G <20',lambda x:x<20),('Vorige FP/G 20-34.9',lambda x:20<=x<35),('Vorige FP/G 35-49.9',lambda x:35<=x<50),('Vorige FP/G 50+',lambda x:x>=50)]
    gp=[metrics(label,[r for r in details if fn(float(r.get('Vorige GP',0)))]) for label,fn in gp_bands]
    fp=[metrics(label,[r for r in details if fn(float(r.get('Vorige FP/G',0)))]) for label,fn in fp_bands]
    return gp,fp


def backtest_threshold_matrix():
    rows=[]
    for threshold in (15,30,45,60):
        groups,details=historical_projection_backtest(threshold)
        overall=next((r for r in groups if r['Groep']=='Alles'),None)
        if overall:
            rows.append({'Min. games':threshold,'N':overall['N'],'MAE':overall['MAE'],'RMSE':overall['RMSE'],'Bias':overall['Bias']})
    return rows




def crystal_style():
    """Pure CSS skin; no external images, JS, or additional dependencies."""
    st.markdown("""<style>
    .stApp {background:radial-gradient(ellipse at 80% 0%,#152653 0%,#080f25 40%,#080d1d 100%);color:#eaf0ff}
    [data-testid="stSidebar"] {background:linear-gradient(165deg,#111e3d,#090e20)}
    [data-testid="stSidebar"] * {color:#e4edff}
    .block-container {padding-top:1.2rem;max-width:1550px}
    .cb-hero {display:flex;align-items:center;gap:20px;padding:15px 20px 20px;
      background:linear-gradient(110deg,rgba(17,33,76,.9),rgba(55,20,99,.46));
      border:1px solid #4558a4;border-radius:20px;margin-bottom:12px;box-shadow:0 8px 30px #03071588}
    .cb-hero h1 {font-size:clamp(2.3rem,4vw,4rem);letter-spacing:-.045em;margin:0;
      background:linear-gradient(90deg,#b7e5ff,#cbb1ff,#fb9dff);background-clip:text;-webkit-background-clip:text;color:transparent}
    .cb-hero p {letter-spacing:.25em;font-size:.77rem;color:#c5d4ff;margin:0}
    .cb-orb {font-size:3.2rem;width:80px;height:80px;display:grid;place-items:center;
      border-radius:50%;background:radial-gradient(circle at 30% 25%,#8ac9ff,#6942ce 55%,#171743);
      border:2px solid #8a92ff;box-shadow:0 0 32px #754fffaa,inset 0 0 15px #9cceff}
    .cb-tag {color:#9ce8ff;letter-spacing:normal;border:1px solid #6878c0;border-radius:12px;padding:2px 8px}
    [data-testid="stMetric"] {background:#111f3c;border:1px solid #2c4779;border-radius:14px;padding:13px}
    [data-testid="stMetricValue"] {color:#d8c5ff}
    div.stButton>button[kind="primary"] {background:linear-gradient(90deg,#3158e9,#a02fdf);border:0}
    .cb-court-shell {background:#0a1730;border:1px solid #31539b;border-radius:18px;padding:18px;margin:16px 0;box-shadow:0 10px 28px #0007}
    .cb-section {font-weight:750;color:#eaf2ff;font-size:1.25rem;margin-bottom:12px}
    .cb-court {position:relative;overflow:hidden;min-height:400px;border:9px solid #9b6237;
      border-radius:10px;background:repeating-linear-gradient(90deg,#b67a4b 0px,#b67a4b 68px,#c38b57 69px,#c38b57 136px);
      box-shadow:inset 0 0 75px #25101a70}
    .cb-court:before {content:'';position:absolute;inset:16px;border:2px solid #f9e6c9a8;pointer-events:none}
    .cb-circle {position:absolute;width:150px;height:150px;border:3px solid #f7e6c6b8;border-radius:50%;
      left:50%;top:50%;transform:translate(-50%,-50%);pointer-events:none}
    .cb-half {position:absolute;top:0;bottom:0;left:50%;border-left:2px solid #f7e6c6b8;pointer-events:none}
    .cb-lineup {position:relative;z-index:1;display:grid;grid-template-columns:repeat(6,minmax(0,1fr));
      grid-template-rows:repeat(2,minmax(160px,1fr));gap:10px;align-items:center;padding:26px 15px;min-height:390px}
    .cb-player {grid-column:span 2;justify-self:center;width:min(100%,215px);padding:13px 7px;
      background:linear-gradient(160deg,#10244ceF,#090e23f5);border:1px solid #7b85d2;
      border-radius:13px;box-shadow:0 7px 18px #090b25b0;text-align:center;color:white}
    .cb-player.bc {grid-column:span 3}
    .cb-player .name {font-weight:800;font-size:clamp(.75rem,1.25vw,1rem);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
    .cb-player .detail {color:#bbcbef;font-size:.77rem;margin-top:3px}
    .cb-player .score {color:#f8e69a;font-size:1.3rem;font-weight:800}
    .cb-pill {display:inline-block;background:#5234c2;border-radius:8px;padding:2px 10px;font-size:.7rem;margin-top:4px}
    .cb-bench {display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:10px;margin-top:12px}
    .cb-bench .cb-player {width:100%;border-color:#3c568f}
    .cb-bench .cb-player .score {font-size:1.1rem}
    @media(max-width:850px){.cb-lineup{gap:5px;padding:15px 5px}.cb-player{padding:8px 3px}.cb-bench{grid-template-columns:repeat(2,minmax(0,1fr))}}
    /* Court and bench need independent grid spans. */
    .cb-bench .cb-player, .cb-bench .cb-player.bc, .cb-bench .cb-player.fc {
      grid-column:auto !important; width:100%;min-width:0;box-sizing:border-box;
      padding:12px 5px;overflow:hidden;
    }
    .cb-bench .cb-player .name {font-size:clamp(.7rem,1vw,.92rem)}
    .cb-bench .cb-player .detail {font-size:.7rem}
    /* Streamlit's default light widgets were illegible on our dark theme. */
    [data-testid="stSidebar"] input,
    [data-testid="stSidebar"] [data-baseweb="select"] > div,
    [data-testid="stSidebar"] [data-baseweb="input"] > div,
    [data-testid="stSidebar"] [data-baseweb="base-input"] {
      background:#172747 !important;color:#f0f5ff !important;border-color:#45629c !important;
    }
    [data-testid="stSidebar"] input::placeholder {color:#a9bce1 !important}
    [data-testid="stSidebar"] button {color:#f4f6ff !important;background:#182a4e !important;border:1px solid #4a65a4 !important}
    [data-testid="stSidebar"] button:disabled {opacity:.48}
    [data-testid="stSidebar"] [data-testid="stTooltipIcon"] button {
      background:transparent !important;border:0 !important;
      padding:0 !important;box-shadow:none !important;
      color:#aabce5 !important;min-width:0 !important;
    }
    [data-testid="stSidebar"] [data-baseweb="select"] * {color:#f0f5ff !important}
    [data-testid="stSidebar"] [data-testid="stWidgetLabel"] * {color:#e8efff !important}
    div[data-testid="stTextInput"] input,
    div[data-testid="stSelectbox"] [data-baseweb="select"] > div {
      background:#172747 !important;color:#f0f5ff !important;border-color:#45629c !important;
    }
    .nba-wrap {border-color:#354b78 !important}
    .nba-table td {color:#eaf1ff !important;border-bottom-color:#263b60 !important}
    .nba-table tr:nth-child(odd) {background:#101c36 !important}
    .nba-table tr:nth-child(even) {background:#172747 !important}
    .nba-table tr:hover {background:#254071 !important}
    @media(max-width:1000px){.cb-bench{grid-template-columns:repeat(3,minmax(0,1fr))}}
    @media(max-width:600px){.cb-bench{grid-template-columns:repeat(2,minmax(0,1fr))}}
    
    /* v4.0.4: explicit contrast for Streamlit's light default widgets on dark skin */
    [data-testid="stMetricLabel"] *, [data-testid="stMetricLabel"] {color:#d9e6ff !important;opacity:1 !important}
    [data-testid="stExpander"] summary, [data-testid="stExpander"] summary * {color:#eef4ff !important}
    [data-testid="stExpander"] {background:#142443 !important;border:1px solid #344b78 !important}
    div.stButton>button {background:#20345c !important;color:#f3f7ff !important;border:1px solid #5775b6 !important}
    div.stButton>button * {color:inherit !important}
    div.stButton>button[kind="primary"] {background:linear-gradient(90deg,#3158e9,#a02fdf) !important;color:#fff !important;border:0 !important}
    div.stButton>button:disabled {background:#1b2944 !important;color:#a4b4d1 !important;opacity:.65 !important}
    [data-testid="stTextInput"] input, [data-testid="stNumberInput"] input,
    [data-testid="stSelectbox"] [data-baseweb="select"],
    [data-testid="stMultiSelect"] [data-baseweb="select"] {background:#1a2c50 !important;color:#f4f7ff !important}
    [data-testid="stTextInput"] input::placeholder {color:#9eb1d0 !important}
    [data-testid="stAlert"] {color:#f2f5ff !important}
    /* v4.0.8: contrast polish; CSS only, zero network or computation. */
    [data-testid="stAlert"] {background:#17294b !important;border:1px solid #49669f !important;border-radius:12px !important}
    [data-testid="stAlert"] *, [data-testid="stAlert"] p {color:#eaf2ff !important;opacity:1 !important}
    [data-testid="stAlert"] [data-testid="stMarkdownContainer"] * {color:#eaf2ff !important}
    [data-testid="stTabs"] [role="tablist"] {border-bottom:1px solid #435a90 !important;gap:8px !important}
    [data-testid="stTabs"] button[role="tab"] {color:#cbd9f7 !important;opacity:1 !important;font-weight:650 !important}
    [data-testid="stTabs"] button[role="tab"] * {color:inherit !important;opacity:1 !important}
    [data-testid="stTabs"] button[role="tab"][aria-selected="true"] {color:#fff !important;border-bottom:3px solid #b67cff !important}
    /* v4.3.2: prominent, accessible navigation pills, no layout/model changes. */
    [data-testid="stTabs"] [role="tablist"] {
        display:flex !important; flex-wrap:wrap !important; gap:10px !important;
        padding:10px !important; background:#101f3e !important;
        border:1px solid #4967a0 !important; border-radius:14px !important;
        margin:10px 0 14px !important;
    }
    [data-testid="stTabs"] button[role="tab"] {
        background:#20365e !important; color:#e8f0ff !important;
        border:1px solid #6682b7 !important; border-radius:10px !important;
        padding:10px 17px !important; min-height:45px !important;
        font-size:14px !important; font-weight:750 !important;
        box-shadow:none !important; transition:background .16s,border-color .16s !important;
    }
    [data-testid="stTabs"] button[role="tab"]:hover {
        background:#304b80 !important; border-color:#a7c4ff !important;
    }
    [data-testid="stTabs"] button[role="tab"][aria-selected="true"] {
        background:linear-gradient(110deg,#365fe8,#8b37c9) !important;
        color:#ffffff !important; border:1px solid #d4a8ff !important;
        box-shadow:0 0 0 1px #6e73d2 !important;
    }
    [data-testid="stTabs"] button[role="tab"]:focus-visible {
        outline:3px solid #f6d581 !important; outline-offset:2px !important;
    }
    [data-testid="stTabs"] [data-baseweb="tab-highlight"],
    [data-testid="stTabs"] [data-baseweb="tab-border"] {display:none !important}
    /* v4.3.3: BaseWeb tabs use data-baseweb=tab, not button[role=tab].
       Cover both Streamlit/BaseWeb variants and force child text contrast. */
    div[data-testid="stTabs"] div[data-baseweb="tab-list"],
    div[data-testid="stTabs"] [role="tablist"] {
      display:flex !important;flex-wrap:wrap !important;gap:10px !important;
      padding:10px !important;background:#101f3e !important;
      border:1px solid #506ba7 !important;border-radius:14px !important;
      margin:8px 0 16px !important;
    }
    div[data-testid="stTabs"] button[data-baseweb="tab"],
    div[data-testid="stTabs"] [data-baseweb="tab-list"] > button,
    div[data-testid="stTabs"] [role="tab"] {
      background:#243c68 !important;color:#f4f7ff !important;
      border:1px solid #7997d2 !important;border-radius:10px !important;
      padding:11px 19px !important;min-height:48px !important;
      opacity:1 !important;font-size:15px !important;font-weight:750 !important;
      box-shadow:0 2px 8px #0004 !important;
    }
    div[data-testid="stTabs"] button[data-baseweb="tab"] *,
    div[data-testid="stTabs"] [role="tab"] * {
      color:#f4f7ff !important;opacity:1 !important;
      -webkit-text-fill-color:#f4f7ff !important;
    }
    div[data-testid="stTabs"] button[data-baseweb="tab"]:hover,
    div[data-testid="stTabs"] [role="tab"]:hover {
      background:#38578e !important;border-color:#b8cfff !important;
    }
    div[data-testid="stTabs"] button[data-baseweb="tab"][aria-selected="true"],
    div[data-testid="stTabs"] [role="tab"][aria-selected="true"] {
      background:linear-gradient(115deg,#375ff0,#9138d6) !important;
      color:#fff !important;border:1px solid #dfb9ff !important;
      box-shadow:0 0 0 1px #877be5,0 4px 14px #0005 !important;
    }
    div[data-testid="stTabs"] [data-baseweb="tab-highlight"],
    div[data-testid="stTabs"] [data-baseweb="tab-border"] {display:none !important}
    [data-testid="stExpander"] summary {background:#172a4c !important}
    [data-testid="stExpander"] summary:hover {background:#20375f !important}
    [data-testid="stExpander"] summary * {color:#f4f7ff !important;opacity:1 !important}
    [data-testid="stMetricDelta"] * {opacity:1 !important}
    .cb-orb svg {filter:drop-shadow(0 0 8px #92c8ff99)}
    .cb-player {transition:transform .15s ease,border-color .15s ease}
    .cb-player:hover {transform:translateY(-2px);border-color:#b9a5ff}
    /* v4.0.10: Streamlit header, widget help icon and floating tooltips. */
    header[data-testid="stHeader"],
    [data-testid="stHeader"] {
      background:#0e1933 !important;color:#eaf0ff !important;
      border-bottom:1px solid #314878 !important;
    }
    [data-testid="stToolbar"], [data-testid="stDecoration"],
    [data-testid="stStatusWidget"] {background-color:transparent !important;color:#eaf0ff !important}
    [data-testid="stToolbar"] button, [data-testid="stToolbar"] svg,
    [data-testid="stHeader"] button, [data-testid="stHeader"] svg {color:#dce7ff !important;fill:currentColor}
    /* Help icons are actual icon buttons, not input widgets. */
    [data-testid="stWidgetLabel"] [data-testid="stTooltipIcon"] button,
    [data-testid="stWidgetLabel"] [data-testid="stTooltipIcon"] {
      background:transparent !important;border:0 !important;
      box-shadow:none !important;padding:0 !important;
      width:auto !important;min-width:0 !important;height:auto !important;
      color:#aabce5 !important;
    }
    [data-testid="stWidgetLabel"] [data-testid="stTooltipIcon"] svg {
      color:#aabce5 !important;fill:none !important;width:15px !important;height:15px !important;
    }
    /* Popovers live in a portal outside .stApp. */
    [data-baseweb="tooltip"], [role="tooltip"] {
      background:#1a2b4d !important;color:#f0f5ff !important;
      border:1px solid #536da8 !important;border-radius:9px !important;
    }
    [data-baseweb="tooltip"] *, [role="tooltip"] * {color:#f0f5ff !important}
    /* v4.0.9: stronger specificity for Streamlit 1.x light-theme defaults. */
    .stApp [data-testid="stButton"] button,
    .stApp [data-testid="stBaseButton-secondary"],
    .stApp [data-testid="stBaseButton-primary"] {
      background:#20365f !important;color:#f5f8ff !important;
      -webkit-text-fill-color:#f5f8ff !important;
      border:1px solid #607cb6 !important;
    }
    .stApp [data-testid="stButton"] button p,
    .stApp [data-testid="stButton"] button span {color:inherit !important;-webkit-text-fill-color:inherit !important}
    .stApp [data-testid="stButton"] button[kind="primary"],
    .stApp [data-testid="stBaseButton-primary"] {
      background:linear-gradient(90deg,#3158e9,#a02fdf) !important;
      border-color:transparent !important;
    }
    .stApp [data-testid="stButton"] button:disabled,
    .stApp [data-testid="stBaseButton-secondary"]:disabled {
      background:#192641 !important;color:#b1bfda !important;
      -webkit-text-fill-color:#b1bfda !important;opacity:.78 !important;
    }
    .stApp [data-testid="stWidgetLabel"],
    .stApp [data-testid="stWidgetLabel"] p,
    .stApp label[data-testid="stWidgetLabel"] {color:#e7efff !important;opacity:1 !important}
    .stApp [data-testid="stSelectbox"] [data-baseweb="select"] > div,
    .stApp [data-testid="stSelectbox"] [data-baseweb="select"] input,
    .stApp [data-testid="stSelectbox"] [data-baseweb="select"] span {
      background-color:#192c50 !important;color:#f5f8ff !important;
      -webkit-text-fill-color:#f5f8ff !important;
    }
    .stApp [data-testid="stSelectbox"] [data-baseweb="select"] svg {fill:#c8d8ff !important}
    [data-baseweb="popover"] [role="listbox"] {background:#192c50 !important}
    [data-baseweb="popover"] [role="option"] {color:#f5f8ff !important;background:#192c50 !important}
    [data-baseweb="popover"] [role="option"]:hover {background:#304b7b !important}
    /* v4.0.11: native Streamlit/BaseWeb tooltips may be rendered in a
       light portal outside .stApp. Force readable DARK text on light surface,
       rather than white text on the default white floating tooltip. */
    div[data-baseweb="tooltip"], div[role="tooltip"],
    div[data-baseweb="popover"] div[role="tooltip"],
    div[data-baseweb="popover"]:has([data-testid="stTooltipContent"]) {
      background:#f2f5fc !important;
      color:#17243e !important;
      -webkit-text-fill-color:#17243e !important;
      border-color:#8fa5cf !important;
    }
    div[data-baseweb="tooltip"] *, div[role="tooltip"] *,
    [data-testid="stTooltipContent"], [data-testid="stTooltipContent"] * {
      color:#17243e !important;
      -webkit-text-fill-color:#17243e !important;
      opacity:1 !important;
    }

    /* v4.1 portrait court and collectible cards; no solver changes */
    .cb-tc-shell{background:#0a1730;border:1px solid #3556a0;border-radius:18px;padding:18px;margin:16px 0}
    .cb-tc-layout{display:grid;grid-template-columns:minmax(0,1.05fr) minmax(260px,.95fr);gap:18px;align-items:stretch}
    .cb-tc-court{position:relative;min-height:670px;max-width:610px;width:100%;margin:0 auto;overflow:hidden;border:9px solid #a36a3e;border-radius:15px;background:repeating-linear-gradient(90deg,#bb8453 0 49px,#c58f5c 50px 99px);box-shadow:inset 0 0 50px #4b281a66}
    .cb-tc-court:before{content:'';position:absolute;inset:14px;border:3px solid #f7e4c8c9;border-radius:3px;pointer-events:none}
    .cb-tc-arc{position:absolute;left:7%;top:4%;width:86%;height:58%;border:3px solid #f7e4c8c9;border-radius:0 0 50% 50%;border-top:0;pointer-events:none}
    .cb-tc-paint{position:absolute;top:0;left:30%;width:40%;height:28%;border:3px solid #f7e4c8c9;border-top:0;background:#a76e49aa;pointer-events:none}
    .cb-tc-hoop{position:absolute;top:10%;left:calc(50% - 21px);width:42px;height:42px;border:3px solid #fff0d3;border-radius:50%;pointer-events:none}
    .cb-tc-slot{position:absolute;z-index:2;width:clamp(115px,23%,160px);transform:translate(-50%,-50%)}
    .cb-tc-card{position:relative;min-width:0;overflow:hidden;padding:6px 7px 8px;text-align:center;color:#fff;background:linear-gradient(150deg,#243452,#0a1228 70%);border:2px solid #e6bc65;border-radius:13px;box-shadow:0 8px 18px #080a20ad,inset 0 0 0 1px #8a663a}
    .cb-tc-portrait{position:relative;height:76px;display:flex;align-items:center;justify-content:center;background:radial-gradient(ellipse at 50% 75%,#6e65b8,#142449 75%);border-radius:9px 9px 3px 3px;overflow:hidden}
    .cb-tc-photo{position:absolute;inset:0;z-index:2;width:100%;height:100%;object-fit:contain;object-position:bottom;filter:drop-shadow(0 2px 3px #060b19aa)}
    .cb-tc-monogram{font-size:2rem;font-weight:900;color:#cbdcff;text-shadow:0 2px 12px #c0a5ff99}
    .cb-tc-name{font-size:clamp(.67rem,1.05vw,.9rem);font-weight:850;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;margin-top:5px}
    .cb-tc-meta{font-size:.68rem;color:#c3d6ff;white-space:nowrap}
    .cb-tc-score{font-weight:900;font-size:1.17rem;color:#ffe49b;margin-top:3px}
    .cb-tc-score small{font-size:.55rem;color:#d5e2ff}
    .cb-tc-bench{display:flex;flex-direction:column;gap:9px;min-width:0}
    .cb-tc-bench-title{font-size:1.08rem;font-weight:800;color:#edf4ff;margin-bottom:2px}
    .cb-tc-reserve{display:grid;grid-template-columns:64px minmax(0,1fr) auto;grid-template-rows:auto auto;column-gap:11px;align-items:center;text-align:left;padding:7px 12px;min-height:85px;border-color:#667fc1}
    .cb-tc-reserve .cb-tc-portrait{grid-row:1/3;grid-column:1;height:65px}
    .cb-tc-reserve .cb-tc-monogram{font-size:1.35rem}
    .cb-tc-reserve .cb-tc-name{grid-column:2;grid-row:1;font-size:.95rem;margin:0}
    .cb-tc-reserve .cb-tc-meta{grid-column:2;grid-row:2;font-size:.73rem}
    .cb-tc-reserve .cb-tc-score{grid-column:3;grid-row:1/3;white-space:nowrap}
    /* v4.2.2 compact full-squad view: keep all ten cards in one visual block */
    .cb-tc-shell{padding:12px 16px;margin:10px 0}
    .cb-tc-shell .cb-section{margin-bottom:7px}
    .cb-tc-layout{grid-template-columns:minmax(0,1fr) minmax(290px,1fr);gap:14px;align-items:start}
    .cb-tc-court{min-height:490px;max-width:510px;border-width:6px}
    .cb-tc-court:before{inset:10px;border-width:2px}
    .cb-tc-arc{border-width:2px}
    .cb-tc-slot{width:clamp(100px,24%,132px)}
    .cb-tc-card{padding:4px 5px 5px;border-width:1.5px;border-radius:10px}
    .cb-tc-portrait{height:47px}
    .cb-tc-monogram{font-size:1.5rem}
    .cb-tc-name{font-size:.76rem;margin-top:3px}
    .cb-tc-meta{font-size:.60rem}
    .cb-tc-score{font-size:1rem;margin-top:1px}
    .cb-tc-bench{gap:6px}
    .cb-tc-bench-title{font-size:.97rem;margin-bottom:1px}
    .cb-tc-reserve{grid-template-columns:49px minmax(0,1fr) auto;min-height:70px;padding:4px 9px}
    .cb-tc-reserve .cb-tc-portrait{height:53px}
    .cb-tc-reserve .cb-tc-name{font-size:.84rem}
    .cb-tc-reserve .cb-tc-meta{font-size:.66rem}
    .cb-tc-reserve .cb-tc-score{font-size:1rem}
    /* v4.2.2: compact team dashboard, hide initials when portrait loads */
    .cb-tc-photo + .cb-tc-monogram{display:none}
    .cb-tc-shell{padding:10px 14px;margin:8px 0 12px}
    .cb-tc-court{min-height:405px;max-width:460px}
    .cb-tc-slot{width:clamp(96px,23%,124px)}
    .cb-tc-portrait{height:43px}
    .cb-tc-reserve{min-height:62px;padding:3px 8px}
    .cb-tc-reserve .cb-tc-portrait{height:46px}
    .cb-tc-bench{gap:5px}
    @media(max-width:900px){.cb-tc-layout{grid-template-columns:1fr}.cb-tc-court{min-height:490px}.cb-tc-bench{display:grid;grid-template-columns:repeat(2,minmax(0,1fr))}.cb-tc-bench-title{grid-column:1/-1}}
    @media(max-width:560px){.cb-tc-court{min-height:450px}.cb-tc-slot{width:29%}.cb-tc-portrait{height:40px}.cb-tc-meta{font-size:.55rem}.cb-tc-bench{grid-template-columns:1fr}}
    </style>""",unsafe_allow_html=True)


# Verified NBA person IDs, deliberately keyed by full player name rather than
# Fantasy element ID (the two numbering systems are unrelated).
# Extend this dictionary when additional identities are confirmed.
CRYSTAL_NBA_HEADSHOTS = {
    'lebron james':2544, 'james harden':201935, 'stephen curry':201939,
    'kevin durant':201142, 'nikola jokic':203999, 'giannis antetokounmpo':203507,
    'victor wembanyama':1641705, 'shai gilgeous-alexander':1628983,
    'trae young':1629027, 'damian lillard':203081, 'tyrese haliburton':1630169,
    'tyrese maxey':1630178, 'anthony edwards':1630162, 'luka doncic':1629029,
    'jayson tatum':1628369, 'jaylen brown':1627759, 'devin booker':1626164,
    'donovan mitchell':1628378, 'ja morant':1629630, "de'aaron fox":1628368,
    'domantas sabonis':1627734, 'bam adebayo':1628389, 'karl-anthony towns':1626157,
    'anthony davis':203076, 'kyrie irving':202681, 'paolo banchero':1631094,
    'chet holmgren':1631096, 'jalen brunson':1628973, 'jalen williams':1631114,
    'lamelo ball':1630163, 'scottie barnes':1630567, 'cade cunningham':1630595,
    'franz wagner':1630532, 'evan mobley':1630596, 'jarrett allen':1628386,
    'myles turner':1626167, 'pascal siakam':1627783, 'jimmy butler':202710,
    'brandon ingram':1627742, 'zion williamson':1629627, 'jaren jackson jr.':1628991,
    'desmond bane':1630217, 'rudy gobert':203497, 'draymond green':203110,
    'kawhi leonard':202695, 'paul george':202331, 'bradley beal':203078,
    'nikola vucevic':202696, 'chris paul':101108, 'klay thompson':202691,
    'austin reaves':1630559, 'darius garland':1629636, 'coby white':1629632,
    'tyler herro':1629639, 'jerami grant':203924, 'keegan murray':1631099,
    'jeremy sochan':1631110, 'thomas bryant':1628418, 'lonnie walker iv':1629022,
    'dorian finney-smith':1627827, 'george payton ii':1627780,
}


def crystal_player_headshot_id(p):
    for key in ('nba_id','nba_player_id','nba_person_id'):
        value=p.get(key)
        if value is not None and str(value).isdigit():
            return int(value)
    full=str(p.get('first_name') or '').strip()+' '+str(p.get('second_name') or '').strip()
    full=full.strip().casefold()
    if full in CRYSTAL_NBA_HEADSHOTS:
        return CRYSTAL_NBA_HEADSHOTS[full]
    # Only unambiguous full names; never guess from abbreviated web_name.
    web=str(p.get('web_name') or '').strip().casefold()
    if ' ' in web and web in CRYSTAL_NBA_HEADSHOTS:
        return CRYSTAL_NBA_HEADSHOTS[web]
    return None


def crystal_court(result,teams,type_names,events,currency):
    """Portrait court with photo cards, using verified NBA person IDs only."""
    selected=result.get('selected') or []
    if not selected:return
    lineups=result.get('lineups') or {}
    event=next((e for e in events if e['id'] in lineups),None)
    starters=list(lineups.get(event['id'],[])) if event else []
    if len(starters)!=5:
        starters=sorted(selected,key=lambda p:-num(p.get('_proj')))[:5]
    starter_ids={int(p['id']) for p in starters}
    bench=[p for p in selected if int(p['id']) not in starter_ids]

    def card(p, reserve=False):
        pos=escape(str(type_names.get(int(p['element_type']),'?')))
        name=escape(str(p.get('web_name') or 'Speler'))
        team=escape(str(teams.get(int(p['team']),{}).get('short_name','?')))
        price=num(p.get('now_cost'))/currency
        proj=num(p.get('_proj'))
        nba_id=crystal_player_headshot_id(p)
        initials=''.join(part[:1] for part in str(p.get('web_name') or '?').replace('.',' ').split()[:2]).upper()
        fallback=f'<span class="cb-tc-monogram">{escape(initials)}</span>'
        if nba_id:
            # The monogram remains visible if the image request fails.
            photo=(f'<img class="cb-tc-photo" src="https://cdn.nba.com/headshots/nba/latest/1040x760/{nba_id}.png" '
                   f'loading="lazy" alt="Portret van {name}" '
                   f'onerror="this.style.display=\'none\'" />'+fallback)
        else:
            photo=fallback
        return (f'<div class="cb-tc-card {"cb-tc-reserve" if reserve else ""}">'
                f'<div class="cb-tc-portrait">{photo}</div>'
                f'<div class="cb-tc-name" title="{name}">{name}</div>'
                f'<div class="cb-tc-meta">{team} · {pos} · ${price:.1f}m</div>'
                f'<div class="cb-tc-score">{proj:.1f} <small>PTS</small></div></div>')

    # FC above, BC below, preserving the actual solver-chosen formation.
    fc=sorted((p for p in starters if int(p['element_type'])!=1),key=lambda p:-num(p.get('_proj')))
    bc=sorted((p for p in starters if int(p['element_type'])==1),key=lambda p:-num(p.get('_proj')))
    ordered=fc+bc
    coords=([(50,16),(23,34),(77,34)] if len(fc)==3 else [(30,21),(70,21)]) + ([(30,73),(70,73)] if len(bc)==2 else [(18,72),(50,78),(82,72)])
    placed=''.join(f'<div class="cb-tc-slot" style="left:{x}%;top:{y}%">{card(p)}</div>' for p,(x,y) in zip(ordered,coords))
    bench_html=''.join(card(p,True) for p in sorted(bench,key=lambda p:(int(p['element_type']),-num(p.get('_proj')))))
    heading='🏀 Starting 5'+(f' · {escape(event.get("name", ""))}' if event else '')
    st.markdown('<div class="cb-tc-shell"><div class="cb-section">'+heading+'</div>'
      '<div class="cb-tc-layout"><div class="cb-tc-court"><div class="cb-tc-arc"></div>'
      '<div class="cb-tc-paint"></div><div class="cb-tc-hoop"></div>'+placed+'</div>'
      '<div class="cb-tc-bench"><div class="cb-tc-bench-title">🪑 Bank · 5 reserves</div>'+bench_html+
      '</div></div></div>',unsafe_allow_html=True)

# Startup trace is diagnostic only; it does not change projections or caching.
import threading, traceback, faulthandler
_CB_TRACE_PATH = Path(__file__).with_name('crystal_ball_startup_trace.log')

def cb_trace(stage, start=None):
    elapsed = '' if start is None else f' | +{time.perf_counter()-start:.2f}s'
    line = f"[Crystal Ball] {datetime.now().strftime('%H:%M:%S')} | {stage}{elapsed}"
    print(line, flush=True)
    try:
        with _CB_TRACE_PATH.open('a', encoding='utf-8') as fh:
            fh.write(line+'\n')
    except OSError:
        pass

def cb_watchdog(start, stop_event):
    if stop_event.wait(20):
        return
    cb_trace('WAARSCHUWING: opstart duurt >20s; Python-stack volgt', start)
    try:
        with _CB_TRACE_PATH.open('a', encoding='utf-8') as fh:
            faulthandler.dump_traceback(file=fh, all_threads=True)
    except Exception as exc:
        cb_trace(f'Stack kon niet worden opgeslagen: {exc}')

def app():
    global RISK_STRATEGY
    st.set_page_config(page_title='Crystal Ball | NBA Fantasy',page_icon='🔮',layout='wide')
    crystal_style()
    _page_start=time.perf_counter()
    cb_trace('START app()',_page_start)
    _cb_stop=threading.Event()
    threading.Thread(target=cb_watchdog,args=(_page_start,_cb_stop),daemon=True).start()
    st.markdown('<div class="cb-hero"><span class="cb-orb" role="img" aria-label="Kristallen basketbal"><svg viewBox="0 0 100 100" width="66" height="66" aria-hidden="true"><defs><radialGradient id="cbGlass" cx="30%" cy="22%" r="85%"><stop offset="0" stop-color="#f0e5ff"/><stop offset=".27" stop-color="#a7d7ff"/><stop offset=".66" stop-color="#7655ee"/><stop offset="1" stop-color="#281456"/></radialGradient></defs><circle cx="50" cy="50" r="43" fill="url(#cbGlass)" stroke="#b7e9ff" stroke-width="2"/><g fill="none" stroke="#e5e4ff" stroke-width="2.4" opacity=".85"><path d="M50 7 V93 M7 50 H93 M20 20 Q79 50 20 80 M80 20 Q21 50 80 80"/></g><ellipse cx="35" cy="24" rx="19" ry="9" fill="#fff" opacity=".18" transform="rotate(-30 35 24)"/></svg></span><div><h1>Crystal Ball</h1><p>FANTASY NBA OPTIMIZER <span class="cb-tag">v4.4</span></p></div></div>',unsafe_allow_html=True)
    st.caption('🔮 Live NBA Fantasy • statistische projecties • 4 Gameweeks • lokale SQLite • optimaliseren op verzoek')
    # Create main navigation immediately under the hero. The team placeholder is
    # filled later, but remains first in visual order.
    tab_team_main,tab_projection_main,tab_backtest_main,tab_inspector_main,tab_diag_main=st.tabs(['🏀 Team','🧪 Projection Lab','📊 Historical Backtest Lab','🔬 Projectie Inspector','⚙️ Technische Diagnose'])
    team_dashboard=tab_team_main.container()

    with tab_diag_main:
        st.info('**Voor technische controle.** Hier staan de modelversie, historische databronnen, cache, laadtijden en solverstatistieken. Je hoeft dit niet te gebruiken om een team te kiezen.')
    with tab_diag_main, st.expander('ℹ️ Versie en model',expanded=False):
        st.caption(f'BUILD {BUILD_ID} • rekenmodel ongewijzigd ten opzichte van v3.18.3')

    if 'result260' not in st.session_state: st.session_state.result260=None
    if 'strategy_compare318' not in st.session_state: st.session_state.strategy_compare318=None
    if 'curve260' not in st.session_state: st.session_state.curve260=None
    if 'compare260' not in st.session_state: st.session_state.compare260=None
    if 'scenarios300' not in st.session_state: st.session_state.scenarios300=None
    if 'scenario_warning300' not in st.session_state: st.session_state.scenario_warning300=[]
    if 'team_eval381' not in st.session_state: st.session_state.team_eval381=None
    if 'season_plan314' not in st.session_state: st.session_state.season_plan314=None

    _load_start=time.perf_counter()
    try:data=load_live()
    except Exception as exc:
        st.error(f'NBA Fantasy API kon niet worden geladen: {exc}'); st.stop()

    cb_trace('NBA API geladen',_page_start)
    _api_seconds=time.perf_counter()-_load_start
    rules=data.get('game_config',{}).get('rules') or data.get('game_settings',{})
    currency=int(rules.get('ui_currency_multiplier',10) or 10)
    players=data['elements']
    teams={int(t['id']):t for t in data['teams']}
    types=data['element_types']
    type_names={int(t['id']):t.get('singular_name_short') or t.get('singular_name') for t in types}

    # Historical loaders touch SQLite and may download source files. Do not repeat
    # them on every multiselect/slider rerun; manual refresh remains available.
    if st.sidebar.button('↻ Historische data verversen',key='cb_refresh_history'):
        st.session_state.pop('cb_history_status',None)
        fantasy_history_by_key.cache_clear()
        history_by_key.cache_clear()
        _fantasy_history_index.cache_clear()
        _durability_history_index.cache_clear()
        st.session_state.pop('cb_projection_cache',None)
        st.session_state.cb_force_history_refresh=True
    _history_start=time.perf_counter()
    with st.status('Data voorbereiden…',expanded=False) as ds:
        if 'cb_history_status' not in st.session_state:
            if st.session_state.pop('cb_force_history_refresh',False):
                st.session_state.cb_history_status=(load_three_year_durability(players,force=True),load_three_year_fantasy())
            else:
                st.session_state.cb_history_status=(local_durability_status(),local_fantasy_status())
        durability_status,fantasy_status=st.session_state.cb_history_status
        ok_years=[y for y,v in durability_status.items() if v['ok']]
        if len(ok_years)==3:
            ds.write('✓ 3-jaars durability geladen: '+', '.join(str(y) for y in ok_years))
        else:
            ds.write('⚠ Durability onvolledig: '+', '.join(str(y) for y in ok_years) if ok_years else '⚠ Geen durability-seizoenen geladen')
        fp_years=[y for y,v in fantasy_status.items() if v['ok']]
        ds.write('✓ Multi-season fantasy: '+(', '.join(str(y) for y in fp_years) if fp_years else 'API fallback actief'))
        stamp=None  # Snapshot only on explicit optimization, not on UI interactions
        ds.write(f'✓ {len(players)} spelers geladen')
        ds.update(label='Data gereed',state='complete',expanded=False)

    # Make missing historical data actionable without blocking normal startup.
    if not any(v.get('count',0) for v in durability_status.values()):
        with tab_diag_main: st.warning('Historische durability ontbreekt in de lokale database. Projecties gebruiken tijdelijk een fallback. Klik links op **Historische data verversen** om de seizoenen eenmalig te downloaden en op te slaan; dit kan even duren.')
    if not any(v.get('count',0) for v in fantasy_status.values()):
        with tab_diag_main: st.info('Ook de gereconstrueerde fantasyhistorie ontbreekt lokaal. De NBA Fantasy API blijft beschikbaar, maar dit is niet dezelfde databasis als een volledige 3-seizoenenprojectie.')
    _history_seconds=time.perf_counter()-_history_start
    cb_trace('Historische status gereed; start coverage',_page_start)
    coverage=durability_coverage(players)
    cov_text=' • '.join(
        f"{'2023-24' if y==2024 else '2024-25' if y==2025 else '2025-26'}: {v['matched']}/{v['total']} ({v['pct']:.1f}%)"
        for y,v in coverage.items())
    cb_trace('Coverage gereed; start 3Y check',_page_start)
    complete3=sum(1 for p in players if len(three_year_games(p))==3)
    complete_pct=100*complete3/len(players) if players else 0.0
    cov_text += f" • volledige 3Y: {complete3}/{len(players)} ({complete_pct:.1f}%)"
    cb_trace('3Y check gereed',_page_start)
    with tab_diag_main:
        if all(v['pct']>=80 for v in coverage.values()):
            st.success('✓ Historische durability actief — '+cov_text)
        else:
            st.warning('⚠️ Durability-dekking — '+cov_text+
                       '. Ontbrekende speler/seizoenen blijven zichtbaar als onvolledige historie.')
        with st.expander('Durability-bronnen'):
            for y,v in durability_status.items():
                cv=coverage[y]
                st.write(f"{y}: {'✓' if v['ok'] else '✕'} bronregels {v['count']} • huidige spelers {cv['matched']}/{cv['total']} ({cv['pct']:.1f}%) • {v['source']}")
                if v.get('error') and not v['ok']: st.caption(v['error'])
            st.markdown('**Multi-season fantasy-projectie**')
            for y,v in fantasy_status.items():
                st.write(f"{y}: {'✓' if v['ok'] else '⚠'} {v['count']} regels • {v['source']}")
                if v.get('error') and not v['ok']: st.caption(v['error'])

    global CURRENT_SEASON_STARTED, CURRENT_FINISHED_GAMEDAYS
    finished_events=[e for e in data['events'] if e.get('finished')]
    CURRENT_FINISHED_GAMEDAYS=len(finished_events)
    CURRENT_SEASON_STARTED=CURRENT_FINISHED_GAMEDAYS>0

    nxt=next((e for e in data['events'] if e.get('is_next')),None) or next((e for e in data['events'] if not e.get('finished')),None)
    if not nxt: st.warning('Geen volgende Gameday gevonden.'); st.stop()
    first_gw=parse_gw(nxt['name'])
    target=set(range(first_gw,first_gw+HORIZON_GWS))
    events=[e for e in data['events'] if parse_gw(e.get('name')) in target and not e.get('finished')]
    cb_trace('Start fixtures',_page_start)
    fixtures=load_fixtures(tuple(e['id'] for e in events))
    cb_trace('Fixtures gereed',_page_start)
    active={eid:{int(f[k]) for f in fs for k in ('team_h','team_a') if f.get(k) is not None} for eid,fs in fixtures.items()}
    b2b=build_b2b(events,active)
    weights={e['id']:GW_DECAY**max(0,(parse_gw(e['name']) or first_gw)-first_gw) for e in events}

    selectable=[p for p in players if p.get('can_select',True) and not p.get('removed',False)]
    cb_trace('Start market prior',_page_start)
    configure_no_history_market_prior(selectable,currency)
    cb_trace('Market prior gereed',_page_start)
    # Normalize both supported representations to {team_id: team_dict}.
    if isinstance(teams, dict):
        team_by_id={int(k):v for k,v in teams.items()}
    else:
        team_by_id={int(t['id']):t for t in teams}
    # Cache expensive base projections between widget-only reruns. Recompute when
    # the API snapshot, historical refresh or risk strategy changes.
    _projection_start=time.perf_counter()
    _proj_signature=(RISK_STRATEGY,currency,CURRENT_FINISHED_GAMEDAYS,
        tuple((int(p['id']),p.get('now_cost'),p.get('form'),p.get('minutes'),
               p.get('points_per_game'),p.get('status'),p.get('chance_of_playing_next_round'))
              for p in selectable),id(st.session_state.get('cb_history_status')))
    _cached=st.session_state.get('cb_projection_cache')
    if _cached is None or _cached[0]!=_proj_signature:
        _proj_values={int(p['id']):final_projection(p,currency,1.0) for p in selectable}
        st.session_state.cb_projection_cache=(_proj_signature,_proj_values)
        _proj_cache_hit=False
    else:
        _proj_values=_cached[1]
        _proj_cache_hit=True
    for p in selectable:
        p['_proj']=_proj_values[int(p['id'])]
    cb_trace('Projecties gereed',_page_start)
    _projection_seconds=time.perf_counter()-_projection_start

    st.sidebar.header('Optimalisatie')
    RISK_STRATEGY=st.sidebar.radio('Teamstrategie',['Safe','Balanced','Aggressive'],index=1,horizontal=True)
    st.sidebar.caption('Safe = betrouwbaarheid · Balanced = gemiddelde projectie · Aggressive = upside')
    mode=st.sidebar.radio('Modus',['🌱 Start seizoen','🔄 Season Manager'],horizontal=False)
    max_budget=num(rules.get('squad_total_spend'),1000)/currency
    name_to_id={f"{p['web_name']} — {teams[int(p['team'])]['short_name']} — ${num(p['now_cost'])/currency:.1f}m":p['id'] for p in selectable}
    id_to_name={v:k for k,v in name_to_id.items()}

    # v3.17.2: profile controls live in a Streamlit fragment. Saving/deleting a profile
    # reruns only this tiny sidebar fragment, not the expensive API/history preparation above.
    @st.fragment
    def render_team_profiles():
        st.sidebar.markdown('### 💾 Teamprofielen')
        profiles=list_team_profiles()
        selected_profile=st.sidebar.selectbox('Profiel', ['— kies profiel —']+profiles, key='profile_select317')
        pc1,pc2=st.sidebar.columns(2)
        if pc1.button('📂 Laden',use_container_width=True,key='profile_load317'):
            if selected_profile.startswith('—'):
                st.sidebar.warning('Kies eerst een opgeslagen profiel.')
                prof=None
            else:
                prof=load_team_profile(selected_profile)
            if prof:
                st.session_state['locked_names317']=[id_to_name[i] for i in prof['roster'] if i in id_to_name][:10]
                st.session_state['excluded_names317']=[id_to_name[i] for i in prof['excluded'] if i in id_to_name][:30]
                st.session_state['budget_cap317']=min(float(max_budget),max(70.0,float(prof.get('budget') or max_budget)))
                st.session_state['tdc_free_transfers']=int(prof.get('free_transfers') if prof.get('free_transfers') is not None else 2)
                st.session_state['tdc_allow_paid']=bool(prof.get('allow_paid',False))
                st.session_state['tdc_max_transfers']=int(prof.get('max_test_transfers') or max(0,int(st.session_state['tdc_free_transfers'])))
                st.session_state['profile_name317']=selected_profile
                st.session_state['profile_flash3172']=f'Profiel {selected_profile} geladen.'
                st.session_state.result260=None; st.session_state.team_eval381=None; st.session_state.season_plan314=None
                # IMPORTANT: this button lives inside a fragment, while the lock/exclusion widgets
                # live in the full app. Force an APP rerun so the restored session-state values are
                # consumed by those widgets immediately.
                st.rerun(scope='app')
        if pc2.button('🗑️ Wissen',use_container_width=True,key='profile_delete317'):
            if selected_profile.startswith('—'):
                st.sidebar.warning('Kies eerst een opgeslagen profiel.')
                return
            delete_team_profile(selected_profile)
            st.session_state['profile_flash3172']=f'Profiel {selected_profile} gewist.'
            st.session_state['profile_select317']='— kies profiel —'
            st.rerun(scope='fragment')

        profile_name=st.sidebar.text_input('Naam voor opslaan',value=st.session_state.get('profile_name317','Mijn team'),key='profile_name_input317')
        if st.sidebar.button('💾 Opslaan',use_container_width=True,key='profile_save317_top'):
            _locked_names=list(st.session_state.get('locked_names317',[]) or [])
            _locked_ids=[name_to_id[n] for n in _locked_names if n in name_to_id]
            _excluded_names=list(st.session_state.get('excluded_names317',[]) or [])
            _excluded_ids=[name_to_id[n] for n in _excluded_names if n in name_to_id]
            if len(_locked_ids)!=10:
                st.sidebar.error('Selecteer precies 10 spelers om een teamprofiel op te slaan.')
            else:
                try:
                    _budget=float(st.session_state.get('budget_cap317',max_budget))
                    _free=int(st.session_state.get('tdc_free_transfers',2))
                    _paid=bool(st.session_state.get('tdc_allow_paid',False))
                    _maxt=int(st.session_state.get('tdc_max_transfers',max(0,_free))) if _paid else max(0,_free)
                    save_team_profile(profile_name,_locked_ids,_excluded_ids,_budget,_free,_paid,_maxt)
                    clean_name=profile_name.strip()
                    st.session_state['profile_name317']=clean_name
                    st.session_state['profile_flash3172']=f'Profiel {clean_name} opgeslagen.'
                    # Do not assign profile_select317 here: Streamlit forbids changing a widget key
                    # after that widget has already been instantiated in the same fragment run.
                    st.rerun(scope='fragment')
                except Exception as ex:
                    st.sidebar.error(str(ex))
        flash=st.session_state.pop('profile_flash3172',None)
        if flash:
            st.sidebar.success(flash)

    cb_trace('Start profielwidgets',_page_start)
    st.sidebar.caption('🌐 Friends Edition: profielen blijven alleen tijdens deze browsersessie bewaard.')
    render_team_profiles()
    cb_trace('Profielwidgets gereed',_page_start)
    if 'budget_cap317' not in st.session_state:
        st.session_state['budget_cap317']=float(max_budget)
    budget=st.sidebar.slider('Maximaal budget ($m)',70.0,float(max_budget),step=.5,key='budget_cap317')
    current_names=[]
    free_transfers=2
    allow_paid=False
    max_test_transfers=4
    if mode.startswith('🔄'):
        current_names=st.sidebar.multiselect('🏀 Huidige 10 spelers',list(name_to_id),
            placeholder='Selecteer je huidige selectie', max_selections=10)
        if current_names and len(current_names)!=10:
            st.sidebar.warning(f'Huidige selectie: {len(current_names)}/10 spelers')
    st.sidebar.markdown('### 👤 Mijn team / locks')
    locked_names=st.sidebar.multiselect('🔒 Mijn team / spelers vastzetten',list(name_to_id),placeholder='Kies maximaal 10 spelers',max_selections=10,key='locked_names317')
    st.sidebar.caption(f'Mijn team: {len(locked_names)}/10 spelers')
    st.sidebar.markdown('### 🚫 Uitsluiten')
    excluded_names=st.sidebar.multiselect('✕ Spelers uitsluiten',list(name_to_id),placeholder='Alleen spelers die je niet wilt',max_selections=30,key='excluded_names317')
    st.sidebar.caption('Veilige selectie: bulk-selectie is begrensd om per ongeluk alles selecteren te voorkomen.')
    locked=[name_to_id[n] for n in locked_names]; excluded=[name_to_id[n] for n in excluded_names]
    current_roster=[name_to_id[n] for n in current_names]
    # UX: een complete 10-player lock-selectie mag direct als huidig team dienen.
    # Zo hoef je dezelfde selectie niet opnieuw in Season Manager in te voeren.
    effective_roster = current_roster if len(current_roster)==10 else (locked if len(locked)==10 else [])
    # v3.15.3: Transfer Decision Center controls are always visible.
    # This also supports Start seizoen + 10 locks, which is how hand-picked teams are often evaluated.
    st.sidebar.markdown('### 🔄 Transfer Decision Center')
    free_transfers=st.sidebar.number_input('Gratis transfers',min_value=0,max_value=10,value=2,step=1,key='tdc_free_transfers')
    allow_paid=st.sidebar.checkbox('Betaalde transfers testen (-100 pt per extra transfer)',value=False,key='tdc_allow_paid')
    if allow_paid:
        min_max=max(1,int(free_transfers)+1)
        max_test_transfers=st.sidebar.number_input('Maximaal aantal transfers testen',min_value=min_max,max_value=10,value=max(4,min_max),step=1,key='tdc_max_transfers')
    else:
        max_test_transfers=max(0,int(free_transfers))
    st.sidebar.caption('Profielen worden lokaal in SQLite bewaard en blijven bestaan wanneer je alleen het .py-bestand vervangt.')
    st.sidebar.caption('Transferplan vergelijkt HOLD met elk toegestaan transferplafond en kiest op netto objective na eventuele puntenstraf.')
    if locked_names:
        st.sidebar.caption('🔒 Locked: '+', '.join(n.split(' — ')[0] for n in locked_names))
    if set(locked)&set(excluded):
        st.sidebar.error('Een speler kan niet tegelijk vastgezet en uitgesloten zijn.'); st.stop()
    if mode.startswith('🔄') and len(effective_roster)!=10:
        st.info('Season Manager: selecteer precies 10 huidige spelers óf gebruik 10 spelers onder Mijn team / locks.')
    elif len(current_roster)!=10 and len(locked)==10:
        st.caption('🔄 Transferplan gebruikt automatisch je 10 spelers uit Mijn team / locks als huidige selectie.')

    with tab_team_main:
        a,b,c,d,e,f=st.columns(6)
        run=a.button('🚀 Optimaliseer team',type='primary',use_container_width=True)
        season_plan=b.button('🔄 Transferplan',use_container_width=True,
                             disabled=len(effective_roster)!=10,
                             help='Gebruikt je 10 huidige Season Manager-spelers, of automatisch je 10 locks. Vergelijkt HOLD en transferopties.')
        evaluator=c.button('🩺 Team Evaluator',use_container_width=True,disabled=len(locked)!=10,
                           help='Lock precies je 10 hand-picked spelers. De evaluator vergelijkt ze met vrij optimum en beste 1/2/3 swaps.')
        scenarios=d.button('🧪 Premiumscenario’s',use_container_width=True)
        compare=e.button('⚖️ Vrij vs locks',use_container_width=True,disabled=not bool(locked))
        clear=f.button('Wis resultaat',use_container_width=True)
        strategy_compare=st.button('🎯 Vergelijk Safe / Balanced / Aggressive (vrij optimum)', use_container_width=True,
            help='Draait drie vrije optimalisaties met hetzelfde budget en dezelfde exclusions. Locks worden bewust genegeerd zodat je ziet wat de strategie werkelijk verandert.')
        if clear:
            st.session_state.result260=None; st.session_state.curve260=None; st.session_state.team_eval381=None; st.session_state.season_plan314=None; st.session_state.strategy_compare318=None; st.rerun()

        if strategy_compare:
            # Compare the three philosophies on a genuinely free optimum. A loaded profile often
            # populates 10 locks; respecting those locks would necessarily produce the same team.
            selected_strategy=RISK_STRATEGY
            comparison={}
            with st.status('Drie strategieën doorrekenen…',expanded=True) as cs:
                try:
                    for strat in ('Safe','Balanced','Aggressive'):
                        RISK_STRATEGY=strat
                        cs.write(f'{strat}: optimaliseren…')
                        rr=optimize(players,teams,types,rules,events,active,currency,weights,first_gw,b2b,budget,[],excluded)
                        rr['_budget_cap']=budget
                        comparison[strat]=rr
                    st.session_state.strategy_compare318=comparison
                    cs.update(label='Strategievergelijking gereed',state='complete',expanded=False)
                except Exception as exc:
                    cs.update(label='Strategievergelijking gestopt',state='error',expanded=True)
                    st.error(f'Vergelijken mislukt: {exc}')
                finally:
                    RISK_STRATEGY=selected_strategy

        cb_trace('UI klaar voor optimalisatie',_page_start)
        _cb_stop.set()
        _prep_seconds=time.perf_counter()-_page_start
        with tab_diag_main, st.expander('⏱️ Laadtijden · diagnose',expanded=True):
            st.caption(f'NBA API/cache: {_api_seconds:.2f}s · historische data: {_history_seconds:.2f}s · '
                       f'projecties: {_projection_seconds:.2f}s ({"cache" if _proj_cache_hit else "nieuw"}) · '
                       f'pagina tot optimalisatie: {_prep_seconds:.2f}s')
            st.caption('De CBC-solvertijd wordt apart getoond na het optimaliseren. '
                       'Netwerk, Streamlit-weergave en browser-rendering kunnen extra tijd kosten.')

        if run:
            t0=time.perf_counter()
            with st.status('Team optimaliseren…',expanded=True) as os:
                os.write('CBC gestart. 2-stage MILP + role/usage gestart: compacte shortlist + CBC (20s, 3% gap).')
                try:
                    result=optimize(players,teams,types,rules,events,active,currency,weights,first_gw,b2b,budget,locked,excluded,
                                    current_roster if mode.startswith('🔄') and len(current_roster)==10 else None,
                                    free_transfers,allow_paid)
                    result['_elapsed']=time.perf_counter()-t0
                    result['_budget_cap']=budget
                    st.session_state.result260=result
                    save_run(result,'Max punten',budget)
                    save_snapshot(players,currency)
                    os.write(f"✓ {result['status']} in {result['_elapsed']:.1f} sec.")
                    os.update(label='Optimalisatie gereed',state='complete',expanded=False)
                except Exception as exc:
                    os.update(label='Optimalisatie gestopt',state='error',expanded=True)
                    st.error(f'Geen resultaat: {exc}')


        if season_plan:
            roster_ids=set(effective_roster)
            by_id={p['id']:p for p in selectable}
            with st.status('Season Manager 2.0 — transferopties doorrekenen…',expanded=True) as sm:
                try:
                    t_sm=time.perf_counter()
                    # One shared candidate universe makes HOLD/1/2/paid directly comparable.
                    seed=optimize(players,teams,types,rules,events,active,currency,weights,first_gw,b2b,budget,
                                  [],excluded,None,2,False)
                    fixed_pool=set(seed.get('candidate_ids',[])) | roster_ids
                    hold=optimize(players,teams,types,rules,events,active,currency,weights,first_gw,b2b,budget,
                                  list(roster_ids),excluded,None,2,False,fixed_pool)
                    plans=[]
                    # Decision Center: compare exact transfer counts on the same pool.
                    # Without paid transfers we stop at the number of free transfers (max 2 is useful in NBA Fantasy).
                    # Paid scenarios subtract 100 points for every exact transfer above free_transfers.
                    max_limit = int(max_test_transfers) if allow_paid else int(free_transfers)
                    for limit in range(1, max_limit+1):
                        paid_scenario = bool(allow_paid and limit > int(free_transfers))
                        rr=optimize(players,teams,types,rules,events,active,currency,weights,first_gw,b2b,budget,
                                    [],excluded,roster_ids,free_transfers,paid_scenario,fixed_pool,
                                    exact_transfers=limit)
                        label=f'Exact {limit} transfer' + ('' if limit==1 else 's')
                        if paid_scenario: label += f' (-{100*(limit-int(free_transfers))} pt)'
                        plans.append({'label':label,'limit':limit,'paid':paid_scenario,'result':rr})
                        penalty=rr.get('transfer_penalty',0.0)
                        sm.write(f"✓ {label}: netto {rr['objective']:.1f}" + (f" • kosten {penalty:.0f} pt" if penalty else ''))
                    st.session_state.season_plan314={'hold':hold,'plans':plans,'budget':budget,
                                                     'free_transfers':free_transfers,'allow_paid':allow_paid,'max_test_transfers':max_test_transfers,'pool':len(fixed_pool),
                                                     'elapsed':time.perf_counter()-t_sm,
                                                     'roster_ids':list(roster_ids)}
                    sm.update(label='Transferplan gereed',state='complete',expanded=False)
                except Exception as exc:
                    st.session_state.season_plan314=None
                    sm.update(label='Transferplan gestopt',state='error',expanded=True)
                    st.error(f'Transferplan kon niet afronden: {exc}')


        # v3.15.3: Transfer Decision Center — exact transfer counts, direct onder de actieknoppen.
        if st.session_state.season_plan314:
            sp=st.session_state.season_plan314
            current_tdc=(int(free_transfers), bool(allow_paid), int(max_test_transfers))
            result_tdc=(int(sp.get('free_transfers',2)), bool(sp.get('allow_paid',False)), int(sp.get('max_test_transfers',sp.get('free_transfers',2))))
            if current_tdc != result_tdc:
                st.warning('⚠️ Transferinstellingen zijn gewijzigd sinds deze berekening. Klik opnieuw op **Transferplan**. Het oude resultaat wordt bewust verborgen zodat AAN/UIT nooit meer naast een verouderde tabel staat.')
            else:
                hold=sp['hold']; names_by_id={p['id']:p['web_name'] for p in selectable}
                st.subheader('🔄 Transfer Decision Center')
                st.caption(f"Zelfde ${sp['budget']:.1f}m budget, 4GW-horizon en {sp['pool']} kandidaten • {sp['free_transfers']} gratis transfer(s) • betaalde transfers: {'AAN t/m '+str(sp.get('max_test_transfers')) if sp.get('allow_paid') else 'UIT'} • evaluatie {sp['elapsed']:.1f}s. Punten zijn modelverwachtingen.")
                rows=[{'Scenario':'HOLD','Transfers':0,'Bruto objective':round(hold.get('gross_objective',hold['objective']),1),
                       'Transferkosten':0,'Netto objective':round(hold['objective'],1),'Netto winst vs HOLD':0.0,
                       'OUT':'—','IN':'—','Budget $m':round(hold['spent'],1)}]
                scenario_items=[{'label':'HOLD','result':hold,'limit':0,'paid':False}]
                for item in sp['plans']:
                    rr=item['result']; outs=rr.get('transfers_out',[]); ins=rr.get('transfers_in',[])
                    penalty=float(rr.get('transfer_penalty',0.0) or 0.0)
                    rows.append({'Scenario':item['label'],'Transfers':len(ins),
                                 'Bruto objective':round(rr.get('gross_objective',rr['objective']+penalty),1),
                                 'Transferkosten':round(penalty,1),'Netto objective':round(rr['objective'],1),
                                 'Netto winst vs HOLD':round(rr['objective']-hold['objective'],1),
                                 'OUT':', '.join(names_by_id.get(pid,str(pid)) for pid in outs) or '—',
                                 'IN':', '.join(p['web_name'] for p in ins) or '—',
                                 'Budget $m':round(rr['spent'],1)})
                    scenario_items.append(item)
                html_table(rows,390)

                best=max(scenario_items,key=lambda x:x['result']['objective'])
                br=best['result']; actual_transfers=len(br.get('transfers_in',[]))
                bout=', '.join(names_by_id.get(pid,str(pid)) for pid in br.get('transfers_out',[])) or 'geen'
                bin_=', '.join(p['web_name'] for p in br.get('transfers_in',[])) or 'geen'
                gain=float(br['objective']-hold['objective']); penalty=float(br.get('transfer_penalty',0.0) or 0.0)
                paid_count=max(0,actual_transfers-int(sp['free_transfers']))

                # Marginal value of the recommended transfer count versus the best lower-transfer scenario.
                lower=[x for x in scenario_items if len(x['result'].get('transfers_in',[])) < actual_transfers]
                prev=max(lower,key=lambda x:x['result']['objective']) if lower else {'result':hold}
                marginal=float(br['objective']-prev['result']['objective'])

                if best['label']=='HOLD' or gain <= 0.01:
                    st.success('Modeladvies: HOLD — geen doorgerekende transferoptie verbetert de netto 4GW-objective.')
                else:
                    free_used=min(actual_transfers,int(sp['free_transfers']))
                    paid_txt=f' + {paid_count} betaalde' if paid_count else ''
                    st.success(f"Modeladvies: {actual_transfers} transfer(s) ({free_used} gratis{paid_txt}) • OUT: {bout} • IN: {bin_} • netto +{gain:.1f} pt vs HOLD")
                    c1,c2,c3,c4=st.columns(4)
                    c1.metric('Netto winst vs HOLD',f'+{gain:.1f} pt')
                    c2.metric('Transferkosten',f'-{penalty:.0f} pt' if penalty else '0 pt')
                    c3.metric('Marginale winst laatste stap',f'+{marginal:.1f} pt')
                    c4.metric('Transfers',f'{actual_transfers} ({paid_count} betaald)')
                    if paid_count:
                        gross_gain=(br.get('gross_objective',br['objective']+penalty)-hold.get('gross_objective',hold['objective']))
                        st.info(f'Betaalde-transfercheck: bruto verbetering +{gross_gain:.1f} pt, straf -{penalty:.0f} pt → netto +{gain:.1f} pt vs HOLD.')
                    elif actual_transfers < int(sp['free_transfers']):
                        st.info(f'Je hebt {sp["free_transfers"]} gratis transfers, maar het model gebruikt er slechts {actual_transfers}: een extra wijziging voegt binnen deze 4GW-horizon geen netto waarde toe.')
                st.caption('Decision Center vergelijkt HOLD met exact 1, 2, 3, ... transfers op exact dezelfde kandidaatpool. Betaalde transfers kosten 100 punten per transfer boven het ingestelde aantal gratis transfers. Verkoopwaarde/sell-on wordt nog niet individueel gereconstrueerd; budget gebruikt de huidige NBA Fantasy-prijzen.')


                # v3.16.0 — Decision Dashboard: turn the winning transfer scenario into an action plan.
                st.markdown('---')
                st.subheader('🎯 Decision Dashboard')
                st.caption('Van transferbeslissing naar uitvoering: huidige selectie → OUT/IN → nieuwe selectie → Starting 5 + captain per Gameday → verwachte winst per Gameweek.')

                # Team before/after, always based on the exact HOLD and recommended scenario above.
                hold_ids={int(p['id']) for p in hold.get('selected',[])}
                rec_ids={int(p['id']) for p in br.get('selected',[])}
                action_out=sorted(hold_ids-rec_ids,key=lambda pid:names_by_id.get(pid,str(pid)))
                action_in=sorted(rec_ids-hold_ids,key=lambda pid:names_by_id.get(pid,str(pid)))
                d1,d2,d3,d4=st.columns(4)
                d1.metric('Actie', 'HOLD' if not action_in else f'{len(action_in)} transfer(s)')
                d2.metric('4GW netto winst', f'{gain:+.1f} pt')
                d3.metric('Nieuw budget', f"${br['spent']:.1f}m")
                d4.metric('Transferstraf', f'-{penalty:.0f} pt' if penalty else '0 pt')
                if action_in:
                    st.success('UIT: '+', '.join(names_by_id.get(pid,str(pid)) for pid in action_out)+'  →  IN: '+', '.join(names_by_id.get(pid,str(pid)) for pid in action_in))
                else:
                    st.success('Geen transfer uitvoeren: HOLD is het beste netto scenario binnen de ingestelde regels.')

                by_player_id={int(p['id']):p for p in selectable}
                before_rows=[]; after_rows=[]
                for pid in sorted(hold_ids,key=lambda x:(int(by_player_id[x]['element_type']),-by_player_id[x]['_proj'])):
                    pp=by_player_id[pid]
                    before_rows.append({'Speler':pp['web_name'],'Team':teams[int(pp['team'])]['short_name'],'Pos':type_names[int(pp['element_type'])],
                                        'Prijs':round(num(pp['now_cost'])/currency,1),'Projectie':round(pp['_proj'],2),
                                        'Status':'OUT' if pid in action_out else 'HOLD'})
                for pid in sorted(rec_ids,key=lambda x:(int(by_player_id[x]['element_type']),-by_player_id[x]['_proj'])):
                    pp=by_player_id[pid]
                    after_rows.append({'Speler':pp['web_name'],'Team':teams[int(pp['team'])]['short_name'],'Pos':type_names[int(pp['element_type'])],
                                       'Prijs':round(num(pp['now_cost'])/currency,1),'Projectie':round(pp['_proj'],2),
                                       'Status':'IN' if pid in action_in else 'HOLD'})
                left,right=st.columns(2)
                with left:
                    st.markdown('#### Huidige selectie')
                    html_table(before_rows,430)
                with right:
                    st.markdown('#### Selectie na advies')
                    html_table(after_rows,430)

                # Exact objective breakdown by Gameweek for HOLD and the recommended scenario.
                def _gw_breakdown(rr):
                    out={}
                    for ev in events:
                        eid=ev['id']; gw=parse_gw(ev.get('name'))
                        if gw is None: continue
                        w=weights.get(eid,1.0)
                        starters=rr.get('lineups',{}).get(eid,[])
                        start=sum((event_projection(pp,ev,first_gw,currency,b2b) if int(pp['team']) in active.get(eid,set()) else 0.0)*w for pp in starters)
                        cap=rr.get('captains',{}).get(eid)
                        capb=(event_projection(cap,ev,first_gw,currency,b2b)*w if cap is not None and int(cap['team']) in active.get(eid,set()) else 0.0)
                        z=out.setdefault(gw,{'start':0.0,'cap':0.0,'total':0.0})
                        z['start']+=start; z['cap']+=capb; z['total']+=start+capb
                    return out
                hb=_gw_breakdown(hold); rb=_gw_breakdown(br)
                gw_rows=[]
                for gw in sorted(set(hb)|set(rb)):
                    h=hb.get(gw,{'total':0.0}); r=rb.get(gw,{'total':0.0})
                    gw_rows.append({'Gameweek':gw,'HOLD':round(h['total'],1),'Na advies':round(r['total'],1),'Δ':round(r['total']-h['total'],1)})
                st.markdown('#### Verwachte impact per Gameweek')
                html_table(gw_rows,300)
                st.caption('Gameweek-punten hierboven gebruiken exact dezelfde 4GW-weging als de optimizer. De som van de verschillen is de bruto modelwinst; eventuele betaalde-transferstraf wordt daarna van de 4GW-objective afgetrokken.')

                st.markdown('#### Uitvoeringsplan per Gameday')
                for ev in events:
                    eid=ev['id']; starters=br.get('lineups',{}).get(eid,[]); cap=br.get('captains',{}).get(eid)
                    playing=[pp for pp in starters if int(pp['team']) in active.get(eid,set())]
                    idle=[pp for pp in starters if int(pp['team']) not in active.get(eid,set())]
                    bc=sum(1 for pp in starters if int(pp['element_type'])==1); fc=len(starters)-bc
                    pts=sum(event_projection(pp,ev,first_gw,currency,b2b) if int(pp['team']) in active.get(eid,set()) else 0.0 for pp in starters)
                    cappts=event_projection(cap,ev,first_gw,currency,b2b) if cap is not None and int(cap['team']) in active.get(eid,set()) else 0.0
                    capname=cap['web_name'] if cap else '—'
                    with st.expander(f"{ev.get('name','')} • {bc} BC + {fc} FC • {len(playing)}/5 met fixture • ⭐ {capname} • {pts+cappts:.1f} ongewogen model-FP"):
                        st.markdown('**Starting 5:** '+(', '.join(pp['web_name'] for pp in starters) or '—'))
                        st.markdown('**Captain:** '+capname)
                        if idle:
                            st.warning('Starter(s) zonder fixture (0 FP): '+', '.join(pp['web_name'] for pp in idle))
                        bench=[pp for pp in br.get('selected',[]) if int(pp['id']) not in {int(x['id']) for x in starters}]
                        st.caption('Bank: '+(', '.join(pp['web_name'] for pp in bench) or '—'))

        if evaluator:
            user_ids=set(locked)
            by_id={p['id']:p for p in selectable}
            with st.status('Hand-picked team benchmarken en repareren…',expanded=True) as es:
                try:
                    t_eval=time.perf_counter()
                    # Seed the same compact universe the normal optimizer uses, then force every
                    # comparison onto that identical universe for apples-to-apples objectives.
                    free_seed=optimize(players,teams,types,rules,events,active,currency,weights,first_gw,b2b,budget,[],excluded)
                    fixed_pool=set(free_seed.get('candidate_ids',[])) | user_ids
                    free=optimize(players,teams,types,rules,events,active,currency,weights,first_gw,b2b,budget,
                                  [],excluded,None,2,False,fixed_pool)
                    exact=optimize(players,teams,types,rules,events,active,currency,weights,first_gw,b2b,budget,
                                   list(user_ids),excluded,None,2,False,fixed_pool)
                    repairs=[]
                    for swaps,keep_n in ((1,9),(2,8),(3,7)):
                        rr=optimize(players,teams,types,rules,events,active,currency,weights,first_gw,b2b,budget,
                                    [],excluded,None,2,False,fixed_pool,user_ids,keep_n)
                        rr_ids={p['id'] for p in rr['selected']}
                        outs=sorted(user_ids-rr_ids,key=lambda pid:(by_id.get(pid,{}).get('web_name') or ''))
                        ins=sorted(rr_ids-user_ids,key=lambda pid:(by_id.get(pid,{}).get('web_name') or ''))
                        repairs.append({'swaps':swaps,'result':rr,'out':outs,'in':ins})
                        es.write(f'Beste max {swaps} swap(s): {rr["objective"]:.1f}')
                    st.session_state.team_eval381={'exact':exact,'free':free,'repairs':repairs,
                                                   'budget':budget,'pool':len(fixed_pool),
                                                   'elapsed':time.perf_counter()-t_eval}
                    es.update(label='Team Evaluator gereed',state='complete',expanded=False)
                except Exception as exc:
                    st.session_state.team_eval381=None
                    es.update(label='Team Evaluator gestopt',state='error',expanded=True)
                    st.error(f'Team Evaluator kon niet afronden: {exc}')

        if scenarios:
            # Dynamic premium set from current selectable pool. No hardcoded star names.
            role_all=role_adjustment_map(selectable,currency)
            ranked=[]
            for p in selectable:
                price=num(p.get('now_cost'))/currency
                if price < 12.0:
                    continue
                rp=final_projection(p,currency,role_all.get(p['id'],{}).get('factor',1.0))
                ranked.append((rp,price,p))
            ranked.sort(key=lambda x:(x[0],x[1]),reverse=True)
            premium=[x[2] for x in ranked[:20]]
            by_id={p['id']:p for p in selectable}
            rows=[]
            with st.status('Dynamische premiumscenario’s doorrekenen met één vaste kandidaatpool…',expanded=True) as ss:
                free_seed=optimize(players,teams,types,rules,events,active,currency,weights,first_gw,b2b,budget,[],excluded)
                fixed_pool=set(free_seed.get('candidate_ids',[]))
                fixed_pool.update(p['id'] for p in premium)
                # Also preserve every manual lock in the comparison universe.
                fixed_pool.update(locked)
                specs=[('Vrij',[])]
                for p in premium:
                    specs.append((p.get('web_name') or full_name(p),[p['id']]))
                # Re-run every scenario, including free, on the identical pool with exact CBC gap.
                for label,lks in specs:
                    r=optimize(players,teams,types,rules,events,active,currency,weights,first_gw,b2b,budget,
                               lks,excluded,None,2,False,fixed_pool)
                    rows.append({'Scenario':label,'Objective':round(r.get('objective',0),2),
                                 'Startpunten':round(r.get('roster_points',0),2),
                                 'Captain':round(r.get('captain_bonus',0),2),
                                 'Budget':round(r.get('spent',0),1),'Kandidaten':r.get('candidate_count'),
                                 'Tijd':round(r.get('timing',{}).get('total',0),2),
                                 'Prijs':round(num(by_id[lks[0]].get('now_cost'))/currency,1) if lks else '',
                                 'Projectie':round(next((x[0] for x in ranked if lks and x[2]['id']==lks[0]),0),2) if lks else '',
                                 '_result':r})
                    ss.write(f"{label}: {r['objective']:.2f}")
                free_row=rows[0]
                free_ids={p['id'] for p in free_row['_result'].get('selected',[])}
                free_obj=free_row['Objective']
                for row in rows:
                    if row['Scenario']=='Vrij':
                        row['In vrij team']='—'
                        row['Opportunity cost']=0.0
                    else:
                        pid=next((p['id'] for p in premium if (p.get('web_name') or full_name(p))==row['Scenario']),None)
                        row['In vrij team']='JA' if pid in free_ids else 'NEE'
                        row['Opportunity cost']=round(max(0.0,free_obj-row['Objective']),2)
                inconsistent=[]
                for row in rows[1:]:
                    pid=next((p['id'] for p in premium if (p.get('web_name') or full_name(p))==row['Scenario']),None)
                    if pid in free_ids and abs(row['Objective']-free_row['Objective'])>0.05:
                        inconsistent.append(row['Scenario'])
                st.session_state.scenarios300=rows
                st.session_state.scenario_warning300=inconsistent
                ss.update(label='Premiumscenario’s gereed',state='complete',expanded=False)

        if compare:
            with st.status('Vrij optimum vergelijken met locks…',expanded=True) as cs:
                t0=time.perf_counter()
                free=optimize(players,teams,types,rules,events,active,currency,weights,first_gw,b2b,budget,[],excluded)
                fixed=optimize(players,teams,types,rules,events,active,currency,weights,first_gw,b2b,budget,locked,excluded)
                st.session_state.compare260={'free':free,'fixed':fixed,'elapsed':time.perf_counter()-t0}
                cs.update(label='Vergelijking gereed',state='complete',expanded=False)

        scenario_rows=st.session_state.get('scenarios300')
        if scenario_rows:
            st.subheader('🧪 Premiumscenario’s — dynamische top 20, zelfde model/budget/4GW-horizon')
            best=max(r['Objective'] for r in scenario_rows)
            display=[]
            ordered=sorted(scenario_rows,key=lambda x:(x.get('Opportunity cost',max(0,best-x['Objective'])),
                                                       -num(x.get('Projectie'),0)))
            for r in ordered:
                oc=r.get('Opportunity cost',round(max(0,best-r['Objective']),2))
                display.append({'Scenario':r['Scenario'],'In vrij team':r.get('In vrij team',''),
                                'Prijs $m':r.get('Prijs',''),'Projectie':r.get('Projectie',''),
                                'Opportunity cost':oc,'Objective':r['Objective'],
                                'Startpunten':r['Startpunten'],'Captain':r['Captain'],
                                'Budget $m':r['Budget'],'Tijd s':r['Tijd']})
            html_table(display)
            pool_sizes={r.get('Kandidaten') for r in scenario_rows}
            if len(pool_sizes)==1:
                st.success(f"✓ Zuivere vergelijking: alle scenario’s gebruiken exact dezelfde {next(iter(pool_sizes))} kandidaten.")
            warnings=st.session_state.get('scenario_warning300',[])
            if warnings:
                st.error('Scenario-sanitycheck mislukt voor: '+', '.join(warnings))
            else:
                st.caption('Sanitycheck OK: een speler die al in het vrije optimum zit, verliest geen punten wanneer hij wordt vastgezet.')
            st.caption('Dit is een modelvergelijking: een lagere scenario-score betekent dat de verplichte ster volgens dezelfde aannames opportunity cost veroorzaakt in de overige rosterplaatsen.')
    with tab_projection_main:
        st.info('**Wat doe je hier?** Zoek en vergelijk spelers op verwachte punten, prijs en betrouwbaarheid. Bijvoorbeeld: is Haliburton statistisch aantrekkelijker dan Lillard?')
        st.subheader('👥 Alle NBA Fantasy-spelers')
        st.caption(f'{len(selectable)} selecteerbare spelers geladen. Dit is de volledige zichtbare pool; de interne solver-shortlist staat hier los van.')
        st.caption('Projectie gebruikt nu een adaptive baseline: seizoen-PPG als anchor en maximaal 35% actuele form naarmate de echte minutensample groeit. Waarde/prijs blijft alleen een verkenningsmaatstaf.')
        q=st.text_input('Zoek speler',placeholder='Bijv. Jokic, Doncic, Maxey',key='player_search320').strip().lower()
        fpos=st.selectbox('Positie',['Alle','BC','FC'],key='player_pos320')
        sort_mode=st.selectbox('Sorteer op',['Projectie','Waarde/prijs','Prijs','Durability%'],key='player_sort330')
        explorer_rows=[]
        role_ex=role_adjustment_map(selectable,currency)
        for p in selectable:
            nm=p.get('web_name') or full_name(p)
            posname='BC' if int(p.get('element_type',0))==1 else 'FC'
            if q and q not in nm.lower() and q not in full_name(p).lower(): continue
            if fpos!='Alle' and posname!=fpos: continue
            price=max(.1,num(p.get('now_cost'))/currency)
            bd=projection_breakdown(p,currency,role_ex)
            proj=bd['Projectie']
            value=proj/price
            explorer_rows.append({'Speler':nm,'Team':team_by_id.get(int(p['team']),{}).get('short_name',team_by_id.get(int(p['team']),{}).get('name','?')),'Pos':posname,
                                  'Prijs':round(price,1),'Projectie':proj,
                                  'Waarde/prijs':round(value,3),'Basis':bd['Basis'],
                                  'Preseason conf.%':bd['Preseason confidence%'],
                                  'Availability%':bd['Availability%'],'Role%':bd['Role%'],
                                  'Durability%':bd['Durability%']})
        explorer_rows.sort(key=lambda r:num(r.get(sort_mode),0),reverse=True)
        html_table(explorer_rows)

        st.subheader('🧪 Projection Lab')
        st.caption('Audit + backtest-gekalibreerde baseline. Historisch FP/G-niveau en vorige GP sturen een begrensde correctie vóór role/durability. Alleen 0Y/API-fallback spelers houden een milde anti-outlier guard.')
        audit_rows=projection_lab_rows(selectable,currency,role_ex)
        # Optimizer Risk = hoe aantrekkelijk de speler voor de solver is × hoeveel data-onzekerheid er is.
        # Nulprojecties krijgen bewust risk 0: slechte data zonder modelimpact is geen optimizer-risico.
        max_proj=max([float(r['Projectie']) for r in audit_rows] or [1.0])
        max_value=max([float(r['Value/prijs']) for r in audit_rows] or [1.0])
        for r in audit_rows:
            attractiveness=0.70*(float(r['Projectie'])/max_proj if max_proj else 0.0)+0.30*(float(r['Value/prijs'])/max_value if max_value else 0.0)
            uncertainty=max(0.0,1.0-float(r['Conf.%'])/100.0)
            r['Optimizer risk']=round(100.0*attractiveness*uncertainty,1) if float(r['Projectie'])>0 else 0.0
        audit_mode=st.selectbox('Projection Lab weergave',['Optimizer Risk — relevantste eerst','Repair picks — audit','Vrij optimaal team — audit','Hoogste projecties','Beste value/prijs','Laagste confidence'],key='projection_lab_sort392')
        result_now=st.session_state.get('result260')
        shown_rows=list(audit_rows)
        if audit_mode=='Optimizer Risk — relevantste eerst':
            shown_rows.sort(key=lambda r:(r['Optimizer risk'],r['Projectie']),reverse=True)
        elif audit_mode=='Repair picks — audit':
            ev=st.session_state.get('team_eval381')
            if not ev:
                st.info('Draai eerst de Team Evaluator met jouw 10 locks. Daarna toont deze weergave alle OUT/IN-spelers uit de 1-, 2- en 3-swap repairs.')
                shown_rows=[]
            else:
                repair_ids=set()
                repair_tag={}
                for rep in ev.get('repairs',[]):
                    n=rep.get('swaps')
                    for pid in rep.get('out',[]):
                        repair_ids.add(pid); repair_tag.setdefault(pid,[]).append(f'{n}-swap OUT')
                    for pid in rep.get('in',[]):
                        repair_ids.add(pid); repair_tag.setdefault(pid,[]).append(f'{n}-swap IN')
                by_pid={p['id']:p for p in selectable}
                names={pid:(by_pid[pid].get('web_name') or full_name(by_pid[pid])) for pid in repair_ids if pid in by_pid}
                shown_rows=[]
                for r in audit_rows:
                    pid=next((pid for pid,nm in names.items() if nm==r['Speler']),None)
                    if pid is not None:
                        rr=dict(r); rr['Repair rol']='; '.join(repair_tag.get(pid,[])); shown_rows.append(rr)
                shown_rows.sort(key=lambda r:(0 if 'IN' in r.get('Repair rol','') else 1,-r['Optimizer risk'],-r['Projectie']))
        elif audit_mode=='Vrij optimaal team — audit':
            if result_now is None:
                st.info('Draai eerst één vrije optimalisatie. Daarna toont deze weergave exact de 10 gekozen spelers met hun projection-risk.')
                shown_rows=[]
            else:
                chosen={p['id'] for p in result_now.get('selected',[])}
                name_by_pid={p['id']:(p.get('web_name') or full_name(p)) for p in selectable}
                chosen_names={name_by_pid[i] for i in chosen if i in name_by_pid}
                shown_rows=[r for r in audit_rows if r['Speler'] in chosen_names]
                shown_rows.sort(key=lambda r:(r['Optimizer risk'],r['Projectie']),reverse=True)
        elif audit_mode=='Hoogste projecties':
            shown_rows.sort(key=lambda r:r['Projectie'],reverse=True)
        elif audit_mode=='Beste value/prijs':
            shown_rows.sort(key=lambda r:r['Value/prijs'],reverse=True)
        else:
            shown_rows.sort(key=lambda r:(r['Conf.%'],-r['Projectie']))
        low=sum(1 for r in audit_rows if r['Niveau']=='LAAG'); mid=sum(1 for r in audit_rows if r['Niveau']=='MIDDEL'); high=sum(1 for r in audit_rows if r['Niveau']=='HOOG')
        risky=sum(1 for r in audit_rows if r['Optimizer risk']>=15)
        a1,a2,a3,a4=st.columns(4); a1.metric('Hoge confidence',high); a2.metric('Middel',mid); a3.metric('Lage confidence',low); a4.metric('Optimizer risk ≥15',risky)
        html_table(shown_rows,500)
        st.caption('Gebruik Repair picks — audit na de Team Evaluator om T. Prince, T. Smith, Lillard, Sabonis en de OUT-spelers direct naast elkaar te controleren. Optimizer risk combineert model-aantrekkelijkheid (projectie + value/prijs) met data-onzekerheid. Een speler met projectie 0 krijgt risk 0: onvolledige data is pas gevaarlijk als de solver de speler aantrekkelijk vindt. Confidence blijft de audit-indicator. De projectie gebruikt nu de v3.10.1 walk-forward bias per historisch FP/G-niveau en vorige GP; 0Y-spelers krijgen een prijs/API-gebaseerde preseason prior in plaats van automatisch 0.')

    with tab_backtest_main:
        st.info('**Wat leer je hier?** Hoe ver zaten onze historische voorspellingen gemiddeld naast de werkelijkheid? Kijk vooral naar MAE (gemiddelde fout) en Bias (structureel te hoog of te laag).')
        st.subheader('🧭 Historical Backtest Lab')
        st.caption('Walk-forward controle van de multi-season fantasy-baseline: een seizoen wordt alleen voorspeld met eerdere seizoenen. Dit valideert de historische PPG-anchor; prijs, toekomstige rol, blessures, fixtures en captaincy worden hier bewust niet achteraf "voorspeld".')
        bt_min_games=st.slider('Backtest: minimum games per seizoen',5,60,15,5,key='backtest_min_games3101')
        bt_groups,bt_details=historical_projection_backtest(bt_min_games)
        if bt_details:
            overall=next((r for r in bt_groups if r['Groep']=='Alles'),bt_groups[0])
            b1,b2,b3,b4=st.columns(4)
            b1.metric('Speler-seizoenen',overall['N']); b2.metric('MAE',f"{overall['MAE']:.2f} FP/G")
            b3.metric('RMSE',f"{overall['RMSE']:.2f} FP/G"); b4.metric('Bias',f"{overall['Bias']:+.2f} FP/G")
            html_table(bt_groups,260)
            st.caption('MAE = gemiddelde absolute fout per wedstrijd. Bias > 0 betekent dat de historische baseline gemiddeld te hoog voorspelde; bias < 0 te laag. RMSE straft grote missers extra.')

            st.markdown('**Robuustheid bij verschillende minimum-gamesgrenzen**')
            html_table(backtest_threshold_matrix(),230)
            st.caption('Dezelfde walk-forward test bij 15/30/45/60 gespeelde wedstrijden. Zo zien we of kleine/injury-seizoenen de foutmaat domineren.')

            gp_diag,fp_diag=backtest_diagnostic_groups(bt_details)
            d1,d2=st.columns(2)
            with d1:
                st.markdown('**Fout naar eerdere beschikbaarheid (GP)**')
                html_table(gp_diag,250)
            with d2:
                st.markdown('**Fout naar eerder fantasy-niveau (FP/G)**')
                html_table(fp_diag,280)
            st.caption('Deze segmenten gebruiken alleen informatie uit de voorafgaande seizoenen. Leeftijd staat niet betrouwbaar in de opgeslagen historische dataset en wordt daarom bewust niet achteraf ingevuld.')

            with st.expander('Grootste historische missers bekijken'):
                html_table(bt_details[:40],540)
        else:
            st.info('Nog onvoldoende historische fantasy-seizoenen in SQLite voor deze backtest. Laad eerst de 3-jaars historie in de app.')


    with tab_inspector_main:
        st.info('**Wat doe je hier?** Kies één speler en ontdek welke historische prestaties, minuten en correcties zijn verwachte score bepalen. Handig als een projectie je verbaast.')
        st.subheader('🔬 Projectie-inspector')
        st.success(f'Projectiemotor: MULTI-SEASON • build {BUILD_ID}')
        st.caption('Deze inspector verandert niets aan de optimizer. Hij maakt exact de projectieberekening zichtbaar die de solver gebruikt.')
        player_options=sorted([(p.get('web_name') or full_name(p),p['id']) for p in selectable],key=lambda x:x[0])
        inspect_name=st.selectbox('Inspecteer speler',[x[0] for x in player_options],key='projection_inspector371')
        inspect_id=next(pid for nm,pid in player_options if nm==inspect_name)
        inspect_p=next(p for p in selectable if p['id']==inspect_id)
        bd=projection_breakdown(inspect_p,currency,role_ex)
        factor_rows=[
            {'Stap':'ENGINE / BUILD','Waarde':BUILD_ID,'Uitleg':'Als deze regel zichtbaar is, draait daadwerkelijk v3.18.3 Risk Calibration Engine'},
            {'Stap':'NBA Fantasy PPG-bron','Waarde':bd['PPG bron'],'Uitleg':'Officiële huidige Fantasy API-anchor'},
            {'Stap':'Fantasy PPG 2023-24','Waarde':bd['FP 23-24'],'Uitleg':'Gereconstrueerd: PTS + REB + 2×AST + 3×BLK + 3×STL'},
            {'Stap':'Fantasy PPG 2024-25','Waarde':bd['FP 24-25'],'Uitleg':'Gereconstrueerd met dezelfde officiële fantasy-scoreformule'},
            {'Stap':'Fantasy PPG 2025-26','Waarde':bd['FP 25-26'],'Uitleg':'Gereconstrueerd met dezelfde officiële fantasy-scoreformule'},
            {'Stap':'Multi-season baseline','Waarde':bd['Multi-season baseline'],'Uitleg':'20/30/50 recencyweging; bij volledige historie 70% reconstructie + 30% Fantasy API-anchor'},
            {'Stap':'Multi-season bron','Waarde':bd['Multi-season bron'],'Uitleg':'Ontbrekende historie wordt niet verzonnen; dan conservatieve API-fallback'},
            {'Stap':'NBA Fantasy form-bron','Waarde':bd['Form bron'],'Uitleg':'Ruwe form uit de Fantasy API'},
            {'Stap':'Minuten-bron','Waarde':bd['Minuten bron'],'Uitleg':'Ruwe seizoenminuten uit de Fantasy API'},
            {'Stap':'Actuele games sample*','Waarde':bd['Games sample*'],'Uitleg':'Afgeleid uit actuele Fantasy total_points ÷ points_per_game; vóór de eerste afgeronde Gameday altijd 0'},
            {'Stap':'Historisch/seizoen-PPG gewicht','Waarde':str(bd['Historisch gewicht%'])+'%','Uitleg':'Stabiele anchor; daalt geleidelijk tot minimaal 65% wanneer form beschikbaar is'},
            {'Stap':'Actuele form gewicht','Waarde':str(bd['Actueel gewicht%'])+'%','Uitleg':'0–35%; groeit over ongeveer 10 geschatte wedstrijden'},
            {'Stap':'Blended baseline','Waarde':bd['Blended baseline'],'Uitleg':'Gewogen combinatie van seizoen-PPG en actuele form'},
            {'Stap':'Minuten reliability','Waarde':str(bd['Minuten reliability%'])+'%','Uitleg':'Bestaande early-sample beveiliging op basis van totale minuten'},
            {'Stap':'Basisprojectie','Waarde':bd['Basis'],'Uitleg':'Adaptive baseline na minuten-reliability'},
            {'Stap':'Historische seizoenen','Waarde':bd['Historische seizoenen'],'Uitleg':'Aantal beschikbare GP-seizoenen uit 2023-24 t/m 2025-26'},
            {'Stap':'Preseason confidence','Waarde':str(bd['Preseason confidence%'])+'%','Uitleg':'Generieke onzekerheidskrimp: 0Y 90%, 1Y 94%, 2Y 97%, 3Y 100%; nooit een speler-specifieke boost'},
            {'Stap':'Availability','Waarde':str(bd['Availability%'])+'%','Uitleg':'Actuele beschikbaarheidsfactor'},
            {'Stap':'Durability rate','Waarde':str(bd['Durability%'])+'%','Uitleg':'Historische beschikbaarheid / games played'},
            {'Stap':'Durability factor','Waarde':str(bd['Durability factor%'])+'%','Uitleg':'Factor die het model werkelijk toepast'},
            {'Stap':'Role factor','Waarde':str(bd['Role%'])+'%','Uitleg':'Huidige role_adjustment_map() factor'},
            {'Stap':'Basis vóór backtest-calibratie','Waarde':bd['Basis vóór backtest-calibratie'],'Uitleg':'Historische/adaptieve baseline vóór empirische correctie'},{'Stap':'Backtest calibration','Waarde':bd['Backtest calibration'],'Uitleg':'Begrensde correctie uit v3.10.1 bias naar historisch FP/G-niveau en vorige GP'},{'Stap':'Voor 0Y safety guard','Waarde':bd['Voor uncertainty-regressie'],'Uitleg':'Gekalibreerde basis × preseason-confidence × availability × durabilityfactor × rolefactor'},{'Stap':'0Y safety penalty','Waarde':bd['Uncertainty penalty'],'Uitleg':'Alleen voor spelers zonder gereconstrueerde historie; geen backtest-boost zonder bewijs'},{'Stap':'Eindprojectie','Waarde':bd['Projectie'],'Uitleg':'Projectie na empirische calibratie en eventuele 0Y safety guard'},
        ]
        html_table(factor_rows)
        st.caption('* Games sample wordt uit de actuele Fantasy API-productie afgeleid en begrensd door het aantal afgeronde Gamedays. Vóór de eerste afgeronde Gameday is de sample altijd 0. Form krijgt maximaal 35% gewicht.')

        strategy_cmp=st.session_state.get('strategy_compare318')
        if strategy_cmp:
            st.subheader('🎯 Safe vs Balanced vs Aggressive — vrij optimum')
            st.caption('Zelfde budget en exclusions; je geladen profiel/locks worden hier bewust niet vastgezet. Safe optimaliseert richting floor/reliability, Balanced op expected value en Aggressive richting ceiling/upside. De strategie-objectives gebruiken verschillende risicoscores en zijn daarom niet bedoeld als onderling vergelijkbare voorspelde punten.')
            summary=[]
            balanced_ids={p['id'] for p in strategy_cmp.get('Balanced',{}).get('selected',[])}
            for strat in ('Safe','Balanced','Aggressive'):
                rr=strategy_cmp.get(strat)
                if not rr: continue
                ids={p['id'] for p in rr.get('selected',[])}
                changed=len(ids.symmetric_difference(balanced_ids))//2 if balanced_ids else 0
                summary.append({'Strategie':strat,'4GW objective':round(rr.get('objective',0),1),
                                'Startpunten':round(rr.get('roster_points',0),1),'Captainbonus':round(rr.get('captain_bonus',0),1),
                                'Budget $m':round(rr.get('spent',0),1),'Spelers anders vs Balanced':changed})
            html_table(summary,240)
            cols=st.columns(3)
            for col,strat in zip(cols,('Safe','Balanced','Aggressive')):
                rr=strategy_cmp.get(strat)
                if not rr: continue
                with col:
                    st.markdown(f'**{strat}**')
                    teamrows=[]
                    for p in sorted(rr.get('selected',[]),key=lambda q:(int(q.get('element_type',0)),-num(q.get('_proj')))):
                        ru=risk_upside_components(p)
                        # Result players retain the strategy-adjusted _proj used by that solve.
                        # Recover the central expectation, then show the same transparent scenario bands.
                        _saved_strategy=RISK_STRATEGY
                        RISK_STRATEGY=strat
                        central=central_projection_from_strategy_value(p,p.get('_proj'))
                        RISK_STRATEGY=_saved_strategy
                        sv=risk_scenario_values(p,central)
                        teamrows.append({'Speler':p.get('web_name') or full_name(p),
                                         'Team':team_by_id.get(int(p['team']),{}).get('short_name','?'),
                                         'Pos':type_names.get(int(p.get('element_type',0)),'?'),
                                         'Prijs':round(num(p.get('now_cost'))/currency,1),
                                         'Floor':round(sv['floor'],1),'Expected':round(sv['expected'],1),
                                         'Ceiling':round(sv['ceiling'],1),'Conf.%':round(100*sv['confidence'])})
                    html_table(teamrows,430)
            safe_ids={p['id'] for p in strategy_cmp.get('Safe',{}).get('selected',[])}
            agg_ids={p['id'] for p in strategy_cmp.get('Aggressive',{}).get('selected',[])}
            byid={p['id']:p for p in selectable}
            def names(ids): return ', '.join(sorted((byid[i].get('web_name') or full_name(byid[i])) for i in ids if i in byid)) or '—'
            st.info('Safe → Aggressive wijzigingen: OUT '+names(safe_ids-agg_ids)+' • IN '+names(agg_ids-safe_ids))

    with team_dashboard:
        result=st.session_state.result260
        if result is None:
            st.success('De app is klaar. Er draait géén solver op de achtergrond.')
            st.info('Kies links je budget / locks / exclusions en klik pas daarna op **Optimaliseer team**.')
            preview=[]
            for p in sorted(selectable,key=lambda q:-q['_proj'])[:30]:
                dd=durability_details(p); gp=dd['games']
                preview.append({'Speler':('🔒 '+p['web_name'] if p['id'] in locked else p['web_name']),'Team':teams[int(p['team'])]['short_name'],
                    'Pos':type_names[int(p['element_type'])],'Prijs':num(p['now_cost'])/currency,
                    'Risk-adjusted':round(p['_proj'],2),'Durability %':('ONBEKEND' if dd['score'] is None else round(dd['score']*100)),
                    'Historie':f"{dd['years_count']}Y",'GP 23-24':gp.get(2024,''),'GP 24-25':gp.get(2025,''),'GP 25-26':gp.get(2026,''),'Bron':dd['source']})
            html_table(preview,560)
            return

        crystal_court(result,teams,type_names,events,currency)

        c1,c2,c3,c4=st.columns(4)
        c1.metric('Budget gebruikt',f"${result['spent']:.1f}m",f"${result['_budget_cap']-result['spent']:.1f}m over")
        c2.metric('Verwachte score · 4 GW',f"{result['objective']:.1f}")
        c3.metric('Startpunten',f"{result['roster_points']:.1f}")
        c4.metric('Captainbonus',f"{result['captain_bonus']:.1f}")
        with tab_diag_main:
            st.subheader('⚙️ Laatste optimalisatie')
            st.metric('Solver',f"{result['status']} • {result.get('_elapsed',0):.1f}s")
            st.caption(f"2-stage optimizer + precompute-cache: {result.get('candidate_count','?')} kandidaten in de MILP i.p.v. {len(selectable)}.")
            tm=result.get('timing',{})
            if tm:
                st.caption("⏱ Profiel — shortlist %.1fs • modelbouw %.1fs • CBC %.1fs • uitlezen %.1fs • totaal %.1fs" %
                           (tm.get('shortlist',0),tm.get('model',0),tm.get('solve',0),tm.get('extract',0),tm.get('total',0)))
        cmp=st.session_state.get('compare260')
        if cmp:
            delta=cmp['fixed']['objective']-cmp['free']['objective']
            st.info(f"⚖️ Vrij optimum: {cmp['free']['objective']:.1f} • Met locks: {cmp['fixed']['objective']:.1f} • verschil {delta:+.1f} punten.")
        tev=st.session_state.get('team_eval381')
        if tev:
            exact=tev['exact']; free=tev['free']
            gap=max(0.0,free['objective']-exact['objective'])
            pct=(100.0*exact['objective']/free['objective']) if free['objective'] else 0.0
            st.subheader('🩺 Team Evaluator — hand-picked vs modeloptimum')
            m1,m2,m3,m4=st.columns(4)
            m1.metric('Jouw team',f"{exact['objective']:.1f}")
            m2.metric('Vrij optimum',f"{free['objective']:.1f}")
            m3.metric('Van modeloptimum',f"{pct:.1f}%")
            m4.metric('Model-gap',f"{gap:.1f} pt")
            st.caption(f"Zelfde ${tev['budget']:.1f}m budget, 4GW-horizon en exact dezelfde {tev['pool']} kandidaten • evaluatie {tev['elapsed']:.1f}s. 'Optimum' betekent optimum binnen dit huidige projectiemodel/kandidaatpool.")
            name_by_id={p['id']:p.get('web_name') or full_name(p) for p in selectable}
            repair_rows=[]
            for item in tev['repairs']:
                rr=item['result']; gain=rr['objective']-exact['objective']
                out_txt=', '.join(name_by_id.get(pid,str(pid)) for pid in item['out']) or '—'
                in_txt=', '.join(name_by_id.get(pid,str(pid)) for pid in item['in']) or '—'
                repair_rows.append({'Max swaps':item['swaps'],'Objective':round(rr['objective'],1),
                                    'Winst vs jouw team':round(gain,1),
                                    'Van vrij optimum %':round(100*rr['objective']/free['objective'],1) if free['objective'] else 0,
                                    'OUT':out_txt,'IN':in_txt,'Budget $m':round(rr['spent'],1)})
            html_table(repair_rows,330)
            st.caption('Repair = maximaal 1, 2 of 3 spelers vervangen; de overige 9/8/7 van jouw hand-picked selectie moeten behouden blijven. Alle cijfers zijn modelprojecties, geen garantie op echte fantasy-punten.')
            st.markdown('### 🔬 Explain the Gap')
            st.caption('Hieronder wordt elke repair opnieuw vergeleken met jouw originele team. Δ startpunten laat het effect van de Starting-5/fixture-rotatie zien; Δ captain laat zien hoeveel van de winst uit een andere captainplanning komt. Omdat lineups samen veranderen, zijn dit scenario-effecten en geen geïsoleerde causale waardes per swap.')
            for item in tev['repairs']:
                rr=item['result']; gain=rr['objective']-exact['objective']
                start_gain=rr.get('roster_points',0)-exact.get('roster_points',0)
                cap_gain=rr.get('captain_bonus',0)-exact.get('captain_bonus',0)
                out_txt=', '.join(name_by_id.get(pid,str(pid)) for pid in item['out']) or '—'
                in_txt=', '.join(name_by_id.get(pid,str(pid)) for pid in item['in']) or '—'
                with st.expander(f"Max {item['swaps']} swap(s) • +{gain:.1f} pt • OUT {out_txt} → IN {in_txt}", expanded=(item['swaps']==1)):
                    g1,g2,g3,g4=st.columns(4)
                    g1.metric('Totale winst',f"+{gain:.1f}")
                    g2.metric('Uit startpunten',f"{start_gain:+.1f}")
                    g3.metric('Uit captainbonus',f"{cap_gain:+.1f}")
                    g4.metric('Budget',f"${rr['spent']:.1f}m")
                    st.markdown('**Netto per swap**')
                    swap_rows=net_swap_summary(exact,rr,item['out'],item['in'],selectable,events,active,currency,weights,first_gw,b2b)
                    html_table(swap_rows,250)
                    st.caption('Netto per swap = productie van de inkomende speler in het gerepareerde team minus productie van de uitgaande speler in jouw originele team. Rotatie/interactie vangt veranderingen bij de spelers die in beide teams blijven. De optimizer lost alle swaps gezamenlijk op; de pairing is dus uitleg, geen losse causaliteit.')
                    st.markdown('**Detail per speler**')
                    detail=gap_explanation(exact,rr,events,active,currency,weights,first_gw,b2b)
                    html_table(detail,360)
                    if abs((start_gain+cap_gain)-gain)>.2:
                        st.caption('Kleine afwijking kan ontstaan door afronding/solvercomponenten.')



        if mode.startswith('🔄') and current_roster:
            names_by_id={p['id']:p['web_name'] for p in selectable}
            tins=', '.join(p['web_name'] for p in result.get('transfers_in',[])) or 'geen'
            tout=', '.join(names_by_id.get(pid,str(pid)) for pid in result.get('transfers_out',[])) or 'geen'
            st.success(f"🔄 Transferadvies — OUT: {tout} • IN: {tins}")

        tab1,tab2,tab3,tab4,tab5,tab6=st.tabs(['📋 Selectie','📈 Value','⭐ Starting 5 & Captain','👥 Players','🗄️ Historie','🗓️ Gameweek Planner'])

        with tab1:
            rows=[]
            for p in sorted(result['selected'],key=lambda p:(int(p['element_type']),-p['_proj'])):
                dd=durability_details(p); gp=dd['games']
                rows.append({'Speler':p['web_name'],'Team':teams[int(p['team'])]['short_name'],
                    'Pos':type_names[int(p['element_type'])],'Prijs':num(p['now_cost'])/currency,
                    'Projectie':round(p['_proj'],2),'Floor':round(p['_proj']*risk_upside_components(p)['floor_factor'],2),
                    'Upside':round(p['_proj']*risk_upside_components(p)['upside_factor'],2),
                    'Confidence%':round(100*risk_upside_components(p)['confidence']),
                    'Role%':round(100*num(p.get('_role_factor',1.0),1.0),1),'Actueel %':round(current_availability(p)*100),
                    'Durability %':('ONBEKEND' if dd['score'] is None else round(dd['score']*100)),'Historie':f"{dd['years_count']}Y",'GP 23-24':gp.get(2024,''),
                    'GP 24-25':gp.get(2025,''),'GP 25-26':gp.get(2026,''),'<60 seizoenen':dd['low_seasons']})
            html_table(rows,560)
            st.download_button('⬇️ Export naar Excel',
                excel_bytes(result,teams,type_names,events,active,currency,weights,first_gw,b2b),
                'Crystal_Ball_v4.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            st.info(f'Risk & Upside strategie: {RISK_STRATEGY}. Floor/upside zijn modelscenario’s rond de centrale projectie, geen garanties. Budget is een maximum, geen doel.')

        with tab2:
            st.write('De budgetcurve wordt **alleen** berekend als jij daarom vraagt.')
            st.caption('Elke budgetstap is een aparte optimalisatie met maximaal 20 seconden CBC-solvertijd.')
            if st.button('📊 Bereken value-analyse'):
                caps=sorted(set([max(70.0,budget-15),max(70.0,budget-10),max(70.0,budget-5),budget]))
                curve=[]; prog=st.progress(0,text='Value-analyse starten…')
                for i,cap in enumerate(caps):
                    t0=time.perf_counter()
                    try:
                        rr=optimize(players,teams,types,rules,events,active,currency,weights,first_gw,b2b,cap,locked,excluded)
                        curve.append({'Budgetcap':cap,'Werkelijk gebruikt':round(rr['spent'],1),
                                      'Objective':round(rr['objective'],2),'Solver':rr['status'],
                                      'Seconden':round(time.perf_counter()-t0,1)})
                    except Exception as exc:
                        curve.append({'Budgetcap':cap,'Werkelijk gebruikt':'','Objective':'',
                                      'Solver':str(exc),'Seconden':round(time.perf_counter()-t0,1)})
                    prog.progress((i+1)/len(caps),text=f'Budget ${cap:.1f}m afgerond')
                prog.empty(); st.session_state.curve260=curve
            if st.session_state.curve260:
                html_table(st.session_state.curve260,360)

        with tab3:
            caprows=[]
            for e in events:
                if e['id'] in result['captains']:
                    p=result['captains'][e['id']]
                    caprows.append({'Gameweek':parse_gw(e['name']),'Gameday':e['name'],
                                    'Captain':p['web_name'],
                                    'Verwachte bonus':round(event_projection(p,e,first_gw,currency,b2b),2)})
            st.subheader('Captain'); html_table(caprows,320)
            chosen=st.selectbox('Bekijk Starting 5',events,format_func=lambda e:e['name'])
            lineup=[]
            for p in result['lineups'][chosen['id']]:
                plays=int(p['team']) in active[chosen['id']]
                lineup.append({'Speler':p['web_name'],'Team':teams[int(p['team'])]['short_name'],
                               'Pos':type_names[int(p['element_type'])],'Speelt':'Ja' if plays else 'Nee',
                               'Projectie':round(event_projection(p,chosen,first_gw,currency,b2b),2) if plays else 0})
            html_table(lineup,360)

        with tab6:
            st.subheader('🗓️ Gameweek Planner — komende 4 Gameweeks')
            st.caption('Compacte uitleg van exact dezelfde Starting 5, captainplanning én 4GW-weging als de optimizer. Spelers zonder fixture kunnen nodig zijn om een geldige 2/3 BC-FC Starting 5 te vormen, maar leveren 0 modelpunten.')
            planner_rows=[]
            for ev in events:
                eid=ev['id']
                starters=result['lineups'].get(eid,[])
                starter_ids={int(p['id']) for p in starters}
                bench=[p for p in result['selected'] if int(p['id']) not in starter_ids]
                cap=result['captains'].get(eid)
                # IMPORTANT: the MILP objective discounts later Gameweeks with `weights[eid]`.
                # The planner must use that exact same coefficient, otherwise its totals can be
                # hundreds of points above the optimizer even though the lineups are identical.
                event_weight=weights.get(eid,1.0)
                start_pts=0.0
                playing_starters=[]
                zero_starters=[]
                for pp in starters:
                    plays=int(pp['team']) in active.get(eid,set())
                    ep=event_projection(pp,ev,first_gw,currency,b2b) if plays else 0.0
                    start_pts+=ep*event_weight
                    (playing_starters if plays else zero_starters).append(pp['web_name'])
                cap_bonus=0.0
                if cap is not None and int(cap['team']) in active.get(eid,set()):
                    cap_bonus=event_projection(cap,ev,first_gw,currency,b2b)*event_weight
                bc=sum(1 for pp in starters if int(pp['element_type'])==1)
                fc=len(starters)-bc
                bench_playing=[pp['web_name'] for pp in bench if int(pp['team']) in active.get(eid,set())]
                bench_idle=[pp['web_name'] for pp in bench if int(pp['team']) not in active.get(eid,set())]
                planner_rows.append({
                    'GW':parse_gw(ev.get('name')), 'Gameday':ev.get('name',''),
                    'Formatie':f'{bc} BC + {fc} FC', 'Spelend':playing_starters,
                    'Geen fixture':zero_starters, 'Bank speelt':bench_playing, 'Bank geen fixture':bench_idle,
                    'Captain':cap['web_name'] if cap else '—', 'Gewicht':event_weight, 'Startpunten_raw':start_pts,
                    'Captainbonus_raw':cap_bonus, 'Totaal_raw':start_pts+cap_bonus,
                })

            # Compact per Gameweek: no horizontal mega-table.
            for gw in sorted({r['GW'] for r in planner_rows if r['GW'] is not None}):
                rr=[r for r in planner_rows if r['GW']==gw]
                gw_start=sum(x['Startpunten_raw'] for x in rr)
                gw_cap=sum(x['Captainbonus_raw'] for x in rr)
                st.markdown(f'### Gameweek {gw} — {gw_start+gw_cap:.1f} gewogen model-FP')
                for r in rr:
                    playing=', '.join(r['Spelend']) or '—'
                    zeros=', '.join(r['Geen fixture']) or '—'
                    bp=', '.join(r['Bank speelt']) or '—'
                    bi=', '.join(r['Bank geen fixture']) or '—'
                    with st.expander(f"{r['Gameday']} • {r['Formatie']} • {len(r['Spelend'])}/5 met fixture • ⭐ {r['Captain']} • {r['Totaal_raw']:.1f} model-FP", expanded=(gw==sorted({x['GW'] for x in planner_rows if x['GW'] is not None})[0])):
                        c1,c2,c3,c4=st.columns(4)
                        c1.metric('Startpunten',f"{r['Startpunten_raw']:.1f}")
                        c2.metric('Captainbonus',f"{r['Captainbonus_raw']:.1f}")
                        c3.metric('Totaal',f"{r['Totaal_raw']:.1f}")
                        c4.metric('Starters met fixture',f"{len(r['Spelend'])}/5")
                        st.caption(f"4GW objective-weging voor deze Gameday: {r['Gewicht']:.4f}")
                        st.markdown(f"**Starting 5 met fixture:** {playing}")
                        if r['Geen fixture']:
                            st.warning(f"Starting 5 zonder fixture (0 FP): {zeros}")
                        st.markdown(f"**Bank met fixture:** {bp}")
                        if r['Bank geen fixture']:
                            st.caption(f"Bank zonder fixture: {bi}")

            if planner_rows:
                calc_start=sum(r['Startpunten_raw'] for r in planner_rows)
                calc_cap=sum(r['Captainbonus_raw'] for r in planner_rows)
                calc_total=calc_start+calc_cap
                model_start=float(result.get('roster_points',0.0))
                model_cap=float(result.get('captain_bonus',0.0))
                model_total=float(result.get('objective',0.0))
                d_start=calc_start-model_start; d_cap=calc_cap-model_cap; d_total=calc_total-model_total
                st.subheader('✅ Planner sanity check')
                checks=[
                    {'Onderdeel':'Startpunten','Planner':round(calc_start,2),'Optimizer':round(model_start,2),'Δ':round(d_start,3)},
                    {'Onderdeel':'Captainbonus','Planner':round(calc_cap,2),'Optimizer':round(model_cap,2),'Δ':round(d_cap,3)},
                    {'Onderdeel':'Totaal','Planner':round(calc_total,2),'Optimizer':round(model_total,2),'Δ':round(d_total,3)},
                ]
                html_table(checks,220)
                if abs(d_start)<=0.15 and abs(d_cap)<=0.15 and abs(d_total)<=0.2:
                    st.success('Planner sluit exact aan op de optimizer: dezelfde lineups, captainplanning én 4GW objective-weging.')
                else:
                    st.error('Planner en optimizer sluiten NIET exact aan. Gebruik de planner niet voor beslissingen totdat deze afwijking is verklaard.')

        with tab4:
            prows=[]
            for p in sorted(selectable,key=lambda p:-p['_proj']):
                dd=durability_details(p); gp=dd['games']
                prows.append({'Speler':p['web_name'],'Team':teams[int(p['team'])]['short_name'],
                    'Pos':type_names[int(p['element_type'])],'Prijs':num(p['now_cost'])/currency,
                    'Base':round(base_projection(p,currency),2),'Risk-adjusted':round(p['_proj'],2),
                    'Actueel %':round(current_availability(p)*100),'Durability %':('ONBEKEND' if dd['score'] is None else round(dd['score']*100)),
                    'Historie':f"{dd['years_count']}Y",'GP 23-24':gp.get(2024,''),'GP 24-25':gp.get(2025,''),'GP 25-26':gp.get(2026,''),
                    '<60 seizoenen':dd['low_seasons'],'Bron':dd['source'],'Status':p.get('status','')})
            html_table(prows,650)

        with tab5:
            con=db_connect()
            snap=con.execute('SELECT COUNT(*),COUNT(DISTINCT captured_at),COUNT(DISTINCT player_id) FROM snapshots').fetchone()
            dh=con.execute('SELECT COUNT(*),COUNT(DISTINCT player_key) FROM durability_seasons').fetchone()
            raw=con.execute('SELECT created_at,mode,budget_cap,spent,objective,roster_points,captain_bonus FROM runs ORDER BY run_id DESC LIMIT 50').fetchall()
            con.close()
            a,b,c,d=st.columns(4)
            a.metric('Snapshot-rijen',snap[0]); b.metric('Meetmomenten',snap[1]); c.metric('Spelers gevolgd',snap[2]); d.metric('3Y durability records',dh[0])
            hist=[{'Datum':r[0],'Modus':r[1],'Budgetcap':r[2],'Gebruikt':r[3],'Objective':r[4],'Startpunten':r[5],'Captainbonus':r[6]} for r in raw]
            html_table(hist,520)
            st.caption(f'Laatste snapshot: {stamp} • database: {DB.name}')

def launch():
    # `python nba_fantasy_v2.4.py` opens the Streamlit app automatically.
    try:
        from streamlit.runtime.scriptrunner import get_script_run_ctx
        running=get_script_run_ctx() is not None
    except Exception: running=False
    if not running:
        subprocess.call([sys.executable,'-m','streamlit','run',str(Path(__file__).resolve())])
    else: app()

if __name__=='__main__': launch()
