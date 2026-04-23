from contextlib import asynccontextmanager
from fastapi import FastAPI, Query, Request, UploadFile, File
from fastapi.responses import JSONResponse, FileResponse
from fastapi.staticfiles import StaticFiles
import os
import glob
import random
import xlrd
import openpyxl
import io

from database import (
    init_db,
    load_chiller_month, save_chiller_cell, save_chiller_inspector, save_chiller_bulk,
    load_chiller_analysis, save_chiller_analysis_cell, save_chiller_analysis_bulk,
    load_date_form, save_date_form_cell, save_date_form_bulk,
    upsert_work_orders, query_work_orders, get_work_order_kpi,
    get_work_order_filter_options,
)


@asynccontextmanager
async def lifespan(app: FastAPI):
    await init_db()
    yield


app = FastAPI(lifespan=lifespan)

TREND_DIR = os.path.normpath(os.path.join(os.path.dirname(__file__), '..', 'HMI_Trend_data'))
DIST_DIR = os.path.normpath(os.path.join(os.path.dirname(__file__), '..', 'dist'))


# ── 냉동기·냉각탑 운전일지 ────────────────────────────────────────────
@app.get('/api/chiller-log')
async def get_chiller_log(year: int = Query(...), month: int = Query(...)):
    data = await load_chiller_month(year, month)
    return JSONResponse(data)


@app.post('/api/chiller-log/cell')
async def save_chiller_cell_route(request: Request):
    body = await request.json()
    if body.get('type') == 'inspector':
        await save_chiller_inspector(body['year'], body['month'], body['day'], body['value'])
    else:
        await save_chiller_cell(body['year'], body['month'], body['day'],
                                body['equipment_id'], body['param_key'], body['value'])
    return JSONResponse({'ok': True})


@app.post('/api/chiller-log/bulk')
async def save_chiller_bulk_route(request: Request):
    body = await request.json()
    await save_chiller_bulk(body['year'], body['month'],
                            body.get('cells', []), body.get('inspectors', []))
    return JSONResponse({'ok': True})


# ── 냉동기 가동 분석 ──────────────────────────────────────────────────
@app.get('/api/chiller-analysis')
async def get_chiller_analysis(year: int = Query(...), month: int = Query(...)):
    data = await load_chiller_analysis(year, month)
    return JSONResponse(data)


@app.post('/api/chiller-analysis/cell')
async def save_analysis_cell(request: Request):
    body = await request.json()
    await save_chiller_analysis_cell(body['year'], body['month'], body['day'],
                                      body['chiller_id'], body['value'])
    return JSONResponse({'ok': True})


@app.post('/api/chiller-analysis/bulk')
async def save_analysis_bulk(request: Request):
    body = await request.json()
    await save_chiller_analysis_bulk(body['year'], body['month'], body.get('records', []))
    return JSONResponse({'ok': True})


# ── 일별 폼 공용 라우트 (FFU / Boiler / Safety) ───────────────────────
@app.get('/api/ffu-log')
async def get_ffu_log(date: str = Query(...)):
    return JSONResponse(await load_date_form('ffu_log', date))

@app.post('/api/ffu-log/cell')
async def save_ffu_log_cell(request: Request):
    body = await request.json()
    await save_date_form_cell('ffu_log', body['date'], body['key'], body['value'])
    return JSONResponse({'ok': True})

@app.post('/api/ffu-log/bulk')
async def save_ffu_log_bulk(request: Request):
    body = await request.json()
    await save_date_form_bulk('ffu_log', body['date'], body.get('records', []))
    return JSONResponse({'ok': True})

@app.get('/api/boiler-log')
async def get_boiler_log(date: str = Query(...)):
    return JSONResponse(await load_date_form('boiler_log', date))

@app.post('/api/boiler-log/cell')
async def save_boiler_log_cell(request: Request):
    body = await request.json()
    await save_date_form_cell('boiler_log', body['date'], body['key'], body['value'])
    return JSONResponse({'ok': True})

@app.post('/api/boiler-log/bulk')
async def save_boiler_log_bulk(request: Request):
    body = await request.json()
    await save_date_form_bulk('boiler_log', body['date'], body.get('records', []))
    return JSONResponse({'ok': True})

@app.get('/api/safety-log')
async def get_safety_log(date: str = Query(...)):
    return JSONResponse(await load_date_form('safety_log', date))

@app.post('/api/safety-log/cell')
async def save_safety_log_cell(request: Request):
    body = await request.json()
    await save_date_form_cell('safety_log', body['date'], body['key'], body['value'])
    return JSONResponse({'ok': True})

@app.post('/api/safety-log/bulk')
async def save_safety_log_bulk(request: Request):
    body = await request.json()
    await save_date_form_bulk('safety_log', body['date'], body.get('records', []))
    return JSONResponse({'ok': True})


# ── HMI 트렌드 API ────────────────────────────────────────────────────
def _read_trend_ws(date: str):
    path = os.path.join(TREND_DIR, f'TREND_{date}.xls')
    if not os.path.exists(path):
        return None
    wb = xlrd.open_workbook(path)
    return wb.sheet_by_name('Raw Data')


def _fill_dummy(series: list, lo: float, hi: float) -> list:
    return [
        round(random.uniform(lo, hi), 3) if (v == '' or v is None) else round(float(v), 3)
        for v in series
    ]


def _ws_to_raw_tags(ws, fab: str, process: str) -> dict:
    """ws에서 fab+process 태그를 raw dict로 반환."""
    tags = {}
    for r in range(1, ws.nrows):
        if ws.cell_value(r, 1) != fab or ws.cell_value(r, 2) != process:
            continue
        name = ws.cell_value(r, 3)
        vtype = ws.cell_value(r, 4)
        unit  = ws.cell_value(r, 29)
        alarm = ws.cell_value(r, 30)
        if name not in tags:
            tags[name] = {'tag': name, 'unit': '', 'alarm_high': None, 'alarm_low': None,
                          'max': [], 'avg': [], 'min': []}
        rec = tags[name]
        if unit: rec['unit'] = unit
        vals = [ws.cell_value(r, c) for c in range(5, 29)]
        if vtype == 'MAX':
            if alarm != '': rec['alarm_high'] = float(alarm)
            rec['max'] = vals
        elif vtype == 'AVG':
            rec['avg'] = vals
        elif vtype == 'MIN':
            if alarm != '': rec['alarm_low'] = float(alarm)
            rec['min'] = vals
    return tags


def _fill_tag(rec: dict) -> dict:
    """빈 시계열을 알람 범위 내 더미로 채우고 MAX>=AVG>=MIN 보장."""
    lo = rec['alarm_low']  if rec['alarm_low']  is not None else 0.0
    hi = rec['alarm_high'] if rec['alarm_high'] is not None else lo + 10.0
    max_v = _fill_dummy(rec['max'], (lo + hi) * 0.55, hi  * 0.95)
    min_v = _fill_dummy(rec['min'], lo * 1.05,        (lo + hi) * 0.45)
    avg_v = _fill_dummy(rec['avg'], (lo + hi) * 0.4,  (lo + hi) * 0.6)
    for i in range(len(max_v)):
        mx, mn = max_v[i], min_v[i]
        if mx < mn: max_v[i], min_v[i] = mn, mx; mx, mn = mn, mx
        avg_v[i] = round((mx + mn) / 2 + random.uniform(-0.05, 0.05) * abs(hi - lo), 3)
    rec['max'], rec['avg'], rec['min'] = max_v, avg_v, min_v
    return rec


@app.get('/api/trend/dates')
async def trend_dates():
    files = sorted(glob.glob(os.path.join(TREND_DIR, 'TREND_????????.xls')))
    return JSONResponse([os.path.basename(f)[6:14] for f in files])


@app.get('/api/trend/filters')
async def trend_filters(date: str = Query('')):
    """최신 파일(또는 지정 날짜)에서 FAB -> [PROCESS] 목록."""
    if not date:
        files = sorted(glob.glob(os.path.join(TREND_DIR, 'TREND_????????.xls')))
        if files:
            date = os.path.basename(files[-1])[6:14]
    ws = _read_trend_ws(date)
    if ws is None:
        return JSONResponse({'error': 'not_found'}, status_code=404)
    fab_map: dict = {}
    for r in range(1, ws.nrows):
        fab_map.setdefault(ws.cell_value(r, 1), set()).add(ws.cell_value(r, 2))
    return JSONResponse({f: sorted(p) for f, p in sorted(fab_map.items())})


@app.get('/api/trend/data')
async def trend_data(
    date: str = Query(''),
    fab: str = Query(''),
    process: str = Query(''),
):
    """단일 날짜, 24시간 시계열."""
    ws = _read_trend_ws(date)
    if ws is None:
        return JSONResponse({'error': 'not_found'}, status_code=404)
    tags = _ws_to_raw_tags(ws, fab, process)
    return JSONResponse([_fill_tag(rec) for rec in tags.values()])


@app.get('/api/trend/range')
async def trend_range(
    start_date: str = Query(''),
    end_date: str = Query(''),
    fab: str = Query(''),
    process: str = Query(''),
):
    """날짜 범위, 일별 집계(MAX의 일최대 / AVG의 일평균 / MIN의 일최소)."""
    from datetime import date as D, timedelta
    try:
        start = D(int(start_date[:4]), int(start_date[4:6]), int(start_date[6:8]))
        end   = D(int(end_date[:4]),   int(end_date[4:6]),   int(end_date[6:8]))
    except (ValueError, IndexError):
        return JSONResponse({'error': 'invalid_dates'}, status_code=400)

    valid_dates = []
    agg: dict = {}   # tag_name -> {meta, max:[], avg:[], min:[]}
    cur = start
    while cur <= end:
        d = cur.strftime('%Y%m%d')
        ws = _read_trend_ws(d)
        if ws is not None:
            valid_dates.append(d)
            for name, rec in _ws_to_raw_tags(ws, fab, process).items():
                _fill_tag(rec)
                if name not in agg:
                    agg[name] = {'tag': rec['tag'], 'unit': rec['unit'],
                                 'alarm_high': rec['alarm_high'], 'alarm_low': rec['alarm_low'],
                                 'max': [], 'avg': [], 'min': []}
                agg[name]['max'].append(round(max(rec['max']), 3))
                agg[name]['avg'].append(round(sum(rec['avg']) / len(rec['avg']), 3))
                agg[name]['min'].append(round(min(rec['min']), 3))
        cur += timedelta(days=1)

    if not valid_dates:
        return JSONResponse({'error': 'no_data'}, status_code=404)
    return JSONResponse({'dates': valid_dates, 'tags': list(agg.values())})


# ── Work Orders ──────────────────────────────────────────────────────
@app.post('/api/work-orders/upload')
async def upload_work_orders(file: UploadFile = File(...)):
    content = await file.read()
    wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True)
    ws = wb.active

    # Find header row (row with 'Wo No')
    header_row = None
    headers = []
    for i, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if row and any(str(c).strip() == 'Wo No' for c in row if c is not None):
            header_row = i
            headers = [str(c).strip() if c is not None else '' for c in row]
            break

    if header_row is None:
        return JSONResponse({'error': 'Header row not found'}, status_code=400)

    col_map = {
        'wo_no':      next((i for i, h in enumerate(headers) if h == 'Wo No'), None),
        'work_type':  next((i for i, h in enumerate(headers) if h == '작업 유형'), None),
        'work_name':  next((i for i, h in enumerate(headers) if h == '작업명'), None),
        'equip_code': next((i for i, h in enumerate(headers) if h == '설비코드'), None),
        'equip_name': next((i for i, h in enumerate(headers) if h == '설비명'), None),
        'equip_type': next((i for i, h in enumerate(headers) if h == '설비종류'), None),
        'location':   next((i for i, h in enumerate(headers) if h == '위치 L4'), None),
        'start_date': next((i for i, h in enumerate(headers) if h == '시작일'), None),
        'end_date':   next((i for i, h in enumerate(headers) if h == '종료일'), None),
        'department': next((i for i, h in enumerate(headers) if h == '작업부서'), None),
        'worker':     next((i for i, h in enumerate(headers) if h == '작업자'), None),
        'writer':     next((i for i, h in enumerate(headers) if h == '작성자'), None),
        'wo_status':  next((i for i, h in enumerate(headers) if h == 'WO 상태'), None),
    }

    records = []
    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        if not row or all(c is None for c in row):
            continue
        def get(key):
            idx = col_map.get(key)
            if idx is None or idx >= len(row):
                return ''
            v = row[idx]
            if v is None:
                return ''
            if hasattr(v, 'strftime'):
                return v.strftime('%Y-%m-%d')
            return str(v).strip()
        wo_no = get('wo_no')
        if not wo_no or wo_no == '예시값':
            continue
        records.append({k: get(k) for k in col_map})

    if not records:
        return JSONResponse({'inserted': 0, 'updated': 0, 'message': '유효한 데이터 없음'})

    result = await upsert_work_orders(records)
    return JSONResponse(result)


@app.get('/api/work-orders')
async def get_work_orders(
    start_date: str = Query(None), end_date: str = Query(None),
    department: str = Query(None), wo_status: str = Query(None),
    work_type: str = Query(None), equip_type: str = Query(None),
    location: str = Query(None), worker: str = Query(None),
    page: int = Query(1), page_size: int = Query(50),
):
    data = await query_work_orders(
        start_date=start_date, end_date=end_date,
        department=department, wo_status=wo_status,
        work_type=work_type, equip_type=equip_type,
        location=location, worker=worker,
        page=page, page_size=page_size,
    )
    return JSONResponse(data)


@app.get('/api/work-orders/kpi')
async def work_order_kpi(
    start_date: str = Query(None),
    end_date: str = Query(None),
):
    data = await get_work_order_kpi(start_date=start_date, end_date=end_date)
    return JSONResponse(data)


@app.get('/api/work-orders/filter-options')
async def work_order_filter_options():
    data = await get_work_order_filter_options()
    return JSONResponse(data)


# ── 에너지 폴더 스캔 ──────────────────────────────────────────────────
import re

# 프로젝트 루트 (backend 상위 폴더)
PROJECT_ROOT = os.path.normpath(os.path.join(os.path.dirname(__file__), '..'))

UTIL_DIR_MAP: dict[str, str] = {
    'LNG':      'gas',
    'STEAM':    'steam',
    'NItrogen': 'nitrogen',
    'NITROGEN': 'nitrogen',
    'Nitrogen': 'nitrogen',
    'ARGON':    'argon',
    'Argon':    'argon',
}

FACTORY_DIR_MAP: dict[str, str] = {
    '1공장': 'f1',
    '2공장': 'f2',
    '3공장': 'f3',
    '신공장': 'fnew',
}


def _parse_year_month(folder_name: str, file_name: str):
    """폴더명 '26년 4월' 또는 파일명 *20260414* 에서 (year, month) 추출"""
    m = re.search(r'(\d{2,4})년\s*(\d{1,2})월', folder_name)
    if m:
        y = int(m.group(1))
        return (y + 2000 if y < 100 else y), int(m.group(2))
    m = re.search(r'(\d{4})(\d{2})\d{2}', file_name)
    if m:
        return int(m.group(1)), int(m.group(2))
    return None, None


def _parse_energy_excel(path: str, year: int, month: int) -> list[dict]:
    """엑셀 파일에서 일별 사용량 파싱. A=날(1-31), B=사용량"""
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        ws = wb.active
        entries = []
        for row in ws.iter_rows(min_row=5, values_only=True):
            if not row or row[0] is None:
                continue
            try:
                day = int(row[0])
                if not (1 <= day <= 31):
                    continue
                # B열(index 1) 우선, 없으면 C열(index 2)
                val = row[1] if (len(row) > 1 and row[1] is not None) else (
                      row[2] if len(row) > 2 else None)
                if val is None:
                    continue
                v = float(val)
                if v <= 0:
                    continue
                date = f"{year}-{str(month).zfill(2)}-{str(day).zfill(2)}"
                entries.append({'date': date, 'value': v})
            except (ValueError, TypeError):
                continue
        wb.close()
        return entries
    except Exception:
        return []


# 스캔 캐시 (5분 TTL)
import time as _time
_energy_cache: dict = {'data': None, 'ts': 0}
_CACHE_TTL = 300  # seconds


def _scan_energy_dirs() -> dict:
    result: dict = {
        'gas':      {'f1': [], 'f2': [], 'f3': [], 'fnew': []},
        'steam':    {'f1': [], 'f2': [], 'f3': [], 'fnew': []},
        'nitrogen': {'f1': [], 'f2': [], 'f3': [], 'fnew': []},
        'argon':    {'f1': [], 'f2': [], 'f3': [], 'fnew': []},
    }

    for util_dir, util_key in UTIL_DIR_MAP.items():
        util_path = os.path.join(PROJECT_ROOT, util_dir)
        if not os.path.isdir(util_path):
            continue
        for factory_dir, factory_key in FACTORY_DIR_MAP.items():
            factory_path = os.path.join(util_path, factory_dir)
            if not os.path.isdir(factory_path):
                continue
            # 직접 파일 + 하위 월별 폴더 모두 스캔
            for root, _dirs, files in os.walk(factory_path):
                folder_name = os.path.basename(root)
                for fname in sorted(files):
                    if not fname.lower().endswith('.xlsx') or fname.startswith('~'):
                        continue
                    year, month = _parse_year_month(folder_name, fname)
                    if year is None:
                        continue
                    fpath = os.path.join(root, fname)
                    entries = _parse_energy_excel(fpath, year, month)
                    result[util_key][factory_key].extend(entries)

    # 날짜 중복 제거 (같은 날짜면 마지막 값 유지)
    for uk in result:
        for fk in result[uk]:
            date_map: dict[str, float] = {}
            for e in result[uk][fk]:
                date_map[e['date']] = e['value']
            result[uk][fk] = [
                {'date': d, 'value': v}
                for d, v in sorted(date_map.items())
            ]
    return result


@app.get('/api/energy-files')
async def get_energy_files(refresh: bool = Query(False)):
    global _energy_cache
    now = _time.time()
    if not refresh and _energy_cache['data'] and (now - _energy_cache['ts']) < _CACHE_TTL:
        return JSONResponse(_energy_cache['data'])
    data = _scan_energy_dirs()
    _energy_cache = {'data': data, 'ts': now}
    return JSONResponse(data)


# ── 프론트엔드 정적 파일 서빙 ─────────────────────────────────────────
# Mount static assets (js, css, images) from dist/assets
if os.path.isdir(os.path.join(DIST_DIR, 'assets')):
    app.mount('/assets', StaticFiles(directory=os.path.join(DIST_DIR, 'assets')), name='assets')


@app.get('/{full_path:path}')
async def serve_frontend(full_path: str):
    file_path = os.path.join(DIST_DIR, full_path)
    if full_path and os.path.isfile(file_path):
        return FileResponse(file_path)
    index = os.path.join(DIST_DIR, 'index.html')
    return FileResponse(index)


# ── 엔트리포인트 ──────────────────────────────────────────────────────
if __name__ == '__main__':
    import uvicorn
    print("=" * 50)
    print("  Maintenance Data Dashboard (FastAPI)")
    print("  http://localhost:5000")
    print("=" * 50)
    uvicorn.run(app, host='0.0.0.0', port=5000)
