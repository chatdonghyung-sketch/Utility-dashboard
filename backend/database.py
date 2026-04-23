import os
from sqlalchemy import text
from sqlalchemy.ext.asyncio import create_async_engine, AsyncSession
from sqlalchemy.orm import sessionmaker

# SQLite DB 파일 위치: 배포 폴더 루트에 mdd_data.db 생성
_DB_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'mdd_data.db'))
DATABASE_URL = f"sqlite+aiosqlite:///{_DB_PATH}"

engine = create_async_engine(DATABASE_URL, echo=False)
async_session = sessionmaker(engine, class_=AsyncSession, expire_on_commit=False)

VALID_TABLES = {'ffu_log', 'boiler_log', 'safety_log'}


async def init_db():
    async with engine.begin() as conn:
        await conn.execute(text('''
            CREATE TABLE IF NOT EXISTS chiller_log (
                year         INTEGER NOT NULL,
                month        INTEGER NOT NULL,
                day          INTEGER NOT NULL,
                equipment_id TEXT NOT NULL,
                param_key    TEXT NOT NULL,
                value        TEXT NOT NULL,
                PRIMARY KEY (year, month, day, equipment_id, param_key)
            )
        '''))
        await conn.execute(text('''
            CREATE TABLE IF NOT EXISTS chiller_inspector (
                year  INTEGER NOT NULL,
                month INTEGER NOT NULL,
                day   INTEGER NOT NULL,
                name  TEXT NOT NULL,
                PRIMARY KEY (year, month, day)
            )
        '''))
        await conn.execute(text('''
            CREATE TABLE IF NOT EXISTS chiller_analysis (
                year       INTEGER NOT NULL,
                month      INTEGER NOT NULL,
                day        INTEGER NOT NULL,
                chiller_id TEXT NOT NULL,
                value      TEXT NOT NULL,
                PRIMARY KEY (year, month, day, chiller_id)
            )
        '''))
        await conn.execute(text('''
            CREATE TABLE IF NOT EXISTS ffu_log (
                date      TEXT NOT NULL,
                field_key TEXT NOT NULL,
                value     TEXT NOT NULL,
                PRIMARY KEY (date, field_key)
            )
        '''))
        await conn.execute(text('''
            CREATE TABLE IF NOT EXISTS boiler_log (
                date      TEXT NOT NULL,
                field_key TEXT NOT NULL,
                value     TEXT NOT NULL,
                PRIMARY KEY (date, field_key)
            )
        '''))
        await conn.execute(text('''
            CREATE TABLE IF NOT EXISTS safety_log (
                date      TEXT NOT NULL,
                field_key TEXT NOT NULL,
                value     TEXT NOT NULL,
                PRIMARY KEY (date, field_key)
            )
        '''))
        await conn.execute(text('''
            CREATE TABLE IF NOT EXISTS work_orders (
                wo_no        TEXT NOT NULL,
                work_type    TEXT,
                work_name    TEXT,
                equip_code   TEXT,
                equip_name   TEXT,
                equip_type   TEXT,
                location     TEXT,
                start_date   TEXT,
                end_date     TEXT,
                department   TEXT,
                worker       TEXT,
                writer       TEXT,
                wo_status    TEXT,
                wo_status_normalized TEXT,
                PRIMARY KEY (wo_no)
            )
        '''))
        await conn.execute(text('CREATE INDEX IF NOT EXISTS idx_wo_start ON work_orders(start_date)'))
        await conn.execute(text('CREATE INDEX IF NOT EXISTS idx_wo_dept ON work_orders(department)'))
        await conn.execute(text('CREATE INDEX IF NOT EXISTS idx_wo_status ON work_orders(wo_status_normalized)'))


# ── 냉동기·냉각탑 운전일지 ────────────────────────────────────────────
async def load_chiller_month(year: int, month: int) -> dict:
    mk = f"{year}-{str(month).zfill(2)}"
    result = {}
    async with async_session() as session:
        rows = await session.execute(
            text('SELECT day, equipment_id, param_key, value FROM chiller_log WHERE year=:y AND month=:m'),
            {'y': year, 'm': month},
        )
        for row in rows:
            key = f"{mk}__{row.day}__{row.equipment_id}__{row.param_key}"
            result[key] = row.value
        rows2 = await session.execute(
            text('SELECT day, name FROM chiller_inspector WHERE year=:y AND month=:m'),
            {'y': year, 'm': month},
        )
        for row in rows2:
            result[f"{mk}__{row.day}__inspector"] = row.name
    return result


async def save_chiller_cell(year: int, month: int, day: int, equipment_id: str, param_key: str, value: str):
    async with async_session() as session:
        await session.execute(text('''
            INSERT INTO chiller_log (year, month, day, equipment_id, param_key, value)
            VALUES (:year, :month, :day, :equipment_id, :param_key, :value)
            ON CONFLICT(year, month, day, equipment_id, param_key) DO UPDATE SET value=excluded.value
        '''), {'year': year, 'month': month, 'day': day,
               'equipment_id': equipment_id, 'param_key': param_key, 'value': value})
        await session.commit()


async def save_chiller_inspector(year: int, month: int, day: int, name: str):
    async with async_session() as session:
        await session.execute(text('''
            INSERT INTO chiller_inspector (year, month, day, name)
            VALUES (:year, :month, :day, :name)
            ON CONFLICT(year, month, day) DO UPDATE SET name=excluded.name
        '''), {'year': year, 'month': month, 'day': day, 'name': name})
        await session.commit()


async def save_chiller_bulk(year: int, month: int, cells: list, inspectors: list):
    async with async_session() as session:
        for c in cells:
            await session.execute(text('''
                INSERT INTO chiller_log (year, month, day, equipment_id, param_key, value)
                VALUES (:year, :month, :day, :equipment_id, :param_key, :value)
                ON CONFLICT(year, month, day, equipment_id, param_key) DO UPDATE SET value=excluded.value
            '''), {'year': year, 'month': month, 'day': c['day'],
                   'equipment_id': c['equipment_id'], 'param_key': c['param_key'], 'value': c['value']})
        for ins in inspectors:
            await session.execute(text('''
                INSERT INTO chiller_inspector (year, month, day, name)
                VALUES (:year, :month, :day, :name)
                ON CONFLICT(year, month, day) DO UPDATE SET name=excluded.name
            '''), {'year': year, 'month': month, 'day': ins['day'], 'name': ins['name']})
        await session.commit()


# ── 냉동기 가동 분석 ──────────────────────────────────────────────────
async def load_chiller_analysis(year: int, month: int) -> dict:
    mk = f"{year}-{str(month).zfill(2)}"
    result = {}
    async with async_session() as session:
        rows = await session.execute(
            text('SELECT day, chiller_id, value FROM chiller_analysis WHERE year=:y AND month=:m'),
            {'y': year, 'm': month},
        )
        for row in rows:
            result[f"{mk}__{row.day}__{row.chiller_id}"] = row.value
    return result


async def save_chiller_analysis_cell(year: int, month: int, day: int, chiller_id: str, value: str):
    async with async_session() as session:
        await session.execute(text('''
            INSERT INTO chiller_analysis (year, month, day, chiller_id, value)
            VALUES (:year, :month, :day, :chiller_id, :value)
            ON CONFLICT(year, month, day, chiller_id) DO UPDATE SET value=excluded.value
        '''), {'year': year, 'month': month, 'day': day, 'chiller_id': chiller_id, 'value': value})
        await session.commit()


async def save_chiller_analysis_bulk(year: int, month: int, records: list):
    async with async_session() as session:
        for r in records:
            await session.execute(text('''
                INSERT INTO chiller_analysis (year, month, day, chiller_id, value)
                VALUES (:year, :month, :day, :chiller_id, :value)
                ON CONFLICT(year, month, day, chiller_id) DO UPDATE SET value=excluded.value
            '''), {'year': year, 'month': month, 'day': int(r['day']),
                   'chiller_id': r['chiller_id'], 'value': str(r['value'])})
        await session.commit()


# ── 일별 key-value 공용 CRUD (FFU / Boiler / Safety) ─────────────────
async def load_date_form(table: str, date: str) -> dict:
    if table not in VALID_TABLES:
        raise ValueError(f"Invalid table: {table}")
    result = {}
    async with async_session() as session:
        rows = await session.execute(
            text(f'SELECT field_key, value FROM {table} WHERE date=:date'),
            {'date': date},
        )
        for row in rows:
            result[f"{date}__{row.field_key}"] = row.value
    return result


async def save_date_form_cell(table: str, date: str, field_key: str, value: str):
    if table not in VALID_TABLES:
        raise ValueError(f"Invalid table: {table}")
    async with async_session() as session:
        await session.execute(text(f'''
            INSERT INTO {table} (date, field_key, value)
            VALUES (:date, :field_key, :value)
            ON CONFLICT(date, field_key) DO UPDATE SET value=excluded.value
        '''), {'date': date, 'field_key': field_key, 'value': value})
        await session.commit()


async def save_date_form_bulk(table: str, date: str, records: list):
    if table not in VALID_TABLES:
        raise ValueError(f"Invalid table: {table}")
    async with async_session() as session:
        for r in records:
            if not r.get('key') or r.get('value') is None:
                continue
            await session.execute(text(f'''
                INSERT INTO {table} (date, field_key, value)
                VALUES (:date, :field_key, :value)
                ON CONFLICT(date, field_key) DO UPDATE SET value=excluded.value
            '''), {'date': date, 'field_key': r['key'], 'value': str(r['value'])})
        await session.commit()


# ── Work Orders ──────────────────────────────────────────────────────
async def upsert_work_orders(records: list) -> dict:
    inserted = 0
    updated = 0
    async with async_session() as session:
        for r in records:
            raw_status = str(r.get('wo_status', '') or '')
            normalized = raw_status.replace(' ', '')
            existing = await session.execute(
                text('SELECT wo_no FROM work_orders WHERE wo_no=:wo_no'),
                {'wo_no': r.get('wo_no', '')}
            )
            if existing.fetchone():
                updated += 1
            else:
                inserted += 1
            await session.execute(text('''
                INSERT INTO work_orders
                    (wo_no, work_type, work_name, equip_code, equip_name, equip_type,
                     location, start_date, end_date, department, worker, writer,
                     wo_status, wo_status_normalized)
                VALUES
                    (:wo_no, :work_type, :work_name, :equip_code, :equip_name, :equip_type,
                     :location, :start_date, :end_date, :department, :worker, :writer,
                     :wo_status, :wo_status_normalized)
                ON CONFLICT(wo_no) DO UPDATE SET
                    work_type=excluded.work_type,
                    work_name=excluded.work_name,
                    equip_code=excluded.equip_code,
                    equip_name=excluded.equip_name,
                    equip_type=excluded.equip_type,
                    location=excluded.location,
                    start_date=excluded.start_date,
                    end_date=excluded.end_date,
                    department=excluded.department,
                    worker=excluded.worker,
                    writer=excluded.writer,
                    wo_status=excluded.wo_status,
                    wo_status_normalized=excluded.wo_status_normalized
            '''), {
                'wo_no':                str(r.get('wo_no', '') or ''),
                'work_type':            str(r.get('work_type', '') or ''),
                'work_name':            str(r.get('work_name', '') or ''),
                'equip_code':           str(r.get('equip_code', '') or ''),
                'equip_name':           str(r.get('equip_name', '') or ''),
                'equip_type':           str(r.get('equip_type', '') or ''),
                'location':             str(r.get('location', '') or ''),
                'start_date':           str(r.get('start_date', '') or ''),
                'end_date':             str(r.get('end_date', '') or ''),
                'department':           str(r.get('department', '') or ''),
                'worker':               str(r.get('worker', '') or ''),
                'writer':               str(r.get('writer', '') or ''),
                'wo_status':            raw_status,
                'wo_status_normalized': normalized,
            })
        await session.commit()
    return {'inserted': inserted, 'updated': updated}


async def query_work_orders(
    start_date: str = None, end_date: str = None,
    department: str = None, wo_status: str = None,
    work_type: str = None, equip_type: str = None,
    location: str = None, worker: str = None,
    page: int = 1, page_size: int = 50,
) -> dict:
    conditions = []
    params = {}
    if start_date:
        conditions.append('start_date >= :start_date')
        params['start_date'] = start_date
    if end_date:
        conditions.append('start_date <= :end_date')
        params['end_date'] = end_date
    if department:
        conditions.append('department = :department')
        params['department'] = department
    if wo_status:
        conditions.append('wo_status_normalized = :wo_status')
        params['wo_status'] = wo_status.replace(' ', '')
    if work_type:
        conditions.append('work_type = :work_type')
        params['work_type'] = work_type
    if equip_type:
        conditions.append('equip_type = :equip_type')
        params['equip_type'] = equip_type
    if location:
        conditions.append('location = :location')
        params['location'] = location
    if worker:
        conditions.append('worker LIKE :worker')
        params['worker'] = f'%{worker}%'

    where = ('WHERE ' + ' AND '.join(conditions)) if conditions else ''
    offset = (page - 1) * page_size

    async with async_session() as session:
        count_row = await session.execute(
            text(f'SELECT COUNT(*) FROM work_orders {where}'), params
        )
        total = count_row.scalar()
        rows = await session.execute(
            text(f'''SELECT wo_no, work_type, work_name, equip_code, equip_name,
                            equip_type, location, start_date, end_date, department,
                            worker, writer, wo_status
                     FROM work_orders {where}
                     ORDER BY start_date DESC, wo_no DESC
                     LIMIT :limit OFFSET :offset'''),
            {**params, 'limit': page_size, 'offset': offset}
        )
        data = [dict(r._mapping) for r in rows]

    return {'total': total, 'page': page, 'page_size': page_size, 'rows': data}


async def get_work_order_kpi(start_date: str = None, end_date: str = None) -> dict:
    conditions = []
    params = {}
    if start_date:
        conditions.append('start_date >= :start_date')
        params['start_date'] = start_date
    if end_date:
        conditions.append('start_date <= :end_date')
        params['end_date'] = end_date
    where = ('WHERE ' + ' AND '.join(conditions)) if conditions else ''

    async with async_session() as session:
        rows = await session.execute(
            text(f'SELECT wo_status_normalized, COUNT(*) as cnt FROM work_orders {where} GROUP BY wo_status_normalized'),
            params
        )
        counts = {r.wo_status_normalized: r.cnt for r in rows}

    confirmed = counts.get('확정', 0)
    drafting  = counts.get('작성중', 0)
    pending   = counts.get('작성대기', 0)
    total     = sum(counts.values())
    return {
        'confirmed': confirmed,
        'drafting':  drafting,
        'pending':   pending,
        'total':     total,
        'other':     total - confirmed - drafting - pending,
    }


async def get_work_order_filter_options() -> dict:
    async with async_session() as session:
        dept_rows   = await session.execute(text("SELECT DISTINCT department FROM work_orders WHERE department != '' ORDER BY department"))
        type_rows   = await session.execute(text("SELECT DISTINCT work_type FROM work_orders WHERE work_type != '' ORDER BY work_type"))
        etype_rows  = await session.execute(text("SELECT DISTINCT equip_type FROM work_orders WHERE equip_type != '' ORDER BY equip_type"))
        loc_rows    = await session.execute(text("SELECT DISTINCT location FROM work_orders WHERE location != '' ORDER BY location"))
        worker_rows = await session.execute(text("SELECT DISTINCT worker FROM work_orders WHERE worker != '' ORDER BY worker"))
        status_rows = await session.execute(text("SELECT DISTINCT wo_status FROM work_orders WHERE wo_status != '' ORDER BY wo_status"))
    return {
        'departments': [r[0] for r in dept_rows],
        'work_types':  [r[0] for r in type_rows],
        'equip_types': [r[0] for r in etype_rows],
        'locations':   [r[0] for r in loc_rows],
        'workers':     [r[0] for r in worker_rows],
        'wo_statuses': [r[0] for r in status_rows],
    }
