import os
import uuid
import json
import sqlite3
import hashlib
import secrets
from datetime import datetime, timedelta
from contextlib import contextmanager

from fastapi import FastAPI, UploadFile, File, Form, Request, HTTPException, Depends
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
import jwt
import openpyxl

# --- Config ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "data.db")
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
SECRET_KEY = os.getenv("SECRET_KEY", "safety-dashboard-secret-key-change-me")
PASSWORD = os.getenv("DASHBOARD_PASSWORD", "cosmax")
JWT_ALGORITHM = "HS256"
JWT_EXPIRE_HOURS = 24

CHANNELS = [
    "정기위험성평가(코스맥스)",
    "정기위험성평가(협력사)",
    "수시위험성평가",
    "안전점검",
    "부서별 위험요소발굴",
    "근로자 제안",
    "5S/EHS평가",
]

os.makedirs(UPLOAD_DIR, exist_ok=True)

app = FastAPI()
security = HTTPBearer(auto_error=False)

# --- Static Files ---
app.mount("/static", StaticFiles(directory=os.path.join(BASE_DIR, "static")), name="static")
app.mount("/uploads", StaticFiles(directory=UPLOAD_DIR), name="uploads")

# --- Database ---
def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    return conn

def init_db():
    conn = get_db()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS records (
            id TEXT PRIMARY KEY,
            channel TEXT NOT NULL,
            month TEXT,
            person TEXT,
            date TEXT,
            location TEXT,
            location_group TEXT,
            content TEXT,
            process TEXT,
            disaster_type TEXT,
            likelihood_before INTEGER DEFAULT 0,
            severity_before INTEGER DEFAULT 0,
            risk_before INTEGER DEFAULT 0,
            grade_before TEXT DEFAULT '-',
            improvement_plan TEXT,
            likelihood_after INTEGER DEFAULT 0,
            severity_after INTEGER DEFAULT 0,
            risk_after INTEGER DEFAULT 0,
            grade_after TEXT DEFAULT '-',
            completion TEXT DEFAULT '미완료',
            week INTEGER DEFAULT 0,
            image TEXT,
            image_after TEXT,
            created_at TEXT
        )
    """)
    conn.commit()
    conn.close()

init_db()

# --- Auth ---
def create_token():
    exp = datetime.utcnow() + timedelta(hours=JWT_EXPIRE_HOURS)
    return jwt.encode({"exp": exp, "sub": "user"}, SECRET_KEY, algorithm=JWT_ALGORITHM)

def verify_token(credentials: HTTPAuthorizationCredentials = Depends(security)):
    if not credentials:
        raise HTTPException(status_code=401, detail="Not authenticated")
    try:
        jwt.decode(credentials.credentials, SECRET_KEY, algorithms=[JWT_ALGORITHM])
    except jwt.ExpiredSignatureError:
        raise HTTPException(status_code=401, detail="Token expired")
    except jwt.InvalidTokenError:
        raise HTTPException(status_code=401, detail="Invalid token")
    return True

# --- Helpers ---
def calc_grade(likelihood, severity):
    if likelihood <= 0 or severity <= 0:
        return 0, "-"
    risk = likelihood * severity
    if risk <= 4:
        grade = "A"
    elif risk <= 8:
        grade = "B"
    elif risk <= 12:
        grade = "C"
    else:
        grade = "D"
    return risk, grade

def extract_location_group(location):
    if not location:
        return ""
    return location.strip()

# --- Routes ---
@app.get("/", response_class=HTMLResponse)
async def root():
    return FileResponse(os.path.join(BASE_DIR, "static", "index.html"))

@app.post("/api/login")
async def login(request: Request):
    body = await request.json()
    pw = body.get("password", "")
    if pw != PASSWORD:
        raise HTTPException(status_code=401, detail="Invalid password")
    token = create_token()
    return {"token": token}

COLUMN_ALIASES = {
    "month": ["월", "점검월", "월도", "개월"],
    "person": ["발굴자", "담당자", "점검자", "이름", "점검원", "작성자", "평가자", "성명", "작업자", "검사자"],
    "date": ["일시", "날짜", "발굴일", "점검일", "평가일", "작성일", "일자", "년월일", "작성일자", "점검날짜"],
    "location": ["장소", "위치", "공장", "점검장소", "점검위치", "작업장소", "작업장", "구분", "부서", "영역", "지역"],
    "content": ["위험요소 내용", "위험요소내용", "내용", "점검내용", "위험내용",
                 "위험요소", "불안전상태", "불안전행동", "지적사항", "발굴내용",
                 "위험성", "유해위험요인", "유해·위험요인", "세부내용", "설명", "비고", "개요"],
    "process": ["공정", "공정명", "작업공정", "작업명", "세부공정", "프로세스", "부공정", "단계"],
    "disaster_type": ["재해유형", "유형", "재해", "사고유형", "위험분류", "분류", "사고분류", "카테고리", "종류"],
    "likelihood_before": ["가능성", "빈도", "발생빈도", "가능성(전)", "빈도(전)", "빈도전", "발생가능성"],
    "severity_before": ["중대성", "심각성", "강도", "중대성(전)", "강도(전)", "심각성(전)", "중대성전", "영향도"],
    "risk_before": ["위험도", "위험도(전)", "위험성", "Risk", "점수", "레벨", "수준", "위험도전"],
    "grade_before": ["등급", "위험등급", "등급(전)", "Risk등급", "등급전", "관리등급", "레벨"],
    "improvement": ["개선대책", "개선", "대책", "개선방안", "조치내용", "개선내용", "조치사항", "안전대책", "예방방법", "조치", "개선계획"],
    "likelihood_after": ["가능성(후)", "빈도(후)", "발생빈도(후)", "빈도후", "가능성후"],
    "severity_after": ["중대성(후)", "심각성(후)", "강도(후)", "중대성후", "심각성후"],
    "risk_after": ["위험도(후)", "위험성(후)", "위험도후", "점수후"],
    "grade_after": ["등급(후)", "위험등급(후)", "등급후"],
    "completion": ["완료", "완료여부", "상태", "조치완료", "이행여부", "진행상태", "이행상태", "처리상태", "완료상태"],
    "week": ["주차", "주", "주수", "주간"],
}

def find_header_row(ws, max_scan=30):
    """Scan rows to find the header row by matching known column names."""
    content_keywords = COLUMN_ALIASES["content"]
    for row in ws.iter_rows(min_row=1, max_row=max_scan):
        cell_texts = []
        for cell in row:
            val = str(cell.value or "").strip().replace("\n", " ")
            cell_texts.append(val)
        # Check if this row has a content-like column
        for text in cell_texts:
            for kw in content_keywords:
                if kw in text:
                    return row[0].row, cell_texts
    return None, []

def map_columns(header_texts):
    """Map column field names to column indices based on header texts."""
    col_map = {}
    for field, aliases in COLUMN_ALIASES.items():
        for idx, text in enumerate(header_texts):
            # Exact match first
            if text in aliases:
                col_map[field] = idx
                break
        # If no exact match, try partial match
        if field not in col_map:
            for alias in aliases:
                for idx, text in enumerate(header_texts):
                    if alias in text and text.strip() != "":
                        col_map[field] = idx
                        break
                if field in col_map:
                    break
    return col_map

def parse_grade_from_text(text):
    """Extract grade letter from text like 'A', 'A등급', 'D(16)' etc."""
    if not text:
        return "-"
    text = text.strip().upper()
    for g in ["A", "B", "C", "D", "E", "F"]:
        if text.startswith(g):
            return g
    return "-"

@app.post("/api/upload")
async def upload_file(
    file: UploadFile = File(...),
    channel: str = Form(...),
    _auth: bool = Depends(verify_token),
):
    if not file.filename.endswith((".xlsx", ".xlsm")):
        raise HTTPException(status_code=400, detail="xlsx/xlsm 파일만 업로드 가능합니다.")

    try:
        raw = await file.read()
        wb = openpyxl.load_workbook(
            filename=__import__("io").BytesIO(raw),
            data_only=True,
        )
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"엑셀 파일을 읽을 수 없습니다: {str(e)}")

    conn = get_db()
    total_count = 0
    error_rows = []

    for sheet_idx, ws in enumerate(wb.worksheets):
        header_row_num, header_texts = find_header_row(ws)
        if header_row_num is None:
            continue

        col_map = map_columns(header_texts)
        if "content" not in col_map:
            continue

        row_num = header_row_num
        for row in ws.iter_rows(min_row=header_row_num + 1, values_only=False):
            row_num += 1
            # Ensure we have enough cells
            max_col_idx = max(col_map.values()) if col_map else 0
            vals = [cell.value if cell.value is not None else "" for cell in row[:max_col_idx + 1]]
            # Pad with empty strings if needed
            while len(vals) <= max_col_idx:
                vals.append("")

            def get_val(field, default=""):
                try:
                    idx = col_map.get(field)
                    if idx is not None and idx < len(vals):
                        v = vals[idx]
                        if v is None or v == "":
                            return default
                        if isinstance(v, datetime):
                            return v.strftime("%Y-%m-%d")
                        return str(v).strip()
                    return default
                except Exception:
                    return default

            def get_int(field, default=0):
                v = get_val(field, "")
                if not v:
                    return default
                try:
                    # Handle various number formats
                    v_clean = str(v).replace(",", "").replace("%", "").strip()
                    return int(float(v_clean))
                except (ValueError, TypeError):
                    return default

            try:
                content_text = get_val("content")
                if not content_text or content_text == "" or content_text.lower() in ["없음", "na", "n/a"]:
                    continue

                month = get_val("month")
                person = get_val("person")
                date_val = get_val("date")
                # Try to normalize date format using stdlib only
                if date_val and isinstance(date_val, str):
                    raw_date = date_val.strip()
                    normalized = raw_date
                    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%y-%m-%d", "%y/%m/%d", "%y.%m.%d"):
                        try:
                            normalized = datetime.strptime(raw_date, fmt).strftime("%Y-%m-%d")
                            break
                        except Exception:
                            continue
                    date_val = normalized
                
                location = get_val("location")
                process = get_val("process")
                disaster_type = get_val("disaster_type")
                improvement = get_val("improvement")
                completion = get_val("completion")
                week = get_int("week")

                # Likelihood / Severity
                lh_before = get_int("likelihood_before")
                sv_before = get_int("severity_before")
                lh_after = get_int("likelihood_after")
                sv_after = get_int("severity_after")

                # Try to read grade directly from Excel, else calculate
                grade_before_text = get_val("grade_before")
                grade_after_text = get_val("grade_after")

                if lh_before > 0 and sv_before > 0:
                    risk_before, grade_before = calc_grade(lh_before, sv_before)
                elif grade_before_text:
                    grade_before = parse_grade_from_text(grade_before_text)
                    risk_before = get_int("risk_before")
                else:
                    risk_before, grade_before = 0, "-"

                if lh_after > 0 and sv_after > 0:
                    risk_after, grade_after = calc_grade(lh_after, sv_after)
                elif grade_after_text:
                    grade_after = parse_grade_from_text(grade_after_text)
                    risk_after = get_int("risk_after")
                else:
                    risk_after, grade_after = 0, "-"

                # Normalize completion
                if completion and "완료" in completion and "미" not in completion:
                    completion = "완료"
                else:
                    completion = "미완료"

                record_id = str(uuid.uuid4())
                conn.execute(
                    """INSERT INTO records (id, channel, month, person, date, location, location_group,
                       content, process, disaster_type, likelihood_before, severity_before,
                       risk_before, grade_before, improvement_plan,
                       likelihood_after, severity_after, risk_after, grade_after,
                       completion, week, created_at)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                    (record_id, channel, month, person, date_val, location,
                     extract_location_group(location), content_text, process, disaster_type,
                     lh_before, sv_before, risk_before, grade_before, improvement,
                     lh_after, sv_after, risk_after, grade_after,
                     completion, week, datetime.utcnow().isoformat()),
                )
                total_count += 1
            except Exception as e:
                error_rows.append({"row": row_num, "error": str(e)})
                continue

    conn.commit()
    conn.close()

    if total_count == 0:
        error_detail = "엑셀 헤더에 '위험요소 내용' 등의 컬럼이 필요합니다."
        if error_rows:
            error_detail += f" (처리 중 {len(error_rows)}개 행 오류)"
        raise HTTPException(status_code=400, detail=error_detail)

    message = f"{total_count}건 업로드 완료 ({channel})"
    if error_rows:
        message += f" (주의: {len(error_rows)}개 행 스킵됨)"
    
    return {"message": message, "uploaded": total_count, "skipped": len(error_rows)}

@app.get("/api/summary")
async def get_summary(request: Request, _auth: bool = Depends(verify_token)):
    params = dict(request.query_params)
    conn = get_db()

    where_clauses = []
    where_params = []

    if params.get("channel"):
        where_clauses.append("channel = ?")
        where_params.append(params["channel"])
    if params.get("year"):
        where_clauses.append("(date LIKE ? OR month LIKE ?)")
        where_params.extend([f"{params['year']}%", f"%{params['year']}%"])
    if params.get("month"):
        where_clauses.append("month = ?")
        where_params.append(params["month"])
    if params.get("location"):
        where_clauses.append("location = ?")
        where_params.append(params["location"])
    if params.get("grade"):
        where_clauses.append("grade_before = ?")
        where_params.append(params["grade"])
    if params.get("disaster_type"):
        where_clauses.append("disaster_type = ?")
        where_params.append(params["disaster_type"])
    if params.get("process"):
        where_clauses.append("process = ?")
        where_params.append(params["process"])
    if params.get("person"):
        where_clauses.append("person = ?")
        where_params.append(params["person"])
    if params.get("week"):
        where_clauses.append("week = ?")
        where_params.append(int(params["week"]))
    if params.get("completion"):
        where_clauses.append("completion = ?")
        where_params.append(params["completion"])
    if params.get("keyword"):
        where_clauses.append("content LIKE ?")
        where_params.append(f"%{params['keyword']}%")

    where_sql = " AND ".join(where_clauses) if where_clauses else "1=1"

    rows = conn.execute(
        f"SELECT * FROM records WHERE {where_sql} ORDER BY created_at DESC",
        where_params,
    ).fetchall()

    # Also fetch ALL records (unfiltered) for filter options
    all_rows = conn.execute("SELECT * FROM records ORDER BY created_at DESC").fetchall()
    conn.close()

    records = [dict(r) for r in rows]
    all_records = [dict(r) for r in all_rows]

    # Stats
    total = len(records)
    grade_a = sum(1 for r in records if r["grade_before"] == "A")
    grade_b = sum(1 for r in records if r["grade_before"] == "B")
    grade_c = sum(1 for r in records if r["grade_before"] == "C")
    grade_d = sum(1 for r in records if r["grade_before"] == "D")
    complete = sum(1 for r in records if r["completion"] == "완료")
    incomplete = total - complete
    improvement_rate = round(complete / total * 100, 1) if total > 0 else 0

    # Repeat detection (by content)
    content_count = {}
    for r in all_records:
        c = (r["content"] or "").strip()
        if c:
            content_count[c] = content_count.get(c, 0) + 1

    repeat_total = 0
    for r in records:
        c = (r["content"] or "").strip()
        r["is_repeat"] = content_count.get(c, 0) >= 2
        r["repeat_count"] = content_count.get(c, 0)
        if r["is_repeat"]:
            repeat_total += 1

    # Location stats
    location_stats = {}
    location_disaster_stats = {}
    for r in records:
        loc = r["location"] or "미분류"
        if loc not in location_stats:
            location_stats[loc] = {"A": 0, "B": 0, "C": 0, "D": 0, "-": 0}
        g = r["grade_before"] or "-"
        if g in location_stats[loc]:
            location_stats[loc][g] += 1

        if loc not in location_disaster_stats:
            location_disaster_stats[loc] = {}
        dt = r["disaster_type"] or "기타"
        location_disaster_stats[loc][dt] = location_disaster_stats[loc].get(dt, 0) + 1

    # Grade cumulative (monthly incomplete remaining)
    grade_cumulative = {}
    months_set = sorted(set(r["month"] or "" for r in all_records if r["month"]))
    # Build cumulative incomplete by grade per month
    for m in months_set:
        if not m:
            continue
        cum = {"A": 0, "B": 0, "C": 0, "D": 0, "total_remaining": 0}
        for r in all_records:
            if not r["month"]:
                continue
            if r["month"] <= m and r["completion"] != "완료":
                g = r["grade_before"] or "-"
                if g in cum:
                    cum[g] += 1
                cum["total_remaining"] += 1
        grade_cumulative[m] = cum

    # Week stats
    week_stats = {}
    for r in records:
        w = str(r["week"] or 0)
        if w == "0":
            continue
        wk = w + "주차"
        week_stats[wk] = week_stats.get(wk, 0) + 1

    # Disaster stats
    disaster_stats = {}
    for r in records:
        dt = r["disaster_type"] or "기타"
        disaster_stats[dt] = disaster_stats.get(dt, 0) + 1

    # Process stats
    process_stats = {}
    for r in records:
        p = r["process"] or "미분류"
        process_stats[p] = process_stats.get(p, 0) + 1

    # Channel stats
    channel_stats = {}
    channel_grade_stats = {}
    for r in records:
        ch = r["channel"]
        channel_stats[ch] = channel_stats.get(ch, 0) + 1
        if ch not in channel_grade_stats:
            channel_grade_stats[ch] = {"A": 0, "B": 0, "C": 0, "D": 0, "complete": 0, "incomplete": 0}
        g = r["grade_before"] or "-"
        if g in channel_grade_stats[ch]:
            channel_grade_stats[ch][g] += 1
        if r["completion"] == "완료":
            channel_grade_stats[ch]["complete"] += 1
        else:
            channel_grade_stats[ch]["incomplete"] += 1

    # Filter options (from all records)
    filters = {
        "channels": sorted(set(r["channel"] for r in all_records if r["channel"])),
        "years": sorted(set(r["date"][:4] for r in all_records if r["date"] and len(r["date"]) >= 4)),
        "months": sorted(set(r["month"] for r in all_records if r["month"])),
        "locations": sorted(set(r["location"] for r in all_records if r["location"])),
        "disaster_types": sorted(set(r["disaster_type"] for r in all_records if r["disaster_type"])),
        "processes": sorted(set(r["process"] for r in all_records if r["process"])),
        "persons": sorted(set(r["person"] for r in all_records if r["person"])),
        "weeks": sorted(set(r["week"] for r in all_records if r["week"] and r["week"] > 0)),
    }

    # Format records for frontend
    output_records = []
    for i, r in enumerate(records, 1):
        content_full = r["content"] or ""
        content_short = content_full[:60] + "..." if len(content_full) > 60 else content_full
        output_records.append({
            "_id": r["id"],
            "no": i,
            "channel": r["channel"],
            "month": r["month"] or "",
            "person": r["person"] or "",
            "date": r["date"] or "",
            "location": r["location"] or "",
            "location_group": r["location_group"] or "",
            "content": content_short,
            "content_full": content_full,
            "process": r["process"] or "",
            "disaster_type": r["disaster_type"] or "",
            "likelihood_before": r["likelihood_before"],
            "severity_before": r["severity_before"],
            "risk_before": r["risk_before"],
            "grade_before": r["grade_before"] or "-",
            "improvement_plan": r["improvement_plan"] or "",
            "likelihood_after": r["likelihood_after"],
            "severity_after": r["severity_after"],
            "risk_after": r["risk_after"],
            "grade_after": r["grade_after"] or "-",
            "completion": r["completion"] or "미완료",
            "week": r["week"],
            "image": r["image"] or "",
            "image_after": r["image_after"] or "",
            "is_repeat": r["is_repeat"],
            "repeat_count": r["repeat_count"],
        })

    return {
        "total": total,
        "grade_a": grade_a,
        "grade_b": grade_b,
        "grade_c": grade_c,
        "grade_d": grade_d,
        "complete": complete,
        "incomplete": incomplete,
        "improvement_rate": improvement_rate,
        "repeat_total": repeat_total,
        "records": output_records,
        "location_stats": location_stats,
        "location_disaster_stats": location_disaster_stats,
        "grade_cumulative": grade_cumulative,
        "week_stats": week_stats,
        "disaster_stats": disaster_stats,
        "process_stats": process_stats,
        "channel_stats": channel_stats,
        "channel_grade_stats": channel_grade_stats,
        "filters": filters,
    }

@app.get("/api/channels/status")
async def channels_status(_auth: bool = Depends(verify_token)):
    conn = get_db()
    counts = {}
    total = 0
    for ch in CHANNELS:
        row = conn.execute("SELECT COUNT(*) as cnt FROM records WHERE channel = ?", (ch,)).fetchone()
        cnt = row["cnt"]
        counts[ch] = cnt
        total += cnt
    conn.close()
    return {"channels": CHANNELS, "counts": counts, "total": total}

@app.post("/api/channels/delete")
async def delete_channel(request: Request, _auth: bool = Depends(verify_token)):
    body = await request.json()
    channel = body.get("channel", "")
    if not channel:
        raise HTTPException(status_code=400, detail="채널명이 필요합니다.")
    conn = get_db()
    cur = conn.execute("DELETE FROM records WHERE channel = ?", (channel,))
    conn.commit()
    deleted = cur.rowcount
    conn.close()
    return {"message": f"[{channel}] {deleted}건 삭제 완료"}

@app.post("/api/record/add")
async def add_record(request: Request, _auth: bool = Depends(verify_token)):
    body = await request.json()
    lh_b = body.get("likelihood_before", 0) or 0
    sv_b = body.get("severity_before", 0) or 0
    lh_a = body.get("likelihood_after", 0) or 0
    sv_a = body.get("severity_after", 0) or 0
    risk_b, grade_b = calc_grade(lh_b, sv_b)
    risk_a, grade_a = calc_grade(lh_a, sv_a)

    record_id = str(uuid.uuid4())
    conn = get_db()
    conn.execute(
        """INSERT INTO records (id, channel, month, person, date, location, location_group,
           content, process, disaster_type, likelihood_before, severity_before,
           risk_before, grade_before, improvement_plan, likelihood_after, severity_after,
           risk_after, grade_after, completion, week, image, image_after, created_at)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
        (record_id, body.get("channel", "안전점검"), body.get("month", ""),
         body.get("person", ""), body.get("date", ""), body.get("location", ""),
         extract_location_group(body.get("location", "")),
         body.get("content", ""), body.get("process", ""),
         body.get("disaster_type", ""), lh_b, sv_b, risk_b, grade_b,
         body.get("improvement_plan", ""), lh_a, sv_a, risk_a, grade_a,
         body.get("completion", "미완료"), body.get("week", 0),
         body.get("image", ""), body.get("image_after", ""),
         datetime.utcnow().isoformat()),
    )
    conn.commit()
    conn.close()
    return {"message": "등록 완료"}

@app.post("/api/record/update")
async def update_record(request: Request, _auth: bool = Depends(verify_token)):
    body = await request.json()
    record_id = body.get("_id", "")
    if not record_id:
        raise HTTPException(status_code=400, detail="레코드 ID가 필요합니다.")

    lh_b = body.get("likelihood_before", 0) or 0
    sv_b = body.get("severity_before", 0) or 0
    lh_a = body.get("likelihood_after", 0) or 0
    sv_a = body.get("severity_after", 0) or 0
    risk_b, grade_b = calc_grade(lh_b, sv_b)
    risk_a, grade_a = calc_grade(lh_a, sv_a)

    conn = get_db()
    conn.execute(
        """UPDATE records SET channel=?, month=?, person=?, date=?, location=?,
           location_group=?, content=?, process=?, disaster_type=?,
           likelihood_before=?, severity_before=?, risk_before=?, grade_before=?,
           improvement_plan=?, likelihood_after=?, severity_after=?, risk_after=?,
           grade_after=?, completion=?, week=?, image=?, image_after=?
           WHERE id=?""",
        (body.get("channel", "안전점검"), body.get("month", ""),
         body.get("person", ""), body.get("date", ""), body.get("location", ""),
         extract_location_group(body.get("location", "")),
         body.get("content", ""), body.get("process", ""),
         body.get("disaster_type", ""), lh_b, sv_b, risk_b, grade_b,
         body.get("improvement_plan", ""), lh_a, sv_a, risk_a, grade_a,
         body.get("completion", "미완료"), body.get("week", 0),
         body.get("image", ""), body.get("image_after", ""),
         record_id),
    )
    conn.commit()
    conn.close()
    return {"message": "수정 완료"}

@app.post("/api/record/delete")
async def delete_record(request: Request, _auth: bool = Depends(verify_token)):
    body = await request.json()
    record_id = body.get("_id", "")
    if not record_id:
        raise HTTPException(status_code=400, detail="레코드 ID가 필요합니다.")
    conn = get_db()
    conn.execute("DELETE FROM records WHERE id = ?", (record_id,))
    conn.commit()
    conn.close()
    return {"message": "삭제 완료"}

@app.post("/api/image/upload")
async def upload_image(
    file: UploadFile = File(...),
    _auth: bool = Depends(verify_token),
):
    ext = os.path.splitext(file.filename)[1].lower()
    if ext not in (".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp"):
        raise HTTPException(status_code=400, detail="이미지 파일만 업로드 가능합니다.")

    filename = f"{uuid.uuid4().hex}{ext}"
    filepath = os.path.join(UPLOAD_DIR, filename)
    content = await file.read()
    with open(filepath, "wb") as f:
        f.write(content)

    return {"url": f"/uploads/{filename}"}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
