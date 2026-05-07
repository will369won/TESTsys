"""
AI智能考试系统 - 后端 v4.0
技术栈: Flask + SQLite + openpyxl
v4.0 新增: 登录页分离/试卷编号/阅卷页/自动阅卷/xlsx保存/统计导出/zip下载/520MB存储限制
v3.1: 教师注册/用户目录管理/密码哈希/试卷数量限制
v3.0: 题目开关/附件、考试设置、多Stata检测、服务器部署
"""

import os
import json
import subprocess
import hashlib
import secrets
import re
import shutil
import time
import random
import zipfile
import io
from datetime import datetime, timedelta
from functools import wraps
from flask import Flask, request, jsonify, send_file, session, g, send_from_directory, Response
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from dotenv import load_dotenv

# ─── App Setup ───
app = Flask(__name__)
app.secret_key = 'testsys-secret-key-change-in-production'
app.config['JSON_AS_ASCII'] = False

# ─── Load .env file ───
_env_path = os.path.join(os.path.dirname(__file__), '.env')
if os.path.exists(_env_path):
    load_dotenv(_env_path)

# ─── Config from env (with .env/file env fallback) ───
DB_PATH = os.environ.get('DB_PATH') or os.path.join(os.path.dirname(__file__), 'testsys.db')

# v3.1 → v3.2: User storage directory
# Priority: env var > .env file > default
_DEFAULT_USERS = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'Users')
USERS_ROOT = os.environ.get('USERS_ROOT') or _DEFAULT_USERS

MAX_EXAMS_PER_TEACHER = int(os.environ.get('MAX_EXAMS_PER_TEACHER', 2))
MAX_FILE_SIZE_MB = int(os.environ.get('MAX_FILE_SIZE_MB', 500))
CLEANUP_DAYS = int(os.environ.get('CLEANUP_DAYS', 7))

# v4.0: Storage limits per exam subfolder
# User/<username>/<exam_title>/ total ≤ 520MB (300MB attachments + 220MB student papers)
MAX_SUBDIR_SIZE_MB = int(os.environ.get('MAX_SUBDIR_SIZE_MB', 520))
MAX_ATTACH_SIZE_MB = int(os.environ.get('MAX_ATTACH_SIZE_MB', 300))
MAX_STUDENT_PAPERS_MB = int(os.environ.get('MAX_STUDENT_PAPERS_MB', 220))

# v3.2: Attachment dir is now dynamic per-exam subfolder (see _get_attachment_dir)

# ─── Database Helpers ───
def get_db():
    if 'db' not in g:
        import sqlite3
        g.db = sqlite3.connect(DB_PATH)
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA journal_mode=WAL")
    return g.db

@app.teardown_appcontext
def close_db(exc):
    db = g.pop('db', None)
    if db is not None:
        db.close()

def init_db():
    """Initialize database tables."""
    db = get_db()
    db.executescript("""
        CREATE TABLE IF NOT EXISTS teachers (
            id TEXT PRIMARY KEY,
            password TEXT NOT NULL,
            username TEXT,
            school TEXT
        );

        CREATE TABLE IF NOT EXISTS students (
            student_id TEXT PRIMARY KEY,
            name TEXT NOT NULL,
            class_name TEXT,
            grade TEXT,
            school TEXT
        );

        CREATE TABLE IF NOT EXISTS exams (
            exam_id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            duration_minutes INTEGER NOT NULL DEFAULT 60,
            created_at TEXT NOT NULL,
            exam_duration INTEGER NOT NULL DEFAULT 60,
            exam_start_time TEXT,
            exam_notice TEXT,
            exam_late_start_limit INTEGER DEFAULT 0,
            exam_late_submit_limit INTEGER DEFAULT 0,
            anti_shuffle_questions INTEGER DEFAULT 0,
            anti_shuffle_options INTEGER DEFAULT 0,
            anti_screen_switch INTEGER DEFAULT 0,
            template_uploaded INTEGER DEFAULT 0
        );

        CREATE TABLE IF NOT EXISTS questions (
            q_id INTEGER PRIMARY KEY AUTOINCREMENT,
            exam_id INTEGER NOT NULL,
            q_type TEXT NOT NULL,
            content TEXT NOT NULL,
            option_a TEXT,
            option_b TEXT,
            option_c TEXT,
            option_d TEXT,
            correct_answer TEXT,
            score REAL NOT NULL DEFAULT 0,
            sort_order INTEGER NOT NULL DEFAULT 0,
            enable_stata INTEGER DEFAULT 0,
            enable_ai INTEGER DEFAULT 0,
            has_attachment INTEGER DEFAULT 0,
            FOREIGN KEY (exam_id) REFERENCES exams(exam_id)
        );

        CREATE TABLE IF NOT EXISTS exam_attachments (
            attachment_id INTEGER PRIMARY KEY AUTOINCREMENT,
            q_id INTEGER NOT NULL,
            original_name TEXT NOT NULL,
            file_path TEXT NOT NULL,
            uploaded_at TEXT NOT NULL,
            FOREIGN KEY (q_id) REFERENCES questions(q_id)
        );

        CREATE TABLE IF NOT EXISTS exam_sessions (
            session_id INTEGER PRIMARY KEY AUTOINCREMENT,
            exam_id INTEGER NOT NULL,
            student_id TEXT NOT NULL,
            student_name TEXT NOT NULL,
            start_time TEXT NOT NULL,
            end_time TEXT,
            status TEXT NOT NULL DEFAULT 'in_progress',
            total_score REAL DEFAULT 0,
            is_graded INTEGER DEFAULT 0,
            terminated_by_teacher INTEGER DEFAULT 0,
            exam_end_time TEXT,
            FOREIGN KEY (exam_id) REFERENCES exams(exam_id),
            FOREIGN KEY (student_id) REFERENCES students(student_id)
        );

        CREATE TABLE IF NOT EXISTS answers (
            answer_id INTEGER PRIMARY KEY AUTOINCREMENT,
            session_id INTEGER NOT NULL,
            q_id INTEGER NOT NULL,
            answer_text TEXT,
            is_correct INTEGER,
            score REAL DEFAULT 0,
            FOREIGN KEY (session_id) REFERENCES exam_sessions(session_id),
            FOREIGN KEY (q_id) REFERENCES questions(q_id)
        );
    """)

    # Seed teacher account (hardcoded, not a registered user)
    cur = db.execute("SELECT id FROM teachers WHERE id = ?", ('001',))
    if cur.fetchone() is None:
        db.execute("INSERT INTO teachers (id, password) VALUES (?, ?)", ('001', '123'))
    db.commit()

    # v2.0: Add new columns if not exist (SQLite 3.35+ supports ALTER TABLE ADD COLUMN)
    try:
        db.execute("ALTER TABLE exam_sessions ADD COLUMN terminated_by_teacher INTEGER DEFAULT 0")
    except Exception:
        pass  # column already exists
    try:
        db.execute("ALTER TABLE exam_sessions ADD COLUMN exam_end_time TEXT")
    except Exception:
        pass  # column already exists

    # v3.0: Add new columns to exams
    for col, default in [
        ('exam_duration', '60'), ('exam_start_time', "''"),
        ('exam_notice', "''"), ('exam_late_start_limit', '0'),
        ('exam_late_submit_limit', '0'), ('anti_shuffle_questions', '0'),
        ('anti_shuffle_options', '0'), ('anti_screen_switch', '0'),
        ('template_uploaded', '0'),
    ]:
        try:
            db.execute(f"ALTER TABLE exams ADD COLUMN {col} DEFAULT {default}")
        except Exception:
            pass

    # v3.0: Add new columns to questions
    for col in ['enable_stata', 'enable_ai', 'has_attachment']:
        try:
            db.execute(f"ALTER TABLE questions ADD COLUMN {col} DEFAULT 0")
        except Exception:
            pass

    # v3.0: Create exam_attachments table
    try:
        db.execute("""
            CREATE TABLE IF NOT EXISTS exam_attachments (
                attachment_id INTEGER PRIMARY KEY AUTOINCREMENT,
                q_id INTEGER NOT NULL,
                original_name TEXT NOT NULL,
                file_path TEXT NOT NULL,
                uploaded_at TEXT NOT NULL,
                FOREIGN KEY (q_id) REFERENCES questions(q_id)
            )
        """)
    except Exception:
        pass

    # v3.0: Ensure attachment directory exists
    attach_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'attachments')
    os.makedirs(attach_dir, exist_ok=True)

    # v3.1: Add new columns to teachers
    for col in ['username', 'school']:
        try:
            db.execute(f"ALTER TABLE teachers ADD COLUMN {col} TEXT")
        except Exception:
            pass  # column already exists

    # v3.1: Ensure user storage root exists
    os.makedirs(USERS_ROOT, exist_ok=True)

    # v3.1: Add teacher_username to exams
    try:
        db.execute("ALTER TABLE exams ADD COLUMN teacher_username TEXT")
    except Exception:
        pass

    # v4.0: Add exam_number (unique 4-digit code) and auto_grade flag
    try:
        db.execute("ALTER TABLE exams ADD COLUMN exam_number TEXT")
    except Exception:
        pass
    try:
        db.execute("ALTER TABLE exams ADD COLUMN auto_grade INTEGER DEFAULT 1")
    except Exception:
        pass
    # v4.0: Add unique index on exam_number
    try:
        db.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_exam_number ON exams(exam_number)")
    except Exception:
        pass

    db.commit()

# ─── Auth Decorator ───
def teacher_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'teacher_id' not in session:
            return jsonify({'error': '未登录'}), 401
        return f(*args, **kwargs)
    return decorated

# ─── v3.1: Password Hashing ───
def _hash_password(password: str, salt: str = None) -> tuple:
    """Hash password with SHA-256 + salt. Returns (hashed, salt)."""
    if salt is None:
        salt = secrets.token_hex(16)
    hashed = hashlib.sha256(f"{salt}{password}".encode()).hexdigest()
    return hashed, salt

def _verify_password(password: str, stored_hash: str, salt: str) -> bool:
    """Verify password against stored hash."""
    h, _ = _hash_password(password, salt)
    return h == stored_hash

# ─── v3.1: User Storage ───
def _user_dir(username: str) -> str:
    """Get user storage directory path."""
    return os.path.join(USERS_ROOT, username)

def _exam_subdir(username: str, exam_title: str) -> str:
    """Get exam subdirectory path."""
    # Sanitize exam title for folder name
    safe = re.sub(r'[<>:"/\\|?*]', '_', exam_title)[:50]
    return os.path.join(_user_dir(username), safe)

def _cleanup_old_exams(username: str):
    """Remove exam subfolders older than CLEANUP_DAYS (7 days)."""
    udir = _user_dir(username)
    if not os.path.exists(udir):
        return
    now = time.time()
    threshold = CLEANUP_DAYS * 24 * 3600
    for entry in os.scandir(udir):
        if entry.is_dir():
            age = now - entry.stat().st_mtime
            if age > threshold:
                try:
                    shutil.rmtree(entry.path)
                except Exception:
                    pass

def _delete_exam_subdir(username: str, exam_title: str):
    """Delete specific exam subdirectory."""
    edir = _exam_subdir(username, exam_title)
    if os.path.exists(edir):
        shutil.rmtree(edir)

def _count_user_exams(username: str) -> int:
    """Count number of exam subdirectories for a user."""
    udir = _user_dir(username)
    if not os.path.exists(udir):
        return 0
    return sum(1 for e in os.scandir(udir) if e.is_dir())


# ════════════════════════════════════════════
# v4.0: Exam Number Generation & Storage Mgmt
# ════════════════════════════════════════════

def _generate_exam_number() -> str:
    """Generate a unique 4-digit exam number. Globally unique."""
    db = get_db()
    for _ in range(100):  # safety limit
        num = ''.join(str(random.randint(0, 9)) for _ in range(4))
        existing = db.execute("SELECT exam_id FROM exams WHERE exam_number = ?", (num,)).fetchone()
        if not existing:
            return num
    raise RuntimeError('Failed to generate unique exam number')


def _calc_subdir_size(subdir: str) -> tuple:
    """Calculate size breakdown of an exam subdirectory.
    Returns (total_mb, attachments_mb, student_papers_mb).
    """
    total = 0
    attach = 0
    student = 0
    if not os.path.exists(subdir):
        return (0, 0, 0)

    attach_dir = os.path.join(subdir, 'attachments')
    student_dir = os.path.join(subdir, 'student_papers')

    for entry in os.scandir(subdir):
        if entry.is_file():
            sz = entry.stat().st_size / (1024 * 1024)
            total += sz
            # Files directly in subdir (not attachments/student) count as attachments
            attach += sz
        elif entry.is_dir():
            dir_sz = sum(
                f.stat().st_size for f in os.scandir(entry.path) if f.is_file()
            ) / (1024 * 1024)
            total += dir_sz
            if entry.name == 'attachments':
                attach += dir_sz
            elif entry.name == 'student_papers':
                student += dir_sz

    return (round(total, 2), round(attach, 2), round(student, 2))


def _check_storage_limit(subdir: str) -> tuple:
    """Check if storage limits are exceeded. Returns (ok: bool, message: str)."""
    total, attach, student = _calc_subdir_size(subdir)
    if total > MAX_SUBDIR_SIZE_MB:
        return (False, f'总存储超过限制 ({total}MB / {MAX_SUBDIR_SIZE_MB}MB)')
    if attach > MAX_ATTACH_SIZE_MB:
        return (False, f'附件存储超过限制 ({attach}MB / {MAX_ATTACH_SIZE_MB}MB)')
    if student > MAX_STUDENT_PAPERS_MB:
        return (False, f'学生试卷存储超过限制 ({student}MB / {MAX_STUDENT_PAPERS_MB}MB)')
    return (True, 'ok')


def _exam_requires_stata(exam_id: int) -> bool:
    """Check if any question in the exam has enable_stata turned on."""
    db = get_db()
    row = db.execute(
        "SELECT COUNT(*) as cnt FROM questions WHERE exam_id = ? AND enable_stata = 1",
        (exam_id,)
    ).fetchone()
    return row['cnt'] > 0 if row else False

# ════════════════════════════════════════════
# Teacher APIs
# ════════════════════════════════════════════

@app.route('/api/teacher/login', methods=['POST'])
def teacher_login():
    data = request.get_json()
    tid = data.get('id', '').strip()
    pwd = data.get('password', '').strip()
    db = get_db()

    # v3.1: Try username login first (registered teachers)
    row = db.execute("SELECT id, password, username FROM teachers WHERE username = ?", (tid,)).fetchone()
    if row and row['password'].startswith('sha256:'):
        # Registered user with hashed password
        parts = row['password'].split(':')
        salt = parts[1]
        stored_hash = parts[2]
        if _verify_password(pwd, stored_hash, salt):
            session['teacher_id'] = row['id']
            session['teacher_username'] = row['username'] or tid
            return jsonify({'ok': True, 'teacher_id': row['id'], 'username': row['username'] or tid})
        return jsonify({'error': '账号或密码错误'}), 401

    # Fallback: hardcoded teacher login (id + plain password)
    row = db.execute("SELECT id FROM teachers WHERE id = ? AND password = ?", (tid, pwd)).fetchone()
    if row:
        session['teacher_id'] = tid
        session.pop('teacher_username', None)
        return jsonify({'ok': True, 'teacher_id': tid})
    return jsonify({'error': '账号或密码错误'}), 401


@app.route('/api/teacher/register', methods=['POST'])
def teacher_register():
    """v3.1: Register new teacher account."""
    data = request.get_json()
    username = data.get('username', '').strip()
    password = data.get('password', '').strip()
    school = data.get('school', '').strip()

    if not username or not password or not school:
        return jsonify({'error': '用户名、密码、学校均不能为空'}), 400

    # Validate username: any printable characters except whitespace
    if not re.match(r'^[\S]+$', username):
        return jsonify({'error': '用户名不能为空或包含空白字符'}), 400

    if len(username) > 30:
        return jsonify({'error': '用户名不能超过30个字符'}), 400

    db = get_db()

    # Check duplicate username
    existing = db.execute("SELECT id FROM teachers WHERE username = ?", (username,)).fetchone()
    if existing:
        return jsonify({'error': '用户名已存在，请选择其他用户名'}), 409

    # Check duplicate id (in case username conflicts with a teacher id)
    existing_id = db.execute("SELECT id FROM teachers WHERE id = ?", (username,)).fetchone()
    if existing_id:
        return jsonify({'error': '该名称已被使用，请选择其他用户名'}), 409

    # Generate unique teacher id
    import random
    teacher_id = 'T' + ''.join([str(random.randint(0, 9)) for _ in range(8)])
    while db.execute("SELECT id FROM teachers WHERE id = ?", (teacher_id,)).fetchone():
        teacher_id = 'T' + ''.join([str(random.randint(0, 9)) for _ in range(8)])

    # Hash password
    hashed, salt = _hash_password(password)
    stored_pw = f"sha256:{salt}:{hashed}"

    # Create user directory
    udir = _user_dir(username)
    os.makedirs(udir, exist_ok=True)

    # Insert into database
    db.execute(
        "INSERT INTO teachers (id, password, username, school) VALUES (?, ?, ?, ?)",
        (teacher_id, stored_pw, username, school)
    )
    db.commit()

    return jsonify({
        'ok': True,
        'teacher_id': teacher_id,
        'username': username,
        'message': f'注册成功！请记住您的教师ID: {teacher_id}，下次登录可使用用户名: {username}'
    })

@app.route('/api/teacher/logout', methods=['POST'])
def teacher_logout():
    session.pop('teacher_id', None)
    session.pop('teacher_username', None)
    return jsonify({'ok': True})

@app.route('/api/teacher/check', methods=['GET'])
def teacher_check():
    logged_in = 'teacher_id' in session
    if logged_in:
        return jsonify({
            'logged_in': True,
            'teacher_id': session.get('teacher_id'),
            'username': session.get('teacher_username'),
            'is_registered': session.get('teacher_username') is not None
        })
    return jsonify({'logged_in': False})

# ─── Exam Template ───

@app.route('/api/template/exam/download', methods=['GET'])
@teacher_required
def download_exam_template():
    wb = Workbook()
    ws = wb.active
    ws.title = "题目模板"
    # Template columns matching upload parser:
    # [题型, 序号, 题干, 参考答案, 分值, 选项A, 选项B, 选项C, 选项D]
    headers = ['题型', '序号', '题干', '参考答案', '分值', '选项A', '选项B', '选项C', '选项D']
    ws.append(headers)
    header_font = Font(bold=True)
    for cell in ws[1]:
        cell.font = header_font

    # Example rows
    examples = [
        ('单选题', 1, '以下哪个是Python的特点？', 'B', 5, '编译型', '解释型', '汇编', '机器码'),
        ('多选题', 2, '以下哪些是Web前端技术？', 'ABD', 5, 'HTML', 'CSS', 'Python', 'JavaScript'),
        ('判断题', 3, 'Python是一种面向对象的语言。', '正确', 5),
        ('简答题', 4, '简述Python中的装饰器是什么。', '', 10),
    ]
    for row in examples:
        ws.append(row)

    path = os.path.join(os.path.dirname(__file__), 'exam_template.xlsx')
    wb.save(path)
    return send_file(path, as_attachment=True, download_name='exam_template.xlsx')

@app.route('/api/template/exam/upload', methods=['POST'])
@teacher_required
def upload_exam_template():
    if 'file' not in request.files:
        return jsonify({'error': '未找到文件'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': '文件名为空'}), 400

    wb = load_workbook(file)
    ws = wb.active
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    if not rows:
        return jsonify({'error': '模板为空'}), 400

    db = get_db()
    # Extract title: use first question's content (first 40 chars) or first cell
    title = ''
    for row in rows:
        if len(row) > 2 and row[2]:
            t = str(row[2]).strip()
            if t:
                title = t[:40]
                break
    if not title:
        for row in rows:
            if len(row) > 0 and row[0]:
                t = str(row[0]).strip()
                if t:
                    title = t[:40]
                    break
    if not title:
        title = f"试卷_{datetime.now().strftime('%Y%m%d_%H%M%S')}"

    # v3.1: Check exam count limit for registered teachers
    username = session.get('teacher_username')
    if username:
        existing = db.execute("SELECT COUNT(*) as cnt FROM exams WHERE teacher_username = ?", (username,)).fetchone()
        if existing['cnt'] >= MAX_EXAMS_PER_TEACHER:
            return jsonify({
                'error': f'每个用户最多上传 {MAX_EXAMS_PER_TEACHER} 份试卷。请先删除旧试卷再上传新试卷。'
            }), 403

    created_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    # v4.0: Generate unique exam number
    exam_number = _generate_exam_number()

    cur = db.execute(
        "INSERT INTO exams (title, duration_minutes, created_at, template_uploaded, teacher_username, exam_number) VALUES (?, 60, ?, 1, ?, ?)",
        (title, created_at, username, exam_number)
    )
    exam_id = cur.lastrowid

    sort_order = 1
    for row in rows:
        # Template columns: [题型, 序号, 题干, 参考答案, 分值, 选项A, 选项B, 选项C, 选项D]
        q_type = str(row[0]).strip() if row[0] else ''
        content = str(row[2]).strip() if len(row) > 2 and row[2] else ''
        if not q_type or not content:
            continue

        db.execute(
            "INSERT INTO questions (exam_id, q_type, content, option_a, option_b, option_c, option_d, correct_answer, score, sort_order, enable_stata, enable_ai, has_attachment) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",
            (exam_id, q_type, content,
             str(row[5]).strip() if len(row) > 5 and row[5] else None,
             str(row[6]).strip() if len(row) > 6 and row[6] else None,
             str(row[7]).strip() if len(row) > 7 and row[7] else None,
             str(row[8]).strip() if len(row) > 8 and row[8] else None,
             str(row[3]).strip() if len(row) > 3 and row[3] else None,
             float(row[4]) if len(row) > 4 and row[4] else 0,
             sort_order, 0, 0, 0)
        )
        sort_order += 1

    # v3.1: Create exam subdirectory for this teacher
    if username:
        edir = _exam_subdir(username, title)
        os.makedirs(edir, exist_ok=True)

    db.commit()
    return jsonify({
        'ok': True, 'exam_id': exam_id, 'title': title,
        'question_count': sort_order - 1,
        'exam_number': exam_number
    })

# ─── Student Template ───

@app.route('/api/template/student/download', methods=['GET'])
@teacher_required
def download_student_template():
    wb = Workbook()
    ws = wb.active
    ws.title = "学生信息"
    ws['A1'] = '学号'
    ws['B1'] = '姓名'
    ws['C1'] = '班级'
    ws['D1'] = '年级'
    ws['E1'] = '学校'
    header_font = Font(bold=True)
    for cell in ws[1]:
        cell.font = header_font

    ws.cell(row=2, column=1, value='2024001')
    ws.cell(row=2, column=2, value='张三')
    ws.cell(row=2, column=3, value='1班')
    ws.cell(row=2, column=4, value='2024级')
    ws.cell(row=2, column=5, value='示例学校')

    path = os.path.join(os.path.dirname(__file__), 'student_template.xlsx')
    wb.save(path)
    return send_file(path, as_attachment=True, download_name='student_template.xlsx')

@app.route('/api/template/student/upload', methods=['POST'])
@teacher_required
def upload_student_template():
    if 'file' not in request.files:
        return jsonify({'error': '未找到文件'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': '文件名为空'}), 400

    wb = load_workbook(file)
    ws = wb.active
    rows = list(ws.iter_rows(min_row=2, values_only=True))

    db = get_db()
    count = 0
    for row in rows:
        student_id = str(row[0]).strip() if row[0] else ''
        name = str(row[1]).strip() if row[1] else ''
        if not student_id or not name:
            continue
        db.execute(
            "INSERT OR REPLACE INTO students (student_id, name, class_name, grade, school) VALUES (?,?,?,?,?)",
            (student_id, name,
             str(row[2]).strip() if row[2] else None,
             str(row[3]).strip() if row[3] else None,
             str(row[4]).strip() if row[4] else None)
        )
        count += 1
    db.commit()
    return jsonify({'ok': True, 'count': count})

# ─── Exam Management ───

@app.route('/api/teacher/exams', methods=['GET'])
@teacher_required
def list_exams():
    db = get_db()
    username = session.get('teacher_username')

    # v4.0: Include exam_number in returned data
    cols = "exam_id, title, duration_minutes, created_at, template_uploaded, " \
            "exam_duration, exam_start_time, exam_notice, teacher_username, exam_number"

    # v3.1: For registered teachers, show only their own exams
    if username:
        rows = db.execute(
            f"SELECT {cols} FROM exams WHERE teacher_username = ? ORDER BY created_at DESC",
            (username,)
        ).fetchall()
    else:
        rows = db.execute(
            f"SELECT {cols} FROM exams WHERE teacher_username IS NULL ORDER BY created_at DESC"
        ).fetchall()

    exams = []
    for r in rows:
        d = dict(r)
        d['question_count'] = db.execute(
            "SELECT COUNT(*) FROM questions WHERE exam_id=?", (r['exam_id'],)
        ).fetchone()[0]
        d['settings_complete'] = bool(d['exam_duration'] and d['exam_start_time'])
        exams.append(d)
    return jsonify(exams)

@app.route('/api/teacher/results', methods=['GET'])
@teacher_required
def list_results():
    db = get_db()
    username = session.get('teacher_username')

    # v3.1: Filter results by teacher ownership
    if username:
        rows = db.execute("""
            SELECT s.session_id, s.exam_id, e.title, s.student_id, s.student_name,
                   s.start_time, s.end_time, s.status, s.total_score, s.is_graded
            FROM exam_sessions s
            JOIN exams e ON s.exam_id = e.exam_id
            WHERE e.teacher_username = ?
            ORDER BY s.start_time DESC
        """, (username,)).fetchall()
    else:
        rows = db.execute("""
            SELECT s.session_id, s.exam_id, e.title, s.student_id, s.student_name,
                   s.start_time, s.end_time, s.status, s.total_score, s.is_graded
            FROM exam_sessions s
            JOIN exams e ON s.exam_id = e.exam_id
            WHERE e.teacher_username IS NULL
            ORDER BY s.start_time DESC
        """).fetchall()
    return jsonify([dict(r) for r in rows])

@app.route('/api/teacher/session/<int:session_id>', methods=['GET'])
@teacher_required
def get_session_detail(session_id):
    db = get_db()
    srow = db.execute("SELECT * FROM exam_sessions WHERE session_id=?", (session_id,)).fetchone()
    if not srow:
        return jsonify({'error': '未找到记录'}), 404

    rows = db.execute("""
        SELECT q.q_id, q.q_type, q.content, q.correct_answer, q.score AS max_score,
               a.answer_text, a.is_correct, a.score AS actual_score
        FROM questions q
        LEFT JOIN answers a ON q.q_id = a.q_id AND a.session_id = ?
        WHERE q.exam_id = ?
        ORDER BY q.sort_order
    """, (session_id, srow['exam_id'])).fetchall()

    questions = []
    for r in rows:
        d = dict(r)
        # Don't expose correct_answer in normal flow, but teacher needs it for grading
        questions.append(d)

    return jsonify({**dict(srow), 'questions': questions})

@app.route('/api/teacher/session/<int:session_id>/grade', methods=['POST'])
@teacher_required
def grade_session(session_id):
    data = request.get_json()
    grades = data.get('grades', {})  # {q_id: score}

    db = get_db()
    for q_id, score in grades.items():
        db.execute(
            "UPDATE answers SET score = ? WHERE session_id = ? AND q_id = ?",
            (float(score), session_id, int(q_id))
        )

    # Recalculate total
    total = db.execute(
        "SELECT COALESCE(SUM(score), 0) FROM answers WHERE session_id = ?", (session_id,)
    ).fetchone()[0]
    db.execute("UPDATE exam_sessions SET total_score = ?, is_graded = 1 WHERE session_id = ?",
               (total, session_id))
    db.commit()
    return jsonify({'ok': True, 'total_score': total})


# ════════════════════════════════════════════
# v3.1: Exam Deletion + Teacher Profile
# ════════════════════════════════════════════

@app.route('/api/teacher/profile', methods=['GET'])
@teacher_required
def teacher_profile():
    """v3.1: Get teacher profile info (username, school, exam count)."""
    tid = session.get('teacher_id')
    username = session.get('teacher_username')
    db = get_db()

    if username:
        # Registered teacher
        row = db.execute("SELECT username, school FROM teachers WHERE username = ?", (username,)).fetchone()
        exam_count = db.execute("SELECT COUNT(*) as cnt FROM exams WHERE teacher_username = ?", (username,)).fetchone()['cnt']

        # v3.1: Cleanup old exam subfolders on access
        _cleanup_old_exams(username)

        # Get exam list for delete dropdown
        exams = db.execute("SELECT exam_id, title, created_at FROM exams WHERE teacher_username = ? ORDER BY created_at DESC", (username,)).fetchall()
        exam_list = [{'exam_id': e['exam_id'], 'title': e['title'], 'created_at': e['created_at']} for e in exams]

        return jsonify({
            'is_registered': True,
            'username': username,
            'school': row['school'] if row else '',
            'exam_count': exam_count,
            'max_exams': MAX_EXAMS_PER_TEACHER,
            'exams': exam_list
        })
    else:
        # Hardcoded teacher (001)
        exam_count = db.execute("SELECT COUNT(*) as cnt FROM exams WHERE teacher_username IS NULL").fetchone()['cnt']
        return jsonify({
            'is_registered': False,
            'username': tid,
            'school': '',
            'exam_count': exam_count,
            'max_exams': MAX_EXAMS_PER_TEACHER,
            'exams': []
        })

@app.route('/api/teacher/exam/<int:exam_id>', methods=['DELETE'])
@teacher_required
def delete_exam(exam_id):
    """v3.1: Delete an exam and its associated user folder."""
    db = get_db()

    # Get exam info
    exam = db.execute("SELECT title, teacher_username FROM exams WHERE exam_id = ?", (exam_id,)).fetchone()
    if not exam:
        return jsonify({'error': '试卷不存在'}), 404

    username = exam['teacher_username']

    # Delete associated user folder
    if username and exam['title']:
        _delete_exam_subdir(username, exam['title'])

    # Delete exam (cascade will handle questions, answers, sessions)
    db.execute("DELETE FROM exam_attachments WHERE q_id IN (SELECT q_id FROM questions WHERE exam_id = ?)", (exam_id,))
    db.execute("DELETE FROM answers WHERE q_id IN (SELECT q_id FROM questions WHERE exam_id = ?)", (exam_id,))
    db.execute("DELETE FROM questions WHERE exam_id = ?", (exam_id,))
    db.execute("DELETE FROM exam_sessions WHERE exam_id = ?", (exam_id,))
    db.execute("DELETE FROM exams WHERE exam_id = ?", (exam_id,))
    db.commit()

    return jsonify({'ok': True, 'message': f'试卷 "{exam["title"]}" 已删除'})

@app.route('/api/teacher/exam/<int:exam_id>/upload-file', methods=['POST'])
@teacher_required
def upload_exam_file(exam_id):
    """v3.1: Upload a file to teacher's exam subfolder (500MB limit)."""
    if 'file' not in request.files:
        return jsonify({'error': '未找到文件'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': '文件名为空'}), 400

    username = session.get('teacher_username')
    if not username:
        return jsonify({'error': '仅注册用户可使用此功能'}), 403

    # Get exam title for folder lookup
    db = get_db()
    exam = db.execute("SELECT title, teacher_username FROM exams WHERE exam_id = ?", (exam_id,)).fetchone()
    if not exam:
        return jsonify({'error': '试卷不存在'}), 404
    if str(session.get('teacher_username', '')).strip() != str(exam['teacher_username']).strip():
        return jsonify({'error': '无权访问此考试'}), 403

    # Check file size (500MB limit)
    file.seek(0, 2)  # seek to end
    size_mb = file.tell() / (1024 * 1024)
    file.seek(0)  # seek back to start

    if size_mb > MAX_FILE_SIZE_MB:
        return jsonify({'error': f'文件大小 ({size_mb:.1f}MB) 超过限制 ({MAX_FILE_SIZE_MB}MB)'}), 400

    # Ensure exam subdirectory exists
    edir = _exam_subdir(username, exam['title'])
    os.makedirs(edir, exist_ok=True)

    # v4.0: Check storage limit before saving
    ok, msg = _check_storage_limit(edir)
    if not ok:
        return jsonify({'error': msg}), 413

    # Save file
    filepath = os.path.join(edir, file.filename)
    file.save(filepath)

    return jsonify({'ok': True, 'filename': file.filename, 'size_mb': round(size_mb, 2)})

@app.route('/api/teacher/exam/<int:exam_id>/files', methods=['GET'])
@teacher_required
def list_exam_files(exam_id):
    """v3.1: List files in teacher's exam subfolder."""
    username = session.get('teacher_username')
    if not username:
        return jsonify({'error': '仅注册用户可使用此功能'}), 403

    db = get_db()
    exam = db.execute("SELECT title, teacher_username FROM exams WHERE exam_id = ?", (exam_id,)).fetchone()
    if not exam:
        return jsonify({'error': '试卷不存在'}), 404
    if str(session.get('teacher_username', '')).strip() != str(exam['teacher_username']).strip():
        return jsonify({'error': '无权访问此考试'}), 403

    edir = _exam_subdir(username, exam['title'])
    if not os.path.exists(edir):
        return jsonify([])

    files = []
    for entry in os.scandir(edir):
        if entry.is_file():
            files.append({
                'name': entry.name,
                'size_mb': round(entry.stat().st_size / (1024 * 1024), 2),
                'modified': datetime.fromtimestamp(entry.stat().st_mtime).strftime('%Y-%m-%d %H:%M:%S')
            })
    return jsonify(files)

# ════════════════════════════════════════════
# Student APIs
# ════════════════════════════════════════════

@app.route('/api/student/login', methods=['POST'])
def student_login():
    """
    v4.0: Student login with exam_number + student_id + name.
    Returns exam info if valid.
    """
    data = request.get_json()
    sid = data.get('student_id', '').strip()
    name = data.get('name', '').strip()
    exam_number = data.get('exam_number', '').strip()

    if not exam_number:
        return jsonify({'error': '请输入试卷编号'}), 400

    db = get_db()

    # Look up exam by exam_number
    exam = db.execute(
        "SELECT exam_id, title, template_uploaded, exam_duration, exam_start_time "
        "FROM exams WHERE exam_number = ?",
        (exam_number,)
    ).fetchone()
    if not exam:
        return jsonify({'error': '试卷编号不存在，请确认后重试'}), 404
    if not exam['template_uploaded']:
        return jsonify({'error': '该试卷尚未完成设置'}), 400

    # Check if exam requires Stata
    requires_stata = _exam_requires_stata(exam['exam_id'])

    # Verify student exists
    row = db.execute(
        "SELECT student_id, name FROM students WHERE student_id = ? AND name = ?",
        (sid, name)
    ).fetchone()
    if not row:
        return jsonify({'error': '学号或姓名不匹配'}), 401

    return jsonify({
        'ok': True,
        'student_id': sid,
        'name': name,
        'exam_id': exam['exam_id'],
        'exam_title': exam['title'],
        'exam_number': exam_number,
        'requires_stata': requires_stata
    })

@app.route('/api/student/exams', methods=['GET'])
def list_available_exams():
    student_id = request.args.get('student_id', '')
    if not student_id:
        return jsonify({'error': '缺少学号'}), 400
    db = get_db()
    # v3.0: Only show exams that have template uploaded AND settings configured
    rows = db.execute(
        "SELECT exam_id, title, exam_duration, exam_start_time, exam_notice, "
        "exam_late_start_limit, exam_late_submit_limit, "
        "anti_shuffle_questions, anti_shuffle_options, anti_screen_switch "
        "FROM exams WHERE template_uploaded = 1 AND exam_duration > 0 AND exam_start_time != '' "
        "ORDER BY created_at DESC"
    ).fetchall()
    exams = []
    for r in rows:
        eid = r['exam_id']
        existing = db.execute(
            "SELECT session_id, status FROM exam_sessions WHERE exam_id = ? AND student_id = ?",
            (eid, student_id)
        ).fetchone()
        exam = dict(r)
        if existing:
            exam['has_session'] = True
            exam['session_id'] = existing['session_id']
            exam['status'] = existing['status']
        else:
            exam['has_session'] = False
        exams.append(exam)
    return jsonify(exams)

@app.route('/api/exam/<int:exam_id>/start', methods=['POST'])
def start_exam(exam_id):
    data = request.get_json()
    student_id = data.get('student_id', '').strip()
    student_name = data.get('student_name', '').strip()

    if not student_id or not student_name:
        return jsonify({'error': '缺少学生信息'}), 400

    db = get_db()

    # v3.0: Check exam settings are configured
    exam_info = db.execute(
        "SELECT exam_id, title, exam_duration, exam_start_time, exam_notice, "
        "exam_late_start_limit, exam_late_submit_limit, "
        "anti_shuffle_questions, anti_shuffle_options, anti_screen_switch "
        "FROM exams WHERE exam_id = ?",
        (exam_id,)
    ).fetchone()
    if not exam_info:
        return jsonify({'error': '考试不存在'}), 404

    # v3.0: Verify teacher has configured exam settings
    if not exam_info['exam_duration'] or not exam_info['exam_start_time']:
        return jsonify({'error': '教师尚未完成考试设置，无法开始考试'}), 400

    # Check for existing in-progress session
    existing = db.execute(
        "SELECT session_id, terminated_by_teacher FROM exam_sessions WHERE exam_id = ? AND student_id = ? AND status = 'in_progress'",
        (exam_id, student_id)
    ).fetchone()

    if existing:
        # v3.0: Check if teacher terminated this session
        if existing['terminated_by_teacher']:
            return jsonify({'error': '你的考试已被教师终止，请联系教师'}), 403
        session_id = existing['session_id']
    else:
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        cur = db.execute(
            "INSERT INTO exam_sessions (exam_id, student_id, student_name, start_time, status) VALUES (?,?,?,?,?)",
            (exam_id, student_id, student_name, now, 'in_progress')
        )
        session_id = cur.lastrowid
        db.commit()

    # Return questions WITH v3.0 settings
    rows = db.execute(
        "SELECT q_id, q_type, content, option_a, option_b, option_c, option_d, score, sort_order, "
        "enable_stata, enable_ai, has_attachment "
        "FROM questions WHERE exam_id = ? ORDER BY sort_order",
        (exam_id,)
    ).fetchall()

    # Get saved answers if resuming
    saved = db.execute(
        "SELECT q_id, answer_text FROM answers WHERE session_id = ?", (session_id,)
    ).fetchall()
    saved_answers = {r['q_id']: r['answer_text'] for r in saved}

    # v3.0: Get attachments for this exam
    attachments = db.execute(
        "SELECT a.attachment_id, a.q_id, a.original_name "
        "FROM exam_attachments a "
        "JOIN questions q ON a.q_id = q.q_id "
        "WHERE q.exam_id = ?",
        (exam_id,)
    ).fetchall()

    return jsonify({
        'ok': True,
        'session_id': session_id,
        'title': exam_info['title'],
        'exam_duration': exam_info['exam_duration'],
        'exam_start_time': exam_info['exam_start_time'],
        'exam_notice': exam_info['exam_notice'],
        'exam_late_start_limit': exam_info['exam_late_start_limit'],
        'exam_late_submit_limit': exam_info['exam_late_submit_limit'],
        'anti_shuffle_questions': bool(exam_info['anti_shuffle_questions']),
        'anti_shuffle_options': bool(exam_info['anti_shuffle_options']),
        'anti_screen_switch': bool(exam_info['anti_screen_switch']),
        'questions': [dict(r) for r in rows],
        'saved_answers': saved_answers,
        'attachments': [dict(a) for a in attachments]
    })

@app.route('/api/exam/session/<int:session_id>/save', methods=['POST'])
def save_answers(session_id):
    data = request.get_json()
    answers = data.get('answers', {})  # {q_id: answer_text}

    db = get_db()
    # Verify session belongs to student
    srow = db.execute("SELECT status, terminated_by_teacher FROM exam_sessions WHERE session_id = ?", (session_id,)).fetchone()
    if not srow:
        return jsonify({'error': '会话不存在'}), 404
    if srow['status'] == 'submitted':
        return jsonify({'error': '考试已提交'}), 400
    # v3.0: Check if terminated by teacher
    if srow['terminated_by_teacher']:
        return jsonify({'error': '考试已被教师终止', 'terminated': True}), 403

    # Support both formats:
    #   dict:  {q_id: answer_text, ...}
    #   list:  [{q_id: int, answer: str}, ...]
    if isinstance(answers, dict):
        ans_items = [(int(k), v) for k, v in answers.items()]
    elif isinstance(answers, list):
        ans_items = [(a.get('q_id'), a.get('answer', '')) for a in answers if 'q_id' in a]
    else:
        ans_items = []

    for q_id, ans in ans_items:
        existing = db.execute(
            "SELECT answer_id FROM answers WHERE session_id = ? AND q_id = ?", (session_id, q_id)
        ).fetchone()
        if existing:
            db.execute("UPDATE answers SET answer_text = ? WHERE session_id = ? AND q_id = ?",
                       (ans, session_id, q_id))
        else:
            db.execute("INSERT INTO answers (session_id, q_id, answer_text) VALUES (?,?,?)",
                       (session_id, q_id, ans))
    db.commit()
    return jsonify({'ok': True})

@app.route('/api/exam/session/<int:session_id>/submit', methods=['POST'])
def submit_exam(session_id):
    data = request.get_json()
    answers = data.get('answers', {})

    db = get_db()
    srow = db.execute(
        "SELECT exam_id, student_id, status FROM exam_sessions WHERE session_id = ?", (session_id,)
    ).fetchone()
    if not srow:
        return jsonify({'error': '会话不存在'}), 404
    if srow['status'] == 'submitted':
        return jsonify({'error': '考试已提交'}), 400

    # Save answers first
    # Support both formats:
    if isinstance(answers, dict):
        ans_items = [(int(k), v) for k, v in answers.items()]
    elif isinstance(answers, list):
        ans_items = [(a.get('q_id'), a.get('answer', '')) for a in answers if 'q_id' in a]
    else:
        ans_items = []

    for q_id, ans in ans_items:
        existing = db.execute(
            "SELECT answer_id FROM answers WHERE session_id = ? AND q_id = ?", (session_id, q_id)
        ).fetchone()
        if existing:
            db.execute("UPDATE answers SET answer_text = ? WHERE session_id = ? AND q_id = ?",
                       (ans, session_id, q_id))
        else:
            db.execute("INSERT INTO answers (session_id, q_id, answer_text) VALUES (?,?,?)",
                       (session_id, q_id, ans))

    # Auto-grade objective questions
    questions = db.execute(
        "SELECT q_id, q_type, correct_answer, score FROM questions WHERE exam_id = ?", (srow['exam_id'],)
    ).fetchall()

    # Build answers dict for grading
    answers_dict = {}
    if isinstance(answers, dict):
        answers_dict = answers
    elif isinstance(answers, list):
        answers_dict = {str(a.get('q_id')): a.get('answer', '') for a in answers if 'q_id' in a}

    total_score = 0
    for q in questions:
        ans = answers_dict.get(str(q['q_id']), '')
        if ans is None:
            ans = ''
        ans = str(ans).strip()

        if q['q_type'] in ('单选题', '选择题', '判断题'):
            is_correct = (ans.upper() == str(q['correct_answer']).strip().upper()) if q['correct_answer'] else False
            score = q['score'] if is_correct else 0
        elif q['q_type'] == '多选题':
            correct = set(str(q['correct_answer']).strip().upper()) if q['correct_answer'] else set()
            given = set(ans.upper())
            is_correct = (given == correct)
            score = q['score'] if is_correct else 0
        else:
            # Subjective - needs manual grading
            is_correct = None
            score = 0

        # Update answer record
        existing = db.execute(
            "SELECT answer_id FROM answers WHERE session_id = ? AND q_id = ?", (session_id, q['q_id'])
        ).fetchone()
        if existing:
            db.execute(
                "UPDATE answers SET is_correct = ?, score = ? WHERE session_id = ? AND q_id = ?",
                (1 if is_correct else (0 if is_correct is False else None), score, session_id, q['q_id'])
            )
        else:
            db.execute(
                "INSERT INTO answers (session_id, q_id, answer_text, is_correct, score) VALUES (?,?,?,?,?)",
                (session_id, q['q_id'], ans, 1 if is_correct else (0 if is_correct is False else None), score)
            )
        total_score += score

    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    db.execute("UPDATE exam_sessions SET end_time = ?, status = 'submitted', total_score = ? WHERE session_id = ?",
               (now, total_score, session_id))
    db.commit()

    # v4.0: Save student answers to xlsx in User/<username>/<exam_title>/student_papers/
    try:
        exam_info = db.execute(
            "SELECT title, teacher_username FROM exams WHERE exam_id = ?",
            (srow['exam_id'],)
        ).fetchone()
        if exam_info and exam_info['teacher_username']:
            username = exam_info['teacher_username']
            edir = _exam_subdir(username, exam_info['title'])
            student_dir = os.path.join(edir, 'student_papers')
            os.makedirs(student_dir, exist_ok=True)

            # Build the xlsx
            wb = Workbook()
            ws = wb.active
            ws.title = "答题记录"
            # Headers matching exam template format
            headers = ['题型', '序号', '题干', '参考答案', '学生答案', '是否正确', '得分', '分值', '选项A', '选项B', '选项C', '选项D']
            ws.append(headers)
            for cell in ws[1]:
                cell.font = Font(bold=True)

            # Get all questions for this exam
            all_q = db.execute(
                "SELECT q_id, q_type, content, option_a, option_b, option_c, option_d, "
                "correct_answer, score, sort_order FROM questions WHERE exam_id = ? ORDER BY sort_order",
                (srow['exam_id'],)
            ).fetchall()

            # Get answers
            all_a = db.execute(
                "SELECT q_id, answer_text, is_correct, score FROM answers WHERE session_id = ?",
                (session_id,)
            ).fetchall()
            ans_map = {a['q_id']: a for a in all_a}

            for seq, q in enumerate(all_q, 1):
                a = ans_map.get(q['q_id'])
                answer_text = a['answer_text'] if a else ''
                is_correct = a['is_correct'] if a else None
                score = a['score'] if a else 0

                correct_str = ''
                if is_correct == 1:
                    correct_str = '正确'
                elif is_correct == 0:
                    correct_str = '错误'

                ws.append([
                    q['q_type'], seq, q['content'],
                    q['correct_answer'] or '', answer_text,
                    correct_str, score, q['score'],
                    q['option_a'] or '', q['option_b'] or '',
                    q['option_c'] or '', q['option_d'] or ''
                ])

            # Total row
            ws.append([])
            ws.append(['总分', '', '', '', '', '', total_score, ''])

            safename = re.sub(r'[<>:"/\\|?*]', '_', str(srow['student_id']))[:50]
            xlsx_path = os.path.join(student_dir, f'{safename}.xlsx')
            wb.save(xlsx_path)
    except Exception as e:
        # Don't fail the request if xlsx save fails, but log it
        print(f"[v4.0] Failed to save student xlsx: {e}")

    return jsonify({'ok': True, 'total_score': total_score, 'is_graded': True})

# ════════════════════════════════════════════
# Monitor APIs (v2.0)
# ════════════════════════════════════════════

@app.route('/api/teacher/monitor', methods=['GET'])
@teacher_required
def monitor_sessions():
    """获取所有在线考试的学生状态"""
    db = get_db()
    rows = db.execute("""
        SELECT s.session_id, s.exam_id, s.student_id, s.student_name,
               s.start_time, s.status, s.total_score,
               s.terminated_by_teacher, s.exam_end_time,
               e.title, e.duration_minutes
        FROM exam_sessions s
        JOIN exams e ON s.exam_id = e.exam_id
        WHERE s.status = 'in_progress'
        ORDER BY s.start_time ASC
    """).fetchall()
    results = []
    for r in rows:
        d = dict(r)
        d['is_terminated'] = bool(d['terminated_by_teacher'])
        results.append(d)
    return jsonify(results)

@app.route('/api/teacher/session/<int:session_id>/terminate', methods=['POST'])
@teacher_required
def terminate_session(session_id):
    """教师终止学生考试"""
    db = get_db()
    srow = db.execute("SELECT status FROM exam_sessions WHERE session_id = ?", (session_id,)).fetchone()
    if not srow:
        return jsonify({'error': '会话不存在'}), 404
    if srow['status'] == 'submitted':
        return jsonify({'error': '考试已提交'}), 400
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    db.execute("UPDATE exam_sessions SET terminated_by_teacher = 1, exam_end_time = ? WHERE session_id = ?",
               (now, session_id))
    db.commit()
    return jsonify({'ok': True, 'message': '已终止该学生的考试'})

@app.route('/api/teacher/session/<int:session_id>/resume', methods=['POST'])
@teacher_required
def resume_session(session_id):
    """教师恢复学生考试"""
    db = get_db()
    srow = db.execute("SELECT * FROM exam_sessions WHERE session_id = ?", (session_id,)).fetchone()
    if not srow:
        return jsonify({'error': '会话不存在'}), 404
    if srow['status'] == 'submitted':
        return jsonify({'error': '考试已提交'}), 400
    if srow['exam_end_time']:
        try:
            end_dt = datetime.strptime(srow['exam_end_time'], '%Y-%m-%d %H:%M:%S')
            if datetime.now() > end_dt:
                return jsonify({'error': '已超过考试终止时间，无法恢复'})
        except Exception:
            pass
    db.execute("UPDATE exam_sessions SET terminated_by_teacher = 0, exam_end_time = NULL WHERE session_id = ?",
               (session_id,))
    db.commit()
    return jsonify({'ok': True, 'message': '已恢复该学生的考试'})

# ════════════════════════════════════════════
# Exam Settings APIs (v3.0)
# ════════════════════════════════════════════

@app.route('/api/exam/<int:exam_id>/settings', methods=['GET'])
@teacher_required
def get_exam_settings(exam_id):
    """获取考试设置"""
    db = get_db()
    row = db.execute("""
        SELECT exam_id, title, exam_duration, exam_start_time, exam_notice,
               exam_late_start_limit, exam_late_submit_limit,
               anti_shuffle_questions, anti_shuffle_options, anti_screen_switch,
               template_uploaded
        FROM exams WHERE exam_id = ?
    """, (exam_id,)).fetchone()
    if not row:
        return jsonify({'error': '考试不存在'}), 404

    d = dict(row)
    # Get questions with their settings
    questions = db.execute("""
        SELECT q_id, q_type, content, score, sort_order,
               enable_stata, enable_ai, has_attachment
        FROM questions WHERE exam_id = ? ORDER BY sort_order
    """, (exam_id,)).fetchall()
    d['questions'] = [dict(q) for q in questions]
    return jsonify(d)

@app.route('/api/exam/<int:exam_id>/settings', methods=['POST'])
@teacher_required
def save_exam_settings(exam_id):
    """保存考试设置"""
    data = request.get_json()
    db = get_db()
    row = db.execute("SELECT exam_id FROM exams WHERE exam_id = ?", (exam_id,)).fetchone()
    if not row:
        return jsonify({'error': '考试不存在'}), 404

    exam_duration = data.get('exam_duration', 60)
    exam_start_time = data.get('exam_start_time', '')
    exam_notice = data.get('exam_notice', '')
    exam_late_start_limit = data.get('exam_late_start_limit', 0)
    exam_late_submit_limit = data.get('exam_late_submit_limit', 0)
    anti_shuffle_questions = data.get('anti_shuffle_questions', 0)
    anti_shuffle_options = data.get('anti_shuffle_options', 0)
    anti_screen_switch = data.get('anti_screen_switch', 0)

    db.execute("""
        UPDATE exams SET
            exam_duration = ?, exam_start_time = ?, exam_notice = ?,
            exam_late_start_limit = ?, exam_late_submit_limit = ?,
            anti_shuffle_questions = ?, anti_shuffle_options = ?, anti_screen_switch = ?
        WHERE exam_id = ?
    """, (
        int(exam_duration), exam_start_time, exam_notice,
        int(exam_late_start_limit), int(exam_late_submit_limit),
        int(anti_shuffle_questions), int(anti_shuffle_options), int(anti_screen_switch),
        exam_id
    ))
    db.commit()
    return jsonify({'ok': True, 'message': '考试设置已保存'})

@app.route('/api/exam/<int:exam_id>/questions-settings', methods=['GET'])
@teacher_required
def get_exam_questions_settings(exam_id):
    """获取考试的题目设置列表（用于考试设置页面）"""
    db = get_db()
    rows = db.execute("""
        SELECT q_id, q_type, content, score, sort_order,
               enable_stata, enable_ai, has_attachment
        FROM questions WHERE exam_id = ? ORDER BY sort_order
    """, (exam_id,)).fetchall()
    return jsonify([dict(r) for r in rows])

@app.route('/api/exam/<int:exam_id>/questions-settings', methods=['POST'])
@teacher_required
def update_question_settings(exam_id):
    """更新题目级别设置（Stata/AI开关、附件标记）"""
    data = request.get_json()
    settings = data.get('settings', [])  # [{q_id, enable_stata, enable_ai, has_attachment}]
    db = get_db()
    for s in settings:
        db.execute("""
            UPDATE questions SET enable_stata = ?, enable_ai = ?, has_attachment = ?
            WHERE q_id = ? AND exam_id = ?
        """, (
            int(s.get('enable_stata', 0)),
            int(s.get('enable_ai', 0)),
            int(s.get('has_attachment', 0)),
            s['q_id'], exam_id
        ))
    db.commit()
    return jsonify({'ok': True})

# ─── Attachment APIs (v3.0 → v3.2) ───
# v3.2: 附件存储到试卷子文件夹 {USERS_ROOT}\<username>\<exam_title>\attachments\

def _get_attachment_dir(exam_id: int) -> str:
    """获取指定试卷的附件目录（按试卷子文件夹）"""
    db = get_db()
    exam = db.execute(
        "SELECT title, teacher_username FROM exams WHERE exam_id = ?",
        (exam_id,)
    ).fetchone()
    if exam and exam['teacher_username']:
        edir = _exam_subdir(exam['teacher_username'], exam['title'])
        attach_dir = os.path.join(edir, 'attachments')
        os.makedirs(attach_dir, exist_ok=True)
        return attach_dir
    # 旧版兼容
    legacy = os.path.join(USERS_ROOT, '_legacy_attachments')
    os.makedirs(legacy, exist_ok=True)
    return legacy

@app.route('/api/attachment/upload', methods=['POST'])
@teacher_required
def upload_attachment():
    """上传题目附件（存到试卷子文件夹）"""
    q_id = request.form.get('q_id')
    if not q_id:
        return jsonify({'error': '缺少题目ID'}), 400
    if 'file' not in request.files:
        return jsonify({'error': '未找到文件'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': '文件名为空'}), 400

    db = get_db()
    # Verify question exists and get exam_id
    qrow = db.execute(
        "SELECT q_id, exam_id FROM questions WHERE q_id = ?", (q_id,)
    ).fetchone()
    if not qrow:
        return jsonify({'error': '题目不存在'}), 404

    # 获取试卷对应的附件目录
    attach_dir = _get_attachment_dir(qrow['exam_id'])

    # Save file
    now = datetime.now().strftime('%Y%m%d_%H%M%S')
    ext = os.path.splitext(file.filename)[1]
    filename = f"q{q_id}_{now}{ext}"
    filepath = os.path.join(attach_dir, filename)
    file.save(filepath)

    # Update has_attachment flag
    db.execute("UPDATE questions SET has_attachment = 1 WHERE q_id = ?", (q_id,))

    # Record in DB
    uploaded_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    cur = db.execute(
        "INSERT INTO exam_attachments (q_id, original_name, file_path, uploaded_at) VALUES (?,?,?,?)",
        (q_id, file.filename, filename, uploaded_at)
    )
    db.commit()
    return jsonify({'ok': True, 'attachment_id': cur.lastrowid, 'filename': file.filename})

@app.route('/api/attachment/download/<int:attachment_id>', methods=['GET'])
def download_attachment(attachment_id):
    """下载附件（解析到试卷子文件夹）"""
    db = get_db()
    row = db.execute(
        """SELECT a.file_path, a.original_name, q.exam_id
           FROM exam_attachments a
           JOIN questions q ON a.q_id = q.q_id
           WHERE a.attachment_id = ?""",
        (attachment_id,)
    ).fetchone()
    if not row:
        return jsonify({'error': '附件不存在'}), 404

    attach_dir = _get_attachment_dir(row['exam_id'])
    filepath = os.path.join(attach_dir, row['file_path'])
    if not os.path.exists(filepath):
        return jsonify({'error': '附件文件已丢失'}), 404

    return send_file(filepath, as_attachment=True, download_name=row['original_name'])

@app.route('/api/exam/<int:exam_id>/attachments', methods=['GET'])
def list_exam_attachments(exam_id):
    """获取考试所有附件（学生端）"""
    db = get_db()
    rows = db.execute("""
        SELECT a.attachment_id, a.q_id, a.original_name, q.has_attachment
        FROM exam_attachments a
        JOIN questions q ON a.q_id = q.q_id
        WHERE q.exam_id = ?
    """, (exam_id,)).fetchall()
    return jsonify([dict(r) for r in rows])

# ════════════════════════════════════════════
# v4.0: Exam Number Lookup & Stata Check (student side)
# ════════════════════════════════════════════

@app.route('/api/exam/by-number/<exam_number>', methods=['GET'])
def get_exam_by_number(exam_number):
    """
    v4.0: Look up exam via 4-digit number.
    Returns exam info including whether Stata is required.
    """
    db = get_db()
    exam = db.execute(
        "SELECT exam_id, title, exam_duration, exam_start_time, exam_notice, "
        "template_uploaded, teacher_username FROM exams WHERE exam_number = ?",
        (exam_number,)
    ).fetchone()
    if not exam:
        return jsonify({'error': '试卷编号不存在'}), 404

    requires_stata = _exam_requires_stata(exam['exam_id'])
    d = dict(exam)
    d['requires_stata'] = requires_stata
    return jsonify(d)


@app.route('/api/exam/<int:exam_id>', methods=['GET'])
def get_exam_info_for_cover(exam_id):
    """
    v4.0: Get basic exam info for the exam cover page (no auth required).
    """
    db = get_db()
    exam = db.execute(
        "SELECT exam_id, title, exam_duration, exam_start_time, exam_notice, "
        "template_uploaded FROM exams WHERE exam_id = ?",
        (exam_id,)
    ).fetchone()
    if not exam:
        return jsonify({'error': '考试不存在'}), 404
    return jsonify(dict(exam))


# ════════════════════════════════════════════
# v4.0: Grading Page APIs
# ════════════════════════════════════════════

@app.route('/api/teacher/exam/<int:exam_id>/students', methods=['GET'])
@teacher_required
def get_exam_students(exam_id):
    """
    v4.0: Get list of all students who took this exam, sorted by student_id.
    Returns cover-page data for the grading page.
    """
    db = get_db()
    students = db.execute(
        "SELECT session_id, student_id, student_name, status, total_score, is_graded, "
        "start_time, end_time FROM exam_sessions WHERE exam_id = ? "
        "ORDER BY student_id ASC",
        (exam_id,)
    ).fetchall()
    return jsonify([dict(s) for s in students])


@app.route('/api/teacher/exam/<int:exam_id>/grading-sessions', methods=['GET'])
@teacher_required
def get_grading_sessions(exam_id):
    """
    v4.0: Get all sessions with full question+answer data for grading.
    Also returns exam settings (auto_grade flag).
    """
    db = get_db()
    exam = db.execute(
        "SELECT exam_id, title, auto_grade FROM exams WHERE exam_id = ?",
        (exam_id,)
    ).fetchone()
    if not exam:
        return jsonify({'error': '考试不存在'}), 404

    # Get all questions
    questions = db.execute(
        "SELECT q_id, q_type, content, option_a, option_b, option_c, option_d, "
        "correct_answer, score, sort_order FROM questions WHERE exam_id = ? "
        "ORDER BY sort_order",
        (exam_id,)
    ).fetchall()

    # Get all sessions (submitted only)
    sessions = db.execute(
        "SELECT session_id, student_id, student_name, total_score, is_graded, status "
        "FROM exam_sessions WHERE exam_id = ? AND status = 'submitted' "
        "ORDER BY student_id ASC",
        (exam_id,)
    ).fetchall()

    # Get all answers for these sessions
    session_ids = [s['session_id'] for s in sessions]
    answers_by_session = {}
    if session_ids:
        placeholders = ','.join('?' for _ in session_ids)
        answers = db.execute(
            f"SELECT session_id, q_id, answer_text, is_correct, score "
            f"FROM answers WHERE session_id IN ({placeholders})",
            session_ids
        ).fetchall()
        for a in answers:
            sid = a['session_id']
            if sid not in answers_by_session:
                answers_by_session[sid] = {}
            answers_by_session[sid][a['q_id']] = {
                'answer_text': a['answer_text'],
                'is_correct': a['is_correct'],
                'score': a['score']
            }

    result_sessions = []
    for s in sessions:
        qlist = []
        for q in questions:
            a = answers_by_session.get(s['session_id'], {}).get(q['q_id'], {})
            qlist.append({
                'q_id': q['q_id'],
                'q_type': q['q_type'],
                'content': q['content'],
                'option_a': q['option_a'],
                'option_b': q['option_b'],
                'option_c': q['option_c'],
                'option_d': q['option_d'],
                'correct_answer': q['correct_answer'],
                'score': q['score'],
                'sort_order': q['sort_order'],
                'answer_text': a.get('answer_text', ''),
                'is_correct': a.get('is_correct'),
                'actual_score': a.get('score', 0)
            })
        result_sessions.append({
            'session_id': s['session_id'],
            'student_id': s['student_id'],
            'student_name': s['student_name'],
            'total_score': s['total_score'],
            'is_graded': s['is_graded'],
            'status': s['status'],
            'questions': qlist
        })

    return jsonify({
        'exam': dict(exam),
        'sessions': result_sessions
    })


@app.route('/api/teacher/exam/<int:exam_id>/auto-grade', methods=['POST'])
@teacher_required
def update_auto_grade_setting(exam_id):
    """v4.0: Toggle auto_grade setting for an exam."""
    data = request.get_json()
    auto_grade = data.get('auto_grade', 1)
    db = get_db()
    db.execute("UPDATE exams SET auto_grade = ? WHERE exam_id = ?",
               (1 if auto_grade else 0, exam_id))
    db.commit()
    return jsonify({'ok': True})


@app.route('/api/teacher/exam/<int:exam_id>/auto-grade', methods=['GET'])
@teacher_required
def get_auto_grade_setting(exam_id):
    """v4.0: Get auto_grade setting."""
    db = get_db()
    row = db.execute("SELECT auto_grade FROM exams WHERE exam_id = ?", (exam_id,)).fetchone()
    if not row:
        return jsonify({'error': '考试不存在'}), 404
    return jsonify({'auto_grade': bool(row['auto_grade'])})


@app.route('/api/teacher/exam/<int:exam_id>/stats/xlsx', methods=['GET'])
@teacher_required
def export_stats_xlsx(exam_id):
    """
    v4.0: Export statistics xlsx with all students' scores.
    Columns: 学号, 姓名, 各题得分..., 总得分
    """
    db = get_db()
    exam = db.execute("SELECT title FROM exams WHERE exam_id = ?", (exam_id,)).fetchone()
    if not exam:
        return jsonify({'error': '考试不存在'}), 404

    questions = db.execute(
        "SELECT q_id, sort_order, score FROM questions WHERE exam_id = ? ORDER BY sort_order",
        (exam_id,)
    ).fetchall()

    sessions = db.execute(
        "SELECT session_id, student_id, student_name, total_score, is_graded "
        "FROM exam_sessions WHERE exam_id = ? AND status = 'submitted' "
        "ORDER BY student_id ASC",
        (exam_id,)
    ).fetchall()

    # Build answers map: {session_id: {q_id: score}}
    q_scores = {}
    for s in sessions:
        answers = db.execute(
            "SELECT q_id, score FROM answers WHERE session_id = ?",
            (s['session_id'],)
        ).fetchall()
        q_scores[s['session_id']] = {a['q_id']: a['score'] for a in answers}

    wb = Workbook()
    ws = wb.active
    ws.title = "得分统计"

    # Header row
    headers = ['学号', '姓名']
    for q in questions:
        headers.append(f'第{q["sort_order"]}题({q["score"]}分)')
    headers.append('总得分')
    headers.append('已批改')
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)

    # Data rows
    for s in sessions:
        row_data = [s['student_id'], s['student_name']]
        for q in questions:
            row_data.append(q_scores.get(s['session_id'], {}).get(q['q_id'], 0))
        row_data.append(s['total_score'])
        row_data.append('是' if s['is_graded'] else '否')
        ws.append(row_data)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    safe_title = re.sub(r'[<>:"/\\|?*]', '_', exam['title'])[:40]
    return send_file(
        output,
        as_attachment=True,
        download_name=f'{safe_title}_成绩统计.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


@app.route('/api/teacher/exam/<int:exam_id>/download-all', methods=['GET'])
@teacher_required
def download_all_papers(exam_id):
    """
    v4.0: Download all student papers as a zip file.
    Reads from User/<username>/<exam_title>/student_papers/
    """
    db = get_db()
    exam = db.execute(
        "SELECT title, teacher_username FROM exams WHERE exam_id = ?",
        (exam_id,)
    ).fetchone()
    if not exam:
        return jsonify({'error': '考试不存在'}), 404

    if not exam['teacher_username']:
        return jsonify({'error': '仅支持注册教师使用此功能'}), 400
    if str(session.get('teacher_username', '')).strip() != str(exam['teacher_username']).strip():
        return jsonify({'error': '无权访问此考试'}), 403

    student_dir = os.path.join(_exam_subdir(exam['teacher_username'], exam['title']), 'student_papers')
    if not os.path.exists(student_dir):
        return jsonify({'error': '暂无学生试卷'}), 404

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for fname in os.listdir(student_dir):
            fpath = os.path.join(student_dir, fname)
            if os.path.isfile(fpath):
                zf.write(fpath, fname)
    buf.seek(0)

    safe_title = re.sub(r'[<>:"/\\|?*]', '_', exam['title'])[:40]
    return send_file(
        buf,
        as_attachment=True,
        download_name=f'{safe_title}_所有学生试卷.zip',
        mimetype='application/zip'
    )


@app.route('/api/teacher/session/<int:session_id>/download/xlsx', methods=['GET'])
@teacher_required
def download_student_xlsx(session_id):
    """
    v4.0: Download a single student's xlsx file.
    """
    db = get_db()
    srow = db.execute(
        "SELECT exam_id, student_id, student_name FROM exam_sessions WHERE session_id = ?",
        (session_id,)
    ).fetchone()
    if not srow:
        return jsonify({'error': '会话不存在'}), 404

    exam = db.execute(
        "SELECT title, teacher_username FROM exams WHERE exam_id = ?",
        (srow['exam_id'],)
    ).fetchone()
    if not exam or not exam['teacher_username']:
        return jsonify({'error': '无法找到对应试卷'}), 404

    safename = re.sub(r'[<>:"/\\|?*]', '_', str(srow['student_id']))[:50]
    xlsx_path = os.path.join(
        _exam_subdir(exam['teacher_username'], exam['title']),
        'student_papers',
        f'{safename}.xlsx'
    )
    if not os.path.exists(xlsx_path):
        return jsonify({'error': '学生答卷文件不存在'}), 404

    return send_file(
        xlsx_path,
        as_attachment=True,
        download_name=f'{srow["student_name"]}_{srow["student_id"]}_答卷.xlsx'
    )


# ════════════════════════════════════════════
# Stata APIs (v2.0 → v3.0 multi-process)
# ════════════════════════════════════════════

# v3.0: Support multiple Stata process names
STATA_PROCESS_NAMES = [
    'StataMP-64.exe',
    'StataMP.exe',
    'StataSE-64.exe',
    'StataSE.exe',
    'StataBE-64.exe',
    'StataBE.exe',
    'stata.exe',
    'stata-64.exe',
]

def _is_stata_running():
    """Check if any Stata process is running on this machine (Windows)."""
    try:
        result = subprocess.run(
            ['tasklist'],
            capture_output=True, text=True, timeout=5
        )
        output_lower = result.stdout.lower()
        for pname in STATA_PROCESS_NAMES:
            if pname.lower() in output_lower:
                return True
        return False
    except Exception:
        return False

def _find_stata_hwnd():
    """Find Stata window handle."""
    try:
        import win32gui
        hwnds = []
        def callback(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if 'Stata' in title:
                    hwnds.append(hwnd)
        win32gui.EnumWindows(callback, None)
        return hwnds[0] if hwnds else None
    except ImportError:
        return None

def _show_window(hwnd, show=True):
    """Show (5) or hide (0) a window."""
    try:
        SW_SHOW = 5
        SW_HIDE = 0
        ctypes.windll.user32.ShowWindow(hwnd, SW_SHOW if show else SW_HIDE)
        return True
    except Exception:
        return False

@app.route('/api/stata/check', methods=['GET'])
def stata_check():
    """检测 Stata 是否已打开"""
    running = _is_stata_running()
    return jsonify({'stata_running': running})

@app.route('/api/stata/show', methods=['POST'])
def stata_show():
    """弹出 Stata 窗口"""
    hwnd = _find_stata_hwnd()
    if hwnd:
        if _show_window(hwnd, True):
            ctypes.windll.user32.SetForegroundWindow(hwnd)
            return jsonify({'ok': True, 'message': 'Stata 窗口已弹出'})
    return jsonify({'ok': False, 'message': '未找到 Stata 窗口，请手动打开'})

@app.route('/api/stata/hide', methods=['POST'])
def stata_hide():
    """缩小 Stata 窗口"""
    hwnd = _find_stata_hwnd()
    if hwnd:
        if _show_window(hwnd, False):
            return jsonify({'ok': True, 'message': 'Stata 窗口已隐藏'})
    return jsonify({'ok': False, 'message': '未找到 Stata 窗口'})

# ─── Exam Status Check (v3.0) ───

@app.route('/api/exam/session/<int:session_id>/status', methods=['GET'])
def check_exam_status(session_id):
    """学生端检查考试状态（是否被终止）"""
    db = get_db()
    srow = db.execute(
        "SELECT status, terminated_by_teacher, exam_end_time FROM exam_sessions WHERE session_id = ?",
        (session_id,)
    ).fetchone()
    if not srow:
        return jsonify({'error': '会话不存在'}), 404
    return jsonify({
        'status': srow['status'],
        'terminated': bool(srow['terminated_by_teacher']),
        'exam_end_time': srow['exam_end_time']
    })

# ════════════════════════════════════════════
# Run
# ════════════════════════════════════════════
if __name__ == '__main__':
    with app.app_context():
        init_db()
    print("AI智能考试系统后端启动中... (v4.0)")
    print("访问 http://localhost:5000")
    app.run(host='0.0.0.0', port=5000, debug=False)
