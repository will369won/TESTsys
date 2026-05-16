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
from openpyxl.worksheet.datavalidation import DataValidation
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
            student_id TEXT NOT NULL,
            name TEXT NOT NULL,
            class_name TEXT,
            grade TEXT,
            school TEXT,
            major TEXT,
            teacher_username TEXT,
            PRIMARY KEY (student_id, teacher_username)
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

    # v5.0: add major column if missing (migration from older schema)
    try:
        db.execute("ALTER TABLE students ADD COLUMN major TEXT")
        db.commit()
    except:
        pass

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
        ('ban_copy', '0'), ('ban_screenshot', '0'),
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

    # v5.0: Add ban_copy and ban_screenshot columns
    try:
        db.execute("ALTER TABLE exams ADD COLUMN ban_copy INTEGER DEFAULT 0")
    except Exception:
        pass
    try:
        db.execute("ALTER TABLE exams ADD COLUMN ban_screenshot INTEGER DEFAULT 0")
    except Exception:
        pass
    # v4.0: Add unique index on exam_number
    try:
        db.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_exam_number ON exams(exam_number)")
    except Exception:
        pass

    # v4.0.3: Add new columns for exam cover info
    for col in ['exam_title', 'school', 'college', 'class_name', 'grade']:
        try:
            db.execute(f"ALTER TABLE exams ADD COLUMN {col} TEXT")
        except Exception:
            pass

    # v4.0.3: Add remark column to questions
    try:
        db.execute("ALTER TABLE questions ADD COLUMN remark TEXT")
    except Exception:
        pass

    # v4.0.3: Add exam_notice_title to exams
    try:
        db.execute("ALTER TABLE exams ADD COLUMN exam_notice_title TEXT")
    except Exception:
        pass

    # v5.0: Add major column to exams
    try:
        db.execute("ALTER TABLE exams ADD COLUMN major TEXT")
    except Exception:
        pass

    # v4.1: Add tab_switch_count column to exam_sessions
    try:
        db.execute("ALTER TABLE exam_sessions ADD COLUMN tab_switch_count INTEGER DEFAULT 0")
    except Exception:
        pass

    # v4.2: Add data_file column to questions
    try:
        db.execute("ALTER TABLE questions ADD COLUMN data_file TEXT")
    except Exception:
        pass

    # v4.4: Add teacher_username to students for per-teacher binding
    # Step 1: add the column (if not exists)
    try:
        db.execute("ALTER TABLE students ADD COLUMN teacher_username TEXT")
    except Exception:
        pass
    # Step 2: fill NULL teacher_username with '001' (existing students owned by built-in teacher)
    try:
        db.execute("UPDATE students SET teacher_username = '001' WHERE teacher_username IS NULL")
    except Exception:
        pass
    # Step 3: recreate table with composite primary key if old schema
    try:
        # Check if primary key is single-column (old schema)
        info = db.execute("PRAGMA table_info(students)").fetchall()
        pk_cols = [r for r in info if r['pk'] > 0]
        if len(pk_cols) == 1:
            # Backup, drop, recreate with composite key
            db.execute("CREATE TABLE IF NOT EXISTS students_new (student_id TEXT NOT NULL, name TEXT NOT NULL, class_name TEXT, grade TEXT, school TEXT, teacher_username TEXT NOT NULL DEFAULT '001', PRIMARY KEY (student_id, teacher_username))")
            db.execute("INSERT OR IGNORE INTO students_new (student_id, name, class_name, grade, school, teacher_username) SELECT student_id, name, class_name, grade, school, COALESCE(teacher_username, '001') FROM students")
            db.execute("DROP TABLE students")
            db.execute("ALTER TABLE students_new RENAME TO students")
    except Exception:
        pass

    # v6.0: Create exam_roster table for binding students to exams
    try:
        db.execute("""
            CREATE TABLE IF NOT EXISTS exam_roster (
                exam_id INTEGER NOT NULL,
                student_id TEXT NOT NULL,
                PRIMARY KEY (exam_id, student_id),
                FOREIGN KEY (exam_id) REFERENCES exams(exam_id) ON DELETE CASCADE
            )
        """)
    except Exception:
        pass

    # Migrate existing exam sessions into exam_roster (backfill)
    try:
        existing_rostered = db.execute("SELECT COUNT(*) FROM exam_roster").fetchone()[0]
        if existing_rostered == 0:
            db.execute("""
                INSERT OR IGNORE INTO exam_roster (exam_id, student_id)
                SELECT DISTINCT exam_id, student_id FROM exam_sessions
            """)
    except Exception:
        pass

    # v6.0: Add bypass_late_start to exam_roster
    try:
        db.execute("ALTER TABLE exam_roster ADD COLUMN bypass_late_start INTEGER DEFAULT 0")
    except Exception:
        pass  # Column may already exist

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

def _strip_image_prefix(text: str) -> str:
    """Remove '图片:' or '图片：' prefix from image filenames."""
    if not text:
        return text
    text = text.strip()
    if text.startswith('图片:'):
        return text[3:].strip()
    if text.startswith('图片：'):
        return text[3:].strip()
    return text

_IMAGE_EXTS = {'.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp', '.svg', '.ico'}

def _looks_like_image(name: str) -> bool:
    """Check if a string looks like an image filename (has image extension)."""
    if not name:
        return False
    _, ext = os.path.splitext(name.strip().lower())
    return ext in _IMAGE_EXTS


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


def _teacher_attach_dir(username: str) -> str:
    """Get teacher's personal attachment storage directory (independent of any exam)."""
    d = os.path.join(_user_dir(username), '_attachments')
    os.makedirs(d, exist_ok=True)
    return d


def _exam_workdir(username: str, exam_id: int) -> str:
    """Get the exam working directory for Stata data files."""
    d = os.path.join(_user_dir(username), '_exam_workdir', str(exam_id))
    os.makedirs(d, exist_ok=True)
    return d


def _session_workdir(username: str, exam_id: int, session_id: int) -> str:
    """Get the per-session Stata working directory (isolated per student)."""
    d = os.path.join(_exam_workdir(username, exam_id), f'session_{session_id}')
    os.makedirs(d, exist_ok=True)
    return d


def _deploy_exam_data(exam_id: int, session_id: int = None):
    """Copy data files (dta/xlsx) referenced by questions to the workdir.
    
    If session_id is provided, copies to the per-session subdirectory
    for student isolation. Otherwise copies to the shared exam workdir.
    """
    db = get_db()
    exam = db.execute(
        "SELECT teacher_username, title FROM exams WHERE exam_id = ?",
        (exam_id,)
    ).fetchone()
    if not exam or not exam['teacher_username']:
        return

    username = exam['teacher_username']

    # Get all questions with data_file set
    questions = db.execute(
        "SELECT q_id, data_file FROM questions WHERE exam_id = ? AND data_file IS NOT NULL AND data_file != ''",
        (exam_id,)
    ).fetchall()

    if not questions:
        return

    target_dir = _session_workdir(username, exam_id, session_id) if session_id else _exam_workdir(username, exam_id)
    teacher_attach_dir = _teacher_attach_dir(username)

    # Also check exam attachment dir
    exam_attach_dir = os.path.join(_exam_subdir(username, exam['title']), 'attachments')

    # Collect all referenced data files
    data_files = set()
    for q in questions:
        for f in q['data_file'].replace('；', ';').split(';'):
            f = f.strip()
            if f:
                data_files.add(f)

    # Copy each data file from teacher's personal storage or exam attachment dir
    for fname in data_files:
        copied = False
        for src_dir in [teacher_attach_dir, exam_attach_dir]:
            src_path = os.path.join(src_dir, fname)
            if os.path.exists(src_path):
                try:
                    shutil.copy2(src_path, os.path.join(target_dir, fname))
                    copied = True
                    break
                except Exception:
                    pass
        if not copied:
            print(f"[v4.2] Data file not found: {fname}")


def _cleanup_session_workdir(exam_id: int, session_id: int):
    """Delete a single session's working directory (after exam submit/expiry)."""
    db = get_db()
    exam = db.execute(
        "SELECT teacher_username FROM exams WHERE exam_id = ?",
        (exam_id,)
    ).fetchone()
    if not exam or not exam['teacher_username']:
        return
    workdir = _session_workdir(exam['teacher_username'], exam_id, session_id)
    if os.path.exists(workdir):
        try:
            shutil.rmtree(workdir)
        except Exception as e:
            print(f"[v4.3] Failed to cleanup session workdir: {e}")


def _cleanup_exam_workdir(exam_id: int):
    """Delete the entire exam working directory (all sessions)."""
    db = get_db()
    exam = db.execute(
        "SELECT teacher_username FROM exams WHERE exam_id = ?",
        (exam_id,)
    ).fetchone()
    if not exam or not exam['teacher_username']:
        return

    workdir = _exam_workdir(exam['teacher_username'], exam_id)
    if os.path.exists(workdir):
        try:
            shutil.rmtree(workdir)
        except Exception as e:
            print(f"[v4.2] Failed to cleanup workdir: {e}")


# ─── v5.0: Option shuffle helper ───

def _shuffle_options(questions, student_id, exam_id, enabled):
    """Deterministically shuffle option_a/b/c/d for 单选题/多选题 per student.

    Seed = hashlib.md5(f'{student_id}:{exam_id}:{q_id}').hexdigest() (full 32-char).
    Shuffles option values and remaps correct_answer to the new letter.
    """
    if not enabled:
        return [dict(q) for q in questions]

    result = []
    for q in questions:
        qd = dict(q)
        if qd['q_type'] in ('单选题', '多选题'):
            # Step 1: collect option keys and values
            keys = []
            vals = []
            for k in ['option_a', 'option_b', 'option_c', 'option_d']:
                if qd.get(k):
                    keys.append(k)
                    vals.append(qd[k])

            if len(keys) >= 2:
                # Step 2: remember the correct answer's value BEFORE shuffle
                ca = qd.get('correct_answer', '').upper()
                correct_val = qd.get(f'option_{ca.lower()}', '')

                # Step 3: shuffle values deterministically
                seed_str = f"{student_id}:{exam_id}:{qd['q_id']}"
                seed_hex = hashlib.md5(seed_str.encode()).hexdigest()
                rng = random.Random(int(seed_hex, 16))
                rng.shuffle(vals)

                # Step 4: assign shuffled values back to original keys
                for i, k in enumerate(keys):
                    qd[k] = vals[i]

                # Step 5: find which NEW key holds the correct value
                if correct_val:
                    for k in keys:
                        if qd[k] == correct_val:
                            qd['correct_answer'] = k[-1].upper()
                            break

        result.append(qd)
    return result


def _shuffle_questions(questions, student_id, exam_id, question_shuffle_enabled, option_shuffle_enabled):
    """v5.1: Shuffle question order per-type + optionally option order.

    Objective questions (单选题/多选题/判断题) are shuffled WITHIN each type,
    keeping questions of the same type grouped together (e.g. all 单选题 come first,
    then all 多选题, then all 判断题). This prevents cross-type mixing.
    Subjective questions (简答题, 综合题, etc.) stay at the end in original order.
    Option shuffle only applies to 单选题/多选题 (not 判断题).
    """
    if not question_shuffle_enabled and not option_shuffle_enabled:
        return [dict(q) for q in questions]

    # Fixed type groups in display order
    type_order = ('单选题', '多选题', '判断题')

    # Collect questions by type
    objective_groups = {}
    for t in type_order:
        objective_groups[t] = []
    subjective_qs = []
    for q in questions:
        if q['q_type'] in type_order:
            objective_groups[q['q_type']].append(q)
        else:
            subjective_qs.append(q)

    # Shuffle each type group independently if enabled
    if question_shuffle_enabled:
        seed_str = f"{student_id}:{exam_id}"
        seed_hex = hashlib.md5(seed_str.encode()).hexdigest()
        rng = random.Random(int(seed_hex, 16))
        for t in type_order:
            group = objective_groups[t]
            if len(group) >= 2:
                rng.shuffle(group)

    # Flatten: all 单选题, then all 多选题, then all 判断题
    objective_qs = []
    for t in type_order:
        objective_qs.extend(objective_groups[t])

    # Shuffle options within each objective question if enabled
    if option_shuffle_enabled:
        objective_qs = _shuffle_options(objective_qs, student_id, exam_id, True)

    # Combine: shuffled objectives (grouped by type), then subjective in original order
    return [dict(q) for q in objective_qs] + [dict(q) for q in subjective_qs]


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


def _execute_stata_via_mcp(command_text: str) -> str:
    """Execute Stata command via stata_mcp API.
    Uses write_dofile to create a do-file, then stata_do to execute it.
    stata_mcp handles log files internally with is_replace_log=True.
    
    stata_do returns a dict like {'log_file_path': {'text': 'path/to/log'}}.
    We read the log file and extract only the actual command output.
    """
    try:
        from stata_mcp.api import stata_do, write_dofile

        # Write the do-file content
        dofile_path = write_dofile(command_text)

        # Execute the do-file
        # is_replace_log=True handles the r(608) error automatically
        result = stata_do(dofile_path, is_replace_log=True)

        # stata_do returns a dict: {'log_file_path': {'text': '...'}}
        # Extract log content from the log file
        if isinstance(result, dict) and 'log_file_path' in result:
            log_path = result['log_file_path'].get('text', '') or result['log_file_path']
            if isinstance(log_path, dict):
                log_path = log_path.get('text', '')
            if log_path and os.path.exists(str(log_path)):
                with open(str(log_path), 'r', encoding='utf-8', errors='replace') as f:
                    content = f.read()
            else:
                return '执行完成（无法读取日志文件）'
        elif isinstance(result, str):
            content = result
        else:
            return f'执行完成（未知返回类型: {type(result).__name__}）'

        if not content:
            return '执行完成（无输出）'

        # Extract only the actual command output from Stata log.
        # Stata log format:
        #   --- separator ---
        #   header lines (name:, log:, log type:, opened on:)
        #   blank
        #   . do "path/to/dofile"
        #   > "
        #   blank
        #   [. cd, . use, etc command lines]
        #   [output lines]
        #   .
        #   end of do-file
        #   [. log close + footer]
        #   --- separator ---
        lines = content.split('\n')
        keep = []
        in_body = False
        for line in lines:
            s = line.rstrip()
            st = s.strip()

            # Skip header lines before the do-file execution
            if not in_body:
                if st == '' or st.startswith('---') or st.startswith('name:') or st.startswith('log:') \
                        or st.startswith('log type:') or st.startswith('opened on:'):
                    continue
                if st.startswith('. do ') or st == '. do' or 'do "' in st:
                    continue
                if st == '> "' or st.startswith('> '):
                    continue
                # First non-header line starts the body
                in_body = True
                # Fall through to body processing

            # In body: collect output
            if st == '':
                keep.append(s)
                continue
            if st == '.':
                continue
            # End conditions: only use explicit log-end markers
            if st.startswith('end of do-file'):
                break
            if st.startswith('. log close') or st.startswith('log close'):
                break
            if st.startswith('closed on:'):
                break
            # Strip leading '. ' prefix from echoed command lines
            if s.startswith('. ') and not s.startswith('. do'):
                s = s[2:]
            # Skip cd command echo and its path output
            if s.startswith('cd "') or s.startswith('cd '):
                continue
            # Skip path echo (after cd, Stata prints the full path)
            if s.startswith('C:\\') or s.startswith('D:\\'):
                continue
            # Skip exit wrapper
            if s.strip() in ('exit, STATA', 'exit'):
                continue
            keep.append(s)

        # Trim trailing/leading empty lines
        while keep and keep[-1].strip() == '':
            keep.pop()
        while keep and keep[0].strip() == '':
            keep.pop(0)

        if not keep:
            return '执行完成（无输出）'

        cleaned = '\n'.join(keep)
        # Truncate to 10000 chars
        if len(cleaned) > 10000:
            cleaned = cleaned[-10000:]
        return cleaned
    except ImportError:
        return 'stata_mcp 未安装，请运行: pip install stata-mcp'
    except Exception as e:
        return f'Stata执行失败: {str(e)}'
    finally:
        _cleanup_stata_batch_logs()


def _cleanup_stata_batch_logs():
    """Remove Stata batch log files left in backend directory."""
    try:
        backend_dir = os.path.dirname(os.path.abspath(__file__))
        for fname in os.listdir(backend_dir):
            if fname.startswith('stata_batch__') and fname.endswith('.log'):
                os.remove(os.path.join(backend_dir, fname))
    except Exception:
        pass

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
        session['teacher_username'] = tid
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

# ─── Template Prescan ───

def _prescan_template(workbook):
    """Prescan a template xlsx for attachment requirements and referenced images.
    Returns dict: {needs_attachment: bool, referenced_images: [str], question_count: int}
    """
    ws = workbook.active
    # Detect header row
    first_row = list(ws.iter_rows(min_row=2, max_row=2, values_only=True))
    if first_row and first_row[0] and str(first_row[0][0]).strip() == '题型':
        rows = list(ws.iter_rows(min_row=3, values_only=True))
    else:
        rows = list(ws.iter_rows(min_row=2, values_only=True))

    if not rows:
        return {'needs_attachment': False, 'referenced_images': [], 'question_count': 0}

    needs_attachment = False
    referenced_images = []
    question_count = 0

    for row in rows:
        q_type = str(row[0]).strip() if row[0] else ''
        content = str(row[1]).strip() if len(row) > 1 and row[1] else ''
        if not q_type or not content:
            continue
        question_count += 1

        # Check 附件 column (col 8, 0-indexed)
        has_attach_val = str(row[8]).strip() if len(row) > 8 and row[8] else ''
        if has_attach_val in ['是', 'yes', '1']:
            needs_attachment = True

        # Check 数据 column (col 10) for data file references
        data_file_raw = str(row[10]).strip() if len(row) > 10 and row[10] else ''
        if data_file_raw:
            for f in data_file_raw.replace('；', ';').split(';'):
                f_clean = f.strip()
                if f_clean and f_clean not in referenced_images:
                    referenced_images.append(f_clean)

        # Parse remark (col 9) for image filenames
        remark_raw = str(row[9]).strip() if len(row) > 9 and row[9] else ''
        if remark_raw:
            for f in remark_raw.replace('；', ';').split(';'):
                f_clean = _strip_image_prefix(f.strip())
                if f_clean and _looks_like_image(f_clean) and f_clean not in referenced_images:
                    referenced_images.append(f_clean)

        # Parse options (cols 2-5) for image filenames
        option_cols = [2, 3, 4, 5]
        for oc in option_cols:
            ov = str(row[oc]).strip() if len(row) > oc and row[oc] else ''
            if ov:
                stripped = _strip_image_prefix(ov)
                if _looks_like_image(stripped) and stripped not in referenced_images:
                    referenced_images.append(stripped)

    return {
        'needs_attachment': needs_attachment,
        'referenced_images': referenced_images,
        'question_count': question_count
    }


@app.route('/api/template/prescan', methods=['POST'])
@teacher_required
def prescan_exam_template():
    """Prescan a template file to check attachment requirements WITHOUT inserting into DB."""
    if 'file' not in request.files:
        return jsonify({'error': '未找到文件'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': '文件名为空'}), 400

    wb = load_workbook(file)
    result = _prescan_template(wb)
    return jsonify({'ok': True, **result})


# ─── Exam Template ───

@app.route('/api/template/exam/download', methods=['GET'])
@teacher_required
def download_exam_template():
    """Download exam template v4.0.6."""
    wb = Workbook()
    ws = wb.active
    ws.title = "题目模板"

    header_font = Font(bold=True, size=11, color='FFFFFF')
    header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
    header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    note_font = Font(size=10, color='333333', bold=True)
    note_align = Alignment(horizontal='left', vertical='top', wrap_text=True)
    body_align = Alignment(horizontal='left', vertical='top', wrap_text=True)

    ws.merge_cells('A1:M1')
    cell = ws['A1']
    cell.value = (
        "填写说明：\n"
        "1. 题型列：使用下拉菜单选择（单选题/多选题/判断题/简答题/综合题）。\n"
        "2. 题干列：填写题目内容。图片在备注列标注文件名（多张用;隔开），图片于附件上传。\n"
        "3. 选项列(A-D)：选择项含图片时，在对应选项单元格填写图片文件名（如 chart.png）。\n"
        "4. 分值列：必填，正整数。\n"
        "5. 附件列：下拉选择\"是\"表示该题需要附件，留空则不需要。\n"
        "6. 参考答案：单选题填大写字母（A），多选题填字母组合（ABD），判断题填\"正确/错误\"。\n"
        "7. 备注列：标注题干图片文件名。\n"
        "8. Stata列：下拉选择[是/否]，默认[是]。\n"
        "9. AI列：下拉选择[是/否]，默认[是]。\n"
        "10. 数据列：填写该题需要的数据文件名(.dta/.xlsx)，多个用;隔开。数据文件通过附件管理上传。"
    )
    cell.font = note_font
    cell.alignment = note_align
    ws.row_dimensions[1].height = 200

    headers = ['题型', '题干', '选项A', '选项B', '选项C', '选项D', '参考答案', '分值', '附件', '备注', '数据', 'Stata', 'AI']
    for col_idx, h in enumerate(headers, 1):
        c = ws.cell(row=2, column=col_idx, value=h)
        c.font = header_font
        c.fill = header_fill
        c.alignment = header_align

    widths = {'A': 14, 'B': 40, 'C': 18, 'D': 18, 'E': 18, 'F': 18,
              'G': 12, 'H': 10, 'I': 8, 'J': 35, 'K': 12, 'L': 8, 'M': 8}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    examples = [
        ('单选题', '以下哪个是Python的特点？', '编译型', '解释型', '汇编', '机器码', 'B', 5, '', '', '', '是', '是'),
        ('多选题', '以下哪些是Web前端技术？', 'HTML', 'CSS', 'Python', 'JavaScript', 'ABD', 5, '', '', '', '是', '是'),
        ('判断题', 'Python是一种面向对象的语言。', '', '', '', '', '正确', 5, '', '', '', '是', '否'),
        ('简答题', '简述Python中的装饰器是什么。', '', '', '', '', '', 10, '', '', '', '是', '是'),
        ('综合题', '阅读以下材料并回答问题...', '', '', '', '', '', 20, '是', 'chart.png', 'data1.dta', '否', '是'),
    ]
    for row_idx, row_data in enumerate(examples, 3):
        for col_idx, value in enumerate(row_data, 1):
            c = ws.cell(row=row_idx, column=col_idx, value=value)
            c.alignment = body_align

    dv_type = DataValidation(type='list', formula1='"单选题,多选题,判断题,简答题,综合题"', allow_blank=True, showDropDown=False)
    dv_type.error = '请从下拉菜单选择题型'
    ws.add_data_validation(dv_type)
    for r in range(3, 105):
        dv_type.add(ws.cell(row=r, column=1))

    dv_attach = DataValidation(type='list', formula1='"是"', allow_blank=True, showDropDown=False)
    dv_attach.error = '请选择"是"或留空'
    ws.add_data_validation(dv_attach)
    for r in range(3, 105):
        dv_attach.add(ws.cell(row=r, column=9))

    dv_stata = DataValidation(type='list', formula1='"是,否"', allow_blank=True, showDropDown=False)
    ws.add_data_validation(dv_stata)
    for r in range(3, 105):
        dv_stata.add(ws.cell(row=r, column=12))

    dv_ai = DataValidation(type='list', formula1='"是,否"', allow_blank=True, showDropDown=False)
    ws.add_data_validation(dv_ai)
    for r in range(3, 105):
        dv_ai.add(ws.cell(row=r, column=13))

    ws.freeze_panes = 'A3'

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

    # Prescan BEFORE inserting anything
    prescan_result = _prescan_template(wb)

    ws = wb.active
    # Detect if row 2 is a header row (contains '题型')
    first_row = list(ws.iter_rows(min_row=2, max_row=2, values_only=True))
    if first_row and first_row[0] and str(first_row[0][0]).strip() == '题型':
        # Row 2 is header row, data starts from row 3
        rows = list(ws.iter_rows(min_row=3, values_only=True))
    else:
        # Compat mode: row 2 is data
        rows = list(ws.iter_rows(min_row=2, values_only=True))
    if not rows:
        return jsonify({'error': '模板为空'}), 400

    db = get_db()
    # v4.0.5: Use uploaded file name as title (strip .xlsx extension)
    original_filename = file.filename
    if original_filename:
        base = os.path.splitext(original_filename)[0]
        if base:
            title = base[:50]  # Limit to 50 chars
        else:
            title = f"试卷_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    else:
        title = f"试卷_{datetime.now().strftime('%Y%m%d_%H%M%S')}"

    # v3.1: Check exam count limit for registered teachers
    username = session.get('teacher_username')
    if username:
        existing = db.execute("SELECT COUNT(*) as cnt FROM exams WHERE teacher_username = ?", (username,)).fetchone()
        if existing['cnt'] >= MAX_EXAMS_PER_TEACHER:
            return jsonify({
                'error': f'每个用户最多上传 {MAX_EXAMS_PER_TEACHER} 份试卷。请先删除旧试卷再上传新试卷。'
            }), 403

    # ── v4.2: Same-name exam overwrite ──
    if username:
        old_exam = db.execute("SELECT exam_id, title FROM exams WHERE teacher_username = ? AND title = ?", (username, title)).fetchone()
        if old_exam:
            old_exam_id = old_exam['exam_id']
            # Delete old exam subdirectory
            _delete_exam_subdir(username, title)
            # Delete old DB records
            db.execute("DELETE FROM exam_attachments WHERE q_id IN (SELECT q_id FROM questions WHERE exam_id = ?)", (old_exam_id,))
            db.execute("DELETE FROM answers WHERE q_id IN (SELECT q_id FROM questions WHERE exam_id = ?)", (old_exam_id,))
            db.execute("DELETE FROM questions WHERE exam_id = ?", (old_exam_id,))
            db.execute("DELETE FROM exam_sessions WHERE exam_id = ?", (old_exam_id,))
            db.execute("DELETE FROM exams WHERE exam_id = ?", (old_exam_id,))
            db.commit()

    created_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    # v4.0: Generate unique exam number
    exam_number = _generate_exam_number()

    cur = db.execute(
        "INSERT INTO exams (title, duration_minutes, created_at, template_uploaded, teacher_username, exam_number) VALUES (?, 60, ?, 1, ?, ?)",
        (title, created_at, username, exam_number)
    )
    exam_id = cur.lastrowid

    # v4.1: Clean up old attachments from previous exam sessions
    # Delete existing attachment files for this exam
    old_attach_rows = db.execute(
        "SELECT a.file_path FROM exam_attachments a "
        "JOIN questions q ON a.q_id = q.q_id WHERE q.exam_id = ?",
        (exam_id,)
    ).fetchall()

    # Delete the physical files
    for row2 in old_attach_rows:
        fp = row2['file_path']
        # Try to find in attachment directory
        try:
            attach_dir = _get_attachment_dir(exam_id)
            fpath = os.path.join(attach_dir, fp)
            if os.path.exists(fpath):
                os.remove(fpath)
        except Exception:
            pass

    # Also delete from exam_attachments table
    db.execute(
        "DELETE FROM exam_attachments WHERE q_id IN (SELECT q_id FROM questions WHERE exam_id = ?)",
        (exam_id,)
    )

    # v4.0.6: Collect all image filenames referenced in template to verify against uploaded attachments
    referenced_images = []  # list of (filename, location_description)

    sort_order = 1
    for row in rows:
        # v4.2 template columns (13 cols):
        # [题型, 题干, 选项A, 选项B, 选项C, 选项D, 参考答案, 分值, 附件, 备注, 数据, Stata, AI]
        # Index:   0     1      2       3       4       5       6        7      8      9     10    11    12
        q_type = str(row[0]).strip() if row[0] else ''
        content = str(row[1]).strip() if len(row) > 1 and row[1] else ''
        if not q_type or not content:
            continue

        # Score validation
        score_val = row[7] if len(row) > 7 and row[7] else None
        if score_val is None or str(score_val).strip() == '':
            return jsonify({'error': f'第{sort_order}题分数未填写，请检查后重新上传'}), 400
        try:
            score = float(score_val)
            if score <= 0:
                return jsonify({'error': f'第{sort_order}题分数未填写完整，请检查后重新上传'}), 400
        except (ValueError, TypeError):
            return jsonify({'error': f'第{sort_order}题分数格式错误，请检查后重新上传'}), 400

        # Parse 附件 column (col 8)
        has_attach_val = str(row[8]).strip() if len(row) > 8 and row[8] else ''
        has_attachment = 1 if has_attach_val in ['是', 'yes', '1'] else 0

        # Parse remarks (col 9): extract all filenames, strip possible '图片:' prefix
        remark_raw = str(row[9]).strip() if len(row) > 9 and row[9] else ''
        remark_clean = ';'.join(
            _strip_image_prefix(f) for f in remark_raw.replace('；', ';').split(';') if _strip_image_prefix(f)
        ) if remark_raw else ''

        # Extract image filenames from options (cols 2-5): parse file references
        option_cols = [2, 3, 4, 5]
        option_labels = ['选项A', '选项B', '选项C', '选项D']
        option_values = []
        for oi, oc in enumerate(option_cols):
            ov = str(row[oc]).strip() if len(row) > oc and row[oc] else ''
            if ov:
                stripped = _strip_image_prefix(ov)
                option_values.append(stripped)
                # If the original had '图片:' prefix but stripped version is an image file, collect it
                if ov != stripped and _looks_like_image(stripped):
                    referenced_images.append((stripped, f'第{sort_order}题{option_labels[oi]}'))
                # If the option cell directly contains an image filename (no prefix, but ends in image ext)
                if ov == stripped and _looks_like_image(stripped):
                    referenced_images.append((stripped, f'第{sort_order}题{option_labels[oi]}'))
            else:
                option_values.append('')

        # Extract image filenames from remark
        if remark_clean:
            for img_name in remark_clean.split(';'):
                img_name = img_name.strip()
                if img_name and _looks_like_image(img_name):
                    referenced_images.append((img_name, f'第{sort_order}题备注'))

        # Read 数据 column (col 10)
        data_file_val = str(row[10]).strip() if len(row) > 10 and row[10] else ''
        data_file_clean = ';'.join(
            f.strip() for f in data_file_val.replace('；', ';').split(';') if f.strip()
        ) if data_file_val else ''

        # Collect data files for consistency check (along with images)
        if data_file_clean:
            for df_name in data_file_clean.split(';'):
                df_name = df_name.strip()
                if df_name:
                    referenced_images.append((df_name, f'第{sort_order}题数据'))

        # Read Stata and AI columns (now at 11, 12)
        # If cell is blank/empty, default to 否 (0)
        enable_stata_val = str(row[11]).strip() if len(row) > 11 and row[11] else '否'
        enable_ai_val = str(row[12]).strip() if len(row) > 12 and row[12] else '否'
        enable_stata = 1 if enable_stata_val in ['是', 'yes', '1'] else 0
        enable_ai = 1 if enable_ai_val in ['是', 'yes', '1'] else 0

        db.execute(
            "INSERT INTO questions (exam_id, q_type, content, option_a, option_b, option_c, option_d, correct_answer, score, sort_order, enable_stata, enable_ai, has_attachment, remark, data_file) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            (exam_id, q_type, content,
             option_values[0] if option_values[0] else None,
             option_values[1] if len(option_values) > 1 and option_values[1] else None,
             option_values[2] if len(option_values) > 2 and option_values[2] else None,
             option_values[3] if len(option_values) > 3 and option_values[3] else None,
             str(row[6]).strip() if len(row) > 6 and row[6] else None,
             score,
             sort_order, enable_stata, enable_ai, has_attachment,
             remark_clean if remark_clean else None,
             data_file_clean if data_file_clean else None)
        )
        sort_order += 1

    # v4.0.6: Check if referenced images match uploaded attachments
    # v4.2: Also check teacher's personal attachment storage and auto-import
    if referenced_images:
        # Get the list of uploaded attachment filenames for this exam
        attach_rows = db.execute(
            "SELECT a.original_name FROM exam_attachments a "
            "JOIN questions q ON a.q_id = q.q_id WHERE q.exam_id = ?",
            (exam_id,)
        ).fetchall()
        uploaded_names = set(a['original_name'] for a in attach_rows)

        # Auto-import from teacher's personal storage if found
        imported_from_personal = []
        teacher_attach_dir = None
        if username:
            teacher_attach_dir = _teacher_attach_dir(username)

        mismatches = []
        for img_name, location in referenced_images:
            if img_name in uploaded_names:
                continue
            # Check teacher's personal storage
            if teacher_attach_dir:
                src_path = os.path.join(teacher_attach_dir, img_name)
                if os.path.exists(src_path):
                    # Auto-import: copy to exam attachment dir + insert record
                    exam_attach_dir = _get_attachment_dir(exam_id)
                    dest_path = os.path.join(exam_attach_dir, img_name)
                    import shutil
                    shutil.copy2(src_path, dest_path)
                    uploaded_names.add(img_name)
                    imported_from_personal.append(img_name)
                    continue
            mismatches.append(f'{location}: 引用图片 "{img_name}" 但附���中不存在')

        if imported_from_personal:
            db.commit()  # save auto-imported records first

        if mismatches:
            # Still commit the data but warn the user
            if username:
                edir = _exam_subdir(username, title)
                os.makedirs(edir, exist_ok=True)
            db.commit()
            return jsonify({
                'ok': True,
                'exam_id': exam_id,
                'title': title,
                'question_count': sort_order - 1,
                'exam_number': exam_number,
                'needs_attachment': prescan_result['needs_attachment'],
                'referenced_images': prescan_result['referenced_images'],
                'attachment_warning': '上传附件与模板要求附件不一致，请修改上传附件名或模板',
                'mismatches': mismatches
            }), 200

    # v3.1: Create exam subdirectory for this teacher
    if username:
        edir = _exam_subdir(username, title)
        os.makedirs(edir, exist_ok=True)

    db.commit()
    return jsonify({
        'ok': True, 'exam_id': exam_id, 'title': title,
        'question_count': sort_order - 1,
        'exam_number': exam_number,
        'needs_attachment': prescan_result['needs_attachment'],
        'referenced_images': prescan_result['referenced_images']
    })

# ─── Teacher Personal Attachment Storage (independent of any exam) ───

@app.route('/api/teacher/attachments/upload', methods=['POST'])
@teacher_required
def upload_teacher_attachments():
    """Upload attachments to teacher's personal storage (no exam needed)."""
    username = session.get('teacher_username')
    if not username:
        return jsonify({'error': '请先注册教师账号'}), 400
    attach_dir = _teacher_attach_dir(username)
    uploaded = []
    for f in request.files.getlist('files'):
        # Overwrite if same name exists
        fpath = os.path.join(attach_dir, f.filename)
        if os.path.exists(fpath):
            os.remove(fpath)
        f.save(fpath)
        uploaded.append({'name': f.filename, 'size': os.path.getsize(fpath)})
    return jsonify({'ok': True, 'files': uploaded})


@app.route('/api/teacher/attachments/list', methods=['GET'])
@teacher_required
def list_teacher_attachments():
    """List all attachments in teacher's personal storage."""
    username = session.get('teacher_username')
    if not username:
        return jsonify({'attachments': []})
    attach_dir = _teacher_attach_dir(username)
    if not os.path.exists(attach_dir):
        return jsonify({'attachments': []})
    files = []
    for f in sorted(os.listdir(attach_dir)):
        fpath = os.path.join(attach_dir, f)
        if os.path.isfile(fpath):
            files.append({'name': f, 'size': os.path.getsize(fpath), 'display_name': f})
    return jsonify({'attachments': files})


@app.route('/api/teacher/attachments/<filename>', methods=['DELETE'])
@teacher_required
def delete_teacher_attachment(filename):
    """Delete an attachment from teacher's personal storage."""
    import urllib.parse
    filename = urllib.parse.unquote(filename)
    username = session.get('teacher_username')
    if not username:
        return jsonify({'error': '未登录'}), 401
    attach_dir = _teacher_attach_dir(username)
    fpath = os.path.join(attach_dir, filename)
    if os.path.exists(fpath):
        os.remove(fpath)
        return jsonify({'ok': True})
    return jsonify({'error': '文件不存在'}), 404


@app.route('/api/teacher/attachments/file/<filename>')
def serve_teacher_attachment(filename):
    """Serve an attachment file from teacher's personal storage."""
    import urllib.parse
    filename = urllib.parse.unquote(filename)
    username = session.get('teacher_username')
    if not username:
        return jsonify({'error': '未登录'}), 401
    attach_dir = _teacher_attach_dir(username)
    # Support lookup by session-stored username
    if not os.path.exists(os.path.join(attach_dir, filename)) and 'teacher_username' in session:
        pass  # already checked above
    fpath = os.path.join(attach_dir, filename)
    if os.path.exists(fpath):
        return send_file(fpath)
    return jsonify({'error': '文件不存在'}), 404


# ─── Exam Attachment Upload (from teacher management page) ───

@app.route('/api/exam/<int:exam_id>/attachments', methods=['POST'])
@teacher_required
def upload_exam_attachments(exam_id):
    """Upload attachments for an exam (from teacher management page)."""
    db = get_db()
    row = db.execute("SELECT exam_id FROM exams WHERE exam_id = ?", (exam_id,)).fetchone()
    if not row:
        return jsonify({'error': '考试不存在'}), 404

    if 'files' not in request.files:
        return jsonify({'error': '未找到文件'}), 400

    files = request.files.getlist('files')
    if not files:
        return jsonify({'error': '文件列表为空'}), 400

    username = session.get('teacher_username')
    saved_count = 0
    for f in files:
        if f.filename == '':
            continue
        # Determine save path
        if username:
            title_row = db.execute("SELECT title FROM exams WHERE exam_id = ?", (exam_id,)).fetchone()
            title = title_row['title'] if title_row else f'exam_{exam_id}'
            subdir = _exam_subdir(username, title)
            attach_dir = os.path.join(subdir, 'attachments')
        else:
            attach_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'attachments')
        os.makedirs(attach_dir, exist_ok=True)

        # Check storage limit
        ok, msg = _check_storage_limit(subdir if username else attach_dir)
        # Save file — overwrite if same filename exists
        file_path = os.path.join(attach_dir, f.filename)
        if os.path.exists(file_path):
            try:
                os.remove(file_path)
            except Exception:
                pass
        f.save(file_path)
        saved_count += 1

    return jsonify({'ok': True, 'saved': saved_count, 'message': f'成功上传 {saved_count} 个文件'})


@app.route('/api/exam/<int:exam_id>/check-attachments-ready', methods=['GET'])
@teacher_required
def check_attachments_ready(exam_id):
    """Check if all questions requiring attachments have them uploaded."""
    db = get_db()
    # Get questions that need attachments
    needs_attach = db.execute(
        "SELECT q_id, sort_order, content FROM questions WHERE exam_id = ? AND has_attachment = 1",
        (exam_id,)
    ).fetchall()

    if not needs_attach:
        return jsonify({'ready': True, 'message': '无需附件'})

    # Get existing attachments for this exam
    existing = db.execute(
        "SELECT q_id, COUNT(*) as cnt FROM exam_attachments a "
        "JOIN questions q ON a.q_id = q.q_id "
        "WHERE q.exam_id = ? GROUP BY a.q_id",
        (exam_id,)
    ).fetchall()
    existing_map = {r['q_id']: r['cnt'] for r in existing}

    missing = []
    for q in needs_attach:
        if q['q_id'] not in existing_map:
            missing.append({
                'q_id': q['q_id'],
                'sort_order': q['sort_order'],
                'content': q['content'][:50] + ('...' if len(q['content']) > 50 else '')
            })

    # Also check if files exist in the attachment directory
    username = session.get('teacher_username')
    if username:
        title_row = db.execute("SELECT title FROM exams WHERE exam_id = ?", (exam_id,)).fetchone()
        title = title_row['title'] if title_row else f'exam_{exam_id}'
        subdir = _exam_subdir(username, title)
        attach_dir = os.path.join(subdir, 'attachments')
        if not os.path.exists(attach_dir) or not os.listdir(attach_dir):
            if missing:
                return jsonify({'ready': False, 'missing': missing,
                    'message': f'还有 {len(missing)} 道题需要附件但未上传'})
            return jsonify({'ready': False, 'message': '附件目录为空，请先上传附件文件'})

    if missing:
        return jsonify({'ready': False, 'missing': missing,
            'message': f'还有 {len(missing)} 道题需要附件但未上传'})

    return jsonify({'ready': True, 'message': '所有需要附件的题目已就绪'})


@app.route('/api/exam/<int:exam_id>/attachments/list', methods=['GET'])
@teacher_required
def list_exam_attachments(exam_id):
    """List all attachment files for an exam."""
    db = get_db()
    row = db.execute("SELECT exam_id FROM exams WHERE exam_id = ?", (exam_id,)).fetchone()
    if not row:
        return jsonify({'error': '考试不存在'}), 404

    username = session.get('teacher_username')
    if username:
        title_row = db.execute("SELECT title FROM exams WHERE exam_id = ?", (exam_id,)).fetchone()
        title = title_row['title'] if title_row else f'exam_{exam_id}'
        subdir = _exam_subdir(username, title)
        attach_dir = os.path.join(subdir, 'attachments')
    else:
        attach_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'attachments')

    files = []
    if os.path.exists(attach_dir):
        for f in sorted(os.listdir(attach_dir), key=lambda x: os.path.getmtime(os.path.join(attach_dir, x)), reverse=True):
            # Extract original filename: stored format is "timestamp_originalname"
            stored_name = f
            display_name = f
            if '_' in f:
                # Format: 1778318519_original_filename.ext
                parts = f.split('_', 1)
                if parts[0].isdigit():
                    display_name = parts[1]
            files.append({
                'name': f,
                'display_name': display_name,
                'size': round(os.path.getsize(os.path.join(attach_dir, f)) / 1024, 1),
                'modified': os.path.getmtime(os.path.join(attach_dir, f)),
            })
    return jsonify({'attachments': files})


@app.route('/api/exam/<int:exam_id>/attachments/<path:filename>', methods=['DELETE'])
@teacher_required
def delete_exam_attachment(exam_id, filename):
    """Delete a specific attachment file."""
    db = get_db()
    row = db.execute("SELECT exam_id FROM exams WHERE exam_id = ?", (exam_id,)).fetchone()
    if not row:
        return jsonify({'error': '考试不存在'}), 404

    username = session.get('teacher_username')
    if username:
        title_row = db.execute("SELECT title FROM exams WHERE exam_id = ?", (exam_id,)).fetchone()
        title = title_row['title'] if title_row else f'exam_{exam_id}'
        subdir = _exam_subdir(username, title)
        attach_dir = os.path.join(subdir, 'attachments')
    else:
        attach_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'attachments')

    filepath = os.path.join(attach_dir, filename)
    # Security: prevent path traversal
    if '..' in filename or not os.path.exists(filepath) or not os.path.isfile(filepath):
        return jsonify({'error': '文件不存在'}), 404

    try:
        os.remove(filepath)
        return jsonify({'ok': True, 'message': '附件已删除'})
    except Exception as e:
        return jsonify({'error': f'删除失败: {str(e)}'}), 500


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
    ws['D1'] = '专业'
    ws['E1'] = '年级'
    ws['F1'] = '学校'
    header_font = Font(bold=True)
    for cell in ws[1]:
        cell.font = header_font

    ws.cell(row=2, column=1, value='2024001')
    ws.cell(row=2, column=2, value='张三')
    ws.cell(row=2, column=3, value='1班')
    ws.cell(row=2, column=4, value='经济学')
    ws.cell(row=2, column=5, value='2024级')
    ws.cell(row=2, column=6, value='示例学校')

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
    username = (session.get('teacher_username') or '').strip()
    count = 0
    for row in rows:
        student_id = str(row[0]).strip() if row[0] else ''
        name = str(row[1]).strip() if row[1] else ''
        if not student_id or not name:
            continue
        db.execute(
            "INSERT OR REPLACE INTO students (student_id, name, class_name, grade, school, major, teacher_username) VALUES (?,?,?,?,?,?,?)",
            (student_id, name,
             str(row[2]).strip() if row[2] else None,
             str(row[3]).strip() if row[3] else None,
             str(row[4]).strip() if row[4] else None,
             str(row[5]).strip() if len(row) > 5 and row[5] else None,
             username)
        )
        count += 1
    db.commit()
    return jsonify({'ok': True, 'count': count})

# ─── Student Management ───

@app.route('/api/teacher/students', methods=['GET'])
@teacher_required
def list_students():
    """List all students belonging to the current teacher."""
    db = get_db()
    username = (session.get('teacher_username') or '').strip()
    search = request.args.get('search', '').strip()

    if search:
        rows = db.execute(
            "SELECT student_id, name, class_name, grade, school, major FROM students "
            "WHERE teacher_username = ? AND (student_id LIKE ? OR name LIKE ? OR class_name LIKE ?) "
            "ORDER BY student_id",
            (username, f'%{search}%', f'%{search}%', f'%{search}%')
        ).fetchall()
    else:
        rows = db.execute(
            "SELECT student_id, name, class_name, grade, school, major FROM students "
            "WHERE teacher_username = ? ORDER BY student_id",
            (username,)
        ).fetchall()
    return jsonify([dict(r) for r in rows])


@app.route('/api/teacher/students/paged', methods=['GET'])
@teacher_required
def list_students_paged():
    """List students with pagination and search."""
    db = get_db()
    username = (session.get('teacher_username') or '').strip()

    page = request.args.get('page', 1, type=int)
    per_page = request.args.get('per_page', 25, type=int)
    search = request.args.get('search', '').strip()

    # Clamp per_page
    per_page = max(10, min(per_page, 100))
    offset = (page - 1) * per_page

    if search:
        like = f'%{search}%'
        count_row = db.execute(
            "SELECT COUNT(*) as cnt FROM students WHERE teacher_username = ? AND "
            "(student_id LIKE ? OR name LIKE ? OR class_name LIKE ?)",
            (username, like, like, like)
        ).fetchone()
        total = count_row['cnt']
        rows = db.execute(
            "SELECT student_id, name, class_name, grade, school, major FROM students "
            "WHERE teacher_username = ? AND (student_id LIKE ? OR name LIKE ? OR class_name LIKE ?) "
            "ORDER BY student_id LIMIT ? OFFSET ?",
            (username, like, like, like, per_page, offset)
        ).fetchall()
    else:
        count_row = db.execute(
            "SELECT COUNT(*) as cnt FROM students WHERE teacher_username = ?",
            (username,)
        ).fetchone()
        total = count_row['cnt']
        rows = db.execute(
            "SELECT student_id, name, class_name, grade, school, major FROM students "
            "WHERE teacher_username = ? ORDER BY student_id LIMIT ? OFFSET ?",
            (username, per_page, offset)
        ).fetchall()

    return jsonify({
        'items': [dict(r) for r in rows],
        'total': total,
        'page': page,
        'per_page': per_page
    })


@app.route('/api/teacher/student', methods=['POST'])
@teacher_required
def add_student():
    """Add a single student to the teacher's roster."""
    data = request.get_json()
    student_id = data.get('student_id', '').strip()
    name = data.get('name', '').strip()
    if not student_id or not name:
        return jsonify({'error': '学号和姓名不能为空'}), 400

    db = get_db()
    username = (session.get('teacher_username') or '').strip()

    # Check if already exists for this teacher
    existing = db.execute(
        "SELECT student_id FROM students WHERE student_id = ? AND teacher_username = ?",
        (student_id, username)
    ).fetchone()
    if existing:
        return jsonify({'error': '该学号已存在'}), 409

    db.execute(
        "INSERT INTO students (student_id, name, class_name, grade, school, major, teacher_username) VALUES (?,?,?,?,?,?,?)",
        (student_id, name,
         data.get('class_name', '').strip() or None,
         data.get('grade', '').strip() or None,
         data.get('school', '').strip() or None,
         data.get('major', '').strip() or None,
         username)
    )
    db.commit()
    return jsonify({'ok': True})


@app.route('/api/teacher/student/<student_id>', methods=['DELETE'])
@teacher_required
def delete_student(student_id):
    """Delete a student from the teacher's roster. Does not cascade to exam data."""
    db = get_db()
    username = (session.get('teacher_username') or '').strip()
    db.execute(
        "DELETE FROM students WHERE student_id = ? AND teacher_username = ?",
        (student_id, username)
    )
    db.commit()
    return jsonify({'ok': True})

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
        d['roster_count'] = db.execute(
            "SELECT COUNT(*) FROM exam_roster WHERE exam_id = ?",
            (r['exam_id'],)
        ).fetchone()[0]
        exams.append(d)
    return jsonify(exams)

@app.route('/api/teacher/results', methods=['GET'])
@teacher_required
def list_results():
    db = get_db()
    username = session.get('teacher_username')
    now = datetime.now()

    # v3.1: Filter results by teacher ownership
    if username:
        rows = db.execute("""
            SELECT s.session_id, s.exam_id, e.title, s.student_id, s.student_name,
                   s.start_time, s.end_time, s.status, s.total_score, s.is_graded,
                   e.exam_start_time, e.exam_duration
            FROM exam_sessions s
            JOIN exams e ON s.exam_id = e.exam_id
            WHERE e.teacher_username = ?
            ORDER BY s.start_time DESC
        """, (username,)).fetchall()
    else:
        rows = db.execute("""
            SELECT s.session_id, s.exam_id, e.title, s.student_id, s.student_name,
                   s.start_time, s.end_time, s.status, s.total_score, s.is_graded,
                   e.exam_start_time, e.exam_duration
            FROM exam_sessions s
            JOIN exams e ON s.exam_id = e.exam_id
            WHERE e.teacher_username IS NULL
            ORDER BY s.start_time DESC
        """).fetchall()

    results = []
    for r in rows:
        d = dict(r)
        # Auto-expire sessions whose exam end time has passed
        if d['status'] == 'in_progress' and d['exam_start_time'] and d['exam_duration']:
            try:
                start_dt = datetime.strptime(d['exam_start_time'], '%Y-%m-%d %H:%M')
                end_dt = start_dt + timedelta(minutes=int(d['exam_duration']))
                if now >= end_dt:
                    now_str = now.strftime('%Y-%m-%d %H:%M:%S')
                    db.execute(
                        "UPDATE exam_sessions SET status = 'submitted', end_time = ? WHERE session_id = ?",
                        (now_str, d['session_id'])
                    )
                    d['status'] = 'submitted'
                    d['end_time'] = now_str
                    try:
                        _cleanup_session_workdir(d['exam_id'], d['session_id'])
                    except Exception:
                        pass
            except (ValueError, TypeError):
                pass
        results.append(d)
    db.commit()
    return jsonify(results)

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
    """v3.1: Get teacher profile info (username, school, exam count).
    Also provides the data needed by teacher-profile.html."""
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
            'ok': True,
            'teacher_id': row['username'] if row else tid,
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
            'ok': True,
            'teacher_id': tid,
            'is_registered': False,
            'username': tid,
            'school': '',
            'exam_count': exam_count,
            'max_exams': MAX_EXAMS_PER_TEACHER,
            'exams': []
        })


@app.route('/api/teacher/profile', methods=['POST'])
@teacher_required
def update_teacher_profile():
    """Update teacher profile (username, school, password)."""
    data = request.get_json()
    db = get_db()
    tid = session.get('teacher_id')
    username = session.get('teacher_username')

    new_username = data.get('username', '').strip()
    new_school = data.get('school', '').strip()
    new_password = data.get('new_password', '').strip()

    if username:
        # Registered user
        updates = []
        params = []

        if new_username and new_username != username:
            # Check duplicate
            existing = db.execute("SELECT id FROM teachers WHERE username = ? AND username != ?", (new_username, username)).fetchone()
            if existing:
                return jsonify({'error': '用户名已存在'}), 409
            updates.append("username = ?")
            params.append(new_username)
            session['teacher_username'] = new_username

        if new_school:
            updates.append("school = ?")
            params.append(new_school)

        if new_password:
            hashed, salt = _hash_password(new_password)
            stored_pw = f"sha256:{salt}:{hashed}"
            updates.append("password = ?")
            params.append(stored_pw)

        if updates:
            sql = "UPDATE teachers SET " + ", ".join(updates) + " WHERE username = ?"
            params.append(username)
            db.execute(sql, params)
            db.commit()

        return jsonify({'ok': True, 'message': '个人信息已更新'})

    # Hardcoded teacher 001 cannot modify (no registered account)
    return jsonify({'error': '默认管理员账号不支持修改，请先注册独立账号'}), 400

@app.route('/api/teacher/exam/<int:exam_id>', methods=['DELETE'])
@teacher_required
def delete_exam(exam_id):
    """Delete an exam and its associated user folder."""
    db = get_db()

    # Get exam info
    exam = db.execute("SELECT title, teacher_username FROM exams WHERE exam_id = ?", (exam_id,)).fetchone()
    if not exam:
        return jsonify({'error': '试卷不存在'}), 404

    username = exam['teacher_username']

    # Auth check: only owner teacher can delete
    sess_username = session.get('teacher_username')
    if username and sess_username and str(username).strip() != str(sess_username).strip():
        return jsonify({'error': '无权删除此试卷'}), 403

    # Delete associated user folder
    if username and exam['title']:
        _delete_exam_subdir(username, exam['title'])

    # Delete exam (cascade will handle questions, answers, sessions)
    db.execute("DELETE FROM exam_attachments WHERE q_id IN (SELECT q_id FROM questions WHERE exam_id = ?)", (exam_id,))
    db.execute("DELETE FROM answers WHERE q_id IN (SELECT q_id FROM questions WHERE exam_id = ?)", (exam_id,))
    db.execute("DELETE FROM questions WHERE exam_id = ?", (exam_id,))
    db.execute("DELETE FROM exam_roster WHERE exam_id = ?", (exam_id,))
    db.execute("DELETE FROM exam_sessions WHERE exam_id = ?", (exam_id,))
    db.execute("DELETE FROM exams WHERE exam_id = ?", (exam_id,))
    db.commit()

    return jsonify({'ok': True, 'message': f'试卷 "{exam["title"]}" 已删除'})

@app.route('/api/teacher/exam/<int:exam_id>/delete', methods=['DELETE'])
@teacher_required
def delete_exam_by_path(exam_id):
    """v4.0.3: Delete exam with full directory cleanup. Only owner teacher can delete."""
    db = get_db()

    exam = db.execute("SELECT title, teacher_username FROM exams WHERE exam_id = ?", (exam_id,)).fetchone()
    if not exam:
        return jsonify({'error': '试卷不存在'}), 404

    username = exam['teacher_username']

    # Auth check: only the owning teacher can delete
    sess_username = session.get('teacher_username')
    if username and sess_username and str(username).strip() != str(sess_username).strip():
        return jsonify({'error': '无权删除此试卷'}), 403

    # Delete exam directory under Users/<username>/<exam_title>/
    if username and exam['title']:
        _delete_exam_subdir(username, exam['title'])

    # Delete DB records
    db.execute("DELETE FROM exam_attachments WHERE q_id IN (SELECT q_id FROM questions WHERE exam_id = ?)", (exam_id,))
    db.execute("DELETE FROM answers WHERE q_id IN (SELECT q_id FROM questions WHERE exam_id = ?)", (exam_id,))
    db.execute("DELETE FROM questions WHERE exam_id = ?", (exam_id,))
    db.execute("DELETE FROM exam_roster WHERE exam_id = ?", (exam_id,))
    db.execute("DELETE FROM exam_sessions WHERE exam_id = ?", (exam_id,))
    db.execute("DELETE FROM exams WHERE exam_id = ?", (exam_id,))
    db.commit()

    return jsonify({'ok': True, 'message': f'试卷 "{exam["title"]}" 及所有相关文件已删除'})

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

    # Look up exam by exam_number (includes teacher_username for student filtering)
    exam = db.execute(
        "SELECT exam_id, title, template_uploaded, exam_duration, exam_start_time, teacher_username "
        "FROM exams WHERE exam_number = ?",
        (exam_number,)
    ).fetchone()
    if not exam:
        return jsonify({'error': '试卷编号不存在，请确认后重试'}), 404
    if not exam['template_uploaded']:
        return jsonify({'error': '该试卷尚未完成设置'}), 400

    # Verify student exists AND belongs to the exam's teacher
    teacher = exam['teacher_username'] or '001'
    row = db.execute(
        "SELECT student_id, name FROM students WHERE student_id = ? AND name = ? AND teacher_username = ?",
        (sid, name, teacher)
    ).fetchone()
    if not row:
        return jsonify({'error': '学号或姓名不匹配'}), 401

    # v6.0: Check if exam has a roster; if so, student must be on it
    roster_check = db.execute(
        "SELECT 1 FROM exam_roster WHERE exam_id = ? AND student_id = ?",
        (exam['exam_id'], sid)
    ).fetchone()
    roster_exists = db.execute(
        "SELECT COUNT(*) as cnt FROM exam_roster WHERE exam_id = ?",
        (exam['exam_id'],)
    ).fetchone()['cnt']
    if roster_exists > 0 and not roster_check:
        return jsonify({'error': '你未被邀请参加本场考试，请联系教师'}), 403

    return jsonify({
        'ok': True,
        'student_id': sid,
        'name': name,
        'exam_id': exam['exam_id'],
        'exam_title': exam['title'],
        'exam_number': exam_number,
        'requires_stata': False,
        'exam_start_time': exam['exam_start_time']
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
        "anti_shuffle_questions, anti_shuffle_options, anti_screen_switch, "
        "ban_copy, ban_screenshot, "
        "teacher_username "
        "FROM exams WHERE exam_id = ?",
        (exam_id,)
    ).fetchone()
    if not exam_info:
        return jsonify({'error': '考试不存在'}), 404

    # v3.0: Verify teacher has configured exam settings
    if not exam_info['exam_duration'] or not exam_info['exam_start_time']:
        return jsonify({'error': '教师尚未完成考试设置，无法开始考试'}), 400

    # v4.0.5: Check exam start time
    now = datetime.now()
    start_time_str = exam_info['exam_start_time']
    try:
        exam_start_dt = datetime.strptime(start_time_str, '%Y-%m-%d %H:%M')
        if now < exam_start_dt:
            return jsonify({'error': f'考试尚未开始，开考时间为 {start_time_str}'}), 403
        # v5.0: Late start limit — reject if past deadline
        late_start_limit = exam_info['exam_late_start_limit']
        if late_start_limit and late_start_limit > 0:
            late_deadline = exam_start_dt + timedelta(minutes=int(late_start_limit))
            if now > late_deadline:
                # v6.0: Check if teacher has bypassed late start limit for this student
                bypass = db.execute(
                    "SELECT bypass_late_start FROM exam_roster WHERE exam_id = ? AND student_id = ?",
                    (exam_id, student_id)
                ).fetchone()
                if not bypass or not bypass['bypass_late_start']:
                    return jsonify({'error': f'已超过开考限时（开考后{late_start_limit}分钟内允许进入），不允许参加考试'}), 403
    except ValueError:
        # If time format is incompatible, try flexible parse
        pass

    # v4.0.6: Check if student already submitted this exam
    already_submitted = db.execute(
        "SELECT session_id FROM exam_sessions WHERE exam_id = ? AND student_id = ? AND status = 'submitted'",
        (exam_id, student_id)
    ).fetchone()
    if already_submitted:
        return jsonify({'error': '你已提交过该试卷，不可重复参加'}), 403

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

    # v4.3: Always deploy data files to per-session workdir (even for resumed sessions)
    _deploy_exam_data(exam_id, session_id)

    # Return questions WITH v3.0 settings
    rows = db.execute(
        "SELECT q_id, q_type, content, option_a, option_b, option_c, option_d, score, sort_order, "
        "enable_stata, enable_ai, has_attachment, remark, data_file "
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

    # v4.2: Collect data files referenced by questions
    data_files = set()
    for r in rows:
        df = r['data_file']
        if df:
            for f in str(df).replace('；', ';').split(';'):
                f = f.strip()
                if f:
                    data_files.add(f)

    # Resolve teacher username and session workdir
    teacher_username = exam_info['teacher_username'] or ''
    workdir_hint = ''
    if teacher_username:
        workdir_hint = _session_workdir(teacher_username, exam_id, session_id)

    # v5.0: Deterministic question + option shuffle per student
    anti_shuffle_qs = bool(exam_info['anti_shuffle_questions'])
    anti_shuffle_opts = bool(exam_info['anti_shuffle_options'])
    shuffled_questions = _shuffle_questions(rows, student_id, exam_id, anti_shuffle_qs, anti_shuffle_opts)

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
        'anti_shuffle_options': anti_shuffle_opts,
        'anti_screen_switch': bool(exam_info['anti_screen_switch']),
        'ban_copy': bool(exam_info['ban_copy']),
        'ban_screenshot': bool(exam_info['ban_screenshot']),
        'questions': shuffled_questions,
        'saved_answers': saved_answers,
        'attachments': [dict(a) for a in attachments],
        'data_files': list(data_files),
        'workdir_hint': workdir_hint
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

    # v5.0: Check late submit limit (earliest allowed submit time)
    exam_info = db.execute(
        "SELECT exam_start_time, exam_late_submit_limit FROM exams WHERE exam_id = ?",
        (srow['exam_id'],)
    ).fetchone()
    if exam_info and exam_info['exam_start_time']:
        late_submit = exam_info['exam_late_submit_limit']
        if late_submit and late_submit > 0:
            try:
                exam_start_dt = datetime.strptime(exam_info['exam_start_time'], '%Y-%m-%d %H:%M')
                earliest_submit = exam_start_dt + timedelta(minutes=int(late_submit))
                if datetime.now() < earliest_submit:
                    remaining = int((earliest_submit - datetime.now()).total_seconds())
                    mins = remaining // 60
                    return jsonify({'error': f'距离可提交还有 {mins} 分钟，不允许提前交卷'}), 403
            except ValueError:
                pass

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
    questions_raw = db.execute(
        "SELECT q_id, q_type, content, option_a, option_b, option_c, option_d, "
        "correct_answer, score FROM questions WHERE exam_id = ?", (srow['exam_id'],)
    ).fetchall()

    # v5.0: Check question + option shuffle; if enabled, shuffle before grading
    exam_cfg = db.execute(
        "SELECT anti_shuffle_questions, anti_shuffle_options FROM exams WHERE exam_id = ?", (srow['exam_id'],)
    ).fetchone()
    shuffle_qs = bool(exam_cfg['anti_shuffle_questions']) if exam_cfg else False
    shuffle_opts = bool(exam_cfg['anti_shuffle_options']) if exam_cfg else False
    questions = _shuffle_questions(questions_raw, srow['student_id'], srow['exam_id'], shuffle_qs, shuffle_opts)

    # Build answers dict for grading
    answers_dict = {}
    if isinstance(answers, dict):
        answers_dict = answers
    elif isinstance(answers, list):
        answers_dict = {str(a.get('q_id')): a.get('answer', '') for a in answers if 'q_id' in a}

    total_score = 0
    for q in questions:
        ans_raw = answers_dict.get(str(q['q_id']), '')
        if ans_raw is None:
            ans_raw = ''
        ans_raw = str(ans_raw).strip()

        # v4.5: Parse JSON answer to extract real answer text (ignore stata fields)
        ans = ans_raw
        try:
            parsed = json.loads(ans_raw)
            if isinstance(parsed, dict):
                ans = parsed.get('choice') or parsed.get('text') or ans_raw
        except (json.JSONDecodeError, TypeError):
            pass
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
            all_q_raw = db.execute(
                "SELECT q_id, q_type, content, option_a, option_b, option_c, option_d, "
                "correct_answer, score, sort_order FROM questions WHERE exam_id = ? ORDER BY sort_order",
                (srow['exam_id'],)
            ).fetchall()

            # v5.0: Shuffle questions + options for xlsx to match what student saw
            exam_cfg = db.execute(
                "SELECT anti_shuffle_questions, anti_shuffle_options FROM exams WHERE exam_id = ?",
                (srow['exam_id'],)
            ).fetchone()
            shuffle_qs = bool(exam_cfg['anti_shuffle_questions']) if exam_cfg else False
            shuffle_opts = bool(exam_cfg['anti_shuffle_options']) if exam_cfg else False
            all_q = _shuffle_questions(all_q_raw, srow['student_id'], srow['exam_id'], shuffle_qs, shuffle_opts)

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

    # v4.3: Clean up this student's per-session Stata workdir after submit
    try:
        _cleanup_session_workdir(srow['exam_id'], session_id)
    except Exception as e:
        print(f"[v4.3] Failed to cleanup session workdir: {e}")

    # v4.2: Lazy cleanup — if all students have submitted, remove entire exam workdir
    try:
        total_sessions = db.execute(
            "SELECT COUNT(*) as cnt FROM exam_sessions WHERE exam_id = ?",
            (srow['exam_id'],)
        ).fetchone()['cnt']
        submitted_sessions = db.execute(
            "SELECT COUNT(*) as cnt FROM exam_sessions WHERE exam_id = ? AND status = 'submitted'",
            (srow['exam_id'],)
        ).fetchone()['cnt']
        if total_sessions == submitted_sessions:
            _cleanup_exam_workdir(srow['exam_id'])
    except Exception as e:
        print(f"[v4.2] Failed to cleanup workdir: {e}")

    return jsonify({'ok': True, 'total_score': total_score, 'is_graded': True})

# ════════════════════════════════════════════
# Monitor APIs (v2.0)
# ════════════════════════════════════════════

@app.route('/api/teacher/monitor', methods=['GET'])
@teacher_required
def monitor_sessions():
    """获取所有在线考试的学生状态"""
    db = get_db()
    now = datetime.now()
    rows = db.execute("""
        SELECT s.session_id, s.exam_id, s.student_id, s.student_name,
               s.start_time, s.status, s.total_score,
               s.terminated_by_teacher, s.exam_end_time,
               e.title, e.duration_minutes,
               e.exam_start_time, e.exam_duration
        FROM exam_sessions s
        JOIN exams e ON s.exam_id = e.exam_id
        WHERE s.status IN ('in_progress', 'submitted')
        ORDER BY s.start_time DESC
    """).fetchall()
    results = []
    for r in rows:
        d = dict(r)
        # Auto-expire sessions whose exam end time has passed
        if d['status'] == 'in_progress' and d['exam_start_time'] and d['exam_duration']:
            try:
                start_dt = datetime.strptime(d['exam_start_time'], '%Y-%m-%d %H:%M')
                end_dt = start_dt + timedelta(minutes=int(d['exam_duration']))
                if now >= end_dt:
                    now_str = now.strftime('%Y-%m-%d %H:%M:%S')
                    db.execute(
                        "UPDATE exam_sessions SET status = 'submitted', end_time = ? WHERE session_id = ?",
                        (now_str, d['session_id'])
                    )
                    d['status'] = 'submitted'
                    d['end_time'] = now_str
                    try:
                        _cleanup_session_workdir(d['exam_id'], d['session_id'])
                    except Exception:
                        pass
            except (ValueError, TypeError):
                pass
        d['is_terminated'] = bool(d['terminated_by_teacher'])
        results.append(d)
    db.commit()
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
def resume_session_old(session_id):
    """教师恢复学生考试 (redirect to new endpoint)"""
    return resume_exam_session(session_id)

# ════════════════════════════════════════════
# Per-Exam Monitor API (v6.0)
# ════════════════════════════════════════════

@app.route('/api/teacher/exam/<int:exam_id>/monitor', methods=['GET'])
@teacher_required
def monitor_exam_students(exam_id):
    """获取某场考试所有考生的状态（包括尚未进入的考生）"""
    db = get_db()
    now = datetime.now()

    # Get exam info
    exam = db.execute(
        "SELECT exam_id, title, exam_start_time, exam_duration, exam_late_start_limit, exam_number "
        "FROM exams WHERE exam_id = ?",
        (exam_id,)
    ).fetchone()
    if not exam:
        return jsonify({'error': '考试不存在'}), 404

    exam_start_time_str = exam['exam_start_time']
    exam_duration = exam['exam_duration'] or 0
    late_start_limit = exam['exam_late_start_limit'] or 0

    # Parse exam times
    exam_start_dt = None
    exam_end_dt = None
    late_deadline = None
    if exam_start_time_str:
        try:
            exam_start_dt = datetime.strptime(exam_start_time_str, '%Y-%m-%d %H:%M')
            if exam_duration > 0:
                exam_end_dt = exam_start_dt + timedelta(minutes=int(exam_duration))
            if late_start_limit > 0:
                late_deadline = exam_start_dt + timedelta(minutes=int(late_start_limit))
        except ValueError:
            pass

    # Get all roster students + their sessions (left join)
    rows = db.execute("""
        SELECT
            r.student_id,
            r.bypass_late_start,
            s.student_id as has_session,
            s.student_name,
            s.session_id,
            s.start_time,
            s.status,
            s.terminated_by_teacher,
            s.exam_end_time
        FROM exam_roster r
        LEFT JOIN (
            SELECT ss.*
            FROM exam_sessions ss
            WHERE ss.exam_id = ?
        ) s ON r.student_id = s.student_id
        ORDER BY r.student_id
    """, (exam_id,)).fetchall()

    results = []
    for r in rows:
        sid = r['student_id']

        # Determine status
        has_session = r['has_session'] is not None
        is_terminated = bool(r['terminated_by_teacher']) if has_session else False
        session_status = r['status'] if has_session else None

        # Auto-expire sessions whose exam end time has passed
        if has_session and session_status == 'in_progress' and exam_end_dt and now >= exam_end_dt:
            now_str = now.strftime('%Y-%m-%d %H:%M:%S')
            db.execute(
                "UPDATE exam_sessions SET status = 'submitted', end_time = ? WHERE session_id = ?",
                (now_str, r['session_id'])
            )
            session_status = 'submitted'
            # Try to clean up
            try:
                _cleanup_session_workdir(exam_id, r['session_id'])
            except Exception:
                pass

        if not has_session:
            # No session yet — determine if "待进入" or "已禁止"
            if late_deadline and now > late_deadline and not r['bypass_late_start']:
                status = 'banned'  # 已禁止
            else:
                status = 'pending'  # 待进入
        elif is_terminated:
            status = 'terminated'  # 已终止
        elif session_status == 'submitted':
            status = 'submitted'  # 已提交
        else:
            status = 'in_progress'  # 考试中

        student_name = r['student_name'] or ''
        # Look up name from students table if not in session
        if not student_name:
            name_row = db.execute(
                "SELECT name FROM students WHERE student_id = ?",
                (sid,)
            ).fetchone()
            if name_row:
                student_name = name_row['name']

        d = {
            'student_id': sid,
            'student_name': student_name,
            'session_id': r['session_id'],
            'status': status,
            'start_time': r['start_time'],
            'bypass_late_start': bool(r['bypass_late_start']),
            'exam_title': exam['title'],
            'exam_number': exam['exam_number'],
            'exam_start_time': exam_start_time_str,
            'exam_end_time': exam_end_dt.strftime('%Y-%m-%d %H:%M:%S') if exam_end_dt else None,
            'late_start_limit': late_start_limit
        }
        results.append(d)

    db.commit()
    return jsonify(results)


@app.route('/api/teacher/exam/<int:exam_id>/student/<student_id>/allow-entry', methods=['POST'])
@teacher_required
def allow_student_entry(exam_id, student_id):
    """Teacher allows a late-start-blocked student to enter the exam."""
    db = get_db()
    exam = db.execute("SELECT exam_id, exam_start_time, exam_duration FROM exams WHERE exam_id = ?", (exam_id,)).fetchone()
    if not exam:
        return jsonify({'error': '考试不存在'}), 404

    # Check if exam has ended
    if exam['exam_start_time'] and exam['exam_duration']:
        try:
            start_dt = datetime.strptime(exam['exam_start_time'], '%Y-%m-%d %H:%M')
            end_dt = start_dt + timedelta(minutes=int(exam['exam_duration']))
            if datetime.now() > end_dt:
                return jsonify({'error': '考试已到截止时间，无法允许进入'}), 400
        except ValueError:
            pass

    # Update exam_roster bypass flag
    db.execute(
        "UPDATE exam_roster SET bypass_late_start = 1 WHERE exam_id = ? AND student_id = ?",
        (exam_id, student_id)
    )
    db.commit()
    return jsonify({'ok': True, 'message': '已允许该生进入考试'})


@app.route('/api/teacher/exam/<int:exam_id>/student/<student_id>/reset', methods=['POST'])
@teacher_required
def reset_student_exam(exam_id, student_id):
    """Teacher resets a submitted student's exam session, keeping the session but
    resetting status back to in_progress and unsetting is_terminated.
    Student can then resume with previous answers visible.
    """
    db = get_db()
    exam = db.execute("SELECT exam_id, exam_start_time, exam_duration FROM exams WHERE exam_id = ?", (exam_id,)).fetchone()
    if not exam:
        return jsonify({'error': '考试不存在'}), 404

    # Check if exam has ended
    if exam['exam_start_time'] and exam['exam_duration']:
        try:
            start_dt = datetime.strptime(exam['exam_start_time'], '%Y-%m-%d %H:%M')
            end_dt = start_dt + timedelta(minutes=int(exam['exam_duration']))
            if datetime.now() > end_dt:
                return jsonify({'error': '考试已到截止时间，无法重置'}), 400
        except ValueError:
            pass

    # Find the student's session
    session_row = db.execute(
        "SELECT session_id, status FROM exam_sessions WHERE exam_id = ? AND student_id = ?",
        (exam_id, student_id)
    ).fetchone()
    if not session_row:
        return jsonify({'error': '该生没有考试记录'}), 404

    # Reset: set status to in_progress, unset terminated
    db.execute(
        "UPDATE exam_sessions SET status = 'in_progress', terminated_by_teacher = 0, exam_end_time = NULL WHERE session_id = ?",
        (session_row['session_id'],)
    )
    db.commit()
    return jsonify({'ok': True, 'message': '已重置该生考试，可重新进入'})


# ════════════════════════════════════════════
# Exam Settings APIs (v3.0)
# ════════════════════════════════════════════

@app.route('/api/exam/<int:exam_id>/settings', methods=['GET'])
@teacher_required
def get_exam_settings(exam_id):
    """Get exam settings including cover info fields."""
    db = get_db()
    row = db.execute("""
        SELECT exam_id, title, exam_title, school, college, class_name, grade,
               exam_duration, exam_start_time, exam_notice, exam_notice_title,
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
    """Save exam settings including cover info fields."""
    data = request.get_json()
    db = get_db()
    row = db.execute("SELECT exam_id FROM exams WHERE exam_id = ?", (exam_id,)).fetchone()
    if not row:
        return jsonify({'error': '考试不存在'}), 404

    exam_title = data.get('exam_title', '')
    school = data.get('school', '')
    college = data.get('college', '')
    class_name = data.get('class_name', '')
    grade = data.get('grade', '')
    major = data.get('major', '')
    exam_duration = data.get('exam_duration', 60)
    exam_start_time = data.get('exam_start_time', '')
    exam_notice = data.get('exam_notice', '')
    ban_copy = data.get('ban_copy', 0)
    ban_screenshot = data.get('ban_screenshot', 0)
    exam_notice_title = data.get('exam_notice_title', '')
    exam_late_start_limit = data.get('exam_late_start_limit', 0)
    exam_late_submit_limit = data.get('exam_late_submit_limit', 0)
    anti_shuffle_questions = data.get('anti_shuffle_questions', 0)
    anti_shuffle_options = data.get('anti_shuffle_options', 0)
    anti_screen_switch = data.get('anti_screen_switch', 0)

    db.execute("""
        UPDATE exams SET
            exam_title = ?, school = ?, college = ?, class_name = ?, grade = ?, major = ?,
            exam_duration = ?, exam_start_time = ?, exam_notice = ?, exam_notice_title = ?,
            exam_late_start_limit = ?, exam_late_submit_limit = ?,
            anti_shuffle_questions = ?, anti_shuffle_options = ?, anti_screen_switch = ?,
            ban_copy = ?, ban_screenshot = ?
        WHERE exam_id = ?
    """, (
        exam_title, school, college, class_name, grade, major,
        int(exam_duration), exam_start_time, exam_notice, exam_notice_title,
        int(exam_late_start_limit), int(exam_late_submit_limit),
        int(anti_shuffle_questions), int(anti_shuffle_options), int(anti_screen_switch),
        int(ban_copy), int(ban_screenshot),
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

@app.route('/api/exam/<int:exam_id>/attachments/file/<filename>')
def serve_attachment_file(exam_id, filename):
    """Serve an attachment file for display in exam pages (images, etc.)."""
    db = get_db()
    exam = db.execute(
        "SELECT title, teacher_username FROM exams WHERE exam_id = ?",
        (exam_id,)
    ).fetchone()
    if not exam:
        return jsonify({'error': '考试不存在'}), 404

    # Look in exam attachment directory
    if exam['teacher_username']:
        attach_dir = _get_attachment_dir(exam_id)
    else:
        attach_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'attachments')

    # Try exact filename
    filepath = os.path.join(attach_dir, filename)
    if os.path.exists(filepath) and os.path.isfile(filepath):
        return send_file(filepath)

    # Try with timestamp prefix: search for files ending with _filename
    if os.path.exists(attach_dir):
        for f in os.listdir(attach_dir):
            if f.endswith('_' + filename) or f.endswith(filename):
                return send_file(os.path.join(attach_dir, f))

    # Also check Users/<teacher>/<exam_title>/attachments/
    if exam['teacher_username']:
        edir = _exam_subdir(exam['teacher_username'], exam['title'])
        attach_dir2 = os.path.join(edir, 'attachments')
        if os.path.exists(attach_dir2):
            filepath2 = os.path.join(attach_dir2, filename)
            if os.path.exists(filepath2) and os.path.isfile(filepath2):
                return send_file(filepath2)
            for f in os.listdir(attach_dir2):
                if f.endswith('_' + filename) or f.endswith(filename):
                    return send_file(os.path.join(attach_dir2, f))

    return jsonify({'error': '文件不存在'}), 404


@app.route('/api/exam/<int:exam_id>/attachments/student', methods=['GET'])
def get_exam_attachments_student(exam_id):
    """获取考试所有附件（学生端）"""
    db = get_db()
    rows = db.execute("""
        SELECT a.attachment_id, a.q_id, a.original_name, q.has_attachment
        FROM exam_attachments a
        JOIN questions q ON a.q_id = q.q_id
        WHERE q.exam_id = ?
    """, (exam_id,)).fetchall()

    # Enrich with file URLs
    result = []
    username = None
    title = None
    exam_info = db.execute(
        "SELECT title, teacher_username FROM exams WHERE exam_id = ?",
        (exam_id,)
    ).fetchone()
    if exam_info:
        username = exam_info['teacher_username']
        title = exam_info['title']

    for r in rows:
        d = dict(r)
        # Generate file URL for display
        filename = d['original_name']
        d['file_url'] = f'/api/exam/{exam_id}/attachments/file/{filename}'
        result.append(d)

    return jsonify(result)

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
        "SELECT exam_id, title, exam_number, exam_duration, exam_start_time, exam_notice, "
        "template_uploaded, teacher_username FROM exams WHERE exam_number = ?",
        (exam_number,)
    ).fetchone()
    if not exam:
        return jsonify({'error': '试卷编号不存在'}), 404

    d = dict(exam)
    d['requires_stata'] = False
    return jsonify(d)


@app.route('/api/exam/<int:exam_id>', methods=['GET'])
def get_exam_info_for_cover(exam_id):
    """
    v4.0.3: Get exam info for the exam cover page (no auth required).
    Returns cover info fields: exam_title, school, college, class_name, grade, exam_notice_title.
    """
    db = get_db()
    exam = db.execute(
        "SELECT exam_id, title, exam_title, school, college, class_name, grade, major, "
        "exam_duration, exam_start_time, exam_notice, exam_notice_title, "
        "exam_late_start_limit, "
        "template_uploaded FROM exams WHERE exam_id = ?",
        (exam_id,)
    ).fetchone()
    if not exam:
        return jsonify({'error': '考试不存在'}), 404
    return jsonify(dict(exam))


# ─── Sample Preview API ───

@app.route('/api/exam/<int:exam_id>/preview', methods=['GET'])
def get_exam_preview(exam_id):
    """Get exam data for the sample paper preview page."""
    db = get_db()
    exam = db.execute(
        "SELECT exam_id, title, exam_title, school, college, class_name, grade, "
        "exam_duration FROM exams WHERE exam_id = ?",
        (exam_id,)
    ).fetchone()
    if not exam:
        return jsonify({'error': '考试不存在'}), 404

    rows = db.execute(
        "SELECT q_id, q_type, content, option_a, option_b, option_c, option_d, "
        "score, sort_order, enable_stata, enable_ai, has_attachment, remark "
        "FROM questions WHERE exam_id = ? ORDER BY sort_order",
        (exam_id,)
    ).fetchall()

    # Get attachments
    attachments = db.execute(
        "SELECT a.attachment_id, a.q_id, a.original_name "
        "FROM exam_attachments a "
        "JOIN questions q ON a.q_id = q.q_id "
        "WHERE q.exam_id = ?",
        (exam_id,)
    ).fetchall()

    return jsonify({
        'exam': dict(exam),
        'questions': [dict(r) for r in rows],
        'attachments': [dict(a) for a in attachments]
    })


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


@app.route('/api/exam/<int:exam_id>/roster', methods=['GET'])
@teacher_required
def get_exam_roster(exam_id):
    """Get the roster of students bound to this exam."""
    db = get_db()
    rows = db.execute(
        "SELECT s.student_id, s.name, s.class_name, s.grade, s.school, s.major "
        "FROM exam_roster r JOIN students s ON r.student_id = s.student_id "
        "WHERE r.exam_id = ? ORDER BY s.student_id",
        (exam_id,)
    ).fetchall()
    return jsonify([dict(r) for r in rows])


@app.route('/api/exam/<int:exam_id>/roster', methods=['POST'])
@teacher_required
def set_exam_roster(exam_id):
    """Set the roster of students for this exam (full replace)."""
    data = request.get_json()
    student_ids = data.get('student_ids', [])
    if not isinstance(student_ids, list):
        return jsonify({'error': '参数错误'}), 400

    db = get_db()
    db.execute("DELETE FROM exam_roster WHERE exam_id = ?", (exam_id,))
    for sid in student_ids:
        db.execute(
            "INSERT OR IGNORE INTO exam_roster (exam_id, student_id) VALUES (?, ?)",
            (exam_id, sid)
        )
    db.commit()
    return jsonify({'ok': True, 'count': len(student_ids)})


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

    # v5.0: Check question + option shuffle settings for this exam
    shuffle_cfg = db.execute(
        "SELECT anti_shuffle_questions, anti_shuffle_options FROM exams WHERE exam_id = ?", (exam_id,)
    ).fetchone()
    shuffle_qs = bool(shuffle_cfg['anti_shuffle_questions']) if shuffle_cfg else False
    shuffle_opts = bool(shuffle_cfg['anti_shuffle_options']) if shuffle_cfg else False

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
        # v5.0: Shuffle questions + options per student (same seed as exam page)
        shuffled = _shuffle_questions(questions, s['student_id'], exam_id, shuffle_qs, shuffle_opts)
        qlist = []
        for q in shuffled:
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
# Stata Execution API (PlanB: server-side)
# ════════════════════════════════════════════

@app.route('/api/stata/execute', methods=['POST'])
def execute_stata_command():
    """Execute Stata command on server and return results.
    Uses stata_mcp for clean, stack-safe execution without IPython dependency.
    Accepts 'exam_id' + 'session_id' to cd to the student's per-session workdir.
    """
    data = request.get_json()
    command = data.get('command', '').strip()
    if not command:
        return jsonify({'error': '请输入Stata命令'}), 400

    # Set working directory to the student's per-session workdir
    exam_id = data.get('exam_id')
    session_id = data.get('session_id')
    workdir = None
    if exam_id and session_id:
        db = get_db()
        exam = db.execute(
            "SELECT teacher_username FROM exams WHERE exam_id = ?",
            (exam_id,)
        ).fetchone()
        if exam and exam['teacher_username']:
            workdir = _session_workdir(exam['teacher_username'], exam_id, int(session_id))
            if not os.path.exists(workdir):
                workdir = None

    # Prepend cd command so Stata finds data files in the student's session workdir
    full_command = command
    if workdir:
        # Use forward slashes for Stata compatibility
        stata_cd = workdir.replace('\\', '/')
        full_command = f'cd "{stata_cd}"\n{command}'

    try:
        log_content = _execute_stata_via_mcp(full_command)
        resp = {
            'ok': True,
            'output': log_content
        }
        return jsonify(resp)
    except Exception as e:
        return jsonify({'error': f'Stata执行失败: {str(e)}'}), 500

# ════════════════════════════════════════════
# AI Chat Proxy API (v4.4)
# ════════════════════════════════════════════

@app.route('/api/ai/chat', methods=['POST'])
def ai_chat_proxy():
    """Proxy AI chat requests to DeepSeek or Qianwen API.
    Students provide their own API key from login page.
    """
    data = request.get_json()
    platform = data.get('platform', '').lower()
    api_key = data.get('api_key', '').strip()
    messages = data.get('messages', [])

    if not api_key:
        return jsonify({'error': '未提供 API Key'}), 400
    if not messages:
        return jsonify({'error': '消息不能为空'}), 400

    # Determine API endpoint based on platform
    if platform == 'deepseek':
        api_url = 'https://api.deepseek.com/v1/chat/completions'
        model = 'deepseek-chat'
    elif platform == 'qianwen':
        api_url = 'https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions'
        model = 'qwen-plus'
    else:
        return jsonify({'error': f'不支持的AI平台: {platform}'}), 400

    try:
        import urllib.request
        payload = json.dumps({
            'model': model,
            'messages': messages,
            'stream': False
        }).encode('utf-8')

        req = urllib.request.Request(api_url, data=payload)
        req.add_header('Content-Type', 'application/json')
        req.add_header('Authorization', f'Bearer {api_key}')
        req.add_header('Accept', 'application/json')

        with urllib.request.urlopen(req, timeout=60) as resp:
            result = json.loads(resp.read().decode('utf-8'))

        if 'choices' in result and len(result['choices']) > 0:
            content = result['choices'][0]['message']['content']
            return jsonify({'ok': True, 'content': content})
        else:
            error_msg = result.get('error', {}).get('message', '未知错误')
            return jsonify({'error': f'AI返回异常: {error_msg}'}), 500

    except urllib.error.HTTPError as e:
        error_body = e.read().decode('utf-8', errors='replace')
        try:
            error_json = json.loads(error_body)
            error_msg = error_json.get('error', {}).get('message', str(e))
        except (ValueError, KeyError):
            error_msg = f'HTTP {e.code}'
        return jsonify({'error': f'AI API错误: {error_msg}'}), 502
    except Exception as e:
        return jsonify({'error': f'AI服务异常: {str(e)}'}), 500


# ════════════════════════════════════════════
# Exam Tab Switch & Check APIs (v4.1)
# ════════════════════════════════════════════

@app.route('/api/exam/<int:exam_id>/session/check', methods=['GET'])
def check_exam_session(exam_id):
    """Check if student has already submitted this exam."""
    student_id = request.args.get('student_id', '')
    if not student_id:
        return jsonify({'error': '缺少学号'}), 400
    db = get_db()
    row = db.execute(
        "SELECT status FROM exam_sessions WHERE exam_id = ? AND student_id = ? AND status = 'submitted'",
        (exam_id, student_id)
    ).fetchone()
    return jsonify({'submitted': row is not None})


@app.route('/api/exam/session/<int:session_id>/tab-switch', methods=['POST'])
def record_tab_switch(session_id):
    """Record a tab switch and auto-submit if >=3 times."""
    data = request.get_json()
    db = get_db()

    # Verify session exists
    session_row = db.execute(
        "SELECT session_id, exam_id, status, tab_switch_count FROM exam_sessions WHERE session_id = ?",
        (session_id,)
    ).fetchone()
    if not session_row:
        return jsonify({'error': '考试会话不存在'}), 404
    if session_row['status'] != 'in_progress':
        return jsonify({'error': '考试已结束'}), 400

    # Increment tab switch count
    count = data.get('count', 1)
    # Update the stored count
    db.execute(
        "UPDATE exam_sessions SET tab_switch_count = ? WHERE session_id = ?",
        (count, session_id)
    )
    db.commit()

    if count >= 3:
        # Auto-submit: mark as submitted with terminated status
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        db.execute(
            "UPDATE exam_sessions SET status = 'submitted', end_time = ?, terminated_by_teacher = 1 WHERE session_id = ?",
            (now, session_id)
        )
        db.commit()
        return jsonify({
            'ok': True,
            'terminated': True,
            'message': '因切屏超过3次，考试已被自动终止并提交'
        })

    return jsonify({
        'ok': True,
        'terminated': False,
        'count': count,
        'message': f'切屏警告（{count}/3）'
    })


@app.route('/api/teacher/exam/session/<int:session_id>/resume', methods=['POST'])
@teacher_required
def resume_exam_session(session_id):
    """Teacher allows a terminated student to resume exam."""
    db = get_db()
    session_row = db.execute(
        "SELECT session_id, exam_id, status, terminated_by_teacher FROM exam_sessions WHERE session_id = ?",
        (session_id,)
    ).fetchone()
    if not session_row:
        return jsonify({'error': '考试会话不存在'}), 404

    # Check if exam has ended
    exam = db.execute(
        "SELECT exam_duration, exam_start_time FROM exams WHERE exam_id = ?",
        (session_row['exam_id'],)
    ).fetchone()

    if exam and exam['exam_start_time']:
        from datetime import datetime
        try:
            start_dt = datetime.strptime(exam['exam_start_time'], '%Y-%m-%d %H:%M')
            duration = int(exam['exam_duration'] or 60)
            end_dt = start_dt + timedelta(minutes=duration)
            if datetime.now() > end_dt:
                return jsonify({'error': '考试已到截止时间，无法恢复'}), 400
        except ValueError:
            pass

    # Resume: reset terminated status and set back to in_progress
    db.execute(
        "UPDATE exam_sessions SET status = 'in_progress', terminated_by_teacher = 0 WHERE session_id = ?",
        (session_id,)
    )
    db.commit()
    return jsonify({'ok': True, 'message': '已同意该生继续考试'})


# ─── Exam Status Check (v3.0) ───

@app.route('/api/exam/session/<int:session_id>/status', methods=['GET'])
def check_exam_status(session_id):
    """学生端检查考试状态（是否被终止/已过期）"""
    db = get_db()
    srow = db.execute(
        "SELECT s.status, s.terminated_by_teacher, s.exam_end_time, s.start_time, "
        "e.exam_start_time, e.exam_duration "
        "FROM exam_sessions s "
        "JOIN exams e ON s.exam_id = e.exam_id "
        "WHERE s.session_id = ?",
        (session_id,)
    ).fetchone()
    if not srow:
        return jsonify({'error': '会话不存在'}), 404

    # Auto-expire if exam end time has passed
    now = datetime.now()
    if srow['status'] == 'in_progress' and srow['exam_start_time'] and srow['exam_duration']:
        try:
            start_dt = datetime.strptime(srow['exam_start_time'], '%Y-%m-%d %H:%M')
            end_dt = start_dt + timedelta(minutes=int(srow['exam_duration']))
            if now >= end_dt:
                now_str = now.strftime('%Y-%m-%d %H:%M:%S')
                db.execute(
                    "UPDATE exam_sessions SET status = 'submitted', end_time = ? WHERE session_id = ?",
                    (now_str, session_id)
                )
                db.commit()
                try:
                    _cleanup_session_workdir(srow['exam_id'], session_id)
                except Exception:
                    pass
                return jsonify({
                    'status': 'submitted',
                    'terminated': bool(srow['terminated_by_teacher']),
                    'exam_end_time': now_str,
                    'auto_expired': True
                })
        except (ValueError, TypeError):
            pass

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
