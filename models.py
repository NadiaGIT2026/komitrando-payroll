# models.py — Database models (SQLite + PostgreSQL support)

import os
from datetime import datetime, date

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATABASE_URL = os.environ.get('DATABASE_URL', '')

# Detect database type
USE_POSTGRES = DATABASE_URL.startswith('postgres')

if USE_POSTGRES:
    import psycopg
    from psycopg.rows import dict_row
    # Fix Render's postgres:// → postgresql://
    if DATABASE_URL.startswith('postgres://'):
        DATABASE_URL = DATABASE_URL.replace('postgres://', 'postgresql://', 1)
else:
    import sqlite3
    DATA_DIR = os.path.join(BASE_DIR, 'data')
    os.makedirs(DATA_DIR, exist_ok=True)
    DB_PATH = os.path.join(DATA_DIR, 'payroll.db')


class DictRow:
    """Make psycopg rows behave like sqlite3.Row."""
    def __init__(self, row_dict):
        self._dict = row_dict
    def __getitem__(self, key):
        if isinstance(key, int):
            return list(self._dict.values())[key]
        return self._dict[key]
    def __contains__(self, key):
        return key in self._dict
    def keys(self):
        return self._dict.keys()
    def get(self, key, default=None):
        return self._dict.get(key, default)


class PgCursorWrapper:
    """Wraps psycopg cursor to return DictRow objects like sqlite3.Row."""
    def __init__(self, cursor):
        self._cur = cursor
    def fetchone(self):
        row = self._cur.fetchone()
        return DictRow(dict(row)) if row else None
    def fetchall(self):
        return [DictRow(dict(r)) for r in self._cur.fetchall()]
    @property
    def lastrowid(self):
        try:
            self._cur.execute("SELECT lastval()")
            return self._cur.fetchone()['lastval']
        except:
            return None
    @property
    def rowcount(self):
        return self._cur.rowcount
    def __iter__(self):
        return iter(self.fetchall())


class PgConnectionWrapper:
    """Wraps psycopg connection to auto-translate SQLite SQL to PostgreSQL."""
    def __init__(self, real_conn):
        self._conn = real_conn
    
    def _translate_sql(self, sql):
        """Convert SQLite-flavored SQL to PostgreSQL."""
        import re
        # ? → %s
        sql = sql.replace('?', '%s')
        # INSERT OR IGNORE → INSERT ... ON CONFLICT DO NOTHING
        sql = re.sub(r'INSERT\s+OR\s+IGNORE\s+INTO', 'INSERT INTO', sql, flags=re.IGNORECASE)
        if 'ON CONFLICT' not in sql.upper() and 'INSERT INTO' in sql.upper():
            # Check if original had OR IGNORE
            pass
        # INSERT OR REPLACE → handled specially
        sql = re.sub(r'INSERT\s+OR\s+REPLACE\s+INTO', 'INSERT INTO', sql, flags=re.IGNORECASE)
        return sql
    
    def execute(self, sql, params=None):
        original_sql = sql
        sql = self._translate_sql(sql)
        
        # Handle INSERT OR IGNORE
        if 'OR IGNORE' in original_sql.upper():
            if 'ON CONFLICT' not in sql.upper():
                if sql.rstrip().endswith(')'):
                    sql += ' ON CONFLICT DO NOTHING'
        
        # Handle INSERT OR REPLACE → use ON CONFLICT with constraint name for PostgreSQL
        if 'OR REPLACE' in original_sql.upper():
            # Match table name precisely: "INTO attendance" vs "INTO payroll"
            sql_lower = original_sql.lower()
            if 'into attendance' in sql_lower:
                sql += ' ON CONFLICT ON CONSTRAINT attendance_employee_id_date_key DO UPDATE SET clock_in=EXCLUDED.clock_in, clock_out=EXCLUDED.clock_out, status=EXCLUDED.status, overtime_hours=EXCLUDED.overtime_hours'
            elif 'into payroll' in sql_lower:
                sql += (' ON CONFLICT ON CONSTRAINT payroll_employee_id_period_key DO UPDATE SET '
                        'work_days=EXCLUDED.work_days, absent_days=EXCLUDED.absent_days, '
                        'base_salary=EXCLUDED.base_salary, gross_salary=EXCLUDED.gross_salary, '
                        'overtime_pay=EXCLUDED.overtime_pay, overtime_total_hours=EXCLUDED.overtime_total_hours, '
                        'total_deductions=EXCLUDED.total_deductions, net_salary=EXCLUDED.net_salary, '
                        'deduction_absence=EXCLUDED.deduction_absence, deduction_late=EXCLUDED.deduction_late, '
                        'bpjs_total_employee=EXCLUDED.bpjs_total_employee, bpjs_total_company=EXCLUDED.bpjs_total_company, '
                        'import_source=EXCLUDED.import_source')
        
        cur = self._conn.cursor(row_factory=dict_row)
        try:
            cur.execute(sql, params or ())
        except Exception as e:
            self._conn.rollback()
            raise e
        return PgCursorWrapper(cur)
    
    def executescript(self, sql):
        """Execute multiple statements."""
        cur = self._conn.cursor()
        cur.execute(sql.encode() if isinstance(sql, str) else sql)
        self._conn.commit()
    
    def commit(self):
        self._conn.commit()
    
    def rollback(self):
        self._conn.rollback()
    
    def close(self):
        self._conn.close()


def get_db():
    if USE_POSTGRES:
        conn = psycopg.connect(DATABASE_URL)
        return PgConnectionWrapper(conn)
    else:
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA foreign_keys=ON")
        return conn


def db_execute(conn, sql, params=None):
    """Execute SQL with compatibility for both SQLite and PostgreSQL."""
    if USE_POSTGRES:
        # Convert ? placeholders to %s for psycopg
        sql = sql.replace('?', '%s')
        # Convert SQLite functions
        sql = sql.replace('CURRENT_TIMESTAMP', 'NOW()')
        cur = conn.cursor(row_factory=dict_row)
        cur.execute(sql, params or ())
        return cur
    else:
        return conn.execute(sql, params or ())


def db_fetchone(conn, sql, params=None):
    """Fetch one row with compatibility."""
    cur = db_execute(conn, sql, params)
    row = cur.fetchone()
    if row is None:
        return None
    if USE_POSTGRES:
        return DictRow(dict(row))
    return row


def db_fetchall(conn, sql, params=None):
    """Fetch all rows with compatibility."""
    cur = db_execute(conn, sql, params)
    rows = cur.fetchall()
    if USE_POSTGRES:
        return [DictRow(dict(r)) for r in rows]
    return rows


def init_db():
    if USE_POSTGRES:
        # Use raw psycopg connection for init (not wrapper)
        raw_conn = psycopg.connect(DATABASE_URL)
        _init_postgres(raw_conn)
        raw_conn.close()
    else:
        conn = get_db()
        _init_sqlite(conn)
        conn.close()


def _init_postgres(conn):
    """Initialize PostgreSQL database."""
    cur = conn.cursor()
    
    cur.execute('''
        CREATE TABLE IF NOT EXISTS factories (
            id SERIAL PRIMARY KEY,
            code TEXT UNIQUE NOT NULL,
            name TEXT NOT NULL,
            address TEXT DEFAULT '',
            is_active INTEGER DEFAULT 1
        )
    ''')
    
    cur.execute("INSERT INTO factories (code, name) VALUES ('F1', 'Pabrik 1 (~1.200 orang)') ON CONFLICT (code) DO NOTHING")
    cur.execute("INSERT INTO factories (code, name) VALUES ('F2', 'Pabrik 2 (~700 orang)') ON CONFLICT (code) DO NOTHING")
    
    cur.execute('''
        CREATE TABLE IF NOT EXISTS employees (
            id SERIAL PRIMARY KEY,
            nik TEXT UNIQUE NOT NULL,
            name TEXT NOT NULL,
            factory_id INTEGER DEFAULT 1,
            department TEXT DEFAULT '',
            position TEXT DEFAULT '',
            section TEXT DEFAULT '',
            join_date TEXT NOT NULL DEFAULT '2025-01-01',
            base_salary REAL NOT NULL DEFAULT 0,
            transport_allowance REAL DEFAULT 0,
            meal_allowance REAL DEFAULT 0,
            ptkp_status TEXT DEFAULT 'TK/0',
            bank_name TEXT DEFAULT '',
            bank_account TEXT DEFAULT '',
            is_active INTEGER DEFAULT 1,
            finger_id TEXT DEFAULT '',
            work_schedule TEXT DEFAULT '',
            created_at TIMESTAMP DEFAULT NOW(),
            updated_at TIMESTAMP DEFAULT NOW(),
            FOREIGN KEY (factory_id) REFERENCES factories(id)
        )
    ''')

    cur.execute('''
        CREATE TABLE IF NOT EXISTS attendance (
            id SERIAL PRIMARY KEY,
            employee_id INTEGER NOT NULL,
            date TEXT NOT NULL,
            clock_in TEXT,
            clock_out TEXT,
            status TEXT DEFAULT 'present',
            overtime_hours REAL DEFAULT 0,
            overtime_type TEXT DEFAULT 'weekday',
            notes TEXT DEFAULT '',
            FOREIGN KEY (employee_id) REFERENCES employees(id),
            UNIQUE(employee_id, date)
        )
    ''')

    cur.execute('CREATE INDEX IF NOT EXISTS idx_emp_factory ON employees(factory_id)')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_emp_finger ON employees(finger_id)')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_att_empdate ON attendance(employee_id, date)')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_att_date ON attendance(date)')

    cur.execute('''
        CREATE TABLE IF NOT EXISTS payroll (
            id SERIAL PRIMARY KEY,
            employee_id INTEGER NOT NULL,
            period TEXT NOT NULL,
            work_days INTEGER DEFAULT 0,
            absent_days INTEGER DEFAULT 0,
            short_hours REAL DEFAULT 0,
            total_work_hours REAL DEFAULT 0,
            overtime_l1 REAL DEFAULT 0,
            overtime_l2 REAL DEFAULT 0,
            overtime_ll1 REAL DEFAULT 0,
            overtime_ll2 REAL DEFAULT 0,
            overtime_ll3 REAL DEFAULT 0,
            overtime_total_hours REAL DEFAULT 0,
            base_salary REAL DEFAULT 0,
            special_allowance REAL DEFAULT 0,
            service_bonus REAL DEFAULT 0,
            attendance_bonus REAL DEFAULT 0,
            transport_allowance REAL DEFAULT 0,
            meal_allowance REAL DEFAULT 0,
            thr REAL DEFAULT 0,
            gross_salary REAL DEFAULT 0,
            overtime_lk1 REAL DEFAULT 0,
            overtime_lk2 REAL DEFAULT 0,
            overtime_ll1_amt REAL DEFAULT 0,
            overtime_ll2_amt REAL DEFAULT 0,
            overtime_ll3_amt REAL DEFAULT 0,
            overtime_pay REAL DEFAULT 0,
            bpjs_kes_company REAL DEFAULT 0,
            bpjs_kes_employee REAL DEFAULT 0,
            bpjs_kes_total REAL DEFAULT 0,
            bpjs_jht_company REAL DEFAULT 0,
            bpjs_jht_employee REAL DEFAULT 0,
            bpjs_jht_total REAL DEFAULT 0,
            bpjs_jkk_company REAL DEFAULT 0,
            bpjs_jkk_employee REAL DEFAULT 0,
            bpjs_jkk_total REAL DEFAULT 0,
            bpjs_jkm_company REAL DEFAULT 0,
            bpjs_jkm_employee REAL DEFAULT 0,
            bpjs_jkm_total REAL DEFAULT 0,
            bpjs_jp_company REAL DEFAULT 0,
            bpjs_jp_employee REAL DEFAULT 0,
            bpjs_jp_total REAL DEFAULT 0,
            bpjs_total_company REAL DEFAULT 0,
            bpjs_total_employee REAL DEFAULT 0,
            bpjs_total_total REAL DEFAULT 0,
            deduction_absence REAL DEFAULT 0,
            deduction_late REAL DEFAULT 0,
            total_deductions REAL DEFAULT 0,
            pph_allowance REAL DEFAULT 0,
            madome REAL DEFAULT 0,
            correction_underpay REAL DEFAULT 0,
            total_additions REAL DEFAULT 0,
            pph21 REAL DEFAULT 0,
            correction_salary REAL DEFAULT 0,
            total_manual_deductions REAL DEFAULT 0,
            net_salary REAL DEFAULT 0,
            payment_method TEXT DEFAULT 'TRANSFER',
            pph21_legacy REAL DEFAULT 0,
            import_source TEXT DEFAULT '',
            created_at TIMESTAMP DEFAULT NOW(),
            FOREIGN KEY (employee_id) REFERENCES employees(id),
            UNIQUE(employee_id, period)
        )
    ''')

    cur.execute('''
        CREATE TABLE IF NOT EXISTS overtime_requests (
            id SERIAL PRIMARY KEY,
            employee_id INTEGER NOT NULL,
            request_date TEXT NOT NULL,
            planned_start TEXT,
            planned_end TEXT,
            planned_hours REAL DEFAULT 0,
            reason TEXT DEFAULT '',
            approved_by TEXT DEFAULT '',
            status TEXT DEFAULT 'pending',
            finger_in TEXT,
            finger_out TEXT,
            actual_hours REAL DEFAULT 0,
            match_status TEXT DEFAULT '',
            ot_cost REAL DEFAULT 0,
            factory_id INTEGER DEFAULT 1,
            notes TEXT DEFAULT '',
            created_at TIMESTAMP DEFAULT NOW(),
            FOREIGN KEY (employee_id) REFERENCES employees(id),
            UNIQUE(employee_id, request_date)
        )
    ''')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_ot_date ON overtime_requests(request_date)')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_ot_emp ON overtime_requests(employee_id)')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_ot_factory ON overtime_requests(factory_id)')

    cur.execute('''
        CREATE TABLE IF NOT EXISTS holidays (
            id SERIAL PRIMARY KEY,
            date TEXT NOT NULL UNIQUE,
            name TEXT DEFAULT '',
            type TEXT DEFAULT '',
            is_cuti_bersama INTEGER DEFAULT 0
        )
    ''')

    cur.execute('''
        CREATE TABLE IF NOT EXISTS leave_balance (
            id SERIAL PRIMARY KEY,
            employee_id INTEGER NOT NULL,
            year INTEGER NOT NULL,
            annual_quota INTEGER DEFAULT 0,
            cb_deduction INTEGER DEFAULT 0,
            carry_over_deduction INTEGER DEFAULT 0,
            used INTEGER DEFAULT 0,
            remaining INTEGER DEFAULT 0,
            updated_at TIMESTAMP DEFAULT NOW(),
            FOREIGN KEY (employee_id) REFERENCES employees(id),
            UNIQUE(employee_id, year)
        )
    ''')

    cur.execute('''
        CREATE TABLE IF NOT EXISTS leave_records (
            id SERIAL PRIMARY KEY,
            employee_id INTEGER NOT NULL,
            leave_date TEXT NOT NULL,
            leave_type TEXT DEFAULT 'cuti',
            status TEXT DEFAULT 'approved',
            notes TEXT DEFAULT '',
            approved_by TEXT DEFAULT '',
            created_at TIMESTAMP DEFAULT NOW(),
            FOREIGN KEY (employee_id) REFERENCES employees(id),
            UNIQUE(employee_id, leave_date)
        )
    ''')

    # ── Ensure named unique constraints exist (for ON CONFLICT) ──
    # attendance: attendance_employee_id_date_key
    cur.execute("""
        SELECT 1 FROM pg_constraint
        WHERE conname = 'attendance_employee_id_date_key'
    """)
    if not cur.fetchone():
        try:
            # Drop anonymous constraint if exists, then create named one
            cur.execute("""
                ALTER TABLE attendance DROP CONSTRAINT IF EXISTS attendance_employee_id_date_key;
                DO $$ BEGIN
                    ALTER TABLE attendance ADD CONSTRAINT attendance_employee_id_date_key
                        UNIQUE (employee_id, date);
                EXCEPTION WHEN duplicate_table THEN NULL;
                END $$;
            """)
        except Exception:
            conn.rollback()

    # payroll: payroll_employee_id_period_key
    cur.execute("""
        SELECT 1 FROM pg_constraint
        WHERE conname = 'payroll_employee_id_period_key'
    """)
    if not cur.fetchone():
        try:
            cur.execute("""
                ALTER TABLE payroll DROP CONSTRAINT IF EXISTS payroll_employee_id_period_key;
                DO $$ BEGIN
                    ALTER TABLE payroll ADD CONSTRAINT payroll_employee_id_period_key
                        UNIQUE (employee_id, period);
                EXCEPTION WHEN duplicate_table THEN NULL;
                END $$;
            """)
        except Exception:
            conn.rollback()

    conn.commit()


def _init_sqlite(conn):
    """Initialize SQLite database (original code)."""
    conn.executescript('''
        CREATE TABLE IF NOT EXISTS factories (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT UNIQUE NOT NULL,
            name TEXT NOT NULL,
            address TEXT DEFAULT '',
            is_active INTEGER DEFAULT 1
        );

        INSERT OR IGNORE INTO factories (code, name) VALUES ('F1', 'Pabrik 1 (~1.200 orang)');
        INSERT OR IGNORE INTO factories (code, name) VALUES ('F2', 'Pabrik 2 (~700 orang)');

        CREATE TABLE IF NOT EXISTS employees (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nik TEXT UNIQUE NOT NULL,
            name TEXT NOT NULL,
            factory_id INTEGER DEFAULT 1,
            department TEXT DEFAULT '',
            position TEXT DEFAULT '',
            section TEXT DEFAULT '',
            join_date TEXT NOT NULL DEFAULT '2025-01-01',
            base_salary REAL NOT NULL DEFAULT 0,
            transport_allowance REAL DEFAULT 0,
            meal_allowance REAL DEFAULT 0,
            ptkp_status TEXT DEFAULT 'TK/0',
            bank_name TEXT DEFAULT '',
            bank_account TEXT DEFAULT '',
            is_active INTEGER DEFAULT 1,
            finger_id TEXT DEFAULT '',
            work_schedule TEXT DEFAULT '',
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (factory_id) REFERENCES factories(id)
        );

        CREATE TABLE IF NOT EXISTS attendance (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_id INTEGER NOT NULL,
            date TEXT NOT NULL,
            clock_in TEXT,
            clock_out TEXT,
            status TEXT DEFAULT 'present',
            overtime_hours REAL DEFAULT 0,
            overtime_type TEXT DEFAULT 'weekday',
            notes TEXT DEFAULT '',
            FOREIGN KEY (employee_id) REFERENCES employees(id),
            UNIQUE(employee_id, date)
        );

        CREATE INDEX IF NOT EXISTS idx_emp_factory ON employees(factory_id);
        CREATE INDEX IF NOT EXISTS idx_emp_finger ON employees(finger_id);
        CREATE INDEX IF NOT EXISTS idx_att_empdate ON attendance(employee_id, date);
        CREATE INDEX IF NOT EXISTS idx_att_date ON attendance(date);

        CREATE TABLE IF NOT EXISTS payroll (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_id INTEGER NOT NULL,
            period TEXT NOT NULL,
            work_days INTEGER DEFAULT 0,
            absent_days INTEGER DEFAULT 0,
            short_hours REAL DEFAULT 0,
            total_work_hours REAL DEFAULT 0,
            overtime_l1 REAL DEFAULT 0,
            overtime_l2 REAL DEFAULT 0,
            overtime_ll1 REAL DEFAULT 0,
            overtime_ll2 REAL DEFAULT 0,
            overtime_ll3 REAL DEFAULT 0,
            overtime_total_hours REAL DEFAULT 0,
            base_salary REAL DEFAULT 0,
            special_allowance REAL DEFAULT 0,
            service_bonus REAL DEFAULT 0,
            attendance_bonus REAL DEFAULT 0,
            transport_allowance REAL DEFAULT 0,
            meal_allowance REAL DEFAULT 0,
            thr REAL DEFAULT 0,
            gross_salary REAL DEFAULT 0,
            overtime_lk1 REAL DEFAULT 0,
            overtime_lk2 REAL DEFAULT 0,
            overtime_ll1_amt REAL DEFAULT 0,
            overtime_ll2_amt REAL DEFAULT 0,
            overtime_ll3_amt REAL DEFAULT 0,
            overtime_pay REAL DEFAULT 0,
            bpjs_kes_company REAL DEFAULT 0,
            bpjs_kes_employee REAL DEFAULT 0,
            bpjs_kes_total REAL DEFAULT 0,
            bpjs_jht_company REAL DEFAULT 0,
            bpjs_jht_employee REAL DEFAULT 0,
            bpjs_jht_total REAL DEFAULT 0,
            bpjs_jkk_company REAL DEFAULT 0,
            bpjs_jkk_employee REAL DEFAULT 0,
            bpjs_jkk_total REAL DEFAULT 0,
            bpjs_jkm_company REAL DEFAULT 0,
            bpjs_jkm_employee REAL DEFAULT 0,
            bpjs_jkm_total REAL DEFAULT 0,
            bpjs_jp_company REAL DEFAULT 0,
            bpjs_jp_employee REAL DEFAULT 0,
            bpjs_jp_total REAL DEFAULT 0,
            bpjs_total_company REAL DEFAULT 0,
            bpjs_total_employee REAL DEFAULT 0,
            bpjs_total_total REAL DEFAULT 0,
            deduction_absence REAL DEFAULT 0,
            deduction_late REAL DEFAULT 0,
            total_deductions REAL DEFAULT 0,
            pph_allowance REAL DEFAULT 0,
            madome REAL DEFAULT 0,
            correction_underpay REAL DEFAULT 0,
            total_additions REAL DEFAULT 0,
            pph21 REAL DEFAULT 0,
            correction_salary REAL DEFAULT 0,
            total_manual_deductions REAL DEFAULT 0,
            net_salary REAL DEFAULT 0,
            payment_method TEXT DEFAULT 'TRANSFER',
            pph21_legacy REAL DEFAULT 0,
            import_source TEXT DEFAULT '',
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (employee_id) REFERENCES employees(id),
            UNIQUE(employee_id, period)
        );
    ''')
    conn.commit()

    # Run migrations for existing databases
    _migrate_sqlite(conn)
    conn.close()


def _migrate_sqlite(conn):
    """Add new columns to existing SQLite tables if they don't exist."""

    conn.execute('''
        CREATE TABLE IF NOT EXISTS overtime_requests (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_id INTEGER NOT NULL,
            request_date TEXT NOT NULL,
            planned_start TEXT,
            planned_end TEXT,
            planned_hours REAL DEFAULT 0,
            reason TEXT DEFAULT '',
            approved_by TEXT DEFAULT '',
            status TEXT DEFAULT 'pending',
            finger_in TEXT,
            finger_out TEXT,
            actual_hours REAL DEFAULT 0,
            match_status TEXT DEFAULT '',
            ot_cost REAL DEFAULT 0,
            factory_id INTEGER DEFAULT 1,
            notes TEXT DEFAULT '',
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (employee_id) REFERENCES employees(id),
            UNIQUE(employee_id, request_date)
        )
    ''')
    conn.execute('CREATE INDEX IF NOT EXISTS idx_ot_date ON overtime_requests(request_date)')
    conn.execute('CREATE INDEX IF NOT EXISTS idx_ot_emp ON overtime_requests(employee_id)')
    conn.execute('CREATE INDEX IF NOT EXISTS idx_ot_factory ON overtime_requests(factory_id)')
    conn.commit()

    conn.execute('''
        CREATE TABLE IF NOT EXISTS holidays (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL UNIQUE,
            name TEXT DEFAULT '',
            type TEXT DEFAULT '',
            is_cuti_bersama INTEGER DEFAULT 0
        )
    ''')
    conn.commit()

    conn.execute('''
        CREATE TABLE IF NOT EXISTS leave_balance (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_id INTEGER NOT NULL,
            year INTEGER NOT NULL,
            annual_quota INTEGER DEFAULT 0,
            cb_deduction INTEGER DEFAULT 0,
            carry_over_deduction INTEGER DEFAULT 0,
            used INTEGER DEFAULT 0,
            remaining INTEGER GENERATED ALWAYS AS (annual_quota - cb_deduction - carry_over_deduction - used) STORED,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (employee_id) REFERENCES employees(id),
            UNIQUE(employee_id, year)
        )
    ''')
    conn.execute('''
        CREATE TABLE IF NOT EXISTS leave_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_id INTEGER NOT NULL,
            leave_date TEXT NOT NULL,
            leave_type TEXT DEFAULT 'cuti',
            status TEXT DEFAULT 'approved',
            notes TEXT DEFAULT '',
            approved_by TEXT DEFAULT '',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (employee_id) REFERENCES employees(id),
            UNIQUE(employee_id, leave_date)
        )
    ''')
    conn.commit()

    _add_column(conn, 'employees', 'section', 'TEXT DEFAULT ""')
    _add_column(conn, 'employees', 'work_schedule', 'TEXT DEFAULT ""')

    new_payroll_cols = [
        ('absent_days', 'INTEGER DEFAULT 0'),
        ('short_hours', 'REAL DEFAULT 0'),
        ('total_work_hours', 'REAL DEFAULT 0'),
        ('overtime_l1', 'REAL DEFAULT 0'),
        ('overtime_l2', 'REAL DEFAULT 0'),
        ('overtime_ll1', 'REAL DEFAULT 0'),
        ('overtime_ll2', 'REAL DEFAULT 0'),
        ('overtime_ll3', 'REAL DEFAULT 0'),
        ('overtime_total_hours', 'REAL DEFAULT 0'),
        ('special_allowance', 'REAL DEFAULT 0'),
        ('service_bonus', 'REAL DEFAULT 0'),
        ('attendance_bonus', 'REAL DEFAULT 0'),
        ('overtime_lk1', 'REAL DEFAULT 0'),
        ('overtime_lk2', 'REAL DEFAULT 0'),
        ('overtime_ll1_amt', 'REAL DEFAULT 0'),
        ('overtime_ll2_amt', 'REAL DEFAULT 0'),
        ('overtime_ll3_amt', 'REAL DEFAULT 0'),
        ('bpjs_kes_total', 'REAL DEFAULT 0'),
        ('bpjs_jht_total', 'REAL DEFAULT 0'),
        ('bpjs_jkk_employee', 'REAL DEFAULT 0'),
        ('bpjs_jkk_total', 'REAL DEFAULT 0'),
        ('bpjs_jkm_employee', 'REAL DEFAULT 0'),
        ('bpjs_jkm_total', 'REAL DEFAULT 0'),
        ('bpjs_jp_total', 'REAL DEFAULT 0'),
        ('bpjs_total_company', 'REAL DEFAULT 0'),
        ('bpjs_total_employee', 'REAL DEFAULT 0'),
        ('bpjs_total_total', 'REAL DEFAULT 0'),
        ('deduction_absence', 'REAL DEFAULT 0'),
        ('deduction_late', 'REAL DEFAULT 0'),
        ('pph_allowance', 'REAL DEFAULT 0'),
        ('madome', 'REAL DEFAULT 0'),
        ('correction_underpay', 'REAL DEFAULT 0'),
        ('total_additions', 'REAL DEFAULT 0'),
        ('correction_salary', 'REAL DEFAULT 0'),
        ('total_manual_deductions', 'REAL DEFAULT 0'),
        ('payment_method', 'TEXT DEFAULT "TRANSFER"'),
        ('pph21_legacy', 'REAL DEFAULT 0'),
        ('import_source', 'TEXT DEFAULT ""'),
    ]
    for col_name, col_def in new_payroll_cols:
        _add_column(conn, 'payroll', col_name, col_def)


def _add_column(conn, table, column, definition):
    """Safely add a column if it doesn't exist (SQLite only)."""
    try:
        conn.execute(f'ALTER TABLE {table} ADD COLUMN {column} {definition}')
        conn.commit()
    except Exception:
        pass


# Initialize on import
init_db()
