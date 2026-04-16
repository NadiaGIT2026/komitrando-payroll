# models.py — Database models (SQLite)

import sqlite3
import os
from datetime import datetime, date

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, 'data')
os.makedirs(DATA_DIR, exist_ok=True)
DB_PATH = os.path.join(DATA_DIR, 'payroll.db')

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn

def init_db():
    conn = get_db()
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

            -- Overtime hours breakdown
            overtime_l1 REAL DEFAULT 0,
            overtime_l2 REAL DEFAULT 0,
            overtime_ll1 REAL DEFAULT 0,
            overtime_ll2 REAL DEFAULT 0,
            overtime_ll3 REAL DEFAULT 0,
            overtime_total_hours REAL DEFAULT 0,

            -- Earnings
            base_salary REAL DEFAULT 0,
            special_allowance REAL DEFAULT 0,
            service_bonus REAL DEFAULT 0,
            attendance_bonus REAL DEFAULT 0,
            transport_allowance REAL DEFAULT 0,
            meal_allowance REAL DEFAULT 0,
            thr REAL DEFAULT 0,
            gross_salary REAL DEFAULT 0,

            -- Overtime amounts
            overtime_lk1 REAL DEFAULT 0,
            overtime_lk2 REAL DEFAULT 0,
            overtime_ll1_amt REAL DEFAULT 0,
            overtime_ll2_amt REAL DEFAULT 0,
            overtime_ll3_amt REAL DEFAULT 0,
            overtime_pay REAL DEFAULT 0,

            -- BPJS Kesehatan
            bpjs_kes_company REAL DEFAULT 0,
            bpjs_kes_employee REAL DEFAULT 0,
            bpjs_kes_total REAL DEFAULT 0,

            -- BPJS JHT
            bpjs_jht_company REAL DEFAULT 0,
            bpjs_jht_employee REAL DEFAULT 0,
            bpjs_jht_total REAL DEFAULT 0,

            -- BPJS JKK
            bpjs_jkk_company REAL DEFAULT 0,
            bpjs_jkk_employee REAL DEFAULT 0,
            bpjs_jkk_total REAL DEFAULT 0,

            -- BPJS JKM
            bpjs_jkm_company REAL DEFAULT 0,
            bpjs_jkm_employee REAL DEFAULT 0,
            bpjs_jkm_total REAL DEFAULT 0,

            -- BPJS JP
            bpjs_jp_company REAL DEFAULT 0,
            bpjs_jp_employee REAL DEFAULT 0,
            bpjs_jp_total REAL DEFAULT 0,

            -- BPJS Total
            bpjs_total_company REAL DEFAULT 0,
            bpjs_total_employee REAL DEFAULT 0,
            bpjs_total_total REAL DEFAULT 0,

            -- Deductions
            deduction_absence REAL DEFAULT 0,
            deduction_late REAL DEFAULT 0,
            total_deductions REAL DEFAULT 0,

            -- Manual adjustments (additions)
            pph_allowance REAL DEFAULT 0,
            madome REAL DEFAULT 0,
            correction_underpay REAL DEFAULT 0,
            total_additions REAL DEFAULT 0,

            -- Manual adjustments (deductions)
            pph21 REAL DEFAULT 0,
            correction_salary REAL DEFAULT 0,
            total_manual_deductions REAL DEFAULT 0,

            -- Final
            net_salary REAL DEFAULT 0,
            payment_method TEXT DEFAULT 'TRANSFER',

            -- Legacy compat
            pph21_legacy REAL DEFAULT 0,

            import_source TEXT DEFAULT '',
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (employee_id) REFERENCES employees(id),
            UNIQUE(employee_id, period)
        );
    ''')
    conn.commit()

    # Run migrations for existing databases
    _migrate(conn)
    conn.close()

def _migrate(conn):
    """Add new columns to existing tables if they don't exist."""

    # Overtime requests table
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

    # Holidays table (may already exist)
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

    # Leave management tables
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

    # Employee table migrations
    _add_column(conn, 'employees', 'section', 'TEXT DEFAULT ""')
    _add_column(conn, 'employees', 'work_schedule', 'TEXT DEFAULT ""')

    # Payroll table migrations - add all new columns
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
    """Safely add a column if it doesn't exist."""
    try:
        conn.execute(f'ALTER TABLE {table} ADD COLUMN {column} {definition}')
        conn.commit()
    except sqlite3.OperationalError:
        pass  # Column already exists

# Initialize on import
init_db()
