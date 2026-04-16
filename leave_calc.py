# leave_calc.py — Cuti (Annual Leave) Management
"""
Rules:
- 12 cuti/year after 1 year of service
- Cuti Bersama (CB) is deducted from annual quota
- Employees < 1 year: if they use CB, it goes negative → deducted from next year's quota
- Carry-over deduction = abs(negative balance from previous year)
"""

from models import get_db
from datetime import date, datetime


def get_employee_tenure_years(join_date_str, ref_date=None):
    """Calculate years of service as of ref_date."""
    if not join_date_str:
        return 0
    ref = ref_date or date.today()
    try:
        jd = datetime.strptime(str(join_date_str)[:10], '%Y-%m-%d').date()
    except (ValueError, TypeError):
        return 0
    delta = ref - jd
    return delta.days / 365.25


def get_annual_quota(join_date_str, year):
    """
    Get annual leave quota.
    - < 1 year of service as of Jan 1 of that year → 0
    - >= 1 year → 12
    """
    ref = date(year, 1, 1)
    years = get_employee_tenure_years(join_date_str, ref)
    if years >= 1:
        return 12
    return 0


def init_leave_balance(year=None):
    """
    Initialize leave_balance for all active employees for given year.
    Also auto-creates leave_records for Cuti Bersama dates.
    """
    if year is None:
        year = date.today().year

    conn = get_db()

    # Get all CB dates for this year
    cb_dates = conn.execute(
        "SELECT date FROM holidays WHERE is_cuti_bersama = 1 AND date LIKE ?",
        (f"{year}-%",)
    ).fetchall()
    cb_date_list = [r['date'] for r in cb_dates]

    employees = conn.execute(
        "SELECT id, join_date FROM employees WHERE is_active = 1"
    ).fetchall()

    stats = {'created': 0, 'updated': 0, 'cb_records': 0}

    for emp in employees:
        emp_id = emp['id']
        join_date = emp['join_date']
        quota = get_annual_quota(join_date, year)

        # Calculate carry-over from previous year
        carry_over = 0
        prev = conn.execute(
            "SELECT remaining FROM leave_balance WHERE employee_id = ? AND year = ?",
            (emp_id, year - 1)
        ).fetchone()
        if prev and prev['remaining'] < 0:
            carry_over = abs(prev['remaining'])

        # Count CB dates (all employees participate in CB regardless of tenure)
        cb_count = len(cb_date_list)

        # Insert or update leave_balance
        existing = conn.execute(
            "SELECT id FROM leave_balance WHERE employee_id = ? AND year = ?",
            (emp_id, year)
        ).fetchone()

        if existing:
            conn.execute("""
                UPDATE leave_balance SET
                    annual_quota = ?,
                    cb_deduction = ?,
                    carry_over_deduction = ?,
                    updated_at = CURRENT_TIMESTAMP
                WHERE employee_id = ? AND year = ?
            """, (quota, cb_count, carry_over, emp_id, year))
            stats['updated'] += 1
        else:
            conn.execute("""
                INSERT INTO leave_balance
                (employee_id, year, annual_quota, cb_deduction, carry_over_deduction, used)
                VALUES (?, ?, ?, ?, ?, 0)
            """, (emp_id, year, quota, cb_count, carry_over))
            stats['created'] += 1

        # Auto-create leave_records for CB dates
        for cb_date in cb_date_list:
            try:
                conn.execute("""
                    INSERT OR IGNORE INTO leave_records
                    (employee_id, leave_date, leave_type, status, notes)
                    VALUES (?, ?, 'cuti_bersama', 'approved', 'Auto: Cuti Bersama')
                """, (emp_id, cb_date))
                stats['cb_records'] += 1
            except Exception:
                pass

    conn.commit()
    conn.close()
    return stats


def use_leave(employee_id, leave_date, leave_type='cuti', notes='', approved_by=''):
    """Record a leave day and update balance."""
    conn = get_db()
    year = int(leave_date[:4])

    # Insert leave record
    conn.execute("""
        INSERT OR REPLACE INTO leave_records
        (employee_id, leave_date, leave_type, status, notes, approved_by)
        VALUES (?, ?, ?, 'approved', ?, ?)
    """, (employee_id, leave_date, leave_type, notes, approved_by))

    # Update used count in balance
    if leave_type == 'cuti':
        used_count = conn.execute("""
            SELECT COUNT(*) FROM leave_records
            WHERE employee_id = ? AND leave_type = 'cuti'
              AND leave_date LIKE ? AND status = 'approved'
        """, (employee_id, f"{year}-%")).fetchone()[0]

        conn.execute("""
            UPDATE leave_balance SET used = ?, updated_at = CURRENT_TIMESTAMP
            WHERE employee_id = ? AND year = ?
        """, (used_count, employee_id, year))

    conn.commit()
    conn.close()


def cancel_leave(employee_id, leave_date):
    """Cancel a leave record."""
    conn = get_db()
    year = int(leave_date[:4])

    conn.execute(
        "DELETE FROM leave_records WHERE employee_id = ? AND leave_date = ?",
        (employee_id, leave_date)
    )

    # Recalculate used count
    used_count = conn.execute("""
        SELECT COUNT(*) FROM leave_records
        WHERE employee_id = ? AND leave_type = 'cuti'
          AND leave_date LIKE ? AND status = 'approved'
    """, (employee_id, f"{year}-%")).fetchone()[0]

    conn.execute("""
        UPDATE leave_balance SET used = ?, updated_at = CURRENT_TIMESTAMP
        WHERE employee_id = ? AND year = ?
    """, (used_count, employee_id, year))

    conn.commit()
    conn.close()


def get_leave_summary(year=None, factory_id=None):
    """Get leave balance summary for all employees."""
    if year is None:
        year = date.today().year

    conn = get_db()
    factory_filter = ''
    params = [year]
    if factory_id:
        factory_filter = ' AND e.factory_id = ?'
        params.append(int(factory_id))

    rows = conn.execute(f"""
        SELECT lb.*, e.nik, e.name, e.join_date, e.department, e.section,
               f.code as factory_code
        FROM leave_balance lb
        JOIN employees e ON lb.employee_id = e.id
        LEFT JOIN factories f ON e.factory_id = f.id
        WHERE lb.year = ? AND e.is_active = 1 {factory_filter}
        ORDER BY f.code, e.name
    """, params).fetchall()
    conn.close()
    return rows


def get_employee_leave_detail(employee_id, year=None):
    """Get detailed leave info for one employee."""
    if year is None:
        year = date.today().year

    conn = get_db()

    balance = conn.execute("""
        SELECT lb.*, e.nik, e.name, e.join_date, e.department
        FROM leave_balance lb
        JOIN employees e ON lb.employee_id = e.id
        WHERE lb.employee_id = ? AND lb.year = ?
    """, (employee_id, year)).fetchone()

    records = conn.execute("""
        SELECT * FROM leave_records
        WHERE employee_id = ? AND leave_date LIKE ?
        ORDER BY leave_date
    """, (employee_id, f"{year}-%")).fetchall()

    conn.close()
    return balance, records
