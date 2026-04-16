# preview_calc.py — Mid-month payroll preview calculation

from datetime import datetime, date, timedelta
from config import (
    MONTHLY_WORK_HOURS, hourly_rate,
    calc_overtime_weekday, calc_overtime_holiday,
    BPJS_KESEHATAN_COMPANY, BPJS_KESEHATAN_EMPLOYEE, BPJS_KESEHATAN_MAX_SALARY,
    BPJS_JHT_COMPANY, BPJS_JHT_EMPLOYEE,
    BPJS_JKK_COMPANY, BPJS_JKM_COMPANY,
    BPJS_JP_COMPANY, BPJS_JP_EMPLOYEE, BPJS_JP_MAX_SALARY,
)
from models import get_db


def _parse_time(t):
    """Parse HH:MM time string to datetime.time, or None."""
    if not t:
        return None
    try:
        return datetime.strptime(t.strip(), '%H:%M').time()
    except (ValueError, AttributeError):
        return None


def _get_holidays(conn, start_date, end_date):
    """Return set of date strings that are holidays in the given range."""
    rows = conn.execute(
        "SELECT date FROM holidays WHERE date BETWEEN ? AND ?",
        (start_date, end_date)
    ).fetchall()
    return {r['date'] for r in rows}


def _working_days_in_range(start_date, end_date, work_schedule, holidays_set):
    """
    Count working days between start_date and end_date (inclusive).
    5-day: Mon-Fri; 6-day: Mon-Sat; Security: every day.
    Holidays are excluded.
    """
    is_6day = '6' in (work_schedule or '') or 'sat' in (work_schedule or '').lower()
    is_security = 'security' in (work_schedule or '').lower()

    d = datetime.strptime(start_date, '%Y-%m-%d').date()
    end = datetime.strptime(end_date, '%Y-%m-%d').date()
    count = 0
    while d <= end:
        ds = d.strftime('%Y-%m-%d')
        wd = d.weekday()  # 0=Mon, 5=Sat, 6=Sun
        if ds not in holidays_set:
            if is_security:
                count += 1  # Satpam works every day in rotation
            elif is_6day:
                if wd < 6:  # Mon-Sat
                    count += 1
            else:
                if wd < 5:  # Mon-Fri
                    count += 1
        d += timedelta(days=1)
    return count


def _is_holiday_or_weekend(date_str, holidays_set):
    """Check if a date is Saturday, Sunday, or a public holiday."""
    d = datetime.strptime(date_str, '%Y-%m-%d').date()
    return d.weekday() >= 5 or date_str in holidays_set


def _calc_overtime_hours(clock_in_str, clock_out_str, date_str, is_all_in, is_satpam, holidays_set):
    """
    Calculate overtime hours for a single attendance record.
    Returns (regular_ot_hours, holiday_ot_hours).
    """
    if is_satpam:
        return 0.0, 0.0

    clock_in = _parse_time(clock_in_str)
    clock_out = _parse_time(clock_out_str)
    if not clock_in or not clock_out:
        return 0.0, 0.0

    is_holiday = _is_holiday_or_weekend(date_str, holidays_set)

    if is_holiday:
        # All working hours on holiday/weekend count as holiday OT
        # Calculate total hours worked
        ci_minutes = clock_in.hour * 60 + clock_in.minute
        co_minutes = clock_out.hour * 60 + clock_out.minute

        if co_minutes <= ci_minutes:
            return 0.0, 0.0

        total_minutes = co_minutes - ci_minutes

        # Dinner break deduction: if works past 20:00 and period crosses 19:00-19:30
        if co_minutes > 20 * 60 and ci_minutes < 19 * 60:
            total_minutes -= 30

        total_hours = total_minutes / 60.0
        if total_hours < 0:
            total_hours = 0

        # For ALL IN on holiday: subtract 4 hours threshold
        if is_all_in:
            total_hours = max(0, total_hours - 4)

        # Round down to nearest 0.5
        total_hours = int(total_hours * 2) / 2.0

        return 0.0, total_hours
    else:
        # Regular workday OT: starts after 17:15
        co_minutes = clock_out.hour * 60 + clock_out.minute
        ot_start = 17 * 60 + 15  # 17:15

        if co_minutes <= ot_start:
            return 0.0, 0.0

        ot_minutes = co_minutes - ot_start

        # Dinner break: deduct 30 min if works past 20:00
        if co_minutes > 20 * 60:
            ot_minutes -= 30

        ot_hours = ot_minutes / 60.0
        if ot_hours < 0:
            ot_hours = 0

        # ALL IN: OT only counts after 4 hours beyond 17:15
        if is_all_in:
            ot_hours = max(0, ot_hours - 4)

        # Round down to nearest 0.5
        ot_hours = int(ot_hours * 2) / 2.0

        return ot_hours, 0.0


def _calc_late(clock_in_str, date_str, holidays_set):
    """
    Check if employee was late. Late = clock_in after 08:00 on a regular workday.
    Returns (is_late: bool, late_minutes: int).
    """
    if _is_holiday_or_weekend(date_str, holidays_set):
        return False, 0

    clock_in = _parse_time(clock_in_str)
    if not clock_in:
        return False, 0

    deadline = 8 * 60  # 08:00
    ci_minutes = clock_in.hour * 60 + clock_in.minute

    if ci_minutes > deadline:
        return True, ci_minutes - deadline
    return False, 0


def calculate_preview(factory_id=1, period_start='2026-03-01', period_end='2026-03-11'):
    """
    Calculate mid-month payroll preview for all employees of a given factory.
    Returns (summary_dict, list_of_employee_dicts, departments, sections).
    """
    conn = get_db()

    holidays_set = _get_holidays(conn, period_start, period_end)

    # Get all active F1 employees
    employees = conn.execute('''
        SELECT e.*, f.code as factory_code
        FROM employees e
        LEFT JOIN factories f ON e.factory_id = f.id
        WHERE e.is_active = 1 AND e.factory_id = ?
        ORDER BY e.department, e.section, e.name
    ''', (factory_id,)).fetchall()

    # Get all attendance for the period
    attendance_rows = conn.execute('''
        SELECT a.*, e.id as eid
        FROM attendance a
        JOIN employees e ON a.employee_id = e.id
        WHERE e.factory_id = ? AND a.date BETWEEN ? AND ?
        ORDER BY a.employee_id, a.date
    ''', (factory_id, period_start, period_end)).fetchall()

    # Build attendance lookup: employee_id -> list of records
    att_by_emp = {}
    for row in attendance_rows:
        eid = row['employee_id']
        if eid not in att_by_emp:
            att_by_emp[eid] = []
        att_by_emp[eid].append(row)

    conn.close()

    results = []
    departments = set()
    sections = set()

    for emp in employees:
        emp_id = emp['id']
        is_all_in = bool(emp['is_all_in'])
        section = (emp['section'] or '').strip().upper()
        is_satpam = 'SATPAM' in section
        work_schedule = emp['work_schedule'] or '5 Days (Mon-Fri)'
        base_salary = emp['base_salary'] or 0
        masa_kerja = emp['masa_kerja'] or 0
        tunjangan_khusus = emp['tunjangan_khusus'] or 0
        bonus_kehadiran = emp['bonus_kehadiran'] or 0
        tunjangan_pph = emp['tunjangan_pph'] or 0
        madome = emp['madome'] or 0
        # Gaji Pokok efektif = Base Salary + Masa Kerja (seniority)
        effective_salary = base_salary + masa_kerja
        department = emp['department'] or ''
        sec = emp['section'] or ''

        departments.add(department)
        sections.add(sec)

        # Working days in the preview period
        working_days = _working_days_in_range(period_start, period_end, work_schedule, holidays_set)

        # Attendance records
        att_records = att_by_emp.get(emp_id, [])
        days_present = len(att_records)

        # Days absent = working days - days present (can be negative if they work weekends, clamp to 0)
        days_absent = max(0, working_days - days_present)

        # Calculate overtime and late for each day
        total_reg_ot = 0.0
        total_hol_ot = 0.0
        late_count = 0
        late_minutes = 0

        for att in att_records:
            reg_ot, hol_ot = _calc_overtime_hours(
                att['clock_in'], att['clock_out'], att['date'],
                is_all_in, is_satpam, holidays_set
            )
            total_reg_ot += reg_ot
            total_hol_ot += hol_ot

            is_late, lm = _calc_late(att['clock_in'], att['date'], holidays_set)
            if is_late:
                late_count += 1
                late_minutes += lm

        total_ot_hours = total_reg_ot + total_hol_ot

        # Overtime pay
        # OT dihitung berdasarkan effective_salary (base + masa kerja)
        ot_pay_regular = calc_overtime_weekday(effective_salary, total_reg_ot)
        ot_pay_holiday = calc_overtime_holiday(effective_salary, total_hol_ot)
        overtime_pay = ot_pay_regular + ot_pay_holiday

        # Bonus Kehadiran: dari data Excel (bukan flat 200K)
        # Diberikan jika hadir penuh (tidak ada alpha)
        attendance_bonus = bonus_kehadiran if days_absent == 0 and days_present > 0 else 0

        # BPJS calculations (berdasarkan effective_salary = base + masa kerja)
        bpjs_kes_base = min(effective_salary, BPJS_KESEHATAN_MAX_SALARY)
        bpjs_kes_company = round(bpjs_kes_base * BPJS_KESEHATAN_COMPANY)
        bpjs_kes_employee = round(bpjs_kes_base * BPJS_KESEHATAN_EMPLOYEE)

        bpjs_jht_company = round(effective_salary * BPJS_JHT_COMPANY)
        bpjs_jht_employee = round(effective_salary * BPJS_JHT_EMPLOYEE)

        bpjs_jkk_company = round(effective_salary * BPJS_JKK_COMPANY)
        bpjs_jkm_company = round(effective_salary * BPJS_JKM_COMPANY)

        bpjs_jp_base = min(effective_salary, BPJS_JP_MAX_SALARY)
        bpjs_jp_company = round(bpjs_jp_base * BPJS_JP_COMPANY)
        bpjs_jp_employee = round(bpjs_jp_base * BPJS_JP_EMPLOYEE)

        bpjs_total_company = bpjs_kes_company + bpjs_jht_company + bpjs_jkk_company + bpjs_jkm_company + bpjs_jp_company
        bpjs_total_employee = bpjs_kes_employee + bpjs_jht_employee + bpjs_jp_employee

        # Absence deduction berdasarkan effective_salary
        if '6' in work_schedule or 'sat' in work_schedule.lower():
            days_per_month = 26
        else:
            days_per_month = 22
        daily_rate = effective_salary / days_per_month if days_per_month > 0 else 0
        absence_deduction = round(daily_rate * days_absent)

        # Late deduction: Rp 5,000 per late incident
        late_deduction = late_count * 5_000

        # Gross salary = Gaji Pokok + Masa Kerja + Tunjangan Khusus + Bonus Kehadiran + OT + Manual
        gross_salary = effective_salary + tunjangan_khusus + attendance_bonus + overtime_pay + tunjangan_pph + madome

        # Total deductions
        total_deductions = bpjs_total_employee + absence_deduction + late_deduction

        # Net salary estimate
        net_salary = gross_salary - total_deductions

        results.append({
            'employee_id': emp_id,
            'nik': emp['nik'],
            'name': emp['name'],
            'department': department,
            'section': sec,
            'is_all_in': is_all_in,
            'is_satpam': is_satpam,
            'work_schedule': work_schedule,
            'base_salary': base_salary,
            'masa_kerja': masa_kerja,
            'effective_salary': effective_salary,
            'tunjangan_khusus': tunjangan_khusus,
            'bonus_kehadiran_rate': bonus_kehadiran,
            'tunjangan_pph': tunjangan_pph,
            'madome': madome,
            'working_days': working_days,
            'days_present': days_present,
            'days_absent': days_absent,
            'total_reg_ot': total_reg_ot,
            'total_hol_ot': total_hol_ot,
            'total_ot_hours': total_ot_hours,
            'ot_pay_regular': ot_pay_regular,
            'ot_pay_holiday': ot_pay_holiday,
            'overtime_pay': overtime_pay,
            'late_count': late_count,
            'late_minutes': late_minutes,
            'late_deduction': late_deduction,
            'attendance_bonus': attendance_bonus,
            'bpjs_kes_company': bpjs_kes_company,
            'bpjs_kes_employee': bpjs_kes_employee,
            'bpjs_jht_company': bpjs_jht_company,
            'bpjs_jht_employee': bpjs_jht_employee,
            'bpjs_jkk_company': bpjs_jkk_company,
            'bpjs_jkm_company': bpjs_jkm_company,
            'bpjs_jp_company': bpjs_jp_company,
            'bpjs_jp_employee': bpjs_jp_employee,
            'bpjs_total_company': bpjs_total_company,
            'bpjs_total_employee': bpjs_total_employee,
            'absence_deduction': absence_deduction,
            'gross_salary': gross_salary,
            'total_deductions': total_deductions,
            'net_salary': net_salary,
        })

    # Build summary
    summary = {
        'total_employees': len(results),
        'total_all_in': sum(1 for r in results if r['is_all_in']),
        'total_satpam': sum(1 for r in results if r['is_satpam']),
        'total_present_days': sum(r['days_present'] for r in results),
        'total_absent_days': sum(r['days_absent'] for r in results),
        'total_ot_hours': sum(r['total_ot_hours'] for r in results),
        'total_reg_ot': sum(r['total_reg_ot'] for r in results),
        'total_hol_ot': sum(r['total_hol_ot'] for r in results),
        'total_overtime_pay': sum(r['overtime_pay'] for r in results),
        'total_base_salary': sum(r['base_salary'] for r in results),
        'total_gross': sum(r['gross_salary'] for r in results),
        'total_bpjs_employee': sum(r['bpjs_total_employee'] for r in results),
        'total_bpjs_company': sum(r['bpjs_total_company'] for r in results),
        'total_deductions': sum(r['total_deductions'] for r in results),
        'total_net': sum(r['net_salary'] for r in results),
        'total_late_count': sum(r['late_count'] for r in results),
        'total_attendance_bonus': sum(r['attendance_bonus'] for r in results),
        'total_tunjangan_khusus': sum(r['tunjangan_khusus'] for r in results),
        'total_tunjangan_pph': sum(r['tunjangan_pph'] for r in results),
        'total_madome': sum(r['madome'] for r in results),
    }

    # Department breakdown
    dept_breakdown = {}
    for r in results:
        dept = r['department'] or '(Tanpa Dept)'
        if dept not in dept_breakdown:
            dept_breakdown[dept] = {
                'count': 0, 'gross': 0, 'net': 0, 'ot_hours': 0, 'overtime_pay': 0
            }
        dept_breakdown[dept]['count'] += 1
        dept_breakdown[dept]['gross'] += r['gross_salary']
        dept_breakdown[dept]['net'] += r['net_salary']
        dept_breakdown[dept]['ot_hours'] += r['total_ot_hours']
        dept_breakdown[dept]['overtime_pay'] += r['overtime_pay']

    summary['dept_breakdown'] = dict(sorted(dept_breakdown.items()))

    return summary, results, sorted(departments), sorted(sections)
