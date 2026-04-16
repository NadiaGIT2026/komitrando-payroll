# payroll_calc.py — Core payroll calculation engine
# 규칙: effective_salary = base_salary + masa_kerja
# OT/BPJS/결근공제 모두 effective_salary 기준
# ALL IN (is_all_in=1): 정상근무일 4시간 초과분부터 OT
# VIP (is_all_in=2): 고정 월급, 출퇴근/OT/공제 없음. TGK 그대로 지급
# 경비(Satpam): OT 적용 안 함

from config import *
from models import get_db
from datetime import datetime, date, timedelta


def _parse_time(t):
    if not t:
        return None
    try:
        return datetime.strptime(t.strip(), '%H:%M').time()
    except (ValueError, AttributeError):
        return None


def _get_holidays(conn, period):
    """Get holidays for a given month period (YYYY-MM)."""
    rows = conn.execute(
        "SELECT date FROM holidays WHERE date LIKE ?",
        (f'{period}%',)
    ).fetchall()
    return {r['date'] for r in rows}


def _is_holiday_or_weekend(date_str, holidays_set):
    d = datetime.strptime(date_str, '%Y-%m-%d').date()
    return d.weekday() >= 5 or date_str in holidays_set


def _working_days_in_month(period, work_schedule, holidays_set, up_to_today=False):
    """Count working days in a month.
    If up_to_today=True, only count up to today (not future days).
    """
    year, month = map(int, period.split('-'))
    is_6day = '6' in (work_schedule or '')
    
    d = date(year, month, 1)
    if month == 12:
        end = date(year + 1, 1, 1) - timedelta(days=1)
    else:
        end = date(year, month + 1, 1) - timedelta(days=1)
    
    # 아직 안 끝난 달이면 오늘까지만 계산
    if up_to_today:
        today = date.today()
        if end > today:
            end = today
    
    count = 0
    while d <= end:
        ds = d.strftime('%Y-%m-%d')
        wd = d.weekday()
        if ds not in holidays_set:
            if is_6day:
                if wd < 6:  # Mon-Sat
                    count += 1
            else:
                if wd < 5:  # Mon-Fri
                    count += 1
        d += timedelta(days=1)
    return count


def _calc_overtime_hours(clock_in_str, clock_out_str, date_str, is_all_in, is_satpam, holidays_set):
    """Calculate overtime hours. Returns (weekday_ot, holiday_ot)."""
    if is_satpam:
        return 0.0, 0.0

    clock_in = _parse_time(clock_in_str)
    clock_out = _parse_time(clock_out_str)
    if not clock_in or not clock_out:
        return 0.0, 0.0

    is_holiday = _is_holiday_or_weekend(date_str, holidays_set)

    if is_holiday:
        ci_minutes = clock_in.hour * 60 + clock_in.minute
        co_minutes = clock_out.hour * 60 + clock_out.minute
        if co_minutes <= ci_minutes:
            return 0.0, 0.0

        total_minutes = co_minutes - ci_minutes
        # 저녁시간 차감: 20:00 이후 퇴근 시 30분
        if co_minutes > 20 * 60 and ci_minutes < 19 * 60:
            total_minutes -= 30

        total_hours = max(0, total_minutes / 60.0)
        if is_all_in:
            total_hours = max(0, total_hours - 4)
        total_hours = int(total_hours * 2) / 2.0  # 0.5단위 내림
        return 0.0, total_hours
    else:
        co_minutes = clock_out.hour * 60 + clock_out.minute
        ot_start = 17 * 60 + 15  # 17:15

        if co_minutes <= ot_start:
            return 0.0, 0.0

        ot_minutes = co_minutes - ot_start
        if co_minutes > 20 * 60:
            ot_minutes -= 30

        ot_hours = max(0, ot_minutes / 60.0)
        if is_all_in:
            ot_hours = max(0, ot_hours - 4)
        ot_hours = int(ot_hours * 2) / 2.0
        return ot_hours, 0.0


def calculate_employee_payroll(employee_id, period, include_thr=False):
    """
    Calculate monthly payroll for one employee.
    period format: 'YYYY-MM'
    """
    conn = get_db()

    emp = conn.execute('SELECT * FROM employees WHERE id = ?', (employee_id,)).fetchone()
    if not emp:
        conn.close()
        return None

    # 출퇴근 데이터
    attendance = conn.execute('''
        SELECT * FROM attendance 
        WHERE employee_id = ? AND date LIKE ?
        ORDER BY date
    ''', (employee_id, f'{period}%')).fetchall()

    holidays_set = _get_holidays(conn, period)

    # === 기본 정보 ===
    base_salary = emp['base_salary'] or 0
    masa_kerja = emp['masa_kerja'] or 0
    effective_salary = base_salary + masa_kerja  # ★ 핵심 규칙
    
    tunjangan_khusus = emp['tunjangan_khusus'] or 0
    bonus_kehadiran = emp['bonus_kehadiran'] or 0
    tunjangan_pph = emp['tunjangan_pph'] or 0
    madome_allowance = emp['madome'] or 0
    transport = emp['transport_allowance'] or 0
    meal = emp['meal_allowance'] or 0
    
    is_all_in_raw = emp['is_all_in'] or 0
    is_vip = (is_all_in_raw == 2)
    is_all_in = (is_all_in_raw == 1)
    section = (emp['section'] or '').strip().upper()
    is_satpam = 'SATPAM' in section
    work_schedule = emp['work_schedule'] or '5 Hari (Senin-Jumat)'

    # === VIP: 고정 월급, 출퇴근/OT/공제 없음 ===
    if is_vip:
        gross = effective_salary + tunjangan_khusus + bonus_kehadiran
        return {
            'employee_id': employee_id,
            'period': period,
            'base_salary': base_salary,
            'masa_kerja': masa_kerja,
            'effective_salary': effective_salary,
            'special_allowance': tunjangan_khusus,
            'attendance_bonus': bonus_kehadiran,
            'gross_salary': gross,
            'overtime_pay': 0,
            'overtime_l1': 0, 'overtime_l2': 0,
            'overtime_ll1': 0, 'overtime_ll2': 0, 'overtime_ll3': 0,
            'overtime_lk1': 0, 'overtime_lk2': 0,
            'overtime_ll1_amt': 0, 'overtime_ll2_amt': 0, 'overtime_ll3_amt': 0,
            'overtime_total_hours': 0,
            'thr': 0,
            'bpjs_kes_employee': 0, 'bpjs_kes_company': 0, 'bpjs_kes_total': 0,
            'bpjs_jht_employee': 0, 'bpjs_jht_company': 0, 'bpjs_jht_total': 0,
            'bpjs_jp_employee': 0, 'bpjs_jp_company': 0, 'bpjs_jp_total': 0,
            'bpjs_jkk_company': 0, 'bpjs_jkk_employee': 0, 'bpjs_jkk_total': 0,
            'bpjs_jkm_company': 0, 'bpjs_jkm_employee': 0, 'bpjs_jkm_total': 0,
            'bpjs_total_company': 0, 'bpjs_total_employee': 0, 'bpjs_total_total': 0,
            'deduction_absence': 0, 'deduction_late': 0,
            'total_deductions': 0,
            'total_manual_deductions': 0,
            'pph21': 0, 'pph_allowance': tunjangan_pph,
            'net_salary': gross + tunjangan_pph,
            'work_days': 0, 'absent_days': 0, 'short_hours': 0,
            'total_work_hours': 0, 'days_present': 0,
            'transport_allowance': transport, 'meal_allowance': meal,
            'service_bonus': masa_kerja,
            'madome': madome_allowance,
            'correction_underpay': 0, 'correction_salary': 0,
            'total_additions': 0,
            'import_source': 'calculated',
            'payment_method': emp['bank_name'] or 'TRANSFER',
            'is_vip': True,
        }

    # === 근무일수 계산 (현재 진행 중인 달이면 오늘까지만) ===
    working_days_month = _working_days_in_month(period, work_schedule, holidays_set, up_to_today=True)
    
    days_present = 0
    days_alpha = 0  # Alpha(무단결근)만 카운트 — BK 감소 대상
    days_absent = 0  # 전체 미출근 (Alpha + Izin 등)
    days_leave = 0   # SID/CT/CB/CN/ML 등 유급 휴가
    short_hours = 0
    total_work_hours = 0
    total_reg_ot = 0.0
    total_hol_ot = 0.0
    late_count = 0
    late_minutes_total = 0

    for att in attendance:
        status = (att['status'] or '').lower()
        leave_code = (att['leave_code'] or '').upper() if 'leave_code' in att.keys() else ''
        
        if status == 'present' and not leave_code:
            days_present += 1
        elif status == 'alpha' or leave_code == 'A':
            days_alpha += 1
        elif leave_code in ('SID', 'S', 'CT', 'CN', 'CK', 'CB', 'CM', 'C', 'ML', 'CG', 'CH'):
            days_leave += 1  # 유급 휴가 — 결근 아님
        elif leave_code == 'I' or status == 'izin':
            days_leave += 1  # 허가 — 결근 아님 (BK 감소 안 함)
        elif status == 'present':
            days_present += 1  # leave_code 있어도 present면 출근
        
        # OT 계산 (finger 데이터 있는 경우만)
        if att['clock_in'] or att['clock_out']:
            reg_ot, hol_ot = _calc_overtime_hours(
                att['clock_in'], att['clock_out'], att['date'],
                is_all_in, is_satpam, holidays_set
            )
            total_reg_ot += reg_ot
            total_hol_ot += hol_ot

            # 지각 체크 (08:00 이후 출근)
            if att['clock_in'] and not _is_holiday_or_weekend(att['date'], holidays_set):
                ci = _parse_time(att['clock_in'])
                if ci and (ci.hour * 60 + ci.minute) > 8 * 60:
                    late_count += 1
                    late_minutes_total += (ci.hour * 60 + ci.minute) - 8 * 60

    # 결근 = Alpha만 (SID/CT/I 등은 결근 아님)
    days_absent = days_alpha
    total_ot_hours = total_reg_ot + total_hol_ot

    # === OT 금액 (effective_salary 기준) ===
    ot_lk1 = calc_overtime_weekday(effective_salary, total_reg_ot)
    ot_lk2 = 0  # L2는 별도 구분 시 사용
    ot_ll1 = calc_overtime_holiday(effective_salary, total_hol_ot)
    ot_ll2 = 0
    ot_ll3 = 0
    overtime_pay = ot_lk1 + ot_ll1

    # === Bonus Kehadiran (Alpha 횟수에 따라 감소, Cuti/SID/Izin 제외) ===
    # Alpha 1일=-20%, 2일=-40%, 3일=-60%, 4일=-80%, 5일+=-100%
    if days_present <= 0 and days_leave <= 0:
        attendance_bonus = 0
    elif days_alpha == 0:
        attendance_bonus = bonus_kehadiran
    elif days_alpha >= 5:
        attendance_bonus = 0
    else:
        reduction = days_alpha * 0.20  # 20% per alpha day
        attendance_bonus = round(bonus_kehadiran * (1 - reduction))

    # === THR ===
    thr = 0
    if include_thr:
        try:
            join_date = datetime.strptime(emp['join_date'], '%Y-%m-%d').date()
            today = date.today()
            months_worked = (today.year - join_date.year) * 12 + (today.month - join_date.month)
            thr = calc_thr(effective_salary, months_worked)
        except (ValueError, TypeError):
            pass

    # === 총 급여 (Gaji Kotor) = effective + tunj_khusus + bonus (OT 별도!) ===
    gross = effective_salary + tunjangan_khusus + attendance_bonus + thr

    # === BPJS (effective_salary 기준) ===
    bpjs_kes_base = min(effective_salary, BPJS_KESEHATAN_MAX_SALARY)
    bpjs_kes_emp = round(bpjs_kes_base * BPJS_KESEHATAN_EMPLOYEE)
    bpjs_kes_comp = round(bpjs_kes_base * BPJS_KESEHATAN_COMPANY)
    bpjs_kes_total = bpjs_kes_comp + bpjs_kes_emp

    bpjs_jht_emp = round(effective_salary * BPJS_JHT_EMPLOYEE)
    bpjs_jht_comp = round(effective_salary * BPJS_JHT_COMPANY)
    bpjs_jht_total = bpjs_jht_comp + bpjs_jht_emp

    bpjs_jp_base = min(effective_salary, BPJS_JP_MAX_SALARY)
    bpjs_jp_emp = round(bpjs_jp_base * BPJS_JP_EMPLOYEE)
    bpjs_jp_comp = round(bpjs_jp_base * BPJS_JP_COMPANY)
    bpjs_jp_total = bpjs_jp_comp + bpjs_jp_emp

    bpjs_jkk_comp = round(effective_salary * BPJS_JKK_COMPANY)
    bpjs_jkm_comp = round(effective_salary * BPJS_JKM_COMPANY)

    bpjs_total_comp = bpjs_kes_comp + bpjs_jht_comp + bpjs_jkk_comp + bpjs_jkm_comp + bpjs_jp_comp
    bpjs_total_emp = bpjs_kes_emp + bpjs_jht_emp + bpjs_jp_emp
    bpjs_total_total = bpjs_total_comp + bpjs_total_emp

    # === 공제 (Potongan) ===
    # 결근 공제 (effective_salary 기준)
    if '6' in work_schedule:
        days_per_month = 26
    else:
        days_per_month = 22
    daily_rate = effective_salary / days_per_month if days_per_month > 0 else 0
    deduction_absence = round(daily_rate * days_absent)
    
    # 지각 공제 = 시급(effective/173) × 지각시간
    hourly = effective_salary / MONTHLY_WORK_HOURS
    deduction_late = round(hourly * late_minutes_total / 60)
    
    total_deductions = deduction_absence + deduction_late

    # === Manual 항목 ===
    # Penambah (추가)
    pph_allowance = tunjangan_pph
    total_additions = pph_allowance + madome_allowance

    # PPh 21 (Gross Kotor × 12 기준)
    annual_gross = gross * 12
    pph21 = calc_pph21_monthly(annual_gross, emp['ptkp_status'] or 'TK/0')

    # Pengurang (차감)
    correction_salary = 0
    total_manual_deductions = pph21 + correction_salary

    # === 최종 급여 (Gaji Diterima) ===
    # Net = Gross + OT + Penambah - BPJS_emp - Potongan - PPH
    net_salary = gross + overtime_pay + total_additions - bpjs_total_emp - total_deductions - total_manual_deductions

    result = {
        'employee_id': employee_id,
        'employee_name': emp['name'],
        'nik': emp['nik'],
        'department': emp['department'],
        'period': period,
        'work_days': days_present,
        'absent_days': days_absent,
        'short_hours': short_hours,
        'total_work_hours': total_work_hours,
        'overtime_l1': total_reg_ot,
        'overtime_l2': 0,
        'overtime_ll1': total_hol_ot,
        'overtime_ll2': 0,
        'overtime_ll3': 0,
        'overtime_total_hours': total_ot_hours,
        'base_salary': base_salary,
        'special_allowance': tunjangan_khusus,
        'service_bonus': masa_kerja,
        'attendance_bonus': attendance_bonus,
        'transport_allowance': transport,
        'meal_allowance': meal,
        'thr': thr,
        'gross_salary': gross,
        'overtime_lk1': ot_lk1,
        'overtime_lk2': ot_lk2,
        'overtime_ll1_amt': ot_ll1,
        'overtime_ll2_amt': ot_ll2,
        'overtime_ll3_amt': ot_ll3,
        'overtime_pay': overtime_pay,
        'bpjs_kes_company': bpjs_kes_comp,
        'bpjs_kes_employee': bpjs_kes_emp,
        'bpjs_kes_total': bpjs_kes_total,
        'bpjs_jht_company': bpjs_jht_comp,
        'bpjs_jht_employee': bpjs_jht_emp,
        'bpjs_jht_total': bpjs_jht_total,
        'bpjs_jkk_company': bpjs_jkk_comp,
        'bpjs_jkk_employee': 0,
        'bpjs_jkk_total': bpjs_jkk_comp,
        'bpjs_jkm_company': bpjs_jkm_comp,
        'bpjs_jkm_employee': 0,
        'bpjs_jkm_total': bpjs_jkm_comp,
        'bpjs_jp_company': bpjs_jp_comp,
        'bpjs_jp_employee': bpjs_jp_emp,
        'bpjs_jp_total': bpjs_jp_total,
        'bpjs_total_company': bpjs_total_comp,
        'bpjs_total_employee': bpjs_total_emp,
        'bpjs_total_total': bpjs_total_total,
        'deduction_absence': deduction_absence,
        'deduction_late': deduction_late,
        'total_deductions': total_deductions,
        'pph_allowance': pph_allowance,
        'madome': madome_allowance,
        'correction_underpay': 0,
        'total_additions': total_additions,
        'pph21': pph21,
        'correction_salary': correction_salary,
        'total_manual_deductions': total_manual_deductions,
        'net_salary': net_salary,
        'payment_method': emp['bank_name'] or 'TRANSFER',
        'import_source': 'calculated',
    }

    conn.close()
    return result


def save_payroll(result):
    """Save calculated payroll to database."""
    conn = get_db()
    conn.execute('''
        INSERT OR REPLACE INTO payroll 
        (employee_id, period, work_days, absent_days, short_hours, total_work_hours,
         overtime_l1, overtime_l2, overtime_ll1, overtime_ll2, overtime_ll3, overtime_total_hours,
         base_salary, special_allowance, service_bonus, attendance_bonus,
         transport_allowance, meal_allowance, thr, gross_salary,
         overtime_lk1, overtime_lk2, overtime_ll1_amt, overtime_ll2_amt, overtime_ll3_amt, overtime_pay,
         bpjs_kes_company, bpjs_kes_employee, bpjs_kes_total,
         bpjs_jht_company, bpjs_jht_employee, bpjs_jht_total,
         bpjs_jkk_company, bpjs_jkk_employee, bpjs_jkk_total,
         bpjs_jkm_company, bpjs_jkm_employee, bpjs_jkm_total,
         bpjs_jp_company, bpjs_jp_employee, bpjs_jp_total,
         bpjs_total_company, bpjs_total_employee, bpjs_total_total,
         deduction_absence, deduction_late, total_deductions,
         pph_allowance, madome, correction_underpay, total_additions,
         pph21, correction_salary, total_manual_deductions,
         net_salary, payment_method, import_source)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
    ''', (
        result['employee_id'], result['period'],
        result['work_days'], result['absent_days'], result['short_hours'], result['total_work_hours'],
        result['overtime_l1'], result['overtime_l2'], result['overtime_ll1'], result['overtime_ll2'], result['overtime_ll3'], result['overtime_total_hours'],
        result['base_salary'], result['special_allowance'], result['service_bonus'], result['attendance_bonus'],
        result['transport_allowance'], result['meal_allowance'], result['thr'], result['gross_salary'],
        result['overtime_lk1'], result['overtime_lk2'], result['overtime_ll1_amt'], result['overtime_ll2_amt'], result['overtime_ll3_amt'], result['overtime_pay'],
        result['bpjs_kes_company'], result['bpjs_kes_employee'], result['bpjs_kes_total'],
        result['bpjs_jht_company'], result['bpjs_jht_employee'], result['bpjs_jht_total'],
        result['bpjs_jkk_company'], result['bpjs_jkk_employee'], result['bpjs_jkk_total'],
        result['bpjs_jkm_company'], result['bpjs_jkm_employee'], result['bpjs_jkm_total'],
        result['bpjs_jp_company'], result['bpjs_jp_employee'], result['bpjs_jp_total'],
        result['bpjs_total_company'], result['bpjs_total_employee'], result['bpjs_total_total'],
        result['deduction_absence'], result['deduction_late'], result['total_deductions'],
        result['pph_allowance'], result['madome'], result['correction_underpay'], result['total_additions'],
        result['pph21'], result['correction_salary'], result['total_manual_deductions'],
        result['net_salary'], result['payment_method'], result['import_source'],
    ))
    conn.commit()
    conn.close()


def run_monthly_payroll(period, include_thr=False, factory_id=None):
    """Run payroll for active employees, optionally filtered by factory."""
    conn = get_db()
    if factory_id:
        employees = conn.execute(
            'SELECT id FROM employees WHERE is_active = 1 AND factory_id = ?',
            [int(factory_id)]
        ).fetchall()
    else:
        employees = conn.execute('SELECT id FROM employees WHERE is_active = 1').fetchall()
    conn.close()

    results = []
    for emp in employees:
        result = calculate_employee_payroll(emp['id'], period, include_thr)
        if result:
            save_payroll(result)
            results.append(result)

    return results
