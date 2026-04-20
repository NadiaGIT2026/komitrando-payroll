# finger_import.py — Import attendance from fingerprint device exports

import pandas as pd
import csv
from datetime import datetime
from models import get_db

def import_from_csv(filepath, format_type='generic'):
    """
    Import fingerprint attendance data from CSV/Excel.
    
    Supported formats:
    - generic: columns [finger_id, date, time, status]
    - zkteco: ZKTeco device export format
    - solution: Solution device format
    - fingerspot: Fingerspot device format
    """
    if filepath.endswith(('.xlsx', '.xls')):
        df = pd.read_excel(filepath)
    else:
        df = pd.read_csv(filepath)
    
    # Normalize column names
    df.columns = [c.strip().lower().replace(' ', '_') for c in df.columns]
    
    if format_type == 'zkteco':
        return _parse_zkteco(df)
    elif format_type == 'solution':
        return _parse_solution(df)
    elif format_type == 'fingerspot':
        return _parse_fingerspot(df)
    else:
        return _parse_generic(df)

def _parse_generic(df):
    """Generic format: finger_id, date, time or finger_id, datetime."""
    records = []
    
    # Try to detect column names
    id_col = _find_column(df, ['finger_id', 'id', 'employee_id', 'nik', 'pin', 'no', 'user_id'])
    
    if 'datetime' in df.columns:
        for _, row in df.iterrows():
            dt = pd.to_datetime(row['datetime'])
            records.append({
                'finger_id': str(row[id_col]).strip(),
                'date': dt.strftime('%Y-%m-%d'),
                'time': dt.strftime('%H:%M:%S'),
            })
    else:
        date_col = _find_column(df, ['date', 'tanggal', 'tgl'])
        time_col = _find_column(df, ['time', 'jam', 'waktu'])
        
        for _, row in df.iterrows():
            d = pd.to_datetime(row[date_col]).strftime('%Y-%m-%d')
            t = str(row[time_col]).strip()
            records.append({
                'finger_id': str(row[id_col]).strip(),
                'date': d,
                'time': t,
            })
    
    return records

def _parse_zkteco(df):
    """ZKTeco format parser."""
    records = []
    id_col = _find_column(df, ['ac-no.', 'ac_no', 'user_id', 'pin', 'no.', 'no'])
    date_col = _find_column(df, ['date', 'tanggal'])
    time_col = _find_column(df, ['time', 'jam'])
    
    for _, row in df.iterrows():
        records.append({
            'finger_id': str(row[id_col]).strip(),
            'date': pd.to_datetime(row[date_col]).strftime('%Y-%m-%d'),
            'time': str(row[time_col]).strip(),
        })
    return records

def _parse_solution(df):
    """Solution device format parser."""
    return _parse_generic(df)  # Usually same as generic

def _parse_fingerspot(df):
    """Fingerspot format parser."""
    records = []
    for _, row in df.iterrows():
        id_col = _find_column(df, ['pin', 'id', 'nik'])
        dt_col = _find_column(df, ['scan_date', 'datetime', 'date'])
        dt = pd.to_datetime(row[dt_col])
        records.append({
            'finger_id': str(row[id_col]).strip(),
            'date': dt.strftime('%Y-%m-%d'),
            'time': dt.strftime('%H:%M:%S'),
        })
    return records

def _find_column(df, candidates):
    """Find matching column name from candidates."""
    for c in candidates:
        if c in df.columns:
            return c
    # Try partial match
    for c in candidates:
        for col in df.columns:
            if c in col:
                return col
    return df.columns[0]  # fallback to first column

def process_attendance(records):
    """
    Process raw scan records into attendance (pair clock_in/clock_out).
    Returns list of attendance entries.
    """
    from collections import defaultdict
    
    # Group by finger_id and date
    grouped = defaultdict(list)
    for r in records:
        key = (r['finger_id'], r['date'])
        grouped[key].append(r['time'])
    
    attendance = []
    for (finger_id, date_str), times in grouped.items():
        times.sort()
        clock_in = times[0] if times else None
        clock_out = times[-1] if len(times) > 1 else None
        
        attendance.append({
            'finger_id': finger_id,
            'date': date_str,
            'clock_in': clock_in,
            'clock_out': clock_out,
        })
    
    return attendance

def _calc_status_and_ot(clock_in, clock_out, date_str, department, is_all_in, holidays_set):
    """
    Calculate attendance status and OT hours based on clock_in/clock_out.
    Returns (status, overtime_hours).
    """
    from datetime import datetime as dt, date as date_cls
    import calendar

    status = 'present'
    overtime_hours = 0.0

    if not clock_in:
        return ('present', 0.0)

    # Determine if it's a holiday/weekend
    try:
        d = date_cls.fromisoformat(date_str)
        is_weekend = d.weekday() >= 5  # Sat=5, Sun=6
        is_holiday = date_str in holidays_set or is_weekend
    except:
        is_holiday = False

    # Parse times
    try:
        ci = dt.strptime(clock_in[:5], '%H:%M')
    except:
        return ('present', 0.0)

    # ── Status: late detection (weekday only) ──
    # Tolerance: 08:15 (15분 tolerance)
    if not is_holiday:
        dept_upper = (department or '').upper()
        # Security (3-shift) — no late check
        if dept_upper not in ('SECURITY', 'SATPAM'):
            late_threshold = dt.strptime('08:15', '%H:%M')
            if ci > late_threshold:
                status = 'late'

    # ── OT calculation ──
    if clock_out and clock_in != clock_out:
        try:
            co = dt.strptime(clock_out[:5], '%H:%M')
        except:
            return (status, 0.0)

        if is_holiday:
            # Holiday: all worked hours count as OT (simplified)
            worked_minutes = (co - ci).total_seconds() / 60.0
            if worked_minutes > 60:  # at least 1 hour to count
                # Deduct 1 hour break if > 4 hours
                if worked_minutes > 240:
                    worked_minutes -= 60
                overtime_hours = max(0, worked_minutes / 60.0)
        else:
            # Weekday: OT = time after normal end
            dept_upper = (department or '').upper()
            normal_end = dt.strptime('17:30', '%H:%M') if dept_upper == 'OFFICE' else dt.strptime('17:00', '%H:%M')
            # OT tolerance: 15분 이내는 OT 아님
            ot_threshold = dt.strptime('17:45', '%H:%M') if dept_upper == 'OFFICE' else dt.strptime('17:15', '%H:%M')

            if co > ot_threshold:
                ot_minutes = (co - normal_end).total_seconds() / 60.0
                # Deduct dinner 30min if clock_out >= 20:00
                if co >= dt.strptime('20:00', '%H:%M'):
                    ot_minutes -= 30
                overtime_hours = max(0, ot_minutes / 60.0)

        # ALL IN: 평일만 3시간 이후부터 OT (휴일은 전체 OT)
        if not is_holiday and is_all_in and overtime_hours > 0:
            overtime_hours = max(0, overtime_hours - 3.0)

        # Cap: 4 hours/day max
        overtime_hours = min(overtime_hours, 4.0)
        overtime_hours = round(overtime_hours, 1)

    return (status, overtime_hours)


def save_attendance(attendance_list):
    """Save processed attendance to database, matching finger_id to employee."""
    conn = get_db()
    saved = 0
    skipped = 0

    # Load holidays set
    holidays_set = set()
    try:
        hrows = conn.execute('SELECT date FROM holidays').fetchall()
        holidays_set = {r['date'] if isinstance(r, dict) else r[0] for r in hrows}
    except:
        pass

    for att in attendance_list:
        # Find employee by finger_id (try finger_id first, then nik)
        emp = conn.execute(
            'SELECT id, department, is_all_in FROM employees WHERE finger_id = ?',
            (att['finger_id'],)
        ).fetchone()

        if not emp:
            emp = conn.execute(
                'SELECT id, department, is_all_in FROM employees WHERE nik = ?',
                (att['finger_id'],)
            ).fetchone()

        if not emp:
            skipped += 1
            continue

        department = emp['department'] if emp['department'] else ''
        is_all_in = bool(emp['is_all_in']) if emp['is_all_in'] else False

        # Calculate status and OT
        status, overtime_hours = _calc_status_and_ot(
            att['clock_in'], att['clock_out'], att['date'],
            department, is_all_in, holidays_set
        )

        try:
            # Check if record exists (don't overwrite SID/leave records)
            existing = conn.execute(
                'SELECT status FROM attendance WHERE employee_id = ? AND date = ?',
                (emp['id'], att['date'])
            ).fetchone()

            if existing and existing['status'] in ('leave', 'sid'):
                # Don't overwrite leave/SID with finger data
                skipped += 1
                continue

            conn.execute('''
                INSERT OR REPLACE INTO attendance
                (employee_id, date, clock_in, clock_out, status, overtime_hours)
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (emp['id'], att['date'], att['clock_in'], att['clock_out'],
                  status, overtime_hours))
            saved += 1
        except Exception:
            skipped += 1

    conn.commit()
    conn.close()
    return {'saved': saved, 'skipped': skipped}

def _parse_attendance_report(df):
    """Parse Attendance Report format: NIK, Nama Karyawan, In (Earliest Scan), Out (Latest Scan), Status."""
    attendance = []
    date_str = None
    
    # Try to extract date from filename or use today
    from datetime import date as date_class
    date_str = date_class.today().strftime('%Y-%m-%d')
    
    for _, row in df.iterrows():
        nik_col = _find_column(df, ['nik', 'pin', 'id'])
        in_col = _find_column(df, ['in_(earliest_scan)', 'in', 'clock_in', 'earliest'])
        out_col = _find_column(df, ['out_(latest_scan)', 'out', 'clock_out', 'latest'])
        
        nik = str(row[nik_col]).strip()
        if not nik or nik == 'nan':
            continue
            
        clock_in = row[in_col] if pd.notna(row[in_col]) else None
        clock_out = row[out_col] if pd.notna(row[out_col]) else None
        
        # Convert time objects to string
        if clock_in and hasattr(clock_in, 'strftime'):
            clock_in = clock_in.strftime('%H:%M:%S')
        elif clock_in:
            clock_in = str(clock_in)
            
        if clock_out and hasattr(clock_out, 'strftime'):
            clock_out = clock_out.strftime('%H:%M:%S')
        elif clock_out:
            clock_out = str(clock_out)
        
        attendance.append({
            'finger_id': nik,
            'date': date_str,
            'clock_in': clock_in,
            'clock_out': clock_out,
        })
    
    return attendance


def import_and_save(filepath, format_type='generic', report_date=None):
    """Full import pipeline: parse → process → save."""
    # Check if it's an Attendance Report format
    if filepath.endswith(('.xlsx', '.xls')):
        df_check = pd.read_excel(filepath)
    else:
        df_check = pd.read_csv(filepath)
    
    cols_lower = [c.strip().lower() for c in df_check.columns]
    
    # Detect Attendance Report format
    if any('nik' in c for c in cols_lower) and any('earliest' in c or 'in' == c for c in cols_lower):
        df_check.columns = [c.strip().lower().replace(' ', '_') for c in df_check.columns]
        attendance = _parse_attendance_report(df_check)
        
        # Override date if provided
        if report_date:
            for att in attendance:
                att['date'] = report_date
        
        result = save_attendance(attendance)
        result['total_records'] = len(attendance)
        result['total_attendance'] = len(attendance)
        return result
    
    # Original flow
    records = import_from_csv(filepath, format_type)
    attendance = process_attendance(records)
    result = save_attendance(attendance)
    result['total_records'] = len(records)
    result['total_attendance'] = len(attendance)
    return result
