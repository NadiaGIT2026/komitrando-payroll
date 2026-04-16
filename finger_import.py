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

def save_attendance(attendance_list):
    """Save processed attendance to database, matching finger_id to employee."""
    conn = get_db()
    saved = 0
    skipped = 0
    
    for att in attendance_list:
        # Find employee by finger_id
        emp = conn.execute(
            'SELECT id FROM employees WHERE finger_id = ?',
            (att['finger_id'],)
        ).fetchone()
        
        if not emp:
            skipped += 1
            continue
        
        try:
            conn.execute('''
                INSERT OR REPLACE INTO attendance 
                (employee_id, date, clock_in, clock_out, status)
                VALUES (?, ?, ?, ?, 'present')
            ''', (emp['id'], att['date'], att['clock_in'], att['clock_out']))
            saved += 1
        except Exception:
            skipped += 1
    
    conn.commit()
    conn.close()
    return {'saved': saved, 'skipped': skipped}

def import_and_save(filepath, format_type='generic'):
    """Full import pipeline: parse → process → save."""
    records = import_from_csv(filepath, format_type)
    attendance = process_attendance(records)
    result = save_attendance(attendance)
    result['total_records'] = len(records)
    result['total_attendance'] = len(attendance)
    return result
