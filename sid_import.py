# sid_import.py — SID (Status Ijin/Dispensasi) Import
"""
Import attendance status (Ijin, Sakit, Alpha, Cuti, etc.) from SID template Excel.
Template format:
  Row 1: Title
  Row 2: Headers (NIK, NAME, FACTORY, 1, 2, 3, ... 31)
  Row 3+: Data — each day cell has status code (H, I, S, A, C, CB, DL, CT, OFF) or blank
"""

import openpyxl
from models import get_db
from datetime import date


# Valid status codes and their DB mapping
STATUS_MAP = {
    'H': 'present',      # Hadir
    'I': 'ijin',         # Ijin (permission)
    'S': 'sakit',        # Sakit (sick)
    'A': 'alpha',        # Alpha (absent without notice)
    'C': 'cuti',         # Cuti (annual leave)
    'CB': 'cuti_bersama',# Cuti Bersama
    'DL': 'dinas_luar',  # Dinas Luar (business trip)
    'CT': 'cuti_tahunan',# Cuti Tahunan
    'OFF': 'off',        # Day off
}


def import_sid_file(filepath, year_month):
    """
    Import SID template Excel into attendance table.
    
    Args:
        filepath: Path to the SID Excel file
        year_month: Period string like '2026-03'
    
    Returns:
        dict with stats: updated, skipped, errors, details
    """
    wb = openpyxl.load_workbook(filepath, read_only=True)
    ws = wb.active
    
    # Parse year/month
    parts = year_month.split('-')
    year = int(parts[0])
    month = int(parts[1])
    
    conn = get_db()
    
    stats = {
        'updated': 0,
        'skipped': 0,
        'errors': [],
        'by_status': {},
    }
    
    rows = list(ws.iter_rows(min_row=3, values_only=False))
    
    for row in rows:
        # Column A = NIK, B = NAME, C = FACTORY, D onwards = days 1-31
        nik_cell = row[0].value
        if not nik_cell:
            continue
        
        nik = str(nik_cell).strip()
        
        # Find employee
        emp = conn.execute(
            'SELECT id FROM employees WHERE nik = ?', (nik,)
        ).fetchone()
        
        if not emp:
            stats['skipped'] += 1
            stats['errors'].append(f'NIK {nik} tidak ditemukan')
            continue
        
        employee_id = emp['id']
        
        # Process each day (columns D=index 3 to AH=index 33)
        for day_idx in range(1, 32):
            col_idx = day_idx + 2  # day 1 = column index 3 (D)
            if col_idx >= len(row):
                break
            
            cell_value = row[col_idx].value
            if not cell_value:
                continue
            
            code = str(cell_value).strip().upper()
            if code not in STATUS_MAP:
                stats['errors'].append(f'NIK {nik}, hari {day_idx}: kode "{code}" tidak valid')
                continue
            
            status = STATUS_MAP[code]
            
            # Skip 'present' — that's handled by fingerprint data
            if status == 'present':
                continue
            
            # Validate date
            try:
                import calendar
                max_day = calendar.monthrange(year, month)[1]
                if day_idx > max_day:
                    break
                att_date = f'{year}-{month:02d}-{day_idx:02d}'
            except ValueError:
                break
            
            # Upsert attendance record
            existing = conn.execute(
                'SELECT id, status FROM attendance WHERE employee_id = ? AND date = ?',
                (employee_id, att_date)
            ).fetchone()
            
            if existing:
                # Only update if current status is 'present' or 'late' (SID overrides)
                if existing['status'] in ('present', 'late'):
                    conn.execute(
                        'UPDATE attendance SET status = ?, notes = ? WHERE id = ?',
                        (status, f'SID import: {code}', existing['id'])
                    )
                    stats['updated'] += 1
                else:
                    # Already has a non-present status, skip
                    stats['skipped'] += 1
                    continue
            else:
                # No attendance record — create one with no clock in/out
                conn.execute(
                    '''INSERT INTO attendance (employee_id, date, clock_in, clock_out, status, notes)
                       VALUES (?, ?, NULL, NULL, ?, ?)''',
                    (employee_id, att_date, status, f'SID import: {code}')
                )
                stats['updated'] += 1
            
            # Track by status
            stats['by_status'][code] = stats['by_status'].get(code, 0) + 1
    
    conn.commit()
    conn.close()
    wb.close()
    
    return stats
