# config.py — Indonesian Labor Law Constants (UU Cipta Kerja & PP 35/2021)

# ============================================================
# BPJS Rates (2026)
# ============================================================
BPJS_KESEHATAN_COMPANY = 0.04      # 4% perusahaan
BPJS_KESEHATAN_EMPLOYEE = 0.01     # 1% karyawan
BPJS_KESEHATAN_MAX_SALARY = 12_000_000  # Batas atas gaji BPJS Kes

BPJS_JHT_COMPANY = 0.037           # 3.7% Jaminan Hari Tua
BPJS_JHT_EMPLOYEE = 0.02           # 2%

BPJS_JKK_COMPANY = 0.0045          # 0.45% Jaminan Kecelakaan Kerja (PT Komitrando actual rate)
BPJS_JKM_COMPANY = 0.003           # 0.3% Jaminan Kematian
BPJS_JP_COMPANY = 0.02             # 2% Jaminan Pensiun
BPJS_JP_EMPLOYEE = 0.01            # 1%
BPJS_JP_MAX_SALARY = 10_042_300    # Batas atas gaji JP (2026 est.)

# ============================================================
# Overtime Rates (PP 35/2021 Pasal 31)
# ============================================================
# Hari kerja biasa (6 hari kerja):
#   Jam pertama: 1.5x upah/jam
#   Jam berikutnya: 2x upah/jam
# Hari libur/istirahat mingguan:
#   8 jam pertama: 2x upah/jam
#   Jam ke-9: 3x upah/jam
#   Jam ke-10+: 4x upah/jam

# Upah per jam = 1/173 x upah sebulan
MONTHLY_WORK_HOURS = 173

def hourly_rate(monthly_salary):
    return monthly_salary / MONTHLY_WORK_HOURS

def calc_overtime_weekday(monthly_salary, hours):
    """Hitung lembur hari kerja biasa."""
    rate = hourly_rate(monthly_salary)
    if hours <= 0:
        return 0
    total = min(hours, 1) * 1.5 * rate
    if hours > 1:
        total += (hours - 1) * 2 * rate
    return round(total)

def calc_overtime_holiday(monthly_salary, hours):
    """Hitung lembur hari libur/istirahat."""
    rate = hourly_rate(monthly_salary)
    if hours <= 0:
        return 0
    total = min(hours, 8) * 2 * rate
    if hours > 8:
        total += min(hours - 8, 1) * 3 * rate
    if hours > 9:
        total += (hours - 9) * 4 * rate
    return round(total)

# ============================================================
# PPh 21 — Progressive Tax (UU HPP 2022)
# ============================================================
PPH21_BRACKETS = [
    (60_000_000,   0.05),
    (250_000_000,  0.15),
    (500_000_000,  0.25),
    (5_000_000_000, 0.30),
    (float('inf'), 0.35),
]

PTKP = {
    'TK/0': 54_000_000,
    'TK/1': 58_500_000,
    'TK/2': 63_000_000,
    'TK/3': 67_500_000,
    'K/0':  58_500_000,
    'K/1':  63_000_000,
    'K/2':  67_500_000,
    'K/3':  72_000_000,
}

def calc_pph21_annual(taxable_income):
    """Hitung PPh 21 tahunan dari PKP."""
    if taxable_income <= 0:
        return 0
    tax = 0
    prev = 0
    for bracket_limit, rate in PPH21_BRACKETS:
        if taxable_income <= prev:
            break
        amount = min(taxable_income, bracket_limit) - prev
        tax += amount * rate
        prev = bracket_limit
    return round(tax)

def calc_pph21_monthly(annual_gross, ptkp_status='TK/0'):
    """Hitung PPh 21 bulanan."""
    ptkp = PTKP.get(ptkp_status, 54_000_000)
    # Biaya jabatan 5% max 6jt/tahun
    biaya_jabatan = min(annual_gross * 0.05, 6_000_000)
    # BPJS yang ditanggung karyawan (sebagai pengurang)
    netto = annual_gross - biaya_jabatan
    pkp = max(netto - ptkp, 0)
    annual_tax = calc_pph21_annual(pkp)
    return round(annual_tax / 12)

# ============================================================
# THR (Tunjangan Hari Raya)
# ============================================================
def calc_thr(monthly_salary, months_worked):
    """THR: >= 12 bulan = 1 bulan gaji, < 12 bulan = proporsional."""
    if months_worked < 1:
        return 0
    if months_worked >= 12:
        return monthly_salary
    return round(monthly_salary * months_worked / 12)

# ============================================================
# Cuti (Annual Leave)
# ============================================================
ANNUAL_LEAVE_DAYS = 12  # Min 12 hari setelah 1 tahun kerja

# ============================================================
# Work Hours
# ============================================================
WORK_HOURS_PER_DAY_6 = 7    # 6 hari kerja: 7 jam/hari
WORK_HOURS_PER_DAY_5 = 8    # 5 hari kerja: 8 jam/hari
WORK_HOURS_PER_WEEK = 40
