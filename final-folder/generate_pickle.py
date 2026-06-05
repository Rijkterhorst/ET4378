#!/usr/bin/env python
# coding: utf-8
"""
generate_pickle.py
------------------
Adapted from 07a_simulation_FINAL_fixed_timing.py (Group20_Python_Scripts_Q10_Q11_fixed_v2.zip).
Produces sim_final.pkl which is required by Final code.py.

Run this script ONCE before running Final code.py.
"""

import pandas as pd, numpy as np, calendar, math, pickle, openpyxl
from pathlib import Path

# ─────────────────────────────────────────────────────────────────────────────
BASE = Path(__file__).parent   # folder this script lives in

INV_FILE     = BASE / 'Inverter Efficiency parameters.xlsx'
WEATHER_FILE = BASE / 'Tampa_FL-hour.csv'
PICKLE_OUT   = BASE / 'sim_final.pkl'

# ── System constants ──────────────────────────────────────────────────────────
YEAR           = 2025          # representative calendar year with 52 Thursdays
N_MOD_SOUTH    = 12
N_MOD_NORTH    = 12
PMAX_STC       = 490           # W
TC_PMAX        = -0.0026       # /°C
NOCT           = 46            # °C  (AIKO Neostar 490W datasheet)
TILT_DEG       = 35
AZ_SOUTH       = 180           # standard compass south (after Az+180 transform)
AZ_NORTH       = 0             # standard compass north
ALBEDO         = 0.2
VMP_STC        = 34.70         # V per module at STC
ISC_STC        = 14.88         # A per module at STC
VOC_STC        = 41.10         # V per module at STC
TC_VMP         = -0.0022       # /°C temperature coefficient of Vmp
TC_VOC         = -0.0022       # /°C temperature coefficient of Voc  (AIKO datasheet)
TC_ISC         =  0.0004       # /°C temperature coefficient of Isc
N_SERIES       = 12             # modules in series per string
N_PARALLEL     = 1             # parallel strings per array

ETA_CC         = 0.95          # charge-controller efficiency (kept for compatibility)
ETA_BATT_RT    = 0.92          # battery round-trip efficiency
ETA_CC_BATT    = 0.95          # CC efficiency for PV → battery
ETA_CC_GRID    = 0.95          # CC efficiency for PV → inverter (grid)
INV_MAX_DC_KW  = 7.113         # kW – max raw PV DC input to system

PDC0           = 6757.76       # W – Sandia model PDC0 (XW Pro 6848)
PAC0           = 6120.0        # W – Sandia model PAC0
PS0            = 42.0          # W – Sandia model self-consumption threshold
C0             = -0.000012     # Sandia model curvature coefficient

# Battery sizing – DoD-based calculation (25.6V / 200Ah)
li_V         = 25.6          # V per battery
li_Ah        = 200           # Ah per battery
_sys_Wh_nodod = 52866                          # Wh – AC blackout energy demand
_DoD          = 0.75
_sys_Wh_min   = _sys_Wh_nodod / (_DoD * 0.961)  # minimum required Wh
_li_Wh        = li_V * li_Ah                    # Wh per battery
_series       = math.ceil(48 / li_V)           # = 2 (two 25.6V batteries to reach 48V bus)
_parallel     = math.ceil(_sys_Wh_min / (_series * _li_Wh))  # = 8
BATT_NOM_KWH  = _series * _parallel * _li_Wh / 1000   # kWh actual installed

SOC_INIT       = 0.95
SOC_MAX        = 0.95
SOC_MIN        = 0.20          # 20 % floor (75 % usable range from 95 % ceiling)
DEG_PER_CYCLE  = 0.0002        # 0.02 % capacity loss per full cycle
P_CRITICAL_KW  = 6.0           # kW critical load during blackout
P_DC0_W        = 6120          # W – PDC0 of XW Pro 6848 (gives 88.8 % blackout load)
DOY1_DOW       = 2             # Jan 1 2025 = Wednesday (Mon=0)

# ═════════════════════════════════════════════════════════════════════════════
# 1. INVERTER EFFICIENCY CURVE  (loaded via openpyxl, identical to original)
# ═════════════════════════════════════════════════════════════════════════════
print("Loading inverter efficiency curve …")
wb_inv = openpyxl.load_workbook(INV_FILE, read_only=True)
ws_inv = wb_inv['Plot Data']
eff_pct, eff_nom = [], []
for row in ws_inv.iter_rows(values_only=True):
    if isinstance(row[0], (int, float)) and isinstance(row[1], (int, float)):
        eff_pct.append(float(row[0]))
        eff_nom.append(float(row[1]) / 100.0)   # convert % → fraction
eff_pct = np.array(eff_pct)
eff_nom = np.array(eff_nom)

def inv_eff(p_dc_w):
    """Return inverter efficiency (fraction) at given DC input power (W)."""
    if p_dc_w <= 0:
        return 0.0
    return float(np.interp(min(100, p_dc_w / P_DC0_W * 100), eff_pct, eff_nom))

# ═════════════════════════════════════════════════════════════════════════════
# 2. WEATHER DATA & POA IRRADIANCE
# ═════════════════════════════════════════════════════════════════════════════
print("Loading weather data …")
wx = pd.read_csv(WEATHER_FILE, sep=';')
wx['Az_std'] = (wx['Az'] + 180) % 360   # convert to standard compass (0 = North)

def poa(hs, az_std, G_bn, G_dh, G_gh, surf_az, tilt):
    """Return POA irradiance [W/m²] for a tilted surface."""
    if hs <= 0:
        return 0.0
    tr  = math.radians(tilt)
    zen = math.radians(90 - hs)
    daz = math.radians(az_std - surf_az)
    beam = max(0.0, G_bn * (math.cos(zen) * math.cos(tr)
                             + math.sin(zen) * math.sin(tr) * math.cos(daz)))
    return max(0.0, beam
               + G_dh * (1 + math.cos(tr)) / 2
               + G_gh * ALBEDO * (1 - math.cos(tr)) / 2)

print("  Computing POA irradiance for both slopes …")
G_s = np.array([poa(wx['hs'][i], wx['Az_std'][i], wx['G_Bn'][i],
                    wx['G_Dh'][i], wx['G_Gh'][i], AZ_SOUTH, TILT_DEG)
                for i in range(8760)])
G_n = np.array([poa(wx['hs'][i], wx['Az_std'][i], wx['G_Bn'][i],
                    wx['G_Dh'][i], wx['G_Gh'][i], AZ_NORTH, TILT_DEG)
                for i in range(8760)])

# ═════════════════════════════════════════════════════════════════════════════
# 3. PV DC POWER  (NOCT temperature model)
# ═════════════════════════════════════════════════════════════════════════════
Ta = wx['Ta'].values

def pv_power(G, Ta_, n):
    if G <= 0:
        return 0.0
    Tm = Ta_ + ((NOCT - 20) / 800) * G
    return n * PMAX_STC * (G / 1000) * max(0, 1 + TC_PMAX * (Tm - 25))

print("  Computing PV DC power …")
P_s_W  = np.array([pv_power(G_s[i], Ta[i], N_MOD_SOUTH) for i in range(8760)])
P_n_W  = np.array([pv_power(G_n[i], Ta[i], N_MOD_NORTH) for i in range(8760)])
P_dc_W = P_s_W + P_n_W

print(f"  Peak DC power     : {P_dc_W.max()/1000:.2f} kW")
print(f"  Annual DC energy  : {P_dc_W.sum()/1000:.0f} kWh")

def _cell_temp(G, Ta_):
    return Ta_ + ((NOCT - 20) / 800) * G

_idx_s   = int(np.argmax(P_s_W))
_idx_n   = int(np.argmax(P_n_W))
_Tm_s    = _cell_temp(G_s[_idx_s], Ta[_idx_s])
_Tm_n    = _cell_temp(G_n[_idx_n], Ta[_idx_n])
_Vmp_s   = N_SERIES   * VMP_STC * (1 + TC_VMP * (_Tm_s - 25))
_Vmp_n   = N_SERIES   * VMP_STC * (1 + TC_VMP * (_Tm_n - 25))
_Voc_s   = N_SERIES   * VOC_STC * (1 + TC_VOC * (_Tm_s - 25))
_Voc_n   = N_SERIES   * VOC_STC * (1 + TC_VOC * (_Tm_n - 25))
_Isc_s   = N_PARALLEL * ISC_STC * (1 + TC_ISC * (_Tm_s - 25))
_Isc_n   = N_PARALLEL * ISC_STC * (1 + TC_ISC * (_Tm_n - 25))
_Ipeak_s = P_s_W[_idx_s] / _Vmp_s
_Ipeak_n = P_n_W[_idx_n] / _Vmp_n

# ═════════════════════════════════════════════════════════════════════════════
# 4. TIME ARRAYS & BLACKOUT MASK
#    Uses DOY from weather file + DOY1_DOW to compute weekday.
#    Hour convention in CSV is 1–24.
# ═════════════════════════════════════════════════════════════════════════════
doy         = wx['Day'].values.astype(int)    # 1–365
hrs         = wx['Hour'].values.astype(int)   # 1–24
mons        = wx['Month'].values.astype(int)  # 1–12

is_blackout = np.zeros(8760, dtype=bool)
for i in range(8760):
    dow = (DOY1_DOW + int(doy[i]) - 1) % 7   # 0=Mon, 3=Thu, 4=Fri
    h   = int(hrs[i])
    if (dow == 3 and h >= 20) or (dow == 4 and h < 4):
        is_blackout[i] = True

print(f"\n  Blackout hours    : {is_blackout.sum()}  (expect {52*8} = 416)")

# ═════════════════════════════════════════════════════════════════════════════
# 5b. LOAD PROFILE  (8760 hours, representative year)
# ═════════════════════════════════════════════════════════════════════════════
MONTHLY_LOAD_KWH = [5702.76, 5150.88, 5702.76, 5518.80, 5702.76, 5518.80,
                    5702.76, 5702.76, 5518.80, 5702.76, 5518.80, 5702.76]
HF_RAW  = np.array([3.81,3.81,3.88,3.96,3.96,4.03,4.19,4.19,
                    4.34,4.34,4.41,4.41,4.49,4.49,4.41,4.41,
                    4.34,4.34,4.26,4.19,4.11,3.96,3.88,3.81])
HF_NORM = HF_RAW / HF_RAW.sum()

load_kW_arr = np.zeros(8760)
for i in range(8760):
    if is_blackout[i]:
        load_kW_arr[i] = P_CRITICAL_KW
    else:
        m = int(mons[i])
        h = int(hrs[i]) - 1            # CSV uses 1–24; factors indexed 0–23
        load_kW_arr[i] = (MONTHLY_LOAD_KWH[m-1] / calendar.monthrange(YEAR, m)[1]) * HF_NORM[h]

# ═════════════════════════════════════════════════════════════════════════════
# 5. BLACKOUT OPERATING POINT
# ═════════════════════════════════════════════════════════════════════════════
# Sandia model inverted: given PAC = 6 kW, solve C0*x² + A*x − PAC = 0 for x = PDC − PS0
_A_san     = PAC0 / (PDC0 - PS0) - C0 * (PDC0 - PS0)
_pac_bo    = P_CRITICAL_KW * 1000.0
_disc      = _A_san**2 - 4 * C0 * (-_pac_bo)
_x_bo      = (-_A_san + math.sqrt(_disc)) / (2 * C0)
_pdc_bo    = _x_bo + PS0                               # DC input at blackout load
bo_inv_eff = _pac_bo / _pdc_bo                         # η = PAC / PDC
DC_per_hour  = P_CRITICAL_KW / bo_inv_eff              # kWh DC drawn per blackout hour
DC_per_event = DC_per_hour * 8                         # kWh DC for one 8-hour event
peak_eta     = 0.0   # computed from Year 1 simulation max eta_inv_h
peak_pdc_kw  = 0.0   # DC input power (kW) at the peak efficiency hour

print(f"  Blackout inv η    : {bo_inv_eff*100:.4f}%  (load = {P_CRITICAL_KW*1000/P_DC0_W*100:.1f}%)")
print(f"  DC per event      : {DC_per_event:.4f} kWh")

# ═════════════════════════════════════════════════════════════════════════════
# 6. THREE-YEAR SIMULATION  (Sandia inverter model + load-tracking dispatch)
# ═════════════════════════════════════════════════════════════════════════════
print("\nRunning 3-year simulation …")

PV_DEGRAD            = [1.0, 0.99, 0.9865]   # Year 1, 2, 3 PV output fractions
ELEC_PRICE_EUR_KWH   = 0.24
BLACKOUT_SAVE_PER_HR = 3 * 45

ann      = []
soc_all  = []
ac_all   = []
mode_all = []

soc_kWh   = SOC_INIT * BATT_NOM_KWH
nom_cap   = BATT_NOM_KWH
cycle_tot = 0

# Pre-allocate Year-1 hourly arrays for battery_load.csv export
_batt_soc_Wh  = np.zeros(8760)
_pv_to_batt   = np.zeros(8760)
_pv_to_load   = np.zeros(8760)
_grid_kWh     = np.zeros(8760)
_batt_to_load = np.zeros(8760)
_eta_inv      = np.zeros(8760)

for yr_idx in range(3):
    yr           = YEAR + yr_idx
    ac_yield     = 0.0
    n_bo         = 0
    bo_ok        = 0
    soc_tr       = []
    ac_tr        = []
    mode_tr      = []
    inv_s        = []
    in_bo_prev       = False
    eta_inv_prev     = 0.92
    elec_save_yr     = 0.0
    blackout_save_yr = 0.0

    bat_max_kWh  = SOC_MAX * nom_cap
    bat_min_kWh  = SOC_MIN * nom_cap

    for i in range(8760):
        load_kW        = load_kW_arr[i]
        pv_dc_kW       = P_dc_W[i] / 1000.0 * PV_DEGRAD[yr_idx]
        bo             = bool(is_blackout[i])
        pv_remaining   = pv_dc_kW
        load_remaining = load_kW
        eta_inv_h      = eta_inv_prev if eta_inv_prev > 0.0 else 0.92
        pv_to_batt     = 0.0
        pv_to_load_kW  = 0.0
        batt_to_load   = 0.0

        if bo and not in_bo_prev:
            n_bo += 1
            if (soc_kWh - bat_min_kWh) >= DC_per_event:
                bo_ok += 1
        in_bo_prev = bo

        # STEP 1 – Battery → Load (blackout only)
        if bo and soc_kWh > bat_min_kWh:
            batt_avail     = soc_kWh - bat_min_kWh
            batt_needed    = load_remaining / eta_inv_h
            batt_used      = min(batt_avail, batt_needed)
            soc_kWh       -= batt_used
            batt_to_load   = batt_used
            load_remaining -= batt_used * eta_inv_h

        # STEP 2 – PV → Battery
        if pv_remaining > 0.0 and soc_kWh < bat_max_kWh:
            room_dc       = (bat_max_kWh - soc_kWh) / (ETA_CC_BATT * ETA_BATT_RT)
            pv_used       = min(pv_remaining, room_dc)
            soc_kWh       = min(soc_kWh + pv_used * ETA_CC_BATT * ETA_BATT_RT, bat_max_kWh)
            pv_to_batt    = pv_used
            pv_remaining -= pv_used

        # STEP 3 – PV → Load (inverter-limited)
        if pv_remaining > 0.0 and load_remaining > 0.0:
            pv_needed      = load_remaining / (ETA_CC_GRID * eta_inv_h)
            pv_used        = min(pv_remaining, pv_needed, INV_MAX_DC_KW)
            pv_to_load_kW  = pv_used
            pv_remaining  -= pv_used
            load_remaining -= pv_used * ETA_CC_GRID * eta_inv_h

        # Inverter efficiency – Sandia model
        P_DC_inv_W = (pv_to_load_kW * ETA_CC_GRID + batt_to_load) * 1000.0
        if P_DC_inv_W > PS0:
            PAC_h = ((PAC0 / (PDC0 - PS0) - C0 * (PDC0 - PS0)) * (P_DC_inv_W - PS0)
                     + C0 * (P_DC_inv_W - PS0) ** 2)
            PAC_h     = max(0.0, min(PAC_h, PAC0))
            eta_inv_h = PAC_h / P_DC_inv_W
            p_ac_out  = PAC_h / 1000.0
        else:
            eta_inv_h = 0.0
            p_ac_out  = 0.0
        eta_inv_prev = eta_inv_h

        # Track peak inverter efficiency for Year 1
        if yr_idx == 0 and P_DC_inv_W > PS0 and eta_inv_h > peak_eta / 100:
            peak_eta    = eta_inv_h * 100
            peak_pdc_kw = P_DC_inv_W / 1000.0

        # Mode and GT accounting
        if bo:
            mode = "BLACKOUT"
        elif p_ac_out > 0.0:
            mode      = "GRID_TIED"
            ac_yield += p_ac_out
            inv_s.append((p_ac_out, eta_inv_h))
        elif pv_to_batt > 0.0:
            mode = "CHARGING"
        else:
            mode = "IDLE"

        soc_kWh = float(np.clip(soc_kWh, bat_min_kWh, bat_max_kWh))

        # Per-year cost savings
        grid_draw = max(load_remaining, 0.0)
        if not bo:
            elec_save_yr += (load_kW - grid_draw) * ELEC_PRICE_EUR_KWH
        else:
            blackout_save_yr += BLACKOUT_SAVE_PER_HR

        # Record Year-1 hourly results for CSV export
        if yr_idx == 0:
            _batt_soc_Wh[i]  = soc_kWh * 1000.0
            _pv_to_batt[i]   = pv_to_batt
            _pv_to_load[i]   = pv_to_load_kW
            _grid_kWh[i]     = max(load_remaining, 0.0)
            _batt_to_load[i] = batt_to_load
            _eta_inv[i]      = eta_inv_h

        soc_tr.append(soc_kWh / nom_cap * 100)
        ac_tr.append(p_ac_out)
        mode_tr.append(mode)

    cycle_tot    += n_bo
    nom_cap_next  = BATT_NOM_KWH * (1 - cycle_tot * DEG_PER_CYCLE)

    wavg = (sum(x[0] * x[1] for x in inv_s) / sum(x[0] for x in inv_s) * 100
            if inv_s else 0.0)

    ann.append({
        'year':            yr,
        'start_soc_pct':   soc_tr[0],
        'end_soc_pct':     soc_tr[-1],
        'nom_cap_kWh':     nom_cap,
        'usable_kWh':      nom_cap * 0.80,
        'ac_yield_kWh':    ac_yield,
        'n_blackouts':     n_bo,
        'bo_ok':           bo_ok,
        'cycle_tot':       cycle_tot,
        'wavg_inv_eff_gt':    wavg,
        'deg_pct':            (1 - nom_cap / BATT_NOM_KWH) * 100,
        'elec_saving_eur':    elec_save_yr,
        'blackout_saving_eur': blackout_save_yr,
        'total_saving_eur':   elec_save_yr + blackout_save_yr,
    })
    soc_all.append(soc_tr)
    ac_all.append(ac_tr)
    mode_all.append(mode_tr)

    print(f"  Year {yr}: yield = {ac_yield:.1f} kWh | "
          f"end SoC = {soc_tr[-1]:.2f}% | "
          f"blackouts = {n_bo} (all covered: {bo_ok == n_bo}) | "
          f"GT η = {wavg:.2f}%")
    if yr_idx == 0:
        print(f"  Peak η (from sim) : {peak_eta:.2f}%  at {peak_pdc_kw:.3f} kW DC")

    nom_cap = nom_cap_next

# ═════════════════════════════════════════════════════════════════════════════
# 7. MONTHLY BREAKDOWN  (Year 1)
# ═════════════════════════════════════════════════════════════════════════════
monthly_gt      = []
monthly_bo_list = []
ac_y1 = np.array(ac_all[0])

for m in range(1, 13):
    mask_gt = (mons == m) & (~is_blackout)
    mask_bo = (mons == m) & (is_blackout)
    monthly_gt.append(float(ac_y1[mask_gt].sum()))
    monthly_bo_list.append(float(ac_y1[mask_bo].sum()))

print(f"\n  Monthly GT sum    : {sum(monthly_gt):.1f} kWh  "
      f"(should match sim: {ann[0]['ac_yield_kWh']:.1f} kWh)")

# ═════════════════════════════════════════════════════════════════════════════
# 8. SAVE PICKLE
# ═════════════════════════════════════════════════════════════════════════════
payload = {
    'ann':          ann,
    'soc_all':      soc_all,
    'ac_all':       ac_all,
    'mode_all':     mode_all,
    'eff_pct':      eff_pct,
    'eff_nom':      eff_nom,
    'G_s':          G_s,
    'G_n':          G_n,
    'P_dc_W':       P_dc_W,
    'is_blackout':  is_blackout,
    'mons':         mons,
    'hrs':          hrs,
    'monthly_gt':   monthly_gt,
    'monthly_bo':   monthly_bo_list,
    'bo_inv_eff':    bo_inv_eff,
    'DC_per_event':  DC_per_event,
    'peak_eta':      peak_eta,
    'peak_pdc_kw':   peak_pdc_kw,
}

with open(PICKLE_OUT, 'wb') as f:
    pickle.dump(payload, f)

print(f"\n✓  Pickle saved → {PICKLE_OUT}")
print("\n─── Key values ──────────────────────────────────────────────────────")
print(f"  Blackout inv η       : {bo_inv_eff*100:.4f}%")
print(f"  DC energy / event    : {DC_per_event:.4f} kWh")
print(f"  Peak inverter η      : {peak_eta:.2f}%")
print(f"  Year-1 GT AC yield   : {ann[0]['ac_yield_kWh']:.1f} kWh")
print(f"  Year-1 GT η (avg)    : {ann[0]['wavg_inv_eff_gt']:.2f}%")
print(f"  Year-3 end SoC       : {ann[2]['end_soc_pct']:.2f}%")
print(f"  Year-3 nom. capacity : {ann[2]['nom_cap_kWh']:.3f} kWh")
final_nom = BATT_NOM_KWH * (1 - 156 * DEG_PER_CYCLE)
print(f"  After-yr3 nom. cap   : {final_nom:.3f} kWh  (81.92 × 0.9688)")

print("\n─── Peak Array Currents ──────────────────────────────────────────────")
print(f"  {'Parameter':<38} {'South':>10}  {'North':>10}")
print(f"  {'-'*62}")
print(f"  {'Peak POA irradiance (W/m²)':<38} {G_s[_idx_s]:>10.1f}  {G_n[_idx_n]:>10.1f}")
print(f"  {'Ambient temperature (°C)':<38} {Ta[_idx_s]:>10.1f}  {Ta[_idx_n]:>10.1f}")
print(f"  {'Cell temperature (°C)':<38} {_Tm_s:>10.1f}  {_Tm_n:>10.1f}")
print(f"  {'Peak DC power (kW)':<38} {P_s_W[_idx_s]/1000:>10.4f}  {P_n_W[_idx_n]/1000:>10.4f}")
_Voc_stc = N_SERIES   * VOC_STC
_Isc_stc = N_PARALLEL * ISC_STC
print(f"  {'Voc – STC (V)':<38} {_Voc_stc:>10.2f}  {_Voc_stc:>10.2f}")
print(f"  {'Voc – temperature derated (V)':<38} {_Voc_s:>10.2f}  {_Voc_n:>10.2f}")
print(f"  {'Vmp – temperature derated (V)':<38} {_Vmp_s:>10.2f}  {_Vmp_n:>10.2f}")
print(f"  {'Isc – STC (A)':<38} {_Isc_stc:>10.2f}  {_Isc_stc:>10.2f}")
print(f"  {'Isc – temperature derated (A)':<38} {_Isc_s:>10.2f}  {_Isc_n:>10.2f}")
print(f"  {'Peak current Imp (A)':<38} {_Ipeak_s:>10.2f}  {_Ipeak_n:>10.2f}")

print("\nNow run  →  Final code.py")

# ═════════════════════════════════════════════════════════════════════════════
# 9. EXPORT battery_load.csv  (same format as load_results.csv)
# ═════════════════════════════════════════════════════════════════════════════
# Electricity savings: grid reduction due to PV, non-blackout hours only
_elec_saving   = np.where(~is_blackout, (load_kW_arr - _grid_kWh) * ELEC_PRICE_EUR_KWH, 0.0)
# Blackout savings: 135 €/hr for every blackout hour the system covers
_blackout_saving = np.where(is_blackout, BLACKOUT_SAVE_PER_HR, 0.0)
_total_saving  = _elec_saving + _blackout_saving

dates_range = pd.date_range(start=f'{YEAR}-01-01', periods=8760, freq='h')
csv_df = pd.DataFrame({
    'Date':                  [d.date() for d in dates_range],
    'Hour':                  [d.hour for d in dates_range],
    'Day':                   [d.strftime('%A') for d in dates_range],
    'Load_kW':               load_kW_arr,
    'BattSoC_Wh':            _batt_soc_Wh,
    'PV_to_Batt_kWh':        _pv_to_batt,
    'PV_to_Load_kWh':        _pv_to_load,
    'PV_total_kWh':          _pv_to_batt + _pv_to_load,
    'Grid_kWh':              _grid_kWh,
    'Batt_to_Load_kWh':      _batt_to_load,
    'eta_inv_vec':           _eta_inv,
    'Elec_Saving_EUR':       _elec_saving,
    'Blackout_Saving_EUR':   _blackout_saving,
    'Total_Saving_EUR':      _total_saving,
})
CSV_OUT = BASE / 'battery_load.csv'
csv_df.to_csv(CSV_OUT, index=False)
print(f"✓  CSV saved  → {CSV_OUT}")

print("\n─── Cost Savings by Year ────────────────────────────────────────────")
print(f"  {'Year':<6} {'PV degr':>8}  {'Elec saving':>14}  {'Blackout saving':>16}  {'Total saving':>14}")
print(f"  {'-'*64}")
for _a, _d in zip(ann, PV_DEGRAD):
    print(f"  {_a['year']:<6} {_d*100:>7.2f}%  "
          f"€ {_a['elec_saving_eur']:>12.2f}  "
          f"€ {_a['blackout_saving_eur']:>14.2f}  "
          f"€ {_a['total_saving_eur']:>12.2f}")
print(f"  {'-'*64}")
print(f"  {'3-yr total':<16}  "
      f"€ {sum(a['elec_saving_eur'] for a in ann):>12.2f}  "
      f"€ {sum(a['blackout_saving_eur'] for a in ann):>14.2f}  "
      f"€ {sum(a['total_saving_eur'] for a in ann):>12.2f}")

# ═════════════════════════════════════════════════════════════════════════════
# 10. BATTERY CURRENT LIMITS
# ═════════════════════════════════════════════════════════════════════════════
_sys_v   = _series   * li_V      # V   – nominal pack voltage  (51.2 V)
_sys_Ah  = _parallel * li_Ah     # Ah  – total pack capacity  (1 600 Ah)

# --- Maximum charging current ---
# Peak PV DC routed to battery (from simulation), then through CC and battery efficiency
_max_pv_to_batt_W  = _pv_to_batt.max() * 1000 * ETA_CC_BATT * ETA_BATT_RT
_I_charge_pv       = _max_pv_to_batt_W / _sys_v            # A
_I_charge_Crate    = 0.5 * _sys_Ah                         # A  (0.5 C limit)
_I_charge_max      = min(_I_charge_pv, _I_charge_Crate)

# --- Maximum discharging current ---
_eta_rated              = PAC0 / PDC0
_I_discharge_inverter   = PDC0 / _sys_v                             # A
_blackout_DC_W          = P_CRITICAL_KW * 1000 / _eta_rated        # W DC
_I_discharge_blackout   = _blackout_DC_W / _sys_v                  # A
_I_discharge_Crate      = 1.0 * _sys_Ah                            # A  (1 C limit)
_I_discharge_max        = min(_I_discharge_inverter, _I_discharge_Crate)

print("\n" + "=" * 55)
print("  BATTERY PACK SPECIFICATION")
print("=" * 55)
print(f"  Configuration         : {_series}S × {_parallel}P  ({_series*_parallel} units total)")
print(f"  Nominal pack voltage  : {_sys_v:.1f} V")
print(f"  Total capacity        : {_sys_Ah:.0f} Ah  /  {_sys_v*_sys_Ah/1000:.2f} kWh")
print(f"  Usable energy (Sys_Wh): {_sys_Wh_min:.0f} Wh")

print()
print("=" * 55)
print("  MAXIMUM CHARGING CURRENT")
print("=" * 55)
print(f"  Peak PV DC → battery   : {_pv_to_batt.max()*1000:.0f} W DC  →  {_max_pv_to_batt_W:.0f} W at battery terminals  @  {_sys_v:.1f} V")
print(f"    → I_charge (PV limit) : {_I_charge_pv:.1f} A")
print(f"  0.5 C rate limit        : {_I_charge_Crate:.1f} A")
print(f"  ► Design max charge I   : {_I_charge_max:.1f} A  "
      f"({'PV-limited' if _I_charge_pv < _I_charge_Crate else 'C-rate limited'})")

print()
print("=" * 55)
print("  MAXIMUM DISCHARGING CURRENT")
print("=" * 55)
print(f"  Inverter rated DC draw  : {PDC0:.1f} W  @  {_sys_v:.1f} V")
print(f"    → I_discharge (inv.)  : {_I_discharge_inverter:.1f} A")
print(f"  Blackout load (6 kW AC) : {_blackout_DC_W:.1f} W DC")
print(f"    → I_discharge (blkout): {_I_discharge_blackout:.1f} A")
print(f"  1 C rate limit          : {_I_discharge_Crate:.1f} A")
print(f"  ► Design max discharge I: {_I_discharge_max:.1f} A  "
      f"({'inverter-limited' if _I_discharge_inverter < _I_discharge_Crate else 'C-rate limited'})")

print()
print("=" * 55)
print("  SUMMARY TABLE")
print("=" * 55)
print(f"  {'Parameter':<30} {'Value':>10}")
print(f"  {'-'*40}")
print(f"  {'Pack voltage':<30} {_sys_v:>9.1f} V")
print(f"  {'Pack capacity':<30} {_sys_Ah:>9.0f} Ah")
print(f"  {'Max charge current':<30} {_I_charge_max:>9.1f} A")
print(f"  {'Max discharge current':<30} {_I_discharge_max:>9.1f} A")
print(f"  {'Max charge power':<30} {_I_charge_max*_sys_v/1000:>9.2f} kW")
print(f"  {'Max discharge power':<30} {_I_discharge_max*_sys_v/1000:>9.2f} kW")
print("=" * 55)
