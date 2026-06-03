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

# ─────────────────────────────────────────────────────────────────────────────
BASE = r'C:\Users\rijkt\OneDrive - Delft University of Technology\SET YR 1\PV systems\ET4378'

INV_FILE     = BASE + r'\Inverter Efficiency parameters.xlsx'
WEATHER_FILE = BASE + r'\Tampa_FL-hour.csv'
PICKLE_OUT   = BASE + r'\sim_final.pkl'

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

ETA_CC         = 0.95          # charge-controller efficiency
ETA_BATT_RT    = 0.92          # battery round-trip efficiency

BATT_NOM_KWH   = 84.48         # 4S×5P Victron LFP 12.8V/330Ah → 20 × 4.224 kWh
SOC_INIT       = 0.95
SOC_MAX        = 0.95
SOC_MIN        = 0.15          # 15 % floor (80 % DoD from 95 % start)
DEG_PER_CYCLE  = 0.0002        # 0.02 % capacity loss per full cycle
P_CRITICAL_KW  = 6.0           # kW critical load during blackout
P_DC0_W        = 6757.76       # W – PDC0 of XW Pro 6848 (gives 88.8 % blackout load)
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
# 5. BLACKOUT OPERATING POINT
# ═════════════════════════════════════════════════════════════════════════════
bo_inv_eff   = inv_eff(P_CRITICAL_KW * 1000)          # fraction at 88.8 % load
DC_per_hour  = P_CRITICAL_KW / bo_inv_eff              # kWh DC drawn per blackout hour
DC_per_event = DC_per_hour * 8                         # kWh DC for one 8-hour event
peak_eta     = float(np.interp(27.5, eff_pct, eff_nom)) * 100  # % at peak-eff point

print(f"  Blackout inv η    : {bo_inv_eff*100:.4f}%  (load = {P_CRITICAL_KW*1000/P_DC0_W*100:.1f}%)")
print(f"  DC per event      : {DC_per_event:.4f} kWh")
print(f"  Peak η (27.5% ld) : {peak_eta:.2f}%")

# ═════════════════════════════════════════════════════════════════════════════
# 6. THREE-YEAR SIMULATION
# ═════════════════════════════════════════════════════════════════════════════
print("\nRunning 3-year simulation …")

ann      = []
soc_all  = []
ac_all   = []
mode_all = []

soc_kWh   = SOC_INIT * BATT_NOM_KWH
nom_cap   = BATT_NOM_KWH
cycle_tot = 0

for yr_idx in range(3):
    yr = YEAR + yr_idx
    ac_yield = 0.0
    n_bo = 0
    bo_ok = 0
    soc_tr  = []
    ac_tr   = []
    mode_tr = []
    inv_s   = []          # (p_ac_out, eta) for GT hours
    in_bo_prev = False

    for i in range(8760):
        soc_pct  = soc_kWh / nom_cap
        p_dc_kw  = P_dc_W[i] / 1000
        p_cc_kw  = p_dc_kw * ETA_CC   # power after charge controller
        bo       = bool(is_blackout[i])

        # Detect start of each blackout event
        if bo and not in_bo_prev:
            n_bo += 1
            if (soc_pct - SOC_MIN) * nom_cap >= DC_per_event:
                bo_ok += 1
        in_bo_prev = bo

        if bo:
            # Battery → inverter → critical load
            soc_kWh  = max(SOC_MIN * nom_cap, soc_kWh - DC_per_hour)
            mode     = "BLACKOUT"
            p_ac_out = P_CRITICAL_KW

        elif soc_pct < SOC_MAX and p_cc_kw > 0:
            # Charging: PV → CC → battery; no GT output
            soc_kWh += min(p_cc_kw * ETA_BATT_RT, (SOC_MAX * nom_cap) - soc_kWh)
            mode     = "CHARGING"
            p_ac_out = 0.0

        elif soc_pct >= SOC_MAX and p_cc_kw > 0:
            # Grid-tied: battery full, PV → CC → inverter → AC
            eta      = inv_eff(p_cc_kw * 1000)
            p_ac_out = p_cc_kw * eta
            ac_yield += p_ac_out
            soc_kWh  = SOC_MAX * nom_cap
            mode     = "GRID_TIED"
            inv_s.append((p_ac_out, eta))

        else:
            # Idle: no PV, no blackout (night)
            mode     = "IDLE"
            p_ac_out = 0.0

        soc_tr.append(soc_kWh / nom_cap * 100)
        ac_tr.append(p_ac_out)
        mode_tr.append(mode)

    # Degrade capacity: each blackout event = 1 full cycle
    cycle_tot    += n_bo
    nom_cap_next  = BATT_NOM_KWH * (1 - cycle_tot * DEG_PER_CYCLE)

    wavg = (sum(x[0] * x[1] for x in inv_s) / sum(x[0] for x in inv_s) * 100
            if inv_s else 0.0)

    ann.append({
        'year':            yr,
        'start_soc_pct':   soc_tr[0],
        'end_soc_pct':     soc_tr[-1],
        'nom_cap_kWh':     nom_cap,              # capacity used during this year
        'usable_kWh':      nom_cap * 0.80,       # 80 % DoD
        'ac_yield_kWh':    ac_yield,
        'n_blackouts':     n_bo,
        'bo_ok':           bo_ok,
        'cycle_tot':       cycle_tot,
        'wavg_inv_eff_gt': wavg,
        'deg_pct':         (1 - nom_cap / BATT_NOM_KWH) * 100,
    })
    soc_all.append(soc_tr)
    ac_all.append(ac_tr)
    mode_all.append(mode_tr)

    print(f"  Year {yr}: yield = {ac_yield:.1f} kWh | "
          f"end SoC = {soc_tr[-1]:.2f}% | "
          f"blackouts = {n_bo} (all covered: {bo_ok == n_bo}) | "
          f"GT η = {wavg:.2f}%")

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
    'bo_inv_eff':   bo_inv_eff,
    'DC_per_event': DC_per_event,
    'peak_eta':     peak_eta,
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
print(f"  After-yr3 nom. cap   : {final_nom:.3f} kWh  (84.48 × 0.9688)")
print("\nNow run  →  Final code.py")
