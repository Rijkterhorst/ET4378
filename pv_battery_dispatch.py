"""
PV + Battery Dispatch Simulation
Translated from MATLAB to Python.

Requirements:
    pip install numpy pandas matplotlib
"""

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from datetime import date
import calendar
import math

# =============================================================================
#  LOAD PROFILE
# =============================================================================

monthly_load = [
    5702.76, 5150.88, 5702.76, 5518.80,
    5702.76, 5518.80, 5702.76, 5702.76,
    5518.80, 5702.76, 5518.80, 5702.76,
]  # kWh per month, Jan–Dec

hourly_factors = [
    3.81, 3.81, 3.88, 3.96, 3.96, 4.03, 4.19, 4.19,
    4.34, 4.34, 4.41, 4.41, 4.49, 4.49, 4.41, 4.41,
    4.34, 4.34, 4.26, 4.19, 4.11, 3.96, 3.88, 3.81,
]

hourly_factors = np.array(hourly_factors)
hourly_factors_norm = hourly_factors / hourly_factors.sum()

year = 2025

# Build hourly date range for the full year
dates = pd.date_range(start=f"{year}-01-01", end=f"{year}-12-31", freq="D")

rows = []
for current_date in dates:
    m = current_date.month
    days_in_month = calendar.monthrange(year, m)[1]
    daily_load = monthly_load[m - 1] / days_in_month
    hourly_load = daily_load * hourly_factors_norm  # shape (24,)
    day_name = current_date.strftime("%A")
    for h in range(24):
        rows.append({
            "Date": current_date.date(),
            "Hour": h,
            "Day": day_name,
            "Load_kW": hourly_load[h],
        })

load_profile = pd.DataFrame(rows)

# Blackout windows: Thursday >= 20:00 and Friday < 04:00
idx_thu = (load_profile["Day"] == "Thursday") & (load_profile["Hour"] >= 20)
idx_fri = (load_profile["Day"] == "Friday")   & (load_profile["Hour"] < 4)

load_profile.loc[idx_thu, "Load_kW"] = 6.0
load_profile.loc[idx_fri, "Load_kW"] = 6.0

# =============================================================================
#  BATTERY COMPARISON CHARTS
# =============================================================================

types      = ["AGM", "DCG", "2V", "12.8 Li-ion", "25.6 Li-ion", "Lead-C"]
tot_cost   = [13580, 11640, 11640, 19800, 15840, 21340]
tot_weight = [1876,  1800,  5640,  580,   624,   2420]
tot_volume = [784,   696,   2439.12, 258.6, 401.76, 10950.72]
cycle_life = [450,   500,   1500,  2500,  2500,  500]

fig, axes = plt.subplots(2, 2, figsize=(14, 9))
fig.suptitle("Battery Type Comparison", fontsize=14)

specs = [
    (axes[0, 0], tot_cost,   "Cost (€)",    "Total Cost"),
    (axes[0, 1], tot_weight, "Weight (kg)",  "Total Weight"),
    (axes[1, 0], tot_volume, "Volume (L)",   "Total Volume"),
    (axes[1, 1], cycle_life, "Cycles",       "Cycle Life"),
]

colors = plt.cm.tab10(np.linspace(0, 1, len(types)))

for ax, data, ylabel, title in specs:
    bars = ax.bar(types, data, color=colors)
    ax.set_ylabel(ylabel)
    ax.set_title(title)
    ax.grid(True, axis="y")
    ax.tick_params(axis="x", rotation=20)

plt.tight_layout()
plt.savefig("battery_comparison.png", dpi=150)
plt.show()

# =============================================================================
#  SYSTEM SIZING
# =============================================================================

sys_v         = 48          # V
sys_Wh_nodod  = 52866       # Wh (without DoD correction)
DoD           = 0.75
Sys_Wh        = sys_Wh_nodod / (DoD * 0.961)
print(f"Corrected system energy: {Sys_Wh:.1f} Wh")

# Li-ion battery sizing
li_V    = 25.6
Li_Ah   = 200
Li_Wh   = li_V * Li_Ah

series         = math.ceil(sys_v / li_V)
parallel       = math.ceil(Sys_Wh / (series * Li_Wh))
total_batteries = series * parallel

print(f"Series: {series}  |  Parallel: {parallel}  |  Total batteries: {total_batteries}")

# =============================================================================
#  SYSTEM / INVERTER PARAMETERS
# =============================================================================

Eff_bat      = 0.92   # battery round-trip efficiency
PDC0         = 6757.76  # W – DC input at rated output
PAC0         = 6120.0   # W – rated AC output
Ps0          = 42.0     # W – self-consumption / standby
C0           = -0.000012  # curvature coefficient
Eff_CC_grid  = 0.95   # charge controller → grid/load efficiency
Eff_cc_batt  = 0.92   # charge controller → battery efficiency

# =============================================================================
#  PV DATA  – computed from weather file (Tampa_FL-hour.csv has no PV column)
# =============================================================================

import math as _math

_wx = pd.read_csv(
    r'C:\Users\rijkt\OneDrive - Delft University of Technology\SET YR 1\PV systems\ET4378\Tampa_FL-hour.csv',
    sep=';')
_wx.columns = _wx.columns.str.strip()
_wx['Az_std'] = (_wx['Az'] + 180) % 360   # → standard compass (0 = North)

# AIKO Neostar 490 W module parameters
_PMAX   = 490.0    # W
_NOCT   = 46.0     # °C  (AIKO datasheet)
_TC     = -0.0026  # /°C
_TILT   = 35.0
_ALBEDO = 0.20
_N_MODS = 24       # 12 south + 12 north

def _poa(hs, az_std, G_bn, G_dh, G_gh, surf_az):
    if hs <= 0:
        return 0.0
    tr  = _math.radians(_TILT)
    zen = _math.radians(90 - hs)
    daz = _math.radians(az_std - surf_az)
    beam = max(0.0, G_bn * (_math.cos(zen)*_math.cos(tr)
                             + _math.sin(zen)*_math.sin(tr)*_math.cos(daz)))
    return max(0.0, beam + G_dh*(1+_math.cos(tr))/2 + G_gh*_ALBEDO*(1-_math.cos(tr))/2)

def _pv(G, Ta):
    if G <= 0:
        return 0.0
    Tc = Ta + ((_NOCT - 20) / 800) * G
    return _PMAX * (G / 1000) * max(0.0, 1 + _TC * (Tc - 25))

_G_s = [_poa(_wx['hs'][i], _wx['Az_std'][i], _wx['G_Bn'][i],
              _wx['G_Dh'][i], _wx['G_Gh'][i], 180) for i in range(8760)]
_G_n = [_poa(_wx['hs'][i], _wx['Az_std'][i], _wx['G_Bn'][i],
              _wx['G_Dh'][i], _wx['G_Gh'][i], 0)   for i in range(8760)]
_Ta  = _wx['Ta'].values

# pvgen: combined DC power [W] from both slopes (12 modules each)
pvgen = np.array([12*_pv(_G_s[i], _Ta[i]) + 12*_pv(_G_n[i], _Ta[i])
                  for i in range(8760)])

# =============================================================================
#  DISPATCH PARAMETERS
# =============================================================================

bat_max = 0.95 * Sys_Wh   # upper SoC limit (Wh)
bat_min = 0.20 * Sys_Wh   # lower SoC limit (Wh)
n_hours = len(load_profile)

inverter_max_DC_kW = 7.489   # kW

# =============================================================================
#  PRE-ALLOCATE OUTPUT ARRAYS
# =============================================================================

batt_soc         = np.zeros(n_hours)
pv_to_battery    = np.zeros(n_hours)
pv_to_load       = np.zeros(n_hours)
energy_from_pv   = np.zeros(n_hours)
grid_to_load     = np.zeros(n_hours)
batt_to_load_kWh = np.zeros(n_hours)
eta_inv_vec      = np.zeros(n_hours)

# Convert blackout boolean series to numpy arrays
is_blackout_arr = (idx_thu | idx_fri).values

# =============================================================================
#  HOURLY DISPATCH LOOP
# =============================================================================

soc = bat_max  # initial SoC

for h in range(n_hours):
    load_kW     = load_profile["Load_kW"].iloc[h]
    pv_kW       = pvgen[h] / 1000.0
    is_blackout = is_blackout_arr[h]

    # Clamp PV to inverter DC input limit
    pv_kW = min(pv_kW, inverter_max_DC_kW)

    pv_remaining   = pv_kW
    load_remaining = load_kW

    # Previous inverter efficiency (fallback for hour 0)
    eta_inv_h = eta_inv_vec[h - 1] if h > 0 else 0.0
    if eta_inv_h == 0:
        eta_inv_h = 0.92

    # ------------------------------------------------------------------
    #  STEP 1 – Battery discharge during blackout
    # ------------------------------------------------------------------
    if is_blackout and soc > bat_min:
        batt_available_kWh = (soc - bat_min) / 1000.0
        batt_needed_kWh    = load_remaining / eta_inv_h
        batt_used_kWh      = min(batt_available_kWh, batt_needed_kWh)
        load_covered_by_batt = batt_used_kWh * eta_inv_h

        soc                    -= batt_used_kWh * 1000.0
        batt_to_load_kWh[h]    = batt_used_kWh
        load_remaining         -= load_covered_by_batt

    # ------------------------------------------------------------------
    #  STEP 2 – PV → Battery
    # ------------------------------------------------------------------
    if pv_remaining > 0 and soc < bat_max:
        room_kWh_batt  = (bat_max - soc) / 1000.0
        room_kWh_pv    = room_kWh_batt / (Eff_cc_batt * Eff_bat)
        pv_used_charge = min(pv_remaining, room_kWh_pv)
        charge_stored  = pv_used_charge * Eff_cc_batt * Eff_bat

        soc              += charge_stored * 1000.0
        pv_to_battery[h]  = pv_used_charge
        pv_remaining      -= pv_used_charge

    # ------------------------------------------------------------------
    #  STEP 3 – PV → Load
    # ------------------------------------------------------------------
    if pv_remaining > 0 and load_remaining > 0:
        pv_needed_for_load  = load_remaining / (Eff_CC_grid * eta_inv_h)
        pv_used_load        = min(pv_remaining, pv_needed_for_load)
        load_covered_by_pv  = pv_used_load * Eff_CC_grid * eta_inv_h

        pv_to_load[h]   = pv_used_load
        pv_remaining    -= pv_used_load
        load_remaining  -= load_covered_by_pv

    # ------------------------------------------------------------------
    #  Compute variable inverter efficiency (Sandia model)
    # ------------------------------------------------------------------
    P_DC_inv = (pv_to_load[h] * Eff_CC_grid + batt_to_load_kWh[h]) * 1000.0  # W
    if P_DC_inv > Ps0:
        PAC_h = (
            (PAC0 / (PDC0 - Ps0) - C0 * (PDC0 - Ps0)) * (P_DC_inv - Ps0)
            + C0 * (P_DC_inv - Ps0) ** 2
        )
        PAC_h     = max(0.0, min(PAC_h, PAC0))
        eta_inv_h = PAC_h / P_DC_inv
    else:
        eta_inv_h = 0.0

    eta_inv_vec[h] = eta_inv_h

    # ------------------------------------------------------------------
    #  STEP 4 – Grid covers remaining load
    # ------------------------------------------------------------------
    grid_to_load[h] = max(load_remaining, 0.0)

    # ------------------------------------------------------------------
    #  STORE RESULTS
    # ------------------------------------------------------------------
    batt_soc[h]       = np.clip(soc, bat_min, bat_max)
    soc               = batt_soc[h]   # carry clamped SoC forward
    energy_from_pv[h] = pv_to_battery[h] + pv_to_load[h]

# =============================================================================
#  ATTACH RESULTS TO TABLE
# =============================================================================

load_profile["BattSoC_Wh"]       = batt_soc
load_profile["PV_to_Batt_kWh"]   = pv_to_battery
load_profile["PV_to_Load_kWh"]   = pv_to_load
load_profile["PV_total_kWh"]     = energy_from_pv
load_profile["Grid_kWh"]         = grid_to_load
load_profile["Batt_to_Load_kWh"] = batt_to_load_kWh
load_profile["eta_inv_vec"]      = eta_inv_vec

# =============================================================================
#  ANNUAL ENERGY SUMMARY
# =============================================================================

print("\n=== Annual Energy Summary ===")
print(f"Total Load              : {load_profile['Load_kW'].sum():.1f} kWh")
print(f"Grid supplied           : {grid_to_load.sum():.1f} kWh")
print(f"PV to Load              : {(pv_to_load * Eff_CC_grid * eta_inv_vec).sum():.1f} kWh")
print(f"Battery to Load         : {(batt_to_load_kWh * eta_inv_vec).sum():.1f} kWh")
print(f"PV to Battery (in)      : {(pv_to_battery * Eff_cc_batt * Eff_bat).sum():.1f} kWh")
pv_curtailed = pvgen.sum() / 1000.0 - energy_from_pv.sum()
print(f"PV curtailed            : {pv_curtailed:.1f} kWh")
nonzero_eta = eta_inv_vec[eta_inv_vec > 0]
print(f"Mean inverter efficiency: {nonzero_eta.mean() * 100:.2f}%")

# Minimum SoC stats
min_soc = batt_soc.min()
print(f"\nMin battery SoC         : {min_soc:.1f} Wh")
print(f"Min battery SoC (frac)  : {min_soc / Sys_Wh:.4f}")

print(f"\nMax PV to Load (single hour): {pv_to_load.max():.4f} kWh")

# =============================================================================
#  EXPORT
# =============================================================================

load_profile.to_csv("load_results.csv", index=False)
print("\nResults saved to load_results.csv")
