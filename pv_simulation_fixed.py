#!/usr/bin/env python
# coding: utf-8

# # ET4378 – Group 20 | Questions 4, 5 & 6
# **Tampa, FL, USA | Load Profile 3 | Both roof slopes used**

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from pathlib import Path
from scipy.interpolate import interp1d
from datetime import date
import calendar
import math

plt.rcParams['figure.figsize'] = (13, 4)
plt.rcParams['figure.dpi'] = 100
print('Libraries loaded.')

# ---
# ## 1. Configuration

# BUG FIX 1: Use raw strings (r"...") for Windows paths to prevent backslash
# escape sequences (e.g. \U, \D, \p being misinterpreted as Unicode escapes).
WEATHER_FILE = Path(r'C:\Users\rijkt\OneDrive - Delft University of Technology\SET YR 1\PV systems\ET4378\Tampa_FL-hour.csv')
INV_FILE     = Path(r'C:\Users\rijkt\OneDrive - Delft University of Technology\SET YR 1\PV systems\ET4378\Inverter Efficiency parameters.xlsx')

# ── Location ──────────────────────────────────────────────────────────────────
LATITUDE  = 27.95

# ── Roof ──────────────────────────────────────────────────────────────────────
ROOF_FULL_WIDTH_M    = 5.0
ROOF_RIDGE_M         = 18.0
ROOF_TILT_DEG        = 35.0
SAFETY_MARGIN_M      = 0.5
CLUSTER_GAP_M        = 0.5
MAX_ROWS_PER_CLUSTER = 4
MAX_COLS_PER_CLUSTER = 4

# ── PV Module: AIKO Neostar 2S ────────────────────────────────────────────────
MODULE_PMAX_STC   = 490.0
MODULE_NOCT       = 46.0
MODULE_TEMP_COEFF = -0.0026
MODULE_EFF_STC    = 0.245
MODULE_LENGTH_M   = 1.762
MODULE_WIDTH_M    = 1.134

# ── Inverter: Schneider XW Pro 6848 ──────────────────────────────────────────
INVERTER_RATED_W  = 6020.0

# ── Charge controller efficiency ─────────────────────────────────────────────
CC_EFFICIENCY     = 0.95

# ── Critical load (Group 20 = Load 3) ────────────────────────────────────────
CRITICAL_LOAD_KW  = 6.0

print('Configuration loaded.')

# ---
# ## 2. Weather Data

weather = pd.read_csv(WEATHER_FILE, sep=';')
weather.columns = weather.columns.str.strip()

date_range = pd.date_range(start='2005-01-01 00:00', periods=8760, freq='h')
weather.index = date_range

print(f'Rows loaded  : {len(weather)}')
print(f'Columns      : {list(weather.columns)}')
print(f'Annual GHI   : {weather["G_Gh"].sum()/1000:.1f} kWh/m²')
print(f'Peak GHI     : {weather["G_Gh"].max():.0f} W/m²')
print(f'Avg temp     : {weather["Ta"].mean():.1f} °C')

fig, axes = plt.subplots(2, 1, figsize=(13, 7))
monthly_ghi  = weather['G_Gh'].resample('ME').sum() / 1000
monthly_temp = weather['Ta'].resample('ME').mean()
months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']

axes[0].bar(range(12), monthly_ghi.values, color='gold', edgecolor='white')
axes[0].set_xticks(range(12)); axes[0].set_xticklabels(months)
axes[0].set_ylabel('GHI [kWh/m²/month]')
axes[0].set_title('Monthly Global Horizontal Irradiation – Tampa, FL')

axes[1].plot(range(12), monthly_temp.values, marker='o', color='tomato')
axes[1].set_xticks(range(12)); axes[1].set_xticklabels(months)
axes[1].set_ylabel('Avg Temperature [°C]')
axes[1].set_title('Monthly Average Ambient Temperature – Tampa, FL')

plt.tight_layout(); plt.show()

# ---
# ## 3. Question 5 – Roof Layout

tilt_rad       = np.radians(ROOF_TILT_DEG)
half_plan_m    = ROOF_FULL_WIDTH_M / 2
slant_m        = half_plan_m / np.cos(tilt_rad)
usable_slant_m = slant_m - 2 * SAFETY_MARGIN_M
usable_ridge_m = ROOF_RIDGE_M - 2 * SAFETY_MARGIN_M

print('── Roof geometry (one slope) ──────────────────────')
print(f'  Half plan width          : {half_plan_m:.2f} m')
print(f'  Slant length             : {slant_m:.3f} m')
print(f'  Usable slant (−2×0.5 m) : {usable_slant_m:.3f} m')
print(f'  Usable ridge (−2×0.5 m) : {usable_ridge_m:.2f} m')

rows_per_cluster = int(usable_slant_m // MODULE_LENGTH_M)
rows_per_cluster = min(rows_per_cluster, MAX_ROWS_PER_CLUSTER)
cluster_slant_m  = rows_per_cluster * MODULE_LENGTH_M

print(f'\n── Along slope (portrait: {MODULE_LENGTH_M} m per module) ─────')
print(f'  Modules that fit         : {usable_slant_m:.3f} / {MODULE_LENGTH_M} = {usable_slant_m/MODULE_LENGTH_M:.2f} → {rows_per_cluster} row(s) per cluster')
print(f'  Cluster height           : {cluster_slant_m:.3f} m')

cols_per_cluster = MAX_COLS_PER_CLUSTER
cluster_width_m  = cols_per_cluster * MODULE_WIDTH_M
n_clusters       = int((usable_ridge_m + CLUSTER_GAP_M) / (cluster_width_m + CLUSTER_GAP_M))
total_width_used = n_clusters * cluster_width_m + (n_clusters - 1) * CLUSTER_GAP_M

print(f'\n── Along ridge (module width: {MODULE_WIDTH_M} m) ──────────────')
print(f'  Cluster width (4 cols)   : {cluster_width_m:.3f} m')
print(f'  Number of clusters       : {n_clusters}')
print(f'  Total width used         : {total_width_used:.3f} m  (of {usable_ridge_m:.1f} m available)')

mods_per_slope = n_clusters * cols_per_cluster * rows_per_cluster
n_slopes       = 2
total_modules  = mods_per_slope * n_slopes
total_power_wp = total_modules * MODULE_PMAX_STC

print(f'\n── Module count ──────────────────────────────────────')
print(f'  Layout per slope         : {n_clusters} clusters × {cols_per_cluster} cols × {rows_per_cluster} row = {mods_per_slope} modules')
print(f'  Both slopes (S + N)      : 2 × {mods_per_slope} = {total_modules} modules')
print(f'  Total installed peak     : {total_modules} × {MODULE_PMAX_STC:.0f} W = {total_power_wp/1000:.2f} kWp')

# ---
# ## 4. Question 4 – POA Irradiance

def compute_poa(weather_df, surface_tilt_deg, surface_azimuth_deg, albedo=0.20):
    tilt    = np.radians(surface_tilt_deg)
    surf_az = np.radians(surface_azimuth_deg)
    sol_az  = np.radians(weather_df['Az'])
    sol_el  = np.radians(weather_df['hs'])

    cos_theta_i = (
        np.sin(sol_el) * np.cos(tilt)
        + np.cos(sol_el) * np.sin(tilt) * np.cos(sol_az - surf_az)
    )
    cos_theta_i = np.clip(cos_theta_i, 0, 1)

    G_beam      = weather_df['G_Bn'] * cos_theta_i
    G_diffuse   = weather_df['G_Dh'] * (1 + np.cos(tilt)) / 2
    G_reflected = weather_df['G_Gh'] * albedo * (1 - np.cos(tilt)) / 2

    return (G_beam + G_diffuse + G_reflected).clip(lower=0)

weather['G_poa_south'] = compute_poa(weather, ROOF_TILT_DEG, surface_azimuth_deg=0)
weather['G_poa_north'] = compute_poa(weather, ROOF_TILT_DEG, surface_azimuth_deg=180)

print('Annual irradiation summary:')
print(f'  GHI  (horizontal)  : {weather["G_Gh"].sum()/1000:.1f} kWh/m²/yr')
print(f'  POA  South slope   : {weather["G_poa_south"].sum()/1000:.1f} kWh/m²/yr')
print(f'  POA  North slope   : {weather["G_poa_north"].sum()/1000:.1f} kWh/m²/yr')

monthly_poa_s = weather['G_poa_south'].resample('ME').sum() / 1000
monthly_poa_n = weather['G_poa_north'].resample('ME').sum() / 1000
monthly_ghi   = weather['G_Gh'].resample('ME').sum() / 1000

x = np.arange(12)
w = 0.25
fig, ax = plt.subplots(figsize=(13, 4))
ax.bar(x - w, monthly_ghi.values,   w, label='GHI (horizontal)', color='gold',       edgecolor='white')
ax.bar(x,     monthly_poa_s.values, w, label='POA – South 35°',  color='darkorange',  edgecolor='white')
ax.bar(x + w, monthly_poa_n.values, w, label='POA – North 35°',  color='steelblue',   edgecolor='white')
ax.set_xticks(x); ax.set_xticklabels(months)
ax.set_ylabel('Irradiation [kWh/m²/month]')
ax.set_title('Monthly GHI vs POA – Tampa, FL (both slopes, 35° tilt)')
ax.legend()
plt.tight_layout(); plt.show()

print(f'South captures {(weather["G_poa_south"].sum()/weather["G_poa_north"].sum()-1)*100:.1f}% more irradiation than North per year.')

# ---
# ## 5. Question 4 – PV Power Output

def calc_pv_power(G_poa, T_amb, n_modules, pmax_stc, noct, gamma):
    T_cell = T_amb + ((noct - 20) / 800) * G_poa
    P_dc   = n_modules * pmax_stc * (G_poa / 1000) * (1 + gamma * (T_cell - 25))
    return P_dc.clip(lower=0)

weather['P_dc_south_W'] = calc_pv_power(
    weather['G_poa_south'], weather['Ta'],
    mods_per_slope, MODULE_PMAX_STC, MODULE_NOCT, MODULE_TEMP_COEFF
)
weather['P_dc_north_W'] = calc_pv_power(
    weather['G_poa_north'], weather['Ta'],
    mods_per_slope, MODULE_PMAX_STC, MODULE_NOCT, MODULE_TEMP_COEFF
)
weather['P_dc_total_W'] = weather['P_dc_south_W'] + weather['P_dc_north_W']

E_dc_south_kwh = weather['P_dc_south_W'].sum() / 1000
E_dc_north_kwh = weather['P_dc_north_W'].sum() / 1000
E_dc_total_kwh = weather['P_dc_total_W'].sum() / 1000

print('─── Annual DC energy output ────────────────────────────')
print(f'  South slope ({mods_per_slope} modules) : {E_dc_south_kwh:,.0f} kWh/yr')
print(f'  North slope ({mods_per_slope} modules) : {E_dc_north_kwh:,.0f} kWh/yr')
print(f'  TOTAL ({total_modules} modules)        : {E_dc_total_kwh:,.0f} kWh/yr')
print(f'  Peak DC power               : {weather["P_dc_total_W"].max()/1000:.2f} kW')
print(f'  Installed peak power        : {total_power_wp/1000:.2f} kWp')
print(f'  Specific yield              : {E_dc_total_kwh/(total_power_wp/1000):.0f} kWh/kWp/yr')

monthly_dc_s = weather['P_dc_south_W'].resample('ME').sum() / 1000
monthly_dc_n = weather['P_dc_north_W'].resample('ME').sum() / 1000

fig, ax = plt.subplots(figsize=(13, 4))
ax.bar(x - w/2, monthly_dc_s.values, w, label=f'South slope ({mods_per_slope} mod)', color='darkorange', edgecolor='white')
ax.bar(x + w/2, monthly_dc_n.values, w, label=f'North slope ({mods_per_slope} mod)', color='steelblue',  edgecolor='white')
ax.set_xticks(x); ax.set_xticklabels(months)
ax.set_ylabel('DC Energy [kWh/month]')
ax.set_title(f'Monthly DC Energy Output – {total_modules} modules total ({mods_per_slope} per slope)')
ax.legend()
plt.tight_layout(); plt.show()

# ---
# ## 6. Question 4 – Inverter Selection

peak_dc_kw  = weather['P_dc_total_W'].max() / 1000
dc_ac_ratio = (total_power_wp / 1000) / (INVERTER_RATED_W / 1000)

print('─── Inverter sizing check ──────────────────────────────')
print(f'  Installed peak DC          : {total_power_wp/1000:.2f} kWp')
print(f'  Simulated peak DC power    : {peak_dc_kw:.2f} kW')
print(f'  Inverter rated power       : {INVERTER_RATED_W/1000:.1f} kW')
print(f'  DC/AC ratio                : {dc_ac_ratio:.2f}')
print(f'  Critical load to cover     : {CRITICAL_LOAD_KW:.1f} kW  ✓')

load_pct         = (weather['P_dc_total_W'].clip(upper=INVERTER_RATED_W) / INVERTER_RATED_W * 100)
daytime_load_pct = load_pct[load_pct > 0]

fig, ax = plt.subplots(figsize=(10, 4))
ax.hist(daytime_load_pct, bins=50, color='steelblue', edgecolor='white')
ax.axvline(CRITICAL_LOAD_KW / INVERTER_RATED_W * 100, color='red', linestyle='--',
           label=f'Blackout load ({CRITICAL_LOAD_KW:.0f} kW = {CRITICAL_LOAD_KW/INVERTER_RATED_W*100:.1f}% rated)')
ax.set_xlabel('Inverter load [% of rated]')
ax.set_ylabel('Hours per year')
ax.set_title('Distribution of Inverter Load – Grid-Tied Mode (daytime hours)')
ax.legend()
plt.tight_layout(); plt.show()

# ---
# ## 7. Question 6 – Inverter Efficiency

inv_eff = pd.read_excel(INV_FILE, sheet_name='Plot Data')
inv_eff.columns = ['pct_rated', 'eff_nominal', 'eff_low', 'eff_high'] + list(inv_eff.columns[4:])
inv_eff = inv_eff[['pct_rated', 'eff_nominal', 'eff_low', 'eff_high']].dropna()
inv_eff = inv_eff[inv_eff['pct_rated'] >= 0]

inv_eff_interp = interp1d(
    inv_eff['pct_rated'],
    inv_eff['eff_nominal'],
    bounds_error=False,
    fill_value=(inv_eff['eff_nominal'].iloc[0], inv_eff['eff_nominal'].iloc[-1])
)

fig, ax = plt.subplots(figsize=(10, 4))
ax.plot(inv_eff['pct_rated'], inv_eff['eff_nominal'], color='steelblue', linewidth=2.5, label='Nominal voltage')
ax.plot(inv_eff['pct_rated'], inv_eff['eff_low'],     color='tomato',    linewidth=1.5, linestyle='--', label='Low voltage')
ax.plot(inv_eff['pct_rated'], inv_eff['eff_high'],    color='green',     linewidth=1.5, linestyle='--', label='High voltage')

blackout_pct = CRITICAL_LOAD_KW * 1000 / INVERTER_RATED_W * 100
blackout_eff = float(inv_eff_interp(blackout_pct))
ax.axvline(blackout_pct, color='red', linestyle=':', linewidth=1.5)
ax.scatter([blackout_pct], [blackout_eff], color='red', zorder=5, s=80)
ax.annotate(f'Blackout point\n{blackout_pct:.1f}% load → {blackout_eff:.1f}%',
            xy=(blackout_pct, blackout_eff),
            xytext=(blackout_pct + 5, blackout_eff - 3),
            fontsize=9, color='red',
            arrowprops=dict(arrowstyle='->', color='red'))

ax.set_xlabel('Output power [% of rated]')
ax.set_ylabel('Efficiency [%]')
ax.set_title('XW Pro 6848 – Efficiency vs Load')
ax.legend()
ax.grid(True, alpha=0.3)
plt.tight_layout(); plt.show()

peak_idx = inv_eff['eff_nominal'].idxmax()
print(f'Peak efficiency : {inv_eff.loc[peak_idx, "eff_nominal"]:.2f}% at {inv_eff.loc[peak_idx, "pct_rated"]:.1f}% load')
print(f'Blackout load   : {blackout_pct:.1f}% rated → efficiency = {blackout_eff:.2f}%')

# Grid-tied average efficiency
P_inv_input_W    = weather['P_dc_total_W'].clip(upper=INVERTER_RATED_W)
pv_hours         = P_inv_input_W > 0
load_pct_hourly  = (P_inv_input_W[pv_hours] / INVERTER_RATED_W) * 100
eff_hourly       = inv_eff_interp(load_pct_hourly.values) / 100.0

energy_weighted_eff = np.average(eff_hourly, weights=P_inv_input_W[pv_hours].values)
simple_mean_eff     = eff_hourly.mean()

print('─── Q6: Inverter Efficiency Analysis ───────────────────')
print(f'  Hours with PV output       : {pv_hours.sum()} h/yr')
print(f'  Avg load (unweighted)      : {load_pct_hourly.mean():.1f}% of rated')
print(f'  Simple mean efficiency     : {simple_mean_eff*100:.2f}%')
print(f'  Energy-weighted efficiency : {energy_weighted_eff*100:.2f}%')
print(f'  Blackout efficiency        : {blackout_eff:.2f}%')

fig, ax = plt.subplots(figsize=(10, 4))
ax.hist(eff_hourly * 100, bins=40, color='steelblue', edgecolor='white', label='Grid-tied hours')
ax.axvline(energy_weighted_eff * 100, color='navy', linewidth=2,
           label=f'Energy-weighted avg: {energy_weighted_eff*100:.2f}%')
ax.axvline(blackout_eff, color='red', linewidth=2, linestyle='--',
           label=f'Blackout fixed point: {blackout_eff:.2f}%')
ax.set_xlabel('Inverter efficiency [%]')
ax.set_ylabel('Hours per year')
ax.set_title('Inverter Efficiency Distribution – Grid-Tied Hours')
ax.legend()
plt.tight_layout(); plt.show()

# =============================================================================
#  LOAD PROFILE
# =============================================================================

monthly_load = [
    5702.76, 5150.88, 5702.76, 5518.80,
    5702.76, 5518.80, 5702.76, 5702.76,
    5518.80, 5702.76, 5518.80, 5702.76,
]

hourly_factors = np.array([
    3.81, 3.81, 3.88, 3.96, 3.96, 4.03, 4.19, 4.19,
    4.34, 4.34, 4.41, 4.41, 4.49, 4.49, 4.41, 4.41,
    4.34, 4.34, 4.26, 4.19, 4.11, 3.96, 3.88, 3.81,
])
hourly_factors_norm = hourly_factors / hourly_factors.sum()

year  = 2025
dates = pd.date_range(start=f"{year}-01-01", end=f"{year}-12-31", freq="D")

rows = []
for current_date in dates:
    m             = current_date.month
    days_in_month = calendar.monthrange(year, m)[1]
    daily_load    = monthly_load[m - 1] / days_in_month
    hourly_load   = daily_load * hourly_factors_norm
    day_name      = current_date.strftime("%A")
    for h in range(24):
        rows.append({
            "Date":    current_date.date(),
            "Hour":    h,
            "Day":     day_name,
            "Load_kW": hourly_load[h],
        })

load_profile = pd.DataFrame(rows)

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

specs  = [
    (axes[0, 0], tot_cost,   "Cost (€)",    "Total Cost"),
    (axes[0, 1], tot_weight, "Weight (kg)",  "Total Weight"),
    (axes[1, 0], tot_volume, "Volume (L)",   "Total Volume"),
    (axes[1, 1], cycle_life, "Cycles",       "Cycle Life"),
]
colors = plt.cm.tab10(np.linspace(0, 1, len(types)))

for ax, data, ylabel, title in specs:
    ax.bar(types, data, color=colors)
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

# Victron LFP Smart 12.8 V / 330 Ah  –  4S × 5P = 20 units  (84.48 kWh nominal)
sys_v           = 51.2   # V  – 4 × 12.8 V
li_V            = 12.8   # V  – per Victron LFP unit
Li_Ah           = 330    # Ah – per unit
Li_Wh           = li_V * Li_Ah                    # 4 224 Wh per unit
series          = 4      # units in series  → 51.2 V
parallel        = 5      # strings in parallel → 1 650 Ah
total_batteries = series * parallel                # 20 units
Sys_Wh          = total_batteries * Li_Wh         # 84 480 Wh nominal

print(f"Battery: {series}S × {parallel}P = {total_batteries} × Victron LFP 12.8V/330Ah")
print(f"Nominal energy: {Sys_Wh/1000:.2f} kWh  |  Bus voltage: {sys_v:.1f} V")

# =============================================================================
#  SYSTEM / INVERTER PARAMETERS
# =============================================================================

Eff_bat      = 0.92
PDC0         = 6757.76
PAC0         = 6120.0
Ps0          = 42.0
C0           = -0.000012
Eff_CC_grid  = 0.95
Eff_cc_batt  = 0.92

# =============================================================================
#  PV DATA
# BUG FIX 2: Use .iloc[h] instead of [h] to avoid KeyError when the Series
# has a DatetimeIndex.  Integer subscript on a DatetimeIndex does NOT give
# positional access; it raises KeyError.
# =============================================================================

pvgen = weather['P_dc_total_W']

# =============================================================================
#  DISPATCH PARAMETERS
# =============================================================================

bat_max         = 0.95 * Sys_Wh
bat_min         = 0.20 * Sys_Wh
n_hours         = len(load_profile)
inverter_max_DC_kW = 7.489

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

is_blackout_arr = (idx_thu | idx_fri).values

# =============================================================================
#  HOURLY DISPATCH LOOP
# =============================================================================

soc = bat_max

for h in range(n_hours):
    load_kW     = load_profile["Load_kW"].iloc[h]
    # BUG FIX 2: use .iloc[h] for positional access on DatetimeIndex Series
    pv_kW       = pvgen.iloc[h] / 1000.0
    is_blackout = is_blackout_arr[h]

    pv_kW = min(pv_kW, inverter_max_DC_kW)

    pv_remaining   = pv_kW
    load_remaining = load_kW

    eta_inv_h = eta_inv_vec[h - 1] if h > 0 else 0.92
    if eta_inv_h == 0:
        eta_inv_h = 0.92

    # ------------------------------------------------------------------
    #  STEP 1 – Battery discharge during blackout
    # ------------------------------------------------------------------
    if is_blackout and soc > bat_min:
        batt_available_kWh   = (soc - bat_min) / 1000.0
        batt_needed_kWh      = load_remaining / eta_inv_h
        batt_used_kWh        = min(batt_available_kWh, batt_needed_kWh)
        load_covered_by_batt = batt_used_kWh * eta_inv_h

        soc                  -= batt_used_kWh * 1000.0
        batt_to_load_kWh[h]   = batt_used_kWh
        load_remaining        -= load_covered_by_batt

    # ------------------------------------------------------------------
    # BUG FIX 3 – Correct dispatch priority: PV → Load BEFORE PV → Battery.
    # Original code charged the battery first, leaving less PV for the load
    # and causing avoidable grid draw.
    # ------------------------------------------------------------------

    # STEP 2 – PV → Load (MOVED BEFORE battery charging)
    if pv_remaining > 0 and load_remaining > 0:
        pv_needed_for_load = load_remaining / (Eff_CC_grid * eta_inv_h)
        pv_used_load       = min(pv_remaining, pv_needed_for_load)
        load_covered_by_pv = pv_used_load * Eff_CC_grid * eta_inv_h

        pv_to_load[h]   = pv_used_load
        pv_remaining    -= pv_used_load
        load_remaining  -= load_covered_by_pv

    # STEP 3 – PV → Battery (excess after load is met)
    if pv_remaining > 0 and soc < bat_max:
        room_kWh_batt  = (bat_max - soc) / 1000.0
        room_kWh_pv    = room_kWh_batt / (Eff_cc_batt * Eff_bat)
        pv_used_charge = min(pv_remaining, room_kWh_pv)
        charge_stored  = pv_used_charge * Eff_cc_batt * Eff_bat

        soc             += charge_stored * 1000.0
        pv_to_battery[h] = pv_used_charge
        pv_remaining    -= pv_used_charge

    # ------------------------------------------------------------------
    #  Inverter efficiency (Sandia model)
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
    soc               = batt_soc[h]
    energy_from_pv[h] = pv_to_battery[h] + pv_to_load[h]

# =============================================================================
#  ATTACH RESULTS
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

min_soc = batt_soc.min()
print(f"\nMin battery SoC         : {min_soc:.1f} Wh")
print(f"Min battery SoC (frac)  : {min_soc / Sys_Wh:.4f}")
print(f"\nMax PV to Load (single hour): {pv_to_load.max():.4f} kWh")

# =============================================================================
#  EXPORT
# =============================================================================

load_profile.to_csv("load_results.csv", index=False)
print("\nResults saved to load_results.csv")
