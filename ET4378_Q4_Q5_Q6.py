#!/usr/bin/env python
# coding: utf-8

# # ET4378 – Group 20 | Questions 4, 5 & 6
# **Tampa, FL, USA | Load Profile 3 | Both roof slopes used**
# 
# This notebook answers:
# - **Q4** – Optimal number of PV modules and most suitable inverter
# - **Q5** – Roof layout: how many modules fit on each slope
# - **Q6** – Average inverter efficiency in grid-tied mode vs blackout

# In[ ]:





# In[1]:


import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from pathlib import Path
from scipy.interpolate import interp1d

plt.rcParams['figure.figsize'] = (13, 4)
plt.rcParams['figure.dpi'] = 100
print('Libraries loaded.')


# ---
# ## 1. Configuration
# All fixed parameters in one place. Change values here and re-run.

# In[2]:


# ── File paths (adjust if your folder structure differs) ──────────────────────
WEATHER_FILE = Path(r'C:\Users\rijkt\OneDrive - Delft University of Technology\SET YR 1\PV systems\ET4378\Tampa_FL-hour.csv')
INV_FILE     = Path(r'C:\Users\rijkt\OneDrive - Delft University of Technology\SET YR 1\PV systems\ET4378\Inverter Efficiency parameters.xlsx')

# ── Location ──────────────────────────────────────────────────────────────────
LATITUDE  = 27.95   # degrees North

# ── Roof (Tampa: 5 x 18 m, 35° tilt) ─────────────────────────────────────────
# The 5 m dimension is the FULL plan width of the building.
# A gable roof means each slope covers half: 5/2 = 2.5 m horizontal.
ROOF_FULL_WIDTH_M  = 5.0    # m  full building width (plan view)
ROOF_RIDGE_M       = 18.0   # m  ridge / building length
ROOF_TILT_DEG      = 35.0   # degrees
SAFETY_MARGIN_M    = 0.5    # m  free on every edge (assignment requirement)
CLUSTER_GAP_M      = 0.5    # m  between consecutive clusters
MAX_ROWS_PER_CLUSTER = 4    # assignment rule: max 4 x 4
MAX_COLS_PER_CLUSTER = 4

# ── PV Module: AIKO Neostar 2S ────────────────────────────────────────────────
MODULE_PMAX_STC   = 490.0   # W  rated power at STC (1000 W/m², 25 °C)
MODULE_NOCT       = 46.0    # °C  Normal Operating Cell Temperature
MODULE_TEMP_COEFF = -0.0026 # /°C  power temperature coefficient γ (−0.26 %/°C)
MODULE_EFF_STC    = 0.245   # 24.5% module efficiency at STC
MODULE_LENGTH_M   = 1.762   # m  along the slope (portrait orientation)
MODULE_WIDTH_M    = 1.134   # m  across the ridge

# ── Inverter: Schneider XW Pro 6848 ──────────────────────────────────────────
INVERTER_RATED_W  = 6020.0  # W  rated AC output power

# ── Charge controller efficiency ─────────────────────────────────────────────
CC_EFFICIENCY     = 0.95

# ── Critical load (Group 20 = Load 3) ────────────────────────────────────────
CRITICAL_LOAD_KW  = 6.0     # kW

print('Configuration loaded.')


# ---
# ## 2. Weather Data
# Load the Tampa hourly weather file and attach a datetime index.

# In[3]:


# Load weather CSV (semicolon-separated)
weather = pd.read_csv(WEATHER_FILE, sep=';')
weather.columns = weather.columns.str.strip()

# Attach an hourly datetime index for the year 2005
date_range = pd.date_range(start='2005-01-01 00:00', periods=8760, freq='h')
weather.index = date_range

print(f'Rows loaded  : {len(weather)}')
print(f'Columns      : {list(weather.columns)}')
print(f'Annual GHI   : {weather["G_Gh"].sum()/1000:.1f} kWh/m²')
print(f'Peak GHI     : {weather["G_Gh"].max():.0f} W/m²')
print(f'Avg temp     : {weather["Ta"].mean():.1f} °C')
weather.head(3)


# In[4]:


fig, axes = plt.subplots(2, 1, figsize=(13, 7))

monthly_ghi  = weather['G_Gh'].resample('M').sum() / 1000
monthly_temp = weather['Ta'].resample('M').mean()
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
# ### How many modules fit on each slope?
# 
# **Roof geometry (gable / double-sided):**
# - Full building width = 5 m → each slope covers **2.5 m horizontally**
# - Slant length per slope = (A/2) / cos(tilt) = 2.5 / cos(35°)
# - Safety margin 0.5 m on every edge; clusters separated by 0.5 m gaps
# - Module in **portrait**: 1.762 m along the slope, 1.134 m across the ridge
# - Max cluster size: 4 rows × 4 columns

# In[5]:


# ── Step 1: Compute roof dimensions ──────────────────────────────────────────
tilt_rad      = np.radians(ROOF_TILT_DEG)
half_plan_m   = ROOF_FULL_WIDTH_M / 2          # 2.5 m – plan width of ONE slope
slant_m       = half_plan_m / np.cos(tilt_rad) # slant length of one slope

usable_slant_m = slant_m  - 2 * SAFETY_MARGIN_M   # remove 0.5 m at ridge + 0.5 m at eave
usable_ridge_m = ROOF_RIDGE_M - 2 * SAFETY_MARGIN_M  # remove 0.5 m at each gable end

print('── Roof geometry (one slope) ──────────────────────')
print(f'  Half plan width          : {half_plan_m:.2f} m')
print(f'  Slant length             : {slant_m:.3f} m')
print(f'  Usable slant (−2×0.5 m) : {usable_slant_m:.3f} m')
print(f'  Usable ridge (−2×0.5 m) : {usable_ridge_m:.2f} m')

# ── Step 2: Rows per cluster (along slope) ────────────────────────────────────
rows_per_cluster = int(usable_slant_m // MODULE_LENGTH_M)  # floor division
rows_per_cluster = min(rows_per_cluster, MAX_ROWS_PER_CLUSTER)
cluster_slant_m  = rows_per_cluster * MODULE_LENGTH_M

print(f'\n── Along slope (portrait: {MODULE_LENGTH_M} m per module) ─────')
print(f'  Modules that fit         : {usable_slant_m:.3f} / {MODULE_LENGTH_M} = {usable_slant_m/MODULE_LENGTH_M:.2f} → {rows_per_cluster} row(s) per cluster')
print(f'  Cluster height           : {cluster_slant_m:.3f} m')

# ── Step 3: Columns per cluster and number of clusters (along ridge) ──────────
cols_per_cluster  = MAX_COLS_PER_CLUSTER  # use maximum (4)
cluster_width_m   = cols_per_cluster * MODULE_WIDTH_M

# Total width consumed per cluster = cluster width + gap (last cluster needs no trailing gap)
# n_clusters × (cluster_width + gap) - gap ≤ usable_ridge
# → n_clusters ≤ (usable_ridge + gap) / (cluster_width + gap)
n_clusters = int((usable_ridge_m + CLUSTER_GAP_M) / (cluster_width_m + CLUSTER_GAP_M))

total_width_used  = n_clusters * cluster_width_m + (n_clusters - 1) * CLUSTER_GAP_M

print(f'\n── Along ridge (module width: {MODULE_WIDTH_M} m) ──────────────')
print(f'  Cluster width (4 cols)   : {cluster_width_m:.3f} m')
print(f'  Number of clusters       : {n_clusters}')
print(f'  Total width used         : {total_width_used:.3f} m  (of {usable_ridge_m:.1f} m available)')

# ── Step 4: Total modules ─────────────────────────────────────────────────────
mods_per_slope = n_clusters * cols_per_cluster * rows_per_cluster
n_slopes       = 2  # south + north
total_modules  = mods_per_slope * n_slopes
total_power_wp = total_modules * MODULE_PMAX_STC

print(f'\n── Module count ──────────────────────────────────────')
print(f'  Layout per slope         : {n_clusters} clusters × {cols_per_cluster} cols × {rows_per_cluster} row = {mods_per_slope} modules')
print(f'  Both slopes (S + N)      : 2 × {mods_per_slope} = {total_modules} modules')
print(f'  Total installed peak     : {total_modules} × {MODULE_PMAX_STC:.0f} W = {total_power_wp/1000:.2f} kWp')


# ---
# ## 4. Question 4 – POA Irradiance (both slopes)
# 
# We compute the Plane-of-Array (POA) irradiance separately for the south-facing and north-facing slopes using the **isotropic sky model**:
# 
# $$G_{\text{POA}} = G_{\text{Bn}} \cdot \cos(\theta_i) + G_{\text{Dh}} \cdot \frac{1+\cos(\beta)}{2} + G_{\text{Gh}} \cdot \rho \cdot \frac{1-\cos(\beta)}{2}$$
# 
# - **South slope**: azimuth = 0° (facing the equator)
# - **North slope**: azimuth = 180° (facing away from equator)

# In[6]:


def compute_poa(weather_df, surface_tilt_deg, surface_azimuth_deg, albedo=0.20):
    """
    Compute hourly Plane-of-Array (POA) irradiance [W/m²].
    
    Parameters
    ----------
    surface_tilt_deg      : tilt from horizontal (degrees)
    surface_azimuth_deg   : 0=South, 90=West, 180=North, -90=East (degrees)
    albedo                : ground reflectance (0.20 = concrete/grass)
    """
    tilt    = np.radians(surface_tilt_deg)
    surf_az = np.radians(surface_azimuth_deg)
    sol_az  = np.radians(weather_df['Az'])
    sol_el  = np.radians(weather_df['hs'])

    # Angle of incidence (beam component)
    cos_theta_i = (
        np.sin(sol_el) * np.cos(tilt)
        + np.cos(sol_el) * np.sin(tilt) * np.cos(sol_az - surf_az)
    )
    cos_theta_i = np.clip(cos_theta_i, 0, 1)  # no back-face irradiance

    G_beam      = weather_df['G_Bn'] * cos_theta_i
    G_diffuse   = weather_df['G_Dh'] * (1 + np.cos(tilt)) / 2
    G_reflected = weather_df['G_Gh'] * albedo * (1 - np.cos(tilt)) / 2

    return (G_beam + G_diffuse + G_reflected).clip(lower=0)


# Compute POA for both slopes
weather['G_poa_south'] = compute_poa(weather, ROOF_TILT_DEG, surface_azimuth_deg=0)
weather['G_poa_north'] = compute_poa(weather, ROOF_TILT_DEG, surface_azimuth_deg=180)

print('Annual irradiation summary:')
print(f'  GHI  (horizontal)  : {weather["G_Gh"].sum()/1000:.1f} kWh/m²/yr')
print(f'  POA  South slope   : {weather["G_poa_south"].sum()/1000:.1f} kWh/m²/yr')
print(f'  POA  North slope   : {weather["G_poa_north"].sum()/1000:.1f} kWh/m²/yr')


# In[7]:


# Monthly POA comparison
monthly_poa_s = weather['G_poa_south'].resample('M').sum() / 1000
monthly_poa_n = weather['G_poa_north'].resample('M').sum() / 1000
monthly_ghi   = weather['G_Gh'].resample('M').sum() / 1000

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
# 
# ### Cell temperature correction (NOCT model)
# $$T_{\text{cell}} = T_{\text{amb}} + \frac{\text{NOCT} - 20}{800} \cdot G_{\text{POA}}$$
# 
# ### DC power output
# $$P_{\text{DC}} = N \cdot P_{\text{STC}} \cdot \frac{G_{\text{POA}}}{1000} \cdot \left[1 + \gamma \cdot (T_{\text{cell}} - 25)\right]$$

# In[8]:


def calc_pv_power(G_poa, T_amb, n_modules, pmax_stc, noct, gamma):
    """
    Return hourly DC power [W] for a string of n_modules.
    
    G_poa    : pd.Series, hourly POA irradiance [W/m²]
    T_amb    : pd.Series, hourly ambient temperature [°C]
    n_modules: number of modules
    pmax_stc : module rated power at STC [W]
    noct     : Normal Operating Cell Temperature [°C]
    gamma    : power temperature coefficient [1/°C]  (negative value)
    """
    T_cell = T_amb + ((noct - 20) / 800) * G_poa
    P_dc   = n_modules * pmax_stc * (G_poa / 1000) * (1 + gamma * (T_cell - 25))
    return P_dc.clip(lower=0)


# DC power from each slope (12 modules each)
weather['P_dc_south_W'] = calc_pv_power(
    weather['G_poa_south'], weather['Ta'],
    mods_per_slope, MODULE_PMAX_STC, MODULE_NOCT, MODULE_TEMP_COEFF
)
weather['P_dc_north_W'] = calc_pv_power(
    weather['G_poa_north'], weather['Ta'],
    mods_per_slope, MODULE_PMAX_STC, MODULE_NOCT, MODULE_TEMP_COEFF
)

# Combined DC power from both slopes
weather['P_dc_total_W'] = weather['P_dc_south_W'] + weather['P_dc_north_W']

# Annual energy
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


# In[9]:


# Monthly DC energy by slope
monthly_dc_s = weather['P_dc_south_W'].resample('M').sum() / 1000
monthly_dc_n = weather['P_dc_north_W'].resample('M').sum() / 1000

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
# 
# The inverter must:
# 1. Handle the **peak DC power** from the array
# 2. Supply the **critical load** (6 kW AC) during blackout
# 3. Operate at 120V / 60Hz (Tampa)
# 
# **Selected inverter: Schneider XW Pro 6848**
# - Rated AC output: 6,800 W
# - Hybrid (grid-tie + battery backup)
# - 120V/60Hz compatible
# 
# The peak DC input will exceed the inverter rating (clipping), which is normal and expected in a system sized for backup power rather than maximum yield.

# In[10]:


peak_dc_kw = weather['P_dc_total_W'].max() / 1000
dc_ac_ratio = (total_power_wp / 1000) / (INVERTER_RATED_W / 1000)

print('─── Inverter sizing check ──────────────────────────────')
print(f'  Installed peak DC          : {total_power_wp/1000:.2f} kWp')
print(f'  Simulated peak DC power    : {peak_dc_kw:.2f} kW')
print(f'  Inverter rated power       : {INVERTER_RATED_W/1000:.1f} kW')
print(f'  DC/AC ratio                : {dc_ac_ratio:.2f}  (>1 means some clipping at peak hours)')
print(f'  Critical load to cover     : {CRITICAL_LOAD_KW:.1f} kW  →  {CRITICAL_LOAD_KW/INVERTER_RATED_W*1000:.1f} kW  ✓ inverter can supply it')

# Load percentage histogram
load_pct = (weather['P_dc_total_W'].clip(upper=INVERTER_RATED_W) / INVERTER_RATED_W * 100)
daytime_load_pct = load_pct[load_pct > 0]  # only hours with PV production

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
# 
# ### Load the efficiency curve from the Excel file

# In[11]:


inv_eff = pd.read_excel(INV_FILE, sheet_name='Plot Data')
inv_eff.columns = ['pct_rated', 'eff_nominal', 'eff_low', 'eff_high'] + list(inv_eff.columns[4:])
inv_eff = inv_eff[['pct_rated', 'eff_nominal', 'eff_low', 'eff_high']].dropna()
inv_eff = inv_eff[inv_eff['pct_rated'] >= 0]

# Build interpolation function (nominal curve)
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

# Mark the blackout operating point
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

# Peak efficiency
peak_idx = inv_eff['eff_nominal'].idxmax()
print(f'Peak efficiency : {inv_eff.loc[peak_idx, "eff_nominal"]:.2f}% at {inv_eff.loc[peak_idx, "pct_rated"]:.1f}% load')
print(f'Blackout load   : {blackout_pct:.1f}% rated → efficiency = {blackout_eff:.2f}%')


# ### Q6 – Average efficiency in grid-tied mode
# 
# In grid-tied mode the inverter converts PV DC power to AC. The operating load varies each hour.
# We weight each hour's efficiency by its energy output to get the **energy-weighted average**.

# In[12]:


# Grid-tied: inverter processes the PV DC output (clipped at rated)
P_inv_input_W = weather['P_dc_total_W'].clip(upper=INVERTER_RATED_W)  # W into inverter

# Hourly load % and efficiency for each hour with PV output
pv_hours = P_inv_input_W > 0
load_pct_hourly = (P_inv_input_W[pv_hours] / INVERTER_RATED_W) * 100
eff_hourly      = inv_eff_interp(load_pct_hourly.values) / 100.0  # as fraction

# Energy-weighted average efficiency
energy_weighted_eff = np.average(eff_hourly, weights=P_inv_input_W[pv_hours].values)

# Simple mean (unweighted) for comparison
simple_mean_eff = eff_hourly.mean()

print('─── Q6: Inverter Efficiency Analysis ───────────────────')
print()
print('GRID-TIED MODE:')
print(f'  Hours with PV output       : {pv_hours.sum()} h/yr')
print(f'  Avg load (unweighted)      : {load_pct_hourly.mean():.1f}% of rated')
print(f'  Simple mean efficiency     : {simple_mean_eff*100:.2f}%')
print(f'  Energy-weighted efficiency : {energy_weighted_eff*100:.2f}%')
print()
print('BLACKOUT MODE:')
print(f'  PV output during blackout  : 0 W  (assignment rule: PV shut off)')
print(f'  Battery supplies load via inverter')
print(f'  Inverter input power       : {CRITICAL_LOAD_KW:.1f} kW AC load ÷ eff')
print(f'  Inverter load              : {blackout_pct:.1f}% of rated ({CRITICAL_LOAD_KW:.0f} kW / {INVERTER_RATED_W/1000:.1f} kW)')
print(f'  Inverter efficiency        : {blackout_eff:.2f}%')
print()
print('EXPLANATION OF DIFFERENCE:')
print(f'  Grid-tied: inverter load varies 0–100% all day → mostly low-to-mid load')
print(f'  Blackout:  inverter always at fixed {blackout_pct:.1f}% load → single operating point')
print(f'  The blackout efficiency ({blackout_eff:.2f}%) is near the peak of the curve,')
print(f'  while the grid-tied average ({energy_weighted_eff*100:.2f}%) is dragged down by')
print(f'  many early-morning/evening hours at very low load where efficiency is lower.')


# In[13]:


# Show efficiency distribution during grid-tied operation
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

print('\n─── Summary for report ──────────────────────────────────')
print(f'  Modules installed          : {total_modules} × AIKO Neostar 2S 490 Wp')
print(f'  Installed peak power       : {total_power_wp/1000:.2f} kWp')
print(f'  Layout per slope           : {n_clusters} clusters × {cols_per_cluster} cols × {rows_per_cluster} row = {mods_per_slope} modules')
print(f'  Total annual DC energy     : {E_dc_total_kwh:,.0f} kWh/yr')
print(f'  Inverter                   : Schneider XW Pro 6848  ({INVERTER_RATED_W/1000:.1f} kW rated)')
print(f'  Inverter eff – grid-tied   : {energy_weighted_eff*100:.2f}%  (energy-weighted average)')
print(f'  Inverter eff – blackout    : {blackout_eff:.2f}%  (fixed {blackout_pct:.1f}% load point)')

