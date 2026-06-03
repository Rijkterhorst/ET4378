"""
Battery Maximum Charge / Discharge Current Analysis
----------------------------------------------------
Run this after (or alongside) pv_battery_dispatch.py, or paste it at the
end of that file.  All parameters must match the dispatch script.

Li-ion pack: 25.6 V modules, 200 Ah each
Configuration: 2S × 4P  (series × parallel)
"""

import numpy as np

# =============================================================================
#  SYSTEM PARAMETERS  (must match dispatch script)
# =============================================================================

# Battery module specs
li_V   = 25.6    # V  – nominal voltage per module
Li_Ah  = 200     # Ah – capacity per module

# Pack configuration (from dispatch script sizing)
series   = 2     # modules in series  → pack voltage
parallel = 4     # strings in parallel → pack capacity

# System voltage & energy
sys_v  = series * li_V          # V   – nominal pack voltage (51.2 V)
sys_Ah = parallel * Li_Ah       # Ah  – total pack capacity (800 Ah)

DoD        = 0.75
Sys_Wh_raw = 52866              # Wh before DoD/efficiency correction
Eff_dod    = 0.961
Sys_Wh     = Sys_Wh_raw / (DoD * Eff_dod)   # Wh – usable energy basis

# Efficiency
Eff_bat      = 0.92   # battery round-trip (charge side)
Eff_cc_batt  = 0.92   # charge controller → battery
Eff_CC_grid  = 0.95   # charge controller → load/grid

# Inverter (Sandia model)
PAC0  = 6120.0    # W – rated AC output
PDC0  = 6757.76   # W – DC input at rated output
Ps0   = 42.0      # W – standby power
C0    = -0.000012

# PV array
inverter_max_DC_kW = 7.489   # kW – max DC input to inverter

# =============================================================================
#  MAXIMUM CHARGING CURRENT
# =============================================================================
# Source 1: PV array routed through charge controller into battery
# Max PV power available at battery terminals = inverter_max_DC_kW × CC efficiency
max_pv_to_batt_W   = inverter_max_DC_kW * 1000 * Eff_cc_batt   # W at battery terminals
I_charge_pv        = max_pv_to_batt_W / sys_v                   # A

# Source 2: If the battery were charged from grid (not modelled in dispatch,
#           but a physical upper bound – limited by CC / charger rating)
# For Li-ion, the typical max charge rate is 0.5 C
C_rate_charge = 0.5
I_charge_Crate = C_rate_charge * sys_Ah                          # A  (0.5 C limit)

# Conservative (design) charging current = minimum of the two limits
I_charge_max = min(I_charge_pv, I_charge_Crate)

# =============================================================================
#  MAXIMUM DISCHARGING CURRENT
# =============================================================================
# The inverter draws at most PDC0 watts from the DC bus.
# During a blackout the full load (6 kW) must be supplied by the battery.
# Battery terminal current = DC power / pack voltage

# Peak inverter DC draw (rated point)
I_discharge_inverter = PDC0 / sys_v                              # A

# Peak load during blackout window (6 kW AC → DC)
# Assume worst-case inverter efficiency ≈ PAC0/PDC0 at rated point
eta_rated = PAC0 / PDC0
blackout_load_W   = 6000                                         # W (AC)
blackout_DC_W     = blackout_load_W / eta_rated                  # W (DC needed)
I_discharge_blackout = blackout_DC_W / sys_v                     # A

# For Li-ion, the typical max discharge rate is 1 C
C_rate_discharge  = 1.0
I_discharge_Crate = C_rate_discharge * sys_Ah                    # A

# Design discharge current = minimum of inverter demand and C-rate limit
I_discharge_max = min(I_discharge_inverter, I_discharge_Crate)

# =============================================================================
#  RESULTS
# =============================================================================
print("=" * 55)
print("  BATTERY PACK SPECIFICATION")
print("=" * 55)
print(f"  Configuration         : {series}S × {parallel}P")
print(f"  Nominal pack voltage  : {sys_v:.1f} V")
print(f"  Total capacity        : {sys_Ah:.0f} Ah  /  {sys_v*sys_Ah/1000:.1f} kWh")
print(f"  Usable energy (Sys_Wh): {Sys_Wh:.0f} Wh")

print()
print("=" * 55)
print("  MAXIMUM CHARGING CURRENT")
print("=" * 55)
print(f"  Max PV power → battery : {max_pv_to_batt_W:.0f} W  @  {sys_v:.1f} V")
print(f"    → I_charge (PV limit) : {I_charge_pv:.1f} A")
print(f"  0.5 C rate limit        : {I_charge_Crate:.1f} A")
print(f"  ► Design max charge I   : {I_charge_max:.1f} A  "
      f"({'PV-limited' if I_charge_pv < I_charge_Crate else 'C-rate limited'})")

print()
print("=" * 55)
print("  MAXIMUM DISCHARGING CURRENT")
print("=" * 55)
print(f"  Inverter rated DC draw  : {PDC0:.1f} W  @  {sys_v:.1f} V")
print(f"    → I_discharge (inv.)  : {I_discharge_inverter:.1f} A")
print(f"  Blackout load (6 kW AC) : {blackout_DC_W:.1f} W DC")
print(f"    → I_discharge (blkout): {I_discharge_blackout:.1f} A")
print(f"  1 C rate limit          : {I_discharge_Crate:.1f} A")
print(f"  ► Design max discharge I: {I_discharge_max:.1f} A  "
      f"({'inverter-limited' if I_discharge_inverter < I_discharge_Crate else 'C-rate limited'})")

print()
print("=" * 55)
print("  SUMMARY TABLE")
print("=" * 55)
print(f"  {'Parameter':<30} {'Value':>10}")
print(f"  {'-'*40}")
print(f"  {'Pack voltage':<30} {sys_v:>9.1f} V")
print(f"  {'Pack capacity':<30} {sys_Ah:>9.0f} Ah")
print(f"  {'Max charge current':<30} {I_charge_max:>9.1f} A")
print(f"  {'Max discharge current':<30} {I_discharge_max:>9.1f} A")
print(f"  {'Max charge power':<30} {I_charge_max*sys_v/1000:>9.2f} kW")
print(f"  {'Max discharge power':<30} {I_discharge_max*sys_v/1000:>9.2f} kW")
print("=" * 55)
