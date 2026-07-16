import pandas as pd
import math

file_path = 'master_readings/TEST READINGS[23].xlsx'
df = pd.read_excel(file_path, sheet_name='Sheet1')

# Row 0 contains the group names
row0 = df.iloc[0]
group_to_cols = {}
for i, col in enumerate(df.columns):
    val = row0[col]
    if pd.notna(val) and isinstance(val, str) and not val.startswith("Grp") and not val.startswith("Specification"):
        spec_col = i - 1
        val_col = i
        group_to_cols[val.strip()] = (spec_col, val_col)

parsed_data = {}
for grp in group_to_cols:
    parsed_data[grp] = {}

# 1. Parse Upper Section (Fixed Values)
# Rows 1 to 17
for grp, (spec_col, val_col) in group_to_cols.items():
    param_col = 1 if spec_col < 38 else 3 
    
    for r in range(1, 18):
        p_val = df.iloc[r, param_col]
        s_val = df.iloc[r, spec_col]
        v_val = df.iloc[r, val_col]
        
        if pd.notna(p_val) and str(p_val).strip() != "":
            param_name = str(p_val).strip()
            # We add it if v_val is not nan, or if it's a known fixed param.
            # Actually, we can just store the pool. If it's NaN, the pool is empty (which means it's randomized)
            pool = [v_val] if pd.notna(v_val) else []
            parsed_data[grp][param_name] = pool

# 2. Parse Lower Section (Randomized Pools)
# Parameter name is in column 5 (index 5)
current_param = None
for r in range(20, len(df)):
    p_val = df.iloc[r, 5]
    if pd.notna(p_val) and str(p_val).strip() != "":
        current_param = str(p_val).strip()
    
    if current_param:
        for grp, (spec_col, val_col) in group_to_cols.items():
            v_val = df.iloc[r, val_col]
            if pd.notna(v_val):
                if current_param not in parsed_data[grp]:
                    parsed_data[grp][current_param] = []
                parsed_data[grp][current_param].append(v_val)

g1 = 'AVSS 0.5f 20/0.18 OD 1.75'
print(f"--- {g1} ---")
for p, pool in parsed_data[g1].items():
    print(f"{p}: {pool}")

g12 = 'FLRY-B 0.35 12/0.20 OD 1.30'
print(f"\n--- {g12} ---")
for p, pool in parsed_data[g12].items():
    print(f"{p}: {pool}")

