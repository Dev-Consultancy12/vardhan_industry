import pandas as pd
from backend.services.excel_generator_v4 import TEST_READINGS_PATH

readings_df = pd.read_excel(TEST_READINGS_PATH, sheet_name='Sheet1')
row0 = readings_df.iloc[0]
group_to_cols = {}
for i, col in enumerate(readings_df.columns):
    val = row0[col]
    if pd.notna(val) and isinstance(val, str) and not val.startswith("Grp") and not val.startswith("Specification"):
        group_to_cols[val.strip()] = (i - 1, i)

parsed_data = {grp: {} for grp in group_to_cols}
for grp, (spec_col, val_col) in group_to_cols.items():
    param_col = 1 if spec_col < 38 else 3 
    for r in range(1, 18):
        p_val = readings_df.iloc[r, param_col]
        s_val = readings_df.iloc[r, spec_col]
        v_val = readings_df.iloc[r, val_col]
        if pd.notna(p_val) and str(p_val).strip() != "":
            pool = [v_val] if pd.notna(v_val) else []
            parsed_data[grp][str(p_val).strip()] = {'spec': str(s_val).strip() if pd.notna(s_val) else "", 'pool': pool}
            
current_param = None
for r in range(20, len(readings_df)):
    p_val = readings_df.iloc[r, 5]
    if pd.notna(p_val) and str(p_val).strip() != "":
        current_param = str(p_val).strip()
    if current_param:
        for grp, (spec_col, val_col) in group_to_cols.items():
            v_val = readings_df.iloc[r, val_col]
            if pd.notna(v_val):
                if current_param not in parsed_data[grp]: parsed_data[grp][current_param] = {'spec': '', 'pool': []}
                parsed_data[grp][current_param]['pool'].append(v_val)

target_grp = "AVSSX 0.5f 19/0.19 OD 1.45"
for k in parsed_data.keys():
    if "AVSSX" in k:
        print(f"Group: {k}")
        print(list(parsed_data[k].keys()))
