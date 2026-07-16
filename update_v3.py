import re

with open('generate_excel_reports_v3.py', 'r') as f:
    content = f.read()

# 1. Replace the entire data parsing and generation logic
# This targets from `try:\n    readings_df = pd.read_excel(TEST_READINGS_PATH, header=None)`
# all the way to the end of the `parameters = [...]` list.

start_marker = "try:\n    readings_df = pd.read_excel(TEST_READINGS_PATH, header=None)"
end_marker = "]\n\ntry:\n    ps_df = pd.read_excel(PACKING_SLIP_PATH)"

new_logic = """category_map = {
    "No of strands": "A",
    "Strands dia": "C",
    "Insulation dia": "C",
    "Conductor resistance": "B",
    "Insulation thickness": "C",
    "Insulation thickness (Nom.)": "C",
    "Temperature rating": "D",
    "Temprature rating": "D",
    "Voltage rating": "D",
    "Volume Resistivity": "B",
    "Withstand voltage": "B",
    "Printing over the cable": "A",
    "Insulator Elongation": "B",
    "Conductor Elongation": "B",
    "Tensile Strength": "B",
    "Shrinkage": "B",
    "Spark Testing": "B",
    "Colour": "A",
    "Appearance": "A"
}

try:
    readings_df = pd.read_excel(TEST_READINGS_PATH, sheet_name='Sheet1')
except Exception as e:
    print(f"Error reading Test Readings: {e}")
    exit(1)

row0 = readings_df.iloc[0]
group_to_cols = {}
for i, col in enumerate(readings_df.columns):
    val = row0[col]
    if pd.notna(val) and isinstance(val, str) and not val.startswith("Grp") and not val.startswith("Specification"):
        spec_col = i - 1
        val_col = i
        group_to_cols[val.strip()] = (spec_col, val_col)

parsed_data = {}
for grp in group_to_cols:
    parsed_data[grp] = {}

for grp, (spec_col, val_col) in group_to_cols.items():
    param_col = 1 if spec_col < 38 else 3 
    
    for r in range(1, 18):
        p_val = readings_df.iloc[r, param_col]
        s_val = readings_df.iloc[r, spec_col]
        v_val = readings_df.iloc[r, val_col]
        
        if pd.notna(p_val) and str(p_val).strip() != "":
            param_name = str(p_val).strip()
            spec_str = str(s_val).strip() if pd.notna(s_val) else ""
            pool = [v_val] if pd.notna(v_val) else []
            parsed_data[grp][param_name] = {'spec': spec_str, 'pool': pool}

current_param = None
for r in range(20, len(readings_df)):
    p_val = readings_df.iloc[r, 5]
    if pd.notna(p_val) and str(p_val).strip() != "":
        current_param = str(p_val).strip()
    
    if current_param:
        for grp, (spec_col, val_col) in group_to_cols.items():
            v_val = readings_df.iloc[r, val_col]
            if pd.notna(v_val):
                if current_param not in parsed_data[grp]:
                    parsed_data[grp][current_param] = {'spec': '', 'pool': []}
                parsed_data[grp][current_param]['pool'].append(v_val)
"""

content = content.replace(content[content.find(start_marker):content.find(end_marker)], new_logic)

# 2. Replace the reading insertion logic inside the main loop
old_insertion = """        readings_table = []
        for p in parameters:
            obs = generate_obs(p['name'], p['fixed'], copper, p['spec'], 8, colour)
            for i in range(cols_to_fill, 8):
                obs[i] = ""
                
            sample_size = "8"
            if "Spark" in p['name']:
                sample_size = "100%"
                
            readings_table.append({
                'parameter': p['name'],
                'specification': p['spec'],
                'sample_size': sample_size,
                'obs': obs,
                'status1': "OK" if cols_to_fill > 0 else "",
                'insp_category': p['cat']
            })"""

new_insertion = """        readings_table = []
        group_data = parsed_data.get(full_desc, {})
        if not group_data:
            print(f"WARNING: Group '{full_desc}' not found in master readings!")
        
        for p_name, p_data in group_data.items():
            spec = p_data.get('spec', '')
            pool = p_data.get('pool', [])
            cat = category_map.get(p_name, "A")
            
            obs = []
            if p_name == "Colour":
                obs = [colour] * 8
            elif len(pool) == 1:
                obs = [pool[0]] * 8
            elif len(pool) > 1:
                for _ in range(8):
                    obs.append(random.choice(pool))
            else:
                obs = ['OK'] * 8
                
            for i in range(cols_to_fill, 8):
                obs[i] = ""
                
            sample_size = "8"
            if "Spark" in p_name:
                sample_size = "100%"
                
            readings_table.append({
                'parameter': p_name,
                'specification': spec,
                'sample_size': sample_size,
                'obs': obs,
                'status1': "OK" if cols_to_fill > 0 else "",
                'insp_category': cat
            })"""

content = content.replace(old_insertion, new_insertion)

# Write back to V3
with open('generate_excel_reports_v3.py', 'w') as f:
    f.write(content)

