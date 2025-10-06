import pandas as pd
import re
from collections import defaultdict

# Step 1: Load the data
df = pd.read_excel(r"inputfile.xlsx")  # Update the path

# Step 2: Extract numeric PTM site positions and preserve original
df["Numeric_Site"] = df["Site"].apply(lambda x: int(re.findall(r"\d+", x)[0]))

# Step 3: Group by protein and store site info
ptm_data = defaultdict(list)
for _, row in df.iterrows():
    protein = row["Protein accession"]
    numeric_site = row["Numeric_Site"]
    site_label = row["Site"]
    mod = row["Modification ype"]
    ptm_data[protein].append((numeric_site, site_label, mod))

# Step 4: Identify hotspots and crosstalk
hotspot_summary = []

for protein, site_info in ptm_data.items():
    positions = [s[0] for s in site_info]
    site_mod_map = defaultdict(list)
    site_label_map = {}

    for num, label, mod in site_info:
        site_mod_map[num].append(mod)
        site_label_map[num] = label

    for central_site in positions:
        window = range(central_site - 7, central_site + 8)
        neighbors = [s for s in positions if s in window and s != central_site]

        # Gather all nearby PTMs as (residue, ptm type)
        nearby_ptms = []
        for pos in neighbors:
            res_label = site_label_map.get(pos)
            mods = site_mod_map.get(pos, [])
            for m in mods:
                nearby_ptms.append((res_label, m))

        # Count how many residues have multiple PTM types (crosstalk)
        crosstalk_sites = [pos for pos in neighbors if len(set(site_mod_map[pos])) > 1]
        crosstalk_count = len(crosstalk_sites)

        # Collect unique PTM types from crosstalk sites
        crosstalk_ptm_types = set()
        for pos in crosstalk_sites:
            crosstalk_ptm_types.update(site_mod_map[pos])

        ptm_count = len(neighbors)
        is_ptm_hotspot = ptm_count >= 5
        is_crosstalk_hotspot = crosstalk_count >= 3

        hotspot_summary.append({
            "Protein Accession": protein,
            "Central Site": site_label_map[central_site],
            "PTM Count (±7)": ptm_count,
            "Crosstalk Residue Count": crosstalk_count,
            "Crosstalk PTM Types": list(crosstalk_ptm_types),
            "Nearby PTMs": nearby_ptms,
            "Is PTM Hotspot": "Yes" if is_ptm_hotspot else "No",
            "Is Crosstalk Hotspot": "Yes" if is_crosstalk_hotspot else "No"
        })

# Step 5: Save to Excel
output_df = pd.DataFrame(hotspot_summary)
output_df.to_excel("output_file.xlsx", index=False)
print("Excel file saved as 'output_file.xlsx'")