import requests
from bs4 import BeautifulSoup
import json
import openpyxl
import pandas as pd

def generate_phosphosite_url(uniprot_id):
    """Generate URL for PhosphoSitePlus page with the given UniProt ID."""
    base_url = "https://www.phosphosite.org/uniprotAccAction?id="
    return f"{base_url}{uniprot_id}"

def fetch_pfam_domain_details(pfam_id):
    """Fetch details about a Pfam domain from the InterPro API."""
    api_url = f"https://www.ebi.ac.uk/interpro/api/entry/pfam/{pfam_id}"
    response = requests.get(api_url, headers={"Accept": "application/json"})
    if response.status_code == 200:
        data = response.json()
        return {
            "id": data.get("metadata", {}).get("accession"),
            "name": data.get("metadata", {}).get("name"),
            "short": data.get("metadata", {}).get("short"),
            "type": data.get("metadata", {}).get("type")
        }
    return None

def fetch_domains_for_uniprot_id(uniprot_id):
    """Fetch domain data from PhosphoSitePlus for a given UniProt ID."""
    url = generate_phosphosite_url(uniprot_id)
    response = requests.get(url)
    response.raise_for_status()
    soup = BeautifulSoup(response.text, 'html.parser')
    domains_param = soup.find('param', {'id': 'Domains', 'name': 'Domains'})

    domains_data = domains_param.get('value', '[]') if domains_param else '[]'
    domains_list = json.loads(domains_data)

    pfamid, start, end, type1, short1, name2 = [], [], [], [], [], []

    for domain in domains_list[1:]:
        pfam = domain.get('PFAM')
        if pfam:
            pfamid.append(pfam)
            domain_details = fetch_pfam_domain_details(pfam)
            if domain_details:
                name = domain_details.get("name")
                if name:
                    short1.append(name.get("short"))
                    name2.append(name.get("name"))
                #pfam_domain_details.append(domain_details)
                type1.append(domain_details.get("type"))
                
            start.append(domain.get('START'))
            end.append(domain.get('END'))
    print(uniprot_id)
    return {
        "uniprot_id": uniprot_id,
        "short_name": short1,
        "start": start,
        "end": end,
        "name": name2,
        "type": type1
    }

# Load UniProt IDs from file
input_file = "input_file.xlsx"  # Replace with your file path
output_file = "output_file.xlsx"
data = pd.read_excel(input_file)
uniprot_ids = data['id'].unique()

# Initialize output workbook
try:
    workbook = openpyxl.load_workbook(output_file)
except FileNotFoundError:
    workbook = openpyxl.Workbook()

sheet = workbook.active
sheet["A1"] = "uniprot_id"
sheet["B1"] = "short_name"
sheet["C1"] = "start"
sheet["D1"] = "end"
sheet["E1"] = "name"
sheet["F1"] = "type"

# Collect domain data for each UniProt ID
row = 2
for uniprot_id in uniprot_ids:
    domain_data = fetch_domains_for_uniprot_id(uniprot_id)
    max_length = max(len(domain_data['short_name']), len(domain_data['start']), 
                     len(domain_data['end']), len(domain_data['name']), len(domain_data['type']))

    for i in range(max_length):
        sheet[f"A{row}"] = domain_data['uniprot_id']
        sheet[f"B{row}"] = domain_data['short_name'][i] if i < len(domain_data['short_name']) else ""
        sheet[f"C{row}"] = domain_data['start'][i] if i < len(domain_data['start']) else ""
        sheet[f"D{row}"] = domain_data['end'][i] if i < len(domain_data['end']) else ""
        sheet[f"E{row}"] = domain_data['name'][i] if i < len(domain_data['name']) else ""
        sheet[f"F{row}"] = domain_data['type'][i] if i < len(domain_data['type']) else ""
        row += 1

# Save the workbook
workbook.save(output_file)
print("Data written to Excel file successfully.")