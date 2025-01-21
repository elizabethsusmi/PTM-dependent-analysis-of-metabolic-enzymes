import pandas as pd


file1_path = r"file1_path.xlsx"  # First file with site information
file2_path = r"file2_path"  # Second file with domain information

sites_df = pd.read_excel(file1_path)  # Contains Uniprot ID, gene symbol, and site information
domains_df = pd.read_excel(file2_path)  # Contains domain information with start-end positions

sites_df = sites_df.dropna(subset=['site'])

domains_df[['start', 'end']] = domains_df[['start', 'end']].astype(int)

matched_data = []

for _, site_row in sites_df.iterrows():
    uniprot_id = site_row['uniprot accession']
    gene_symbol = site_row['gene symbol']
    
    if isinstance(site_row['site'], str) and len(site_row['site']) > 1:
        site_position = int(site_row['site'][1:])
        
        matching_domains = domains_df[
            (domains_df['uniprot_id'] == uniprot_id) &
            (domains_df['start'] <= site_position) &
            (domains_df['end'] >= site_position)
        ]
        
        if not matching_domains.empty:
            for _, domain_row in matching_domains.iterrows():
                matched_data.append({
                    'Uniprot Accession': uniprot_id,
                    'Gene Symbol': gene_symbol,
                    'Enzyme' : site_row['enzyme'],
                    'Site': site_row['site'],
                    'Modification': site_row['type'],
                    'Domain/Region Type': domain_row['type'],
                    'Domain/Region Description': domain_row['name'],
                    'Domain Start': domain_row['start'],
                    'Domain End': domain_row['end']
                })
        else:
            matched_data.append({
                'Uniprot Accession': uniprot_id,
                'Gene Symbol': gene_symbol,
                'Site': site_row['site'],
                'Modification': site_row['type'],
                'Domain/Region Type': None,
                'Domain/Region Description': None,
                'Domain Start': None,
                'Domain End': None
            })

matched_df = pd.DataFrame(matched_data)

output_path = r'output_filepath'
matched_df.to_excel(output_path, index=False)

print("Mapping completed and saved to:", output_path)
