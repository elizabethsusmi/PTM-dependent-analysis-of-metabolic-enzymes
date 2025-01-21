import pandas as pd
from collections import Counter

# Dictionary of PTM modification rules for residues
ptm_rules = {
    'ADP-Ribosylation': ['E', 'D', 'C', 'G','H','K','R','N','S','Y'],
    'Acetylation': ['A','K','C','D','E','G','I','L','M','N','P','Q','R','S','T','V','Y'],
    'Biotinylation': ['K'],
    'Carbamidation': ['C'],
    'Carboxylation': ['K'],
    'Deamidation': ['N', 'Q'],
    'Dephosphorylation': ['S', 'T', 'Y'],
    'Farnesylation': ['C'],
    'GPI-anchor': ['A','D','G','N','S'],
    'Geranylgeranylation': ['C'],
    'Glutarylation': ['K'],
    'Glutathionylation': ['C'],
    'Hydroxylation': ['P', 'K','D','H','L','N','R'],
    'Lipoylation': ['K'],
    'Malonylation': ['K'],
    'Methylation': ['K', 'R', 'C','G','H','L','M','N','Q','S'],
    'N-linked Glycosylation': ['N','C','D','E','F','G','H','I','K','L','M','P','Q','R','S','T','Y','W','V'],
    'Neddylation': ['K'],
    'Nitration': ['Y'],
    'O-linked Glycosylation': ['S', 'T', 'Y','A','D','E','I','G','H','K','L','M','N','P','Q','R','V'],
    'Oxidation': ['C','L','M'],
    'Phosphorylation': ['S', 'T', 'Y','A','C','D','E','F','G','H','I','K','L','M','N','P','Q','R'],
    'Pyruvate': ['K'],
    'S-nitrosylation': ['C'],
    'S-palmitoylation': ['C','A','D','E','G','I','L','M','N','P','Q','R','S','V'],
    'Succinylation': ['K','C'],
    'Sulfoxidation': ['M','A','C','D','E','F','G','H','I','K','L','N','P','Q','R','S','T','V','Y','W'],
    'Sumoylation': ['K','A','S'],
    'Ubiquitination': ['A','C','D','E','F','G','H','I','K','L','N','P','Q','R','S','T','V','W','Y'],
}

def count_amino_acids(sequence):
    """
    Count the frequency of each amino acid in the given sequence.
    """
    sequence = sequence.upper()  # Make sure the sequence is in uppercase
    aa_count = Counter(sequence)  # Count the occurrence of each amino acid
    return aa_count

def calculate_ptm_occurrences(sequence):
    """
    Calculate the possible occurrences of each PTM based on rules and the sequence.
    """
    sequence = sequence.upper()  # Make sure the sequence is in uppercase
    ptm_occurrences = {}

    for ptm, residues in ptm_rules.items():
            count = 0
            for residue in residues:
                count += sequence.count(residue)
            ptm_occurrences[ptm] = count
    
    return ptm_occurrences

def process_excel(input_file, output_file):
    # Read the input Excel file
    df = pd.read_excel(input_file)

    # Create lists to hold the results
    amino_acid_counts = []
    ptm_results = []

    # Iterate through each sequence in the 'Sequence' column (adjust the column name if needed)
    for seq in df['Sequence']:
        aa_count = count_amino_acids(seq)
        ptm_count = calculate_ptm_occurrences(seq)

        # Combine both results
        result = {'Amino Acids': dict(aa_count), 'PTM Occurrences': ptm_count}
        amino_acid_counts.append(aa_count)
        ptm_results.append(ptm_count)

    # Add the results to the DataFrame
    aa_df = pd.DataFrame(amino_acid_counts)
    ptm_df = pd.DataFrame(ptm_results)

    # Add calculated PTM columns to the sequence DataFrame
    ptm_df = pd.DataFrame(ptm_occurrences)
    seq_df = pd.concat([seq_df, ptm_df], axis=1)

# Example usage
input_file = r"E:\PhD work\DataBase\PTM_DATA\analysis_october\data&figures_final\ptm_aa_extraction_example.xlsx"  # File with sequences and gene symbols
output_file = 'mapped_ptm_results.xlsx'  # Final output file

def process_excel(input_file, output_file):
    # Read the input Excel file
    df = pd.read_excel(input_file)

    # Create lists to hold the results
    amino_acid_counts = []
    ptm_results = []

    # Iterate through each sequence in the 'Sequence' column (adjust the column name if needed)
    for seq in df['Sequence']:
        aa_count = count_amino_acids(seq)
        ptm_count = calculate_ptm_occurrences(seq)

        # Combine both results
        result = {'Amino Acids': dict(aa_count), 'PTM Occurrences': ptm_count}
        amino_acid_counts.append(aa_count)
        ptm_results.append(ptm_count)

    # Add the results to the DataFrame
    aa_df = pd.DataFrame(amino_acid_counts)
    ptm_df = pd.DataFrame(ptm_results)

    # Concatenate the original data with the new results
    result_df = pd.concat([df, aa_df, ptm_df], axis=1)

    # Write the result to a new Excel file
    result_df.to_excel(output_file, index=False)

# Example usage:
input_file = r"input_filepath"  # Input Excel file with protein sequences
output_file = 'output_filepath'  # Output Excel file for PTM results
process_excel(input_file, output_file)
