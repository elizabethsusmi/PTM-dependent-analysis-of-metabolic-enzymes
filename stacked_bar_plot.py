import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import seaborn as sns
from matplotlib import rcParams

# Set Helvetica Bold globally
rcParams['font.family'] = 'Helvetica'
rcParams['font.weight'] = 'bold'

# Load data from Excel
file_path = r"inputfile.xlsx"  # Your Excel file path
sheet_name = "Sheet2"  # Your sheet name
data = pd.read_excel(file_path, sheet_name=sheet_name)

# Set gene symbols as the index (optional)
data.set_index("GeneSymbol", inplace=True)

# Calculate total PTM counts per gene for sorting
data['Total'] = data.sum(axis=1)

# Sort data descending by total counts
data_sorted = data.sort_values(by='Total', ascending=False)

# Drop the 'Total' column for plotting cleanly
data_sorted = data_sorted.drop(columns=['Total'])

# Extract PTM types and sorted data values
ptm_types = data_sorted.columns
gene_symbols_sorted = data_sorted.index
values_sorted = data_sorted.values

# Calculate cumulative sums for stacking with sorted data
cumulative_data_sorted = np.cumsum(values_sorted, axis=1)

# Generate unique colors using Seaborn palettes (adjusted for number of PTM types)
palette = sns.color_palette("Set3", n_colors=15) + sns.color_palette("tab20", n_colors=14)
unique_colors = palette[:len(ptm_types)]

# Create figure with constrained layout
fig, ax = plt.subplots(figsize=(15, 8), constrained_layout=True)

# Plot stacked bars with sorted data
for i, ptm in enumerate(ptm_types):
    bottom = cumulative_data_sorted[:, i - 1] if i > 0 else None
    ax.bar(gene_symbols_sorted, values_sorted[:, i], bottom=bottom, label=ptm, color=unique_colors[i % len(unique_colors)])

# Add labels, title and legend
ax.set_xlabel("Gene Symbols", fontsize=12, fontweight='bold')
ax.set_ylabel("Count of PTM Sites", fontsize=12, fontweight='bold')
# ax.set_title("PTM Site Counts Across Genes (Sorted by Total)", fontsize=14, fontweight='bold')
ax.legend(
    title="PTM Types",
    fontsize=8,
    loc="upper right",
    bbox_to_anchor=(1, 1),
    frameon=True
)

# Adjust margins and x-axis to fit bars well
ax.margins(x=0.09)
ax.set_xlim(-0.5, len(gene_symbols_sorted) - 0.5)

# Rotate x-axis ticks for readability
plt.xticks(rotation=90, fontsize=3, fontweight="normal")
plt.tight_layout()
plt.show()
