import pandas as pd
import matplotlib.pyplot as plt
import geopandas as gpd
import numpy as np
import os
from matplotlib.patches import Wedge
from matplotlib.lines import Line2D
import matplotlib.patches as mpatches

# ==========================================
# 1. CONFIGURATION & SETTINGS
# ==========================================

input_file = r"G:\Paper\nema-Nanopore-Sequencing\zoology new zealand and australia\data\Location_Counts_Summary_With_Details.xlsx"
output_dir = os.path.dirname(input_file)

# --- USER SETTINGS ---
# تغییری که خواستید: افزایش سایز به 0.3
SCALE_FACTOR = 0.06       

SHOW_LABELS = False        
FONT_FAMILY = 'Arial'     
LABEL_FONT_SIZE = 9
TITLE_FONT_SIZE = 14

# ضریبی برای تبدیل واحد مختصات نقشه به واحد نقطه (Point) در نمودار
# این عدد باعث می‌شود سایز دایره‌های لجند با دایره‌های نقشه هم‌خوانی داشته باشد
LEGEND_PIXEL_CONVERSION = 45 

# ==========================================
# 2. DATA LOADING & PREPARATION
# ==========================================

print("Loading data...")
try:
    df = pd.read_excel(input_file)
except FileNotFoundError:
    print(f"Error: File not found at {input_file}")
    exit()

# Filter & Clean
df = df[df['Derived_Country'].isin(['Australia', 'New Zealand'])].copy()
df = df[~df['Derived_State'].isin(['Unknown State', 'Unknown Region'])]
df['Gene'] = df['files_info'].astype(str).str.replace('.xlsx', '', regex=False)

# Pivot Table
pivot_df = df.pivot_table(
    index=['Derived_Country', 'Derived_State'],
    columns='Gene',
    values='Count',
    aggfunc='sum'
).fillna(0)
pivot_df['Total_Count'] = pivot_df.sum(axis=1)

# --- COORDINATES ---
state_coordinates = {
    # Australia
    'New South Wales': (146.5, -32.5), 'Queensland': (144.5, -23.0),
    'South Australia': (135.5, -30.0), 'Tasmania': (146.8, -42.0),
    'Victoria': (144.5, -37.0), 'Western Australia': (122.5, -26.0),
    'Northern Territory': (133.5, -19.5), 
    
    # New Zealand
    'Northland': (173.8, -35.2), 'Auckland': (174.8, -36.9),
    'Waikato': (175.5, -38.0), 'Bay of Plenty': (176.8, -38.2),
    'Gisborne': (178.0, -38.6), 'Hawke\'s Bay': (176.8, -39.5),
    'Taranaki': (174.2, -39.3), 'Manawatu-Wanganui': (175.4, -40.0),
    'Wellington': (175.2, -41.2), 'Tasman': (172.8, -41.4),
    'Nelson': (173.2, -41.2), 'Marlborough': (173.6, -41.8),
    'West Coast': (170.8, -42.8), 'Canterbury': (171.2, -43.8),
    'Otago': (169.8, -45.2), 'Southland': (167.8, -45.8)
}

# --- ABBREVIATIONS ---
state_abbr = {
    # Australia
    'New South Wales': 'NSW', 'Queensland': 'QLD', 'South Australia': 'SA',
    'Tasmania': 'TAS', 'Victoria': 'VIC', 'Western Australia': 'WA',
    'Northern Territory': 'NT', 
    
    # New Zealand
    'Northland': 'NTL', 'Auckland': 'AKL', 'Waikato': 'WKO', 'Bay of Plenty': 'BOP',
    'Gisborne': 'GIS', 'Hawke\'s Bay': 'HKB', 'Taranaki': 'TAR', 'Manawatu-Wanganui': 'MW',
    'Wellington': 'WGN', 'Tasman': 'TAS', 'Nelson': 'NSN', 'Marlborough': 'MBH',
    'West Coast': 'WC', 'Canterbury': 'CAN', 'Otago': 'OTA', 'Southland': 'STL'
}

pivot_df['lon'] = pivot_df.index.get_level_values('Derived_State').map(lambda x: state_coordinates.get(x, (None, None))[0])
pivot_df['lat'] = pivot_df.index.get_level_values('Derived_State').map(lambda x: state_coordinates.get(x, (None, None))[1])
pivot_df['abbr'] = pivot_df.index.get_level_values('Derived_State').map(lambda x: state_abbr.get(x, x[:3].upper()))
pivot_df = pivot_df.dropna(subset=['lon', 'lat'])

# ==========================================
# 3. LOAD MAP GEOMETRY
# ==========================================
print("Loading high-res map data...")
map_url = "https://naturalearth.s3.amazonaws.com/10m_cultural/ne_10m_admin_1_states_provinces.zip"
try:
    world_gdf = gpd.read_file(map_url)
    oceania_full = world_gdf[world_gdf['admin'].isin(['Australia', 'New Zealand'])]
except:
    print("Warning: Using low-res fallback.")
    world = gpd.read_file(gpd.datasets.get_path('naturalearth_lowres'))
    oceania_full = world[world['name'].isin(['Australia', 'New Zealand'])]

# ==========================================
# 4. PLOTTING FUNCTION
# ==========================================

def create_gene_map(target_countries, title, filename_suffix, x_lim, y_lim, show_labels):
    print(f"Generating map for: {filename_suffix}...")
    
    # 1. Setup Figure
    fig, ax = plt.subplots(figsize=(14, 10))
    
    col_name = 'admin' if 'admin' in oceania_full.columns else 'name'
    map_subset = oceania_full[oceania_full[col_name].isin(target_countries)]
    data_subset = pivot_df[pivot_df.index.get_level_values('Derived_Country').isin(target_countries)]
    
    # Draw Map
    map_subset.plot(
        ax=ax, color='#F9F9F9', edgecolor='#555555', linewidth=0.8, linestyle='--'
    )
    
    ax.set_xlim(x_lim)
    ax.set_ylim(y_lim)
    ax.set_title(title, fontsize=TITLE_FONT_SIZE, fontname=FONT_FAMILY, pad=20)
    ax.set_xlabel("Longitude", fontname=FONT_FAMILY)
    ax.set_ylabel("Latitude", fontname=FONT_FAMILY)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)

    # 2. Colors
    gene_cols = [c for c in pivot_df.columns if c not in ['Total_Count', 'lon', 'lat', 'abbr']]
    colors = plt.cm.get_cmap('tab10')(np.linspace(0, 1, len(gene_cols)))
    
    # 3. Draw Pies
    def draw_pie(ax, x, y, ratios, size, colors):
        current_angle = 0
        for ratio, color in zip(ratios, colors):
            if ratio > 0:
                theta1 = current_angle
                theta2 = current_angle + (ratio * 360)
                # alpha=0.9 to make colors pop, linewidth=0.5 for clear separation
                wedge = Wedge((x, y), size, theta1, theta2, facecolor=color, alpha=0.9, edgecolor='black', linewidth=0.5)
                ax.add_patch(wedge)
                current_angle = theta2

    for idx, row in data_subset.iterrows():
        total = row['Total_Count']
        if total == 0: continue
        
        ratios = [row[gene] / total for gene in gene_cols]
        # Main sizing formula
        radius = np.sqrt(total) * SCALE_FACTOR
        
        draw_pie(ax, row['lon'], row['lat'], ratios, radius, colors)
        
        if show_labels:
            ax.text(
                row['lon'], row['lat'] + radius + (0.5 * (SCALE_FACTOR/0.08)), # Adjust offset based on scale
                row['abbr'], fontsize=LABEL_FONT_SIZE, fontname=FONT_FAMILY, 
                ha='center', va='bottom', fontweight='bold', color='#333333'
            )

    # 4. LEGENDS (Dynamic Calculation)
    
    # A. Genes
    color_handles = [mpatches.Patch(color=colors[i], label=gene_cols[i]) for i in range(len(gene_cols))]
    first_legend = plt.legend(
        handles=color_handles, title="Genes", loc='upper right', bbox_to_anchor=(1.18, 1), frameon=False
    )
    ax.add_artist(first_legend)
    
    # B. Size Reference (UPDATED LOGIC)
    size_values = [10, 50, 100]
    size_handles = []
    for val in size_values:
        # Step 1: Calculate the radius in data units (same as map)
        radius_data = np.sqrt(val) * SCALE_FACTOR
        
        # Step 2: Convert data radius to scatter point size (Area)
        # s (size) in scatter is in points^2. 
        # We use LEGEND_PIXEL_CONVERSION to map Data Units -> Points
        radius_points = radius_data * LEGEND_PIXEL_CONVERSION
        scatter_area = radius_points ** 2
        
        handle = plt.scatter([], [], s=scatter_area, color='gray', alpha=0.5, label=str(val), edgecolor='black')
        size_handles.append(handle)
        
    plt.legend(
        handles=size_handles, title="Sample Count", loc='upper right', 
        bbox_to_anchor=(1.18, 0.75), labelspacing=2.0, frameon=False, borderpad=1
    )
    
    # 5. Save
    out_png = os.path.join(output_dir, f"Map_{filename_suffix}.png")
    out_pdf = os.path.join(output_dir, f"Map_{filename_suffix}.pdf")
    
    plt.savefig(out_png, dpi=1200, bbox_inches='tight')
    plt.savefig(out_pdf, bbox_inches='tight')
    print(f"Saved: {out_png}")
    plt.close()

# ==========================================
# 5. EXECUTION
# ==========================================

# 1. Combined
create_gene_map(
    target_countries=['Australia', 'New Zealand'],
    title="Gene Distribution: Australia & New Zealand",
    filename_suffix="Combined_AU_NZ",
    x_lim=(110, 180), y_lim=(-50, -10), show_labels=SHOW_LABELS
)

# 2. Australia
create_gene_map(
    target_countries=['Australia'],
    title="Gene Distribution: Australia",
    filename_suffix="Australia_Only",
    x_lim=(112, 155), y_lim=(-45, -10), show_labels=SHOW_LABELS
)

# 3. New Zealand
create_gene_map(
    target_countries=['New Zealand'],
    title="Gene Distribution: New Zealand",
    filename_suffix="NewZealand_Only",
    x_lim=(165, 180), y_lim=(-48, -33), show_labels=SHOW_LABELS
)

print("All maps generated successfully.")