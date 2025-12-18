
```markdown
# Host Ecology and Geospatial Distribution Pipeline

This repository contains the computational framework and source code associated with the study **"Noshadi et al. (2025)"**.
Link https://github.com/ArashNoshadi/long-read-metabarcoding.git 

The pipeline is designed to analyze nucleotide sequence metadata retrieved from NCBI GenBank, specifically focusing on samples from **Australia and New Zealand**. It performs automated host categorization based on ecological roles and generates high-resolution geospatial visualizations (pie charts overlaid on state maps) to illustrate gene distribution patterns.

---

## 📋 Table of Contents
- [Overview](#overview)
- [Features](#features)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
  - [1. Host Categorization](#1-host-categorization)
  - [2. Geospatial Visualization](#2-geospatial-visualization)
- [Data Structure](#data-structure)
- [Citation](#citation)
- [License](#license)

---

## Overview

This project provides a robust workflow for:
1.  **Data Curation:** Filtering GenBank data for Australia and New Zealand using geotags.
2.  **Taxonomic Validation:** Cleaning misannotations and sequencing artifacts (as detailed in Noshadi et al., 2026).
3.  **Ecological Analysis:** Classifying host organisms into a hierarchical structure (Plants, Animals, Insects).
4.  **Mapping:** Plotting gene distribution/prevalence using pie charts projected onto geographic coordinates.

---

## Features

### 🧬 Host Categorization Module
An automated dictionary-based classifier that sorts host organisms into three primary domains and specific subgroups:
* **Animals:** Terrestrial (Mammals, Birds, Reptiles, Amphibians) & Aquatic (Fish, Invertebrates).
* **Plants:** Herbaceous (Crops, Vegetables) & Woody (Trees, Shrubs).
* **Insects:** Functional groups (Herbivores, Predators/Parasitoids, Pollinators, Vectors).

### 🗺️ Geospatial Visualization Module
* **High-Resolution Mapping:** Layers data at the state and city level.
* **Pie Chart Overlays:** Generates maps where each Australian/New Zealand state is represented by a pie chart showing the proportion of different genes or host categories found in that region.

---

## Prerequisites

The code is written in **Python**. Ensure you have the following dependencies installed:

* Python (Version 3.1 or higher recommended)
* `pandas` (Data manipulation)
* `matplotlib` (Visualization)
* `basemap` or `cartopy` (Geospatial plotting)
* `numpy`

---

## Installation

1.  **Clone the repository:**
    ```bash
    git clone [https://github.com/your-username/your-repo-name.git](https://github.com/your-username/your-repo-name.git)
    cd your-repo-name
    ```

2.  **Install dependencies:**
    ```bash
    pip install pandas matplotlib numpy
    # Note: Installation of Basemap or Cartopy may require specific binaries depending on your OS.
    ```

---


### 2. Geospatial Visualization

Generate the distribution maps with pie charts overlaid on states.

```bash
python run_mapping.py --input data/categorized_data.csv --output results/maps/

```

*The script will output `.png` or `.svg` files for each gene/family analyzed.*

---

## Data Structure

The input data should ideally be a CSV file containing at least the following columns:

| Column Name | Description |
| --- | --- |
| `Accession` | GenBank Accession Number |
| `Organism` | Name of the organism |
| `Host` | Host organism name (to be categorized) |
| `Country` | Country of origin (Australia/New Zealand) |
| `State` | State or Region |
| `Latitude` | Decimal Latitude |
| `Longitude` | Decimal Longitude |
| `Collection_Date` | Date of sampling |

---

## Citation

If you use this code or dataset in your research, please cite the original paper:

> **Noshadi et al. (2026).** [Title of Your Paper]. 
 DOI: [Insert DOI]

---


