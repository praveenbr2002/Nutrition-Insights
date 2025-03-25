#!/usr/bin/env python
# coding: utf-8

# In[3]:


# Import necessary libraries
import pandas as pd

# Load the Excel file
file_path = 'Pulse database for projects 2.xlsx'  # Update the path to your actual file
df = pd.read_excel(file_path)

# Remove any leading/trailing whitespace from column names
df.columns = df.columns.str.strip()

# Step 1: Define the essential amino acids columns we need
essential_amino_acids = ["THR", "VAL", "ILE", "LEU", "HIS", "LYS", "TRP", "M+C", "P+T"]
required_columns = ["SAMPLE", "PROTEIN %"] + essential_amino_acids

# Step 2: Check for the existence of Met, Cys, Phe, and Tyr columns for summation
if 'MET' in df.columns and 'CYS' in df.columns:
    df['M+C'] = df['MET'] + df['CYS']
if 'PHE' in df.columns and 'TYR' in df.columns:
    df['P+T'] = df['PHE'] + df['TYR']

# Select only the required columns and make a copy to avoid SettingWithCopyWarning
df_selected = df[required_columns].copy()

# Step 3: Calculate the Essential Amino Acids (mg/g protein) using the provided formula
for amino_acid in essential_amino_acids:
    df_selected[amino_acid] = (df_selected[amino_acid] * 1000) / df_selected["PROTEIN %"]

# New reference patterns
reference_patterns_1991 = {
    "Infant": {
        "THR": 43, "VAL": 55, "ILE": 46, "LEU": 93, "HIS": 26,
        "LYS": 66, "TRP": 17, "M+C": 42, "P+T": 72
    },
    "Pre_school_child": {
        "THR": 34, "VAL": 35, "ILE": 28, "LEU": 66, "HIS": 19,
        "LYS": 58, "TRP": 11, "M+C": 25, "P+T": 63
    },
    "School_child": {
        "THR": 28, "VAL": 25, "ILE": 30, "LEU": 44, "HIS": 19,
        "LYS": 44, "TRP": 9, "M+C": 22, "P+T": 22
    },
    "Adult": {
        "THR": 9, "VAL": 13, "ILE": 13, "LEU": 19, "HIS": 16,
        "LYS": 16, "TRP": 5, "M+C": 17, "P+T": 19
    }
}

reference_patterns_2007 = {
    "0.5": {
        "THR": 31, "VAL": 43, "ILE": 32, "LEU": 66, "HIS": 20,
        "LYS": 57, "TRP": 8.5, "M+C": 27, "P+T": 52
    },
    "1-2": {
        "THR": 27, "VAL": 42, "ILE": 31, "LEU": 63, "HIS": 18,
        "LYS": 52, "TRP": 7.4, "M+C": 25, "P+T": 46
    },
    "3-10": {
        "THR": 25, "VAL": 40, "ILE": 31, "LEU": 61, "HIS": 16,
        "LYS": 48, "TRP": 6.6, "M+C": 23, "P+T": 41
    },
    "11-14": {
        "THR": 25, "VAL": 40, "ILE": 30, "LEU": 60, "HIS": 16,
        "LYS": 48, "TRP": 6.5, "M+C": 23, "P+T": 41
    },
    "15-18": {
        "THR": 24, "VAL": 40, "ILE": 30, "LEU": 60, "HIS": 16,
        "LYS": 47, "TRP": 6.3, "M+C": 23, "P+T": 40
    },
    "18+": {
        "THR": 23, "VAL": 39, "ILE": 30, "LEU": 59, "HIS": 15,
        "LYS": 45, "TRP": 6, "M+C": 22, "P+T": 38
    }
}

reference_patterns_2013 = {
    "Infant": {
        "THR": 44, "VAL": 55, "ILE": 55, "LEU": 96, "HIS": 21,
        "LYS": 69, "TRP": 17, "M+C": 33, "P+T": 94
    },
    "Child": {
        "THR": 31, "VAL": 43, "ILE": 32, "LEU": 66, "HIS": 20,
        "LYS": 57, "TRP": 8.5, "M+C": 27, "P+T": 52
    },
    "3+": {
        "THR": 28, "VAL": 40, "ILE": 30, "LEU": 61, "HIS": 16,
        "LYS": 48, "TRP": 6.6, "M+C": 23, "P+T": 41
    }
}


# List to hold all reference patterns
all_reference_patterns = {
    "1991": reference_patterns_1991,
    "2007": reference_patterns_2007,
    "2013": reference_patterns_2013
}

# Dictionary to store all tables for each reference pattern group
all_tables = {}

# Loop through each reference pattern set
for year, patterns in all_reference_patterns.items():
    all_tables[year] = {}
    for pattern_name, pattern_values in patterns.items():
        # Create a temporary DataFrame for the current pattern
        df_pattern = df_selected[["SAMPLE", "PROTEIN %"]].copy()
        
        # Calculate AAS for each essential amino acid based on the current pattern
        for amino_acid in essential_amino_acids:
            df_pattern[amino_acid] = df_selected[amino_acid] / pattern_values[amino_acid]
        
        # Add PATTERN and AGE columns
        df_pattern["PATTERN"] = year
        df_pattern["AGE"] = pattern_name
        
        # Calculate the limiting amino acid score (minimum AAS) for the current pattern
        df_pattern["AAS"] = df_pattern[essential_amino_acids].min(axis=1)
        
        # Identify the limiting amino acid for each row
        df_pattern["LIMITING AMINO ACID"] = df_pattern[essential_amino_acids].idxmin(axis=1)
        
        # Round the results and store the DataFrame in the dictionary
        df_pattern = df_pattern.round(2)
        all_tables[year][pattern_name] = df_pattern

# Save all tables to a single Excel file with separate sheets
output_file = "amino_acid_scores_new.xlsx"

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    for year, patterns in all_tables.items():
        for pattern_name, table in patterns.items():
            sheet_name = f"{year}_{pattern_name}"[:31]  # Limit sheet name to 31 characters
            table.to_excel(writer, index=False, sheet_name=sheet_name)

print(f"All tables have been successfully exported to {output_file}.")


# In[ ]:




