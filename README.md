# Topic Volume Filler

A Streamlit tool that automatically fills in **Topic Volume**, **Reference Topic**, and **Match Type** columns in a master Excel sheet by matching topics against keyword data from zip-archived CSV files.

## Features

- 📂 Upload a **master Excel file** (with topics per country tab) and a **zip of keyword CSVs**
- 🔍 Smart topic matching using tokenised phrase matching, reverse phrase lookup, and cosine-similarity fallback
- 🐾 **Brand filtering** — excludes competitor/branded keywords; only Pedigree & Whiskas allowed by default
- 📄 **Source File Name** column — tracks which CSV inside the zip each match came from
- 🌍 Species guard — matches are filtered to relevant species (dog/cat) per row
- 🎨 Colour-coded output with Excel comments for match confidence
- 📊 Match Log tab summarising all matches, skips, and sources

## Requirements

- Python 3.9+
- `streamlit`
- `openpyxl`
- `pandas`

## Running Locally

```bash
pip install -r requirements.txt
streamlit run main.py
```

## Usage

1. Upload your **Master Excel file** (topics sheet with country tabs)
2. Upload a **zip file** containing keyword CSVs (one or more country zips)
3. Configure options in the sidebar (brand filtering, species guard, etc.)
4. Click **Run** and download the filled Excel output

## Supported Brands (Allowed in Matches)

- **Pedigree** (dog food)
- **Whiskas** (cat food)
- **Dentastix** (dental treats)

All other pet brand keywords (Royal Canin, Purina, IAMS, Hills, Sheba, etc.) are filtered out by default.
