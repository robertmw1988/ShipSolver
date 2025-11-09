# Local Data Transformer for Ship Solver

This script transforms the large AllData.csv into a preprocessed format ready for Google Sheets import.

## Setup

1. Install Node.js if you haven't already (https://nodejs.org/)

2. Install dependencies:
```bash
npm init -y
npm install csv-parse csv-stringify
```

3. Prepare your input files:

   - `AllData.csv` - Your full dataset
   - `Ship_Parameters.csv` - A CSV with your ship parameters, needs columns:
     ```
     Ship type,Ship level
     HENERPRISE,7
     ...etc...
     ```

4. Run the script:
```bash
node transform_data.js
```

The script will:
1. Read AllData.csv and Ship_Parameters.csv
2. Filter rows to match your ship parameters
3. Transform the data using the same logic as the Google Apps Script
4. Write the result to TransformedData.csv

## Output

The transformed data will be written to `TransformedData.csv` with:
- Key columns (Ship type, duration type, level, target artifact)
- Value columns for each artifact type/tier/rarity combination
- Drop counts in cells

You can then import this smaller, preprocessed dataset into Google Sheets.