const fs = require('fs');
const path = require('path');
const csv = require('csv-parse/sync');
const { stringify } = require('csv-stringify/sync');

// Configuration (same structure as the Apps Script version)
const CFG = {
    inputFile: path.join(__dirname, 'AllData.csv'),
    outputFile: path.join(__dirname, 'TransformedData.csv'),
    paramsFile: path.join(__dirname, 'Ship_Parameters.csv'),
};

// Read and parse CSV file
function readCSV(filePath) {
    const content = fs.readFileSync(filePath, 'utf-8');
    return csv.parse(content, {
        columns: true,
        skip_empty_lines: true
    });
}

// Write array of arrays to CSV
function writeCSV(filePath, data) {
    const output = stringify(data);
    fs.writeFileSync(filePath, output);
}

function transformData() {
    // Read input data
    console.log('Reading data from:', CFG.inputFile);
    const rows = readCSV(CFG.inputFile);
    
    console.log('Transforming all rows:', rows.length);

    // Define column structure for the pivot
    const keyCols = ['Ship type', 'Ship duration type', 'Ship level', 'Target artifact'];
    const valueCols = ['Artifact type', 'Artifact tier', 'Artifact rarity'];
    
    // Initialize pivot data structures
    const pivot = {};
    const allValueKeys = new Set();
    
    // First pass: Build pivot and collect all value combinations
    for (const row of rows) {
        // Create the key (e.g., "HENERPRISE | SHORT | 7 | PROPHECY_STONE")
        const keyParts = keyCols.map(col => row[col]);
        const keyStr = keyParts.join(' | ');
        
        // Create the value key (e.g., "PROPHECY_STONE | 4 | LEGENDARY")
        const valueStr = valueCols.map(col => row[col]).join(' | ');
        
        // Get drop count
        const drops = parseInt(row['Total drops']) || 0;
        
        // Add to pivot structure
        if (!pivot[keyStr]) {
            pivot[keyStr] = {};
        }
        pivot[keyStr][valueStr] = (pivot[keyStr][valueStr] || 0) + drops;
        
        // Track all possible value combinations
        allValueKeys.add(valueStr);
    }
    
    // Convert value keys to sorted array and parse components
    const valueKeysList = Array.from(allValueKeys).sort();
    const valueComponents = valueKeysList.map(key => {
        const [artifactType, tier, rarity] = key.split(' | ');
        return { artifactType, tier, rarity, fullKey: key };
    });

    // Create the three header rows
    const outputRows = [
        // Row 1: Key column headers + artifact types
        [...keyCols, ...valueComponents.map(v => v.artifactType)],
        // Row 2: Empty in key columns + tiers
        ['', '', '', '', ...valueComponents.map(v => v.tier)],
        // Row 3: Empty in key columns + rarities
        ['', '', '', '', ...valueComponents.map(v => v.rarity)],
        // Row 4: Complete column headers
        [...keyCols, ...valueKeysList]
    ];

    // Create data rows from pivot
    for (const [key, values] of Object.entries(pivot)) {
        const keyParts = key.split(' | ');
        const row = [...keyParts];
        
        // Add value columns in same order as headers
        for (const valueKey of valueKeysList) {
            row.push(values[valueKey] || 0);
        }
        outputRows.push(row);
    }

    // Write output
    console.log('Writing transformed data to:', CFG.outputFile);
    writeCSV(CFG.outputFile, outputRows);
    
    console.log('Transformation complete!');
    console.log('- Input rows:', rows.length);
    console.log('- Output row groups:', outputRows.length - 1);
    console.log('- Value combinations:', valueKeysList.length);
}

// Install dependencies first:
// npm init -y
// npm install csv-parse csv-stringify

try {
    transformData();
} catch (err) {
    console.error('Error:', err.message);
    if (err.code === 'ENOENT') {
        console.error('\nMake sure you have the following files:');
        console.error('1. AllData.csv - The full dataset');
        console.error('2. Ship_Parameters.csv - A CSV with columns: "Ship type,Ship level"');
    }
}