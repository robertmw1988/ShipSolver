# GitHub Copilot Instructions for This Repository

## Project Overview

This project is creating a Google Sheets app script. The app scripts will be used for indexing a data subset from a google sheet 'AllArtifactsData' into multiple other sheets based on certain criteria. The indexed data will then be used for various purposes including analysis, and solving optimization problems using linear programming solvers.

## Coding Standards and Conventions

1.  **Language:** All new code should be written in JavaScript ES5 compatible with Google Apps Script.

2.  **Data Handling:**
    *   Use Google Apps Script's built-in `SpreadsheetApp` service for all spreadsheet interactions.
    *   Ensure efficient data retrieval and manipulation to minimize execution time.

3.  **Naming Conventions:**
    *   Variables : `camelCase`.
    *   utility and helperFunctions : `_camelCase()`.
    *   Sheet functions : `UPPER_SNAKE_CASE()`.
    *   Files : `kebab-case` for non-component files, `PascalCase.tsx` for React components.

## Testing Guidelines

*   Write Diagnostic functions to validate data manipulation.

## Deployment Information

This application is deployed as a Google Sheets add-on. Ensure that any changes made are compatible with the Google Apps Script environment and do not exceed execution time limits.

## Security Considerations

*  Avoid hardcoding sensitive information such as API keys or credentials. Use Google Apps Script's Properties Service for storing such data securely.

## Tasks

1.  **Transform ALL Artifact Data:**
    *   Create a function to read data from 'AllArtifactsData' sheet and transform it into a structured format for further processing.
    *   The function should read the data, pivot it, and write the transformed data to a new sheet.
    *   Rows are defined by: Ship type, Ship duration type, Ship level, and Target artifact
    *   Columns are combinations of: Artifact type, Artifact tier, and Artifact rarity
    *   Values are the sum of Total drops for each combination.
    *   There are columns that are not included in the data set, which are added to the table for completeness. These columns are filled with zeros.
    *   Ensure the transformed data is easily accessible for indexing.



