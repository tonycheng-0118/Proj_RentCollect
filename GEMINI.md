# Project Overview

This is a Google Apps Script project for managing rent collection. It automates the process of tracking rent payments, utility bills, and other costs associated with rental properties. The script imports and parses bank records from various banks, matches payments to tenant contracts, calculates arrears, and generates reports.

## Key Features

*   **Automated Rent Collection:** The script automates the process of tracking rent payments against contracts.
*   **Bank Record Parsing:** It parses bank statements from multiple banks (E.Sun, KTB, CTBC) to identify tenant payments.
*   **Contract Management:** The script manages rental contracts, including start and end dates, rent amount, and deposit.
*   **Utility and Miscellaneous Cost Tracking:** It tracks utility bills and other miscellaneous costs associated with each property.
*   **Reporting:** The script generates reports on rent arrears, contract status, and other relevant information.
*   **Testing:** The project uses the `QUnitGS2` library for unit testing.

## Project Structure

The project is organized into several modules, each responsible for a specific part of the rent collection process:

*   `rentCollect_main.js`: The main script that orchestrates the entire rent collection process.
*   `rentCollect_parser.js`: Parses data from various sources, including bank records, tenant information, and contracts.
*   `rentCollect_contract.js`: Contains the core logic for managing contracts and calculating rent arrears.
*   `rentCollect_report.js`: Generates reports based on the processed data.
*   `rentCollect_import.js`: Handles the import of data from external sources.
*   `rentCollect_misc.js`: Contains miscellaneous utility functions.
*   `rentCollect_pkg.js`: Defines data structures and classes used throughout the project.
*   `tests.js`: Contains unit tests for the project.

## Building and Running

This project is a Google Apps Script project and is meant to be run in the Google Apps Script environment.

### Pushing to Google Apps Script Environment

The project uses `clasp` to manage the project locally. To push the code to the Google Apps Script environment, use the following command:

```bash
clasp push
```

### Running the Script

To run the script, open the project in the Google Apps Script editor and run the `rentCollect_main` function.

## Development Conventions

*   **Coding Style:** The code follows a consistent style, with clear variable names and comments.
*   **Error Handling:** The script includes error handling and logging to help with debugging.
*   **Modularity:** The code is organized into modules to improve maintainability.
*   **Testing:** The project uses `QUnitGS2` for unit testing. New features should be accompanied by corresponding tests.
