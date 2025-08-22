# The Report Automator

This Python script is designed to automatically create analytical Excel reports from sales data files (`.csv` or `.xlsx`). It takes raw data, processes it, and generates a clean, multi-sheet report complete with charts.

## Features

-   **Multi-Format Support:** Reads both `.csv` and `.xlsx` files seamlessly.
-   **Flexible Date Filtering:** Allows filtering the data for any specific date range.
-   **Informative Reports:** Generates multiple reports (e.g., sales by manager, top 5 products) on separate Excel sheets.
-   **Data Visualization:** Automatically builds a column chart to visualize manager performance.
-   **Fully Configurable:** All column names and script settings can be easily adjusted in the `config.ini` file without touching the source code.

## Setup and Configuration

1.  **Clone the repository:**
    ```bash
    git clone https://github.com/freddieBear13/excel-csv-automator-example
    ```

2.  **Navigate to the project directory:**
    ```bash
    cd excel-csv-automator-example
    ```

3.  **Install the required libraries:**
    ```bash
    pip install -r requirements.txt
    ```

4.  **Configure `config.ini`:**
    Open the `config.ini` file and ensure the column names under the `[COLUMNS]` section match the column headers in your source data file.

## Usage

The script is run from the command line.

**Basic Usage (to process the entire file):**
```bash
python report_automator.py your_data_file.xlsx
```

## Usage with Date Filtering
```bash
python report_automator.py your_data_file.csv --start 2025-08-01 --end 2025-08-15
```
This will create a new Excel report file (e.g., your_data_file_report.xlsx) in the same directory as the source file.
