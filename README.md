## Setup
Install all required packages:
```bash
pip install -r requirements.txt
```
## Steps to Run
#### 1. Place the Excel file *excel_sample_data_qae.xlsx* in the *data/* folder.
#### 2. Run the scraper script to download the latest exchange rates and convert Excel sheets to CSV:
```bash
python scrape.py
```
This will generate the following CSVs in the *data/* folder:
- exchange-rates.csv
- sales.csv
- product.csv
- store.csv

#### 3. Run the processing script to process and aggregate the sales data:
```bash
python processing.py
```
#### 4. Result will be exported as *aggregated_sales_report.xlsx* in the *data/* folder.
The report includes 2 sheets:
- By_Region → sales aggregated by store region
- By_Category → sales aggregated by product category
