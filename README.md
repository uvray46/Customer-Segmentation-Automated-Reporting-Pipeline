# Customer-Segmentation-Automated-Reporting-Pipeline 
<img width="1280" height="720" alt="image" src="[https://cdn.prod.website-files.com/6064b31ff49a2d31e0493af1/66693af8a543407c30c675e8_AD_4nXdbVMeQrllNgJGkHvWortq6GLCa8JX45YmYsOiSqgUpmbTD0xtfyD5zkW5Mf7e1Hz0TjkQBdZDUEuqqQOmQoUgEqdmVcBRfidU6zs4c7x3XtPkbnF0awgB6AaY7wtzrvyG4B2iObP9QweeWmtoXakuhZVu7.jpeg]" />

**Author:** James Whitmarsh

**Date:** 09/01/2025 

**Tooling:** PostgreSQL, Python (pandas + SQLAlchemy), Power BI Desktop 

## Introduction  
Automated analytics pipeline that transforms raw customer data into segmented business insights and generates executive-ready Excel and Word reports. Designed to demonstrate scalable data transformation, reporting automation, and repeatable analytics workflows used to support data-driven decision making. 

## This package includes the following:

- `bank.csv`: The marketing dataset
- `automated_marketing_report_script.py`: Python script to generate automated Excel and Word reports
- `automated_marketing_report.xlsx`: (Will be created on run) Excel report including targeting strategy and chart
- `automated_marketing_report.docx`: (Will be created on run) Word report including key insights and visuals

## Requirements

- Python 3.8+
- Libraries:
    - pandas
    - matplotlib
    - python-docx
    - xlsxwriter

## To Run

1. Ensure all files are in the same directory.
2. Open a terminal or command prompt in this directory.
3. Run:

```
pip install pandas matplotlib python-docx xlsxwriter
python automated_marketing_report_script.py
```

4. Two output files will be created:
   - `automated_marketing_report.xlsx`
   - `automated_marketing_report.docx`

These can be shared with stakeholders and executives.
