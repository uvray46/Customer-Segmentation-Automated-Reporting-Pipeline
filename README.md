# Customer-Segmentation-Automated-Reporting-Pipeline 
<img width="1349" height="600" alt="image" src="https://github.com/user-attachments/assets/009df762-f202-46b8-9a2f-c125e535f753" />

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
