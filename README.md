# Customer-Segmentation-Automated-Reporting-Pipeline 
<img width="1280" height="720" alt="image" src="[https://www.bing.com/images/search?view=detailV2&ccid=vRknxL1v&id=AB272097DAB1217C4350F7192D7A25EB73F855FB&thid=OIP.vRknxL1vrWee-lILmtJKEgHaEJ&mediaurl=https%3a%2f%2fimg.freepik.com%2fpremium-photo%2fschematic-data-pipeline-background_1046712-840.jpg&cdnurl=https%3a%2f%2fth.bing.com%2fth%2fid%2fR.bd1927c4bd6fad679efa520b9ad24a12%3frik%3d%252b1X4c%252bslei0Z9w%26pid%3dImgRaw%26r%3d0&exph=351&expw=626&q=data+pipline+visual+background&FORM=IRPRST&ck=CE4271E4B95DF2BD2B107FF9CE7E4E64&selectedIndex=2&itb=0]" />

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
