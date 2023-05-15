# Automated-report-
VBA project

I prepared a macro - auto-generating report which:

1) copy data from regional reports to main report; ranges are dynamic (spotting last used row), 
3) create table from copied data, format date as "dd/mm/yyyy", adjust columns width,
4) create in new workbook pivot table based on data from main report,
5) create pivot chart in the same file,
6) copy data from main report to separate csv file. 

Every regional report looks like this:
![regional_report](https://github.com/korneldata/Automated-report-/blob/e5d58ef5d00cdd9ce74a6f23a05756c290027bc6/regional_report.jpg)

Data from regional reports was copied and formatted to main report file:
![main_report](https://github.com/korneldata/Automated-report-/blob/e5d58ef5d00cdd9ce74a6f23a05756c290027bc6/main_report.jpg)

Based on the main report, pivot table was created in separate worbook:
![pivot_table](https://github.com/korneldata/Automated-report-/blob/e5d58ef5d00cdd9ce74a6f23a05756c290027bc6/pivot_table.jpg)

Then pivot chart was added:
![pivot_chart](https://github.com/korneldata/Automated-report-/blob/e5d58ef5d00cdd9ce74a6f23a05756c290027bc6/pivot_chart.jpg)

And main report data was copied to separate csv file:
![csv](https://github.com/korneldata/Automated-report-/blob/e5d58ef5d00cdd9ce74a6f23a05756c290027bc6/csv.jpg)
