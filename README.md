# PyDF Report Certificate Generator

This is an old Python program that I create to help educators produce an automatic student reports/certificate.

It is running on Python 2.7, using reportlab to produce the PDF file and also xlrd to grab and parse data from an MS Excel document.

Before running the script, make sure the required modules are installed by running this command `pip install -r requirements.txt`

Click "Open" icon to open an Excel file with specific format. For instance "Data Example.xls". After that the program will show the tables. Then click "Generate PDF Report(s)" icon.

The program will generate reports as many as lines in the Excel files. The report will look like this.

Frontside: 
![][reportFrontsied.jpg]

Backside:
![][ReportBackside.jpg]