# catho-pdf-reader

## What does it do?

This application -- a _.NET 7 Worker Service_ -- reads (specific) data from PDF curriculum files from **Catho** and saves it in an Excel file.

The following table presents the data read from the PDF file:

| Data name (in portuguese) | Description                               							  |
| --------------------------|:-----------------------------------------------------------------------:|
| Nome                      | Applicant's name                          							  |
| Local                     | Applicant's residence city + state                					  |
| Celular                   | Applicant's cellphone number              							  |
| Email                     | Applicant's email                         							  |
| LinkedIn                  | Applicant's LinkedIn profile (**default value**)      			  	  |
| Salário                   | Applicant's salary proposal     										  |
| Como soube da vaga?       | How did the applicant know about the job offer? (**default value**)     |


## How does it work?

The process works in the following order:

- The application lists the PDF files to be processed in the \ToProcess directory;
- After listing the files, the application extracts the data from the PDF files;
- After extracting, the application moves the read PDF files to the \Processed directory, and then saves the generated Excel file at the \GeneretedExcels directory.