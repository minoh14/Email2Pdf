This library converts emails to PDF files, so that we can feed them to UiPath Document Understanding for data extraction.

The conversion is done using a python script which borrows most of its code from **[msgtopdf](https://pypi.org/project/msgtopdf/)**.

This library requires the following prerequisites on the robot machine:

* Microsoft Outlook installed
* [wkhtmltopdf(x64)](https://github.com/wkhtmltopdf/packaging/releases/download/0.12.6-1/wkhtmltox-0.12.6-1.msvc2015-win64.exe) installed (and must be on PATH)
* Python 3.10 x64  
  pip install **pywin32**  

Supported file formats:

* Input:  MSG
* Output: PDF

Limitations:

* All the paths provided to this library must be absolute. Relative path leads to exceptions.
