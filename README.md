This library converts emails to PDF files, so that we can feed them to UiPath Document Understanding for data extraction.

The conversion is done using a python script which borrows most of its code from **[msgtopdf](https://pypi.org/project/msgtopdf/)**.

This library requires the following prerequisites on the robot machine:

* Microsoft Outlook installed
* Python 3.10 x64  
  pip install **pywin32**  
  pip install **wkhtmltopdf**

Supported file formats:

* Input:  MSG
* Output: PDF

Limitations:

* All the paths provided to this library must be absolute. Relative path leads to exceptions.
