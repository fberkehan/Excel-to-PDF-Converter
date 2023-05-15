# Excel-to-PDF-Converter
This repository contains a simple Windows Forms application developed with C# in Visual Studio. The primary function of the application is to convert Excel files (.xlsx, .xls, .csv, .xml, etc.) into PDF format.
Excel to PDF Converter
This repository contains a simple Windows Forms application developed with C# in Visual Studio. The primary function of the application is to convert Excel files (.xlsx, .xls, .csv, .xml, etc.) into PDF format.

Features
User-friendly interface with a single button for file selection.
Reads and interprets Excel files, transforming tables into rows and columns formatted for A4 size.
Converts and saves the formatted document as a PDF file.
Automatically saves the converted PDF to a folder named "EXCEL2PDF" on the user's desktop.
If the "EXCEL2PDF" folder does not exist, the application creates it.
Usage
To use the application, simply press the "Select File" button and choose the Excel file you want to convert. The application will handle the conversion and automatically save the PDF in the "EXCEL2PDF" folder on your desktop.

Dependencies
The application uses the following NuGet packages:

ExcelDataReader for reading Excel files.
iTextSharp for creating PDF files.
Build
The application is built using Visual Studio. To create a single .exe file, use the "Publish" feature with the "Produce single file" option enabled.

We hope you find this application useful! For any issues, suggestions or enhancements, feel free to open an issue or submit a pull request.
Installation
Before running the application, make sure you have .NET Framework installed on your machine.

To install the application:

Download the repository.
Navigate to the 'bin' then 'Release' or 'Debug' directory.
Run the 'exceltopdf.exe' executable file.
Contribution
Contributions to this repository are always welcome. Whether you have a feature request, bug report, or a pull request, we appreciate all feedback and contributions. Please don't hesitate to open an issue or submit a pull request.

License
This project is licensed under the MIT License. See the LICENSE file in the repository for more information.

Contact
For any additional information, you can reach out by opening an issue on this repository. We'd be happy to help or answer any questions you may have.

Thank you for visiting this repository, and we look forward to seeing your awesome contributions!

Disclaimer
The code in this repository is provided "as is" without warranty of any kind, either expressed or implied, including limitation warranties of merchantability, fitness for a particular purpose, and noninfringement. In no event shall the authors or copyright holders be liable for any claim, damages, or other liability, whether in an action of contract, tort or otherwise, arising from, out of or in connection with the software or the use or other dealings in the software.
