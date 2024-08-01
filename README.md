How to convert content .MSG to .PDF using Python. 
Step 1: Install Python and Required Libraries
1.	Install Python:
o	If you don't have Python installed, download and install it from Python's official website.
2.	Install Required Libraries:
o	Open the Command Prompt or Terminal.
o	Run the following commands to install pywin32 and pdfkit:
“pip install pywin32 pdfkit”
Step 2: Install wkhtmltopdf
1.	Download and Install wkhtmltopdf:
o	Download wkhtmltopdf from wkhtmltopdf's official website.
o	Install it and make a note of the installation path (usually C:\Program Files (x86)\wkhtmltopdf\bin\wkhtmltopdf.exe). < this part is critical because this is the location the script will call to execute. 
Step 3: Prepare the Python Script
It’s part of the content MSG to PSF folder
 
 
Step 4: Run the Script
1.	Close Microsoft Outlook:
o	Ensure that Microsoft Outlook is completely closed. The script needs Outlook to be closed to access the .msg files.
2.	Run the Script:
o	Open the Command Prompt (Windows) or Terminal (macOS/Linux).
o	Navigate to the directory where you saved the script. For example:
cd "C:\Users\YourUsername\Documents\MSG to PDF\"
3.	Run the script using Python:
python msg_to_pdf_converter.py
It should look something like this: 
C:\Users\Administrator\OneDrive\OneDrive – YourCompany\Documents\MSG to PDF> python msg_to_pdf_converter.py
 

Step 5: Use the Graphical Interface
1.	Select MSG Files:
o	A window will pop up with a button labeled "Select MSG Files".
 
o	Click on this button to open a file dialog.
o	Navigate to the folder containing your .msg files, select the files you want to convert, and click "Open".
2.	Convert to PDF:
o	After selecting the files, the button "Convert to PDF" will become enabled.
o	Click on this button to start the conversion process.
o	The script will convert each .msg file to a .pdf file and display a message when the conversion is complete.

 
Troubleshooting
•	File Permission Errors: Ensure you have read and write permissions for the files and directories you are working with.
•	Outlook Errors: Make sure Outlook is closed before running the script.
•	Library Installation Issues: If you encounter issues with installing the libraries, make sure you have the correct version of Python and pip installed.
Notes
•	Images: The script attempts to include embedded images from the .msg files in the converted PDFs. If some images are not displayed, it may be due to how they are embedded or referenced in the email.
•	File Cleanup: The script cleans up temporary HTML files and any saved images after conversion.


