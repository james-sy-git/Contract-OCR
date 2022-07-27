# Contract-OCR

Optimal character recognition app to perform legal contract parsing.

Created to convert previously scanned (e.g. non-machine readable) contracts, and extract a clause of interest, which is then
packaged and outputted as an Excel spreadsheet. Can also write to an existing spreadsheet.

Uses tesseract (wrapper: pytesseract) to perform OCR, OpenCV and PyMuPDF to handle the file processing and conversion, and OpenPyXL 
to create a spreadsheet output.

Integrates a Tkinter GUI, packaged using PyInstaller in the final release, for usability and portability.

Built Kivy GUI as well, uisng kvlang, but Python 3.8 + PyInstaller + Kivy led to incompatibility issues.

Averaged a 10.5 second runtime per document, approx. 0.42 seconds per page.

UX:

1. Choose non-machine readable PDFs from any drive
2. Enter search query (title of clause to be extracted)
3. Choose working directory where the file and image processing will
   take place, and also where a new Excel spreadsheet will be written
4a. If a new spreadsheet is desired, enter the file name and sheet name
4b. To write to an existing spreadsheet, choose it using the file dialog
5. Send it!!
6. Find desired output in the Excel file -- if the clause is not found in
   a certain PDF, the output cell will be blank
