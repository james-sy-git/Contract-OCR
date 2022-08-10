# Contract-OCR

Optimal character recognition app to perform legal contract parsing.

Created to convert previously scanned (e.g. non-machine readable) contracts, and extract up to 3 clauses of interest, which are then
packaged and outputted as an Excel spreadsheet. Has the ability to combine clauses that are separated by page breaks.

Uses tesseract (wrapper: pytesseract) to perform OCR, OpenCV and PyMuPDF to handle the file processing and conversion, and OpenPyXL 
to create a spreadsheet output.

Integrates a Tkinter GUI, packaged using PyInstaller in the final release, for usability and portability.

Built Kivy GUI as well, uisng kvlang, but Python 3.8 + PyInstaller + Kivy led to incompatibility issues.

Averaged a 10.5 second runtime per document, approx. 0.42 seconds per page.

UX:

1. Choose non-machine readable PDFs from any drive
2. Enter search queries (titles of clauses to be extracted)
3. Choose working directory where the file and image processing will
   take place, and also where a new Excel spreadsheet will be written
4. Enter the file name and sheet name of the new Excel spreadsheet.
5. Send it!!
6. Find desired output in the Excel file -- if the clause is not found in
   a certain PDF, the output cell will be blank
   
Note: Final release does not contain ContractApp.py or Contract.kv

Versions:

v1.0.0 (UR): Base functionality, handling one clause

v1.1.0 (UR): Changed GUI colors, widened buttons

v1.2.0 (UR): Added line-break case for company entity name extraction

v1.3.0: Integrated "lock-out" to disable buttons when parser is running

v2.0.0 (UR): Added three-clause ("greatest hits") functionality and ability to read clauses between page breaks
