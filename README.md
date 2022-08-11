# Contract-OCR

James Sy

Last Updated: August 11, 2022

Optimal character recognition app to perform legal contract parsing.

Created to convert previously scanned (e.g. non-machine readable) contracts, and extract up to 3 clauses of interest, which are then
packaged and outputted as an Excel spreadsheet. Because of the variability of contract language, can take an indefinite amount of 
"synonyms" for clause titles that point to the same information. Has the ability to combine clauses that are separated by page breaks.

Uses tesseract (wrapper: pytesseract) to perform OCR, OpenCV, Regexes, and PyMuPDF to handle the file processing and conversion, and OpenPyXL 
to create a spreadsheet output.

Integrates a Tkinter GUI, packaged using PyInstaller in the final release, for usability and portability.

Built Kivy GUI as well, uisng kvlang, but Python 3.8 + PyInstaller + Kivy led to incompatibility issues.

Averaged a 10.5 second runtime per document, approx. 0.42 seconds per page.

UX:

1. Choose non-machine readable PDFs from any drive
2. Enter search queries (titles of clauses to be extracted), as well as their "synonyms"
3. 
Note: Only one query is required, the other two are optional!

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

v2.1.0 (UR): Implemented "synonym-search", allowing user to enter multiple search keywords for the same information

v2.2.0: FINAL RELEASE; allowed for empty search queries and added checks for Windows-illegal strings
