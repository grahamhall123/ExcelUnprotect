# Excel Unprotect

A simple command line tool that can be used to unprotect xls, xlsx and xlsm documents and sheets.

How to use:

    ExcelUnprotect -f name-of-file.xlsx -o name-of-output.xlsx

It should give you a list of sheets that are protected and prompt to remove. If no output file is given it will create a new file with the same filename with _-unprotected_ appended to it.

*_Note_: It can't be used on spreadsheets that have been password protected to stop them from being opened.*

