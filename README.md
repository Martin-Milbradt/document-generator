# Document Generator

A simple tool that generates Word & PDF documents.
It does this by dynamically filling a Word template with values from an Excel spreadsheet. One document is created for each row in the sheet.

## Guide

- Duplicate & rename [sheet.example.xlsx](sheet.example.xlsx) and [template.example.docx](template.example.docx) to *sheet.xlsx* and *template.docx* respectively or run [document-generator.py](document-generator.py) once to do that automatically.
- Edit the sheet & template to your liking.
- Close sheet & template (files are locked otherwise).
- Run [document-generator.py](document-generator.py).
- The output files are in the newly created directories *pdf* and *word*.
