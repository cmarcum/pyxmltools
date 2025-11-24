# pyxmltools
A library for a collection of python scripts for xml viewing/manipulation

## unzipdocx.py
DOCX files are not true binaries - rather, they consist of a zip archive that contains an xml typsetting source of the file that M$Word renders (among other files). This simple script takes a docx file and unzips it into a new directory, revealing the source files and directory tree. 
