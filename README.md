# pyxmltools
A library for a collection of python scripts for xml viewing/manipulation

## [unzipdocx.py](/code/unzipdocx.py)
DOCX files are not true binaries - rather, they consist of a zip archive that contains an xml typsetting source of the file that M$Word renders (among other files). This simple script takes a docx file and unzips it into a new directory, revealing the source files and directory tree. 

The resulting file structure typically looks like this:

### Root files
* [Content_Types].xml
* _rels/.rels
* docProps/core.xml
* docProps/app.xml
* docProps/thumbnail.jpeg (binary)

### Main WordprocessingML files
* word/document.xml
* word/styles.xml
* word/settings.xml
* word/webSettings.xml
* word/fontTable.xml
* word/theme/theme1.xml
* word/numbering.xml

### Relationships
* word/_rels/document.xml.rels
* word/_rels/fontTable.xml.rels
* word/_rels/settings.xml.rels
* word/_rels/theme1.xml.rels

### Media (binary) (if pressent)
* word/media/image1.jpeg (if present — may or may not exist depending on python-docx behavior)
(In your example output, the only binary was thumbnail.jpeg.)

# [detectai.py](/code/detectai.py)
This python script automates flagging whether docx, xlsx, or pptx files should be reviewed for further evaluation as being AI-, machine-, or script-generated. It takes an M$Office file as an input and prints any detected telltales to the terminal.
