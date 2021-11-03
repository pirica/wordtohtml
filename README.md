# This is an improved class to read a DOCX file and output it in HTML format.

## DESCRIPTION

This improved DOCX to HTML class will recognise nearly all the formatting, themes, images etc. in the original Word DOCX document. The only significant exception is tabs as these are very difficult to replicate in HTLM due to its page width being very flexible. The resultant HTML should look very much like the original.

For update notes see below.

NOTE - will run on (at least) php 7.3.

NOTE:- It will not read the older 'DOC' Word format. In this case it will give an error message saying that it is a 'non zip file'.

FEATURES
 1. It will use the correct font and font size (assuming that a common font is being used).
 2. Text formating - Bold, Underlining, Italic, Strikethrough, Superscript and Subscript are all replicated, along with text alignment : Left, Centre, Right, Justified. Also indented and hanging text.
 3. It will display multi-level lists with the correct alpha-numeric numbering as per the original word document.
 4. It will now recognise paragraphs numbered with the Word paragraph numbering function as well as the list numbering function.
 5. It will display tables with merged cells correctly along with border and cell colours etc.
 6. Both footnote and endnote references are located in the correct place in the text. All the actual footnotes and endnotes are located at the end of the text. Links are provided to jump from a reference to actual note and then back again.
 7. The bookmarks in a 'Table of Contents' or similar provide a link to the correct section of the document as per the original Word document. A return link is also provided.
 8. As many Word documents are quite long a link is provided to jump back to the top of the document. Located at the middle of the right side of the screen.
 9. In the default mode, images are formatted and sized very similarly to the original DOCX word document, which is fine for desktop computers. However an option is provided to allow for external CSS formatting to be used instead (e.g. to allow for better display in mobile devices etc.). In this mode each image is given a unique CSS class name - 'Wimg1' for the first image. 'Wimg2' for the second image, etc. to enable formatting of each image as desired.
 10. In the default mode, tables are replicated as near as possible to the original DOCX word document formatting and size. However an option is provided for the tables to take up 100% of screen width (to allow for better display in mobile devices etc.).
 11. By default images are saved into the 'images' directory, which is automatically created if it does not exist. An option is provided to enable the name of this directory to be changed if desired.


# USAGE

## debug mode 
Will display the various zipped XML files which make up a DOCX file and also display the resultant HTML.
```
$rt = new WordPHP(true);
```

## without debug
```
$rt = new WordPHP(false); or $rt = new WordPHP();
```

## Set output encoding (Default is ISO-8859-1)
Will alter the encoding of the resultant HTML - eg. 'UTF-8', 'windows-1252', etc.
```
$rt = new WordPHP(false, 'desired encoding');
```

## Change directory for images (Default is 'images')
Will change the directory used for any images in the document.
```
$rt = new WordPHP(false, null, 'dir_name');
```

## Read docx file and return the html code - Default mode
```
$text = $rt->readDocument('FILENAME','NN');
```

## Read docx file and return the html code - Option 1
Images can be formatted using external CSS
```
$text = $rt->readDocument('FILENAME','YN');
```

## Read docx file and return the html code - Option 2
Tables will be formatted to 100% of screen width
```
$text = $rt->readDocument('FILENAME','NY');
```

## Read docx file and return the html code - Option 1 and 2
Images can be formatted using external CSS and Tables will be formatted to 100% of screen width
```
$text = $rt->readDocument('FILENAME','YY');
```

## UPDATE NOTES

Version 2.1.4 - Clearance of a minor bug

Version 2.1.3 - Clearance of a few bugs: 1. The endnote references in the text were not in roman numerals as they should be. - 2. In certain circumstances the correct font and font size were not being used. - 3. In certain circumstances a bookmark link did not work.

Version 2.1.2 - Will now recognise higher resolution images (Word deals with these differently to lower resolution images)

Version 2.1.1 - Will now recognise paragraph numbering as well as list numbering.

Version 2.1.0 - Improved method of displaying lists. Can now display footnotes and endnotes and their text references. The bookmarks in a 'Table of Contents' or similar, now link to the appropriate section of the document.
