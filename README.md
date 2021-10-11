# This is an improved class to read a DOCX file and output it in HTML format.

## DESCRIPTION

This improved DOCS to HTML class will recognise nearly all the formatting, themes, images etc. in the original Word DOCX document. The only significant exception is tabs as these are very difficult to replicate in HTLM due to its page width being very flexible. The resultant HTML should look very much like the original.

<ul>
<li>It will use the correct font and font size (assuming that it is not a very obsure font being used).</li>
<li>Test formating - Bold, Underlining, Italic, Strikethrough, Superscript and Subscript are all replicated, along with text alignment : Left, Centre, Right, Justified.</li>
<li>It will display multi-level lists with the correct alpha-numeric numbering as per original.</li>
<li>It will display tables with merged cells correctly along with border and cell colours etc.</li>
<li>In the default mode, images are formatted and sized very similarly to the original word document, which is fine for desktop computers. However an option is provided to allow for external CSS formatting instead (to allow for better display in mobile devices etc.). In this mode each image is given a unique CSS class name - 'Wimg1' for the first image. 'Wimg2' for the second image, etc. to enable formatting of each image if desired.</li>
<li>In the default mode, tables are replicated as near as possible to the original word document formatting and size. However an option is provided where the tables take up 100% of screen width (to allow for better display in mobile devices etc.).</li>
</ul>

# USAGE

## debug mode 
Will display the various XML files which make up a DOCX file.
```
$rt = new WordPHP(true);
```

## without debug
```
$rt = new WordPHP(false); or $rt = new WordPHP();
```

## Set output encoding
Will set the encoding of the resultant HTML - eg. 'UTF-8', 'windows-1252', etc.
```
$rt = new WordPHP(false, 'desired encoding');
```

## Read docx file and returns the html code as $text - Default mode
```
$text = $rt->readDocument(FILENAME);
```

## Read docx file and returns the html code - Option 1
Images can be formatted using external CSS
```
$text = $rt->readDocument(FILENAME,YN);
```

## Read docx file and returns the html code - Option 2
Tables will be formatted to 100% of screen width
```
$text = $rt->readDocument(FILENAME,NY);
```

## Read docx file and returns the html code - Option 1 and 2
Images can be formatted using external CSS and Tables will be formatted to 100% of screen width
```
$text = $rt->readDocument(FILENAME,YY);
```

# NOTE:
Images will be saved in an 'images' folder. This folder will be automatically created if it does not already exist.
