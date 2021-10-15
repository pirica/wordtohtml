# This is an improved class to read a DOCX file and output it in HTML format.

## DESCRIPTION

This improved DOCX to HTML class will recognise nearly all the formatting, themes, images etc. in the original Word DOCX document. The only significant exception is tabs as these are very difficult to replicate in HTLM due to its page width being very flexible. The resultant HTML should look very much like the original.<br>
NOTE:- It will not read the older 'DOC' Word format. In this case it will give an error message saying that it is not a 'non zip file'.

<ul>
  <li>It will use the correct font and font size (assuming that it is not a very obsure font being used).</li>
  <li>Test formating - Bold, Underlining, Italic, Strikethrough, Superscript and Subscript are all replicated, along with text alignment : Left, Centre, Right, Justified. Also indented and hanging text.</li>
  <li>It will display multi-level lists with the correct alpha-numeric numbering as per original.</li>
  <li>It will display tables with merged cells correctly along with border and cell colours etc.</li>
  <li>In the default mode, images are formatted and sized very similarly to the original DOCX word document, which is fine for desktop computers. However an option is provided to allow for external CSS formatting to be used instead (e.g. to allow for better display in mobile devices etc.). In this mode each image is given a unique CSS class name - 'Wimg1' for the first image. 'Wimg2' for the second image, etc. to enable formatting of each image if desired.</li>
  <li>In the default mode, tables are replicated as near as possible to the original DOCX word document formatting and size. However an option is provided for the tables to take up 100% of screen width (to allow for better display in mobile devices etc.).</li>
  <li>By default images are saved into the 'images' directory, which is automatically created if it does not exist. An option is provided to enable the name of the directory to be changed if desired.</li>
</ul>

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
