# This class read a DOCX file and output it to HTML format.

# USAGE

## debug mode
```
$rt = new WordPHP(true);
```

## without debug
```
$rt = new WordPHP(false); or $rt = new WordPHP();
```

## Set output encoding
```
$rt = new WordPHP(false, OUTPUT_ENCODING);
```

## Read docx file and returns the html code
```
$text = $rt->readDocument(FILENAME);
```

# NOTE:
Images will be saved in the 'images' folder. This folder will be automatically created if it does not already exist.
