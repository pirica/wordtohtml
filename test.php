<?php
require_once('wordphp.php');
$rt = new WordPHP(false);
$text = $rt->readDocument('sample.docx','N');
echo $text;
// The following two lines are optional if required to save the html text to a file (newfile.php)
$myfile = fopen("newfile.php", "w") or die("Unable to open file!");
fwrite($myfile, $text);
?>
