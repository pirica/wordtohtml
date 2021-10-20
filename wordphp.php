<?php
class WordPHP
{
	private $debug = false;
	private $file;
	private $styles_xml;
	private $numb_xml;
	private $rels_xml;
	private $doc_xml;
	private $doc_media = [];
	private $last = 'none';
	private $encoding = 'ISO-8859-1';
	private $tmpDir = 'images';
	private $FSFactor = 22; //Font size conversion factor
	private $MTFactor = 13; //Margin and table width conversion factor
	
	/**
	 * CONSTRUCTOR
	 * 
	 * @param Boolean $debug Debug mode or not
	 * @param String $encoding selects alternative encoding if required
	 * @param String $tmpDir selects alternative image directory if required
	 * @return void
	 */
	public function __construct($debug_=null, $encoding=null, $tmpDir=null)
	{
		if($debug_ != null) {
			$this->debug = $debug_;
		}
		if ($encoding != null) {
			$this->encoding = $encoding;
		}
		if ($tmpDir != null) {
			$this->tmpDir = $tmpDir;
		}
	}


	/**
	 * READS The Document and Relationships into separated XML files
	 * 
	 * @param var $object The class variable to set as DOMDocument 
	 * @param var $xml The xml file
	 * @param string $encoding The encoding to be used
	 * @return void
	 */
	private function setXmlParts(&$object, $xml, $encoding)
	{
		$object = new DOMDocument();
		$object->encoding = $encoding;
		$object->preserveWhiteSpace = false;
		$object->formatOutput = true;
		$object->loadXML($xml);
		$object->saveXML();
	}



	/**
	 * READS The Document and Relationships into separated XML files
	 * 
	 * @param String $filename The filename
	 * @return void
	 */
	private function readZipPart($filename)
	{
		$zip = new ZipArchive();
		$_xml = 'word/document.xml';
		$_xml_rels = 'word/_rels/document.xml.rels';
		$_xml_numb = 'word/numbering.xml';
		$_xml_styles = 'word/styles.xml';
		$_xml_fonts = 'word/fontTable.xml';
		$_xml_theme = 'word/theme/theme1.xml';
		$_xml_settings = 'word/settings.xml';
		$_xml_foot = 'word/footnotes.xml';
		$_xml_end = 'word/endnotes.xml';
		
		if (true === $zip->open($filename)) {
			if (($index = $zip->locateName($_xml)) !== false) {
				$xml = $zip->getFromIndex($index);
			}
			//Get the relationships
			if (($index = $zip->locateName($_xml_rels)) !== false) {
				$xml_rels = $zip->getFromIndex($index);
			}
			// load all images if they exist
			for ($i=0; $i<$zip->numFiles;$i++) {
            	$zip_element = $zip->statIndex($i);
            	 if(preg_match("([^\s]+(\.(?i)(jpg|jpeg|png|gif|bmp))$)",$zip_element['name'])) {
            	 	$this->doc_media[$zip_element['name']] = $zip_element['name'];
            	 }
        	}
			//Get the list references from the word numbering file
			if (($index = $zip->locateName($_xml_numb)) !== false) {
				$xml_numb = $zip->getFromIndex($index);
			}
			//Get the style references from the word styles file
			if (($index = $zip->locateName($_xml_styles)) !== false) {
				$xml_styles = $zip->getFromIndex($index);
			}
			//Get the style references from the word fonts file
			if (($index = $zip->locateName($_xml_fonts)) !== false) {
				$xml_fonts = $zip->getFromIndex($index);
			}
			//Get the style references from the word themes file
			if (($index = $zip->locateName($_xml_theme)) !== false) {
				$xml_theme = $zip->getFromIndex($index);
			}
			if (($index = $zip->locateName($_xml_settings)) !== false) {
				$xml_settings = $zip->getFromIndex($index);
			}
			if (($index = $zip->locateName($_xml_foot)) !== false) {
				$xml_foot = $zip->getFromIndex($index);
			}
			if (($index = $zip->locateName($_xml_end)) !== false) {
				$xml_end = $zip->getFromIndex($index);
			}
			$zip->close();
		} else die('non zip file');

		$enc = mb_detect_encoding($xml);
		$this->setXmlParts($this->doc_xml, $xml, $enc);
		$this->setXmlParts($this->rels_xml, $xml_rels, $enc);
		$this->setXmlParts($this->numb_xml, $xml_numb, $enc);
		$this->setXmlParts($this->styles_xml, $xml_styles, $enc);
		$this->setXmlParts($this->fonts_xml, $xml_fonts, $enc);
		$this->setXmlParts($this->theme_xml, $xml_theme, $enc);
		$this->setXmlParts($this->settings_xml, $xml_settings, $enc);
		$this->setXmlParts($this->foot_xml, $xml_foot, $enc);
		$this->setXmlParts($this->end_xml, $xml_end, $enc);
		
		if($this->debug) {
			echo "XML File : word/document.xml<br>";
			echo "<textarea style='width:100%; height: 200px;'>";
			echo $this->doc_xml->saveXML();
			echo "</textarea>";
			echo "<br>XML File : word/_rels/document.xml.rels<br>";
			echo "<textarea style='width:100%; height: 200px;'>";
			echo $this->rels_xml->saveXML();
			echo "</textarea>";
			echo "<br>XML File : word/numbering.xml<br>";
			echo "<textarea style='width:100%; height: 200px;'>";
			echo $this->numb_xml->saveXML();
			echo "</textarea>";
			echo "<br>XML File : word/styles.xml<br>";
			echo "<textarea style='width:100%; height: 200px;'>";
			echo $this->styles_xml->saveXML();
			echo "</textarea>";
			echo "<br>XML File : word/fontTable.xml<br>";
			echo "<textarea style='width:100%; height: 200px;'>";
			echo $this->fonts_xml->saveXML();
			echo "</textarea>";
			echo "<br>XML File : word/theme/theme1.xml<br>";
			echo "<textarea style='width:100%; height: 200px;'>";
			echo $this->theme_xml->saveXML();
			echo "</textarea>";
			echo "<br>XML File : word/settings.xml<br>";
			echo "<textarea style='width:100%; height: 200px;'>";
			echo $this->settings_xml->saveXML();
			echo "</textarea>";
			echo "<br>XML File : word/footnotes.xml<br>";
			echo "<textarea style='width:100%; height: 200px;'>";
			echo $this->foot_xml->saveXML();
			echo "</textarea>";
			echo "<br>XML File : word/endnotes.xml<br>";
			echo "<textarea style='width:100%; height: 200px;'>";
			echo $this->end_xml->saveXML();
			echo "</textarea>";
		}
	}


	/**
	 * Looks up a font in the themes XML file and returns the various fonts
	 * 
	 * @param - None
	 * @returns Array - The major and minor font of the theme
	 */
	private function findfonts()
	{
		$reader1 = new XMLReader();
		$reader1->XML($this->theme_xml->saveXML());
		while ($reader1->read()) {
		// look for required style
			if ($reader1->nodeType == XMLREADER::ELEMENT && $reader1->name == 'a:majorFont') {
				$st1 = new XMLReader;
				$st1->xml(trim($reader1->readOuterXML()));
				while ($st1->read()) {
					if ($st1->name == 'a:latin') {
						$Mfont['major'] = $st1->getAttribute("typeface");
						if (substr($Mfont['major'],0,9) == 'Helvetica'){
							$Mfont['major'] = 'Helvetica';
						}

					}
				}
			}
			if ($reader1->nodeType == XMLREADER::ELEMENT && $reader1->name == 'a:minorFont') {
				$st2 = new XMLReader;
				$st2->xml(trim($reader1->readOuterXML()));
				while ($st2->read()) {
					if ($st2->name == 'a:latin') {
						$Mfont['minor'] = $st2->getAttribute("typeface");
						if (substr($Mfont['minor'],0,9) == 'Helvetica'){
							$Mfont['minor'] = 'Helvetica';
						}
					}
				}
			}
		}
		return $Mfont;
	}


	/**
	 * Looks up the footnotes XML file and returns the footnotes if any exist
	 * 
	 * @returns String - The footnote number and associated text
	 */
	private function footnotes()
	{
		$reader1 = new XMLReader();
		$reader1->XML($this->foot_xml->saveXML());
		$Ftext = array();
		$hyper = '';
		while ($reader1->read()) {
		// look for required style
			$znum = 1;
			if ($reader1->nodeType == XMLREADER::ELEMENT && $reader1->name == 'w:footnote') { //Get footnote
				$Footnum = $reader1->getAttribute("w:id");
				$st2 = new XMLReader;
				$st2->xml(trim($reader1->readOuterXML()));
				while ($st2->read()) {
					if ($st2->nodeType == XMLREADER::ELEMENT && $st2->name == 'w:hyperlink') {
						$hyper = 'Y';
					}
					if ($st2->nodeType == XMLREADER::ELEMENT && $st2->name == 'w:br') {
						$Ftext[$Footnum] .= "<br>";
					}
					if ($st2->name == 'w:t') {
						$tmp = htmlentities($st2->expand()->textContent);
						if ($hyper == 'Y'){
							$tmp = ltrim($tmp);
							$tmp = rtrim($tmp);
							$Ftext[$Footnum] .= "<a href='".$tmp."'>".$tmp."</a>";
							$hyper = '';
						} else {
							$Ftext[$Footnum] .= $tmp;
						}
					}
				}
					
			}
		}
		return $Ftext;
	}


	/**
	 * Looks up the footnotes XML file and returns the endnotes if any exist
	 * 
	 * @returns String - The endnote number and associated text
	 */
	private function endnotes()
	{
		$reader1 = new XMLReader();
		$reader1->XML($this->end_xml->saveXML());
		$Etext = array();
		$hyper = '';
		while ($reader1->read()) {
		// look for required style
			$znum = 1;
			if ($reader1->nodeType == XMLREADER::ELEMENT && $reader1->name == 'w:endnote') { //Get endnote
				$Endnum = $reader1->getAttribute("w:id");
				$st2 = new XMLReader;
				$st2->xml(trim($reader1->readOuterXML()));
				while ($st2->read()) {
					if ($st2->nodeType == XMLREADER::ELEMENT && $st2->name == 'w:hyperlink') {
						$hyper = 'Y';
					}
					if ($st2->name == 'w:t') {
						$tmp = htmlentities($st2->expand()->textContent);
						if ($hyper == 'Y'){
							$tmp = ltrim($tmp);
							$tmp = rtrim($tmp);
							$Etext[$Endnum] .= "<a href='".$tmp."'>".$tmp."</a>";
							$hyper = '';
						} else {
							$Etext[$Endnum] .= $tmp;
						}
					}
				}
					
			}
		}
		return $Etext;
	}


	/**
	 * Looks up a style in the styles XML file and returns the various style parameters
	 * 
	 * @param String $style - The name of the style
	 * @returns Array - The various parameters of the style
	 */
	private function findstyles($style)
	{
		$Rfont = $this->findfonts();
		$reader1 = new XMLReader();
		$reader1->XML($this->styles_xml->saveXML());
		$FontTheme = '';
		$Rstyle = array();
		while ($reader1->read()) {
		// look for required style
			$znum = 1;
			if ($reader1->nodeType == XMLREADER::ELEMENT && $reader1->name == 'w:docDefaults') { //Get document default settings
				$st2 = new XMLReader;
				$st2->xml(trim($reader1->readOuterXML()));
				while ($st2->read()) {
					if($st2->name == "w:spacing") { // Checks for paragraph spacing
						if ($st2->getAttribute("w:before") <>''){
							$Rstyle['DMtop'] =  "margin-top: ".round($st2->getAttribute("w:before")/$this->MTFactor)."px;";
						}
						if ($st2->getAttribute("w:after") <>''){
							$Rstyle['DMbot'] =  "margin-bottom: ".round($st2->getAttribute("w:after")/$this->MTFactor)."px;";
						}
					}
					if($st2->name == "w:sz") {
						$Rstyle['DFontS'] = round($st2->getAttribute("w:val")/$this->FSFactor,2);
					}
					if($st2->name == "w:color") {
						$Rstyle['DColor'] = $st2->getAttribute("w:val");
					}
				}
			}
			if ($reader1->nodeType == XMLREADER::ELEMENT && $reader1->name == 'w:style') { //Get style settings
				if ($reader1->getAttribute("w:styleId") == $style){
					$st1 = new XMLReader;
					$st1->xml(trim($reader1->readOuterXML()));
					while ($st1->read()) {

						if($st1->name == "w:rFonts" and $st1->getAttribute("w:ascii")) {
							$FF = $st1->getAttribute("w:ascii");
							if (substr($FF,0,9) == 'Helvetica'){
								$FF = 'Helvetica';
							}
							$Rstyle['Font'] = " font-family: ".$FF.";";
						}
						if($st1->name == "w:rFonts" and $st1->getAttribute("w:asciiTheme")) {
							$FontTheme = $st1->getAttribute("w:asciiTheme");
						}						
						if($st1->name == "w:sz") {
							$Rstyle['FontS'] = round($st1->getAttribute("w:val")/$this->FSFactor,2);
						}
						if($st1->name == "w:caps") {
							$Rstyle['Caps'] = " text-transform: uppercase;";
						}
						if($st1->name == "w:b") {
							$Rstyle['Bold'] = " font-weight: bold;";
						}
						if($st1->name == "w:u") {
							$Rstyle['Under'] = " text-decoration: underline;";
						}
						if($st1->name == "w:i") {
							$Rstyle['Italic'] = " font-style: italic;";
						}
						if($st1->name == "w:color") {
							$Rstyle['Color'] = $st1->getAttribute("w:val");
						}
						if($st1->name == "w:spacing") { // Checks for paragraph spacing
							if ($st1->getAttribute("w:before") <>''){
								$Rstyle['MPtop'] =  "margin-top: ".round($st1->getAttribute("w:before")/$this->MTFactor)."px;";
							}
							if ($st1->getAttribute("w:after") <>''){
								$Rstyle['MPbot'] =  "margin-bottom: ".round($st1->getAttribute("w:after")/$this->MTFactor)."px;";
							}
						}
						if($st1->name == "w:ind") { // Checks for paragragh indent
							if ($st1->getAttribute("w:left") <>''){
								$Rstyle['Ileft'] =  "padding-left: ".round($st1->getAttribute("w:left")/$this->MTFactor)."px;";
							}
							if ($st1->getAttribute("w:right") <>''){
								$Rstyle['Iright'] =  "padding-right: ".round($st1->getAttribute("w:right")/$this->MTFactor)."px;";
							}
							if ($st1->getAttribute("w:hanging") <>''){
								$Rstyle['Ihang'] =  "text-indent: -".round($st1->getAttribute("w:hanging")/$this->MTFactor)."px;";
							}
							if ($st1->getAttribute("w:firstLine") <>''){
								$Rstyle['Ifirst'] =  "text-indent: ".round($st1->getAttribute("w:firstLine")/$this->MTFactor)."px;";
							}
						}
						if($st1->name == "w:jc") { // Checks for paragragh alignment
							switch($st1->getAttribute("w:val")) {
								case "left":
									$Rstyle['Align'] =  "text-align: left;";
									break;
								case "center":
									$Rstyle['Align'] =  "text-align: center;";
									break;
								case "right":
									$Rstyle['Align'] =  "text-align: right;";
									break;
									case "both":
									$Rstyle['Align'] =  "text-align: justify;";
									break;
							}
						}
						if($st1->name == "w:shd" && $st1->getAttribute("w:fill") != "000000") {
							$Rstyle['Bcolor'] = $st1->getAttribute("w:fill");
						}
						if ($st1->nodeType == XMLREADER::ELEMENT && $st1->name === 'w:tblBorders') { //Get table border style
							$tc2 = new SimpleXMLElement($st1->readOuterXML());
							foreach ($tc2->children('w',true) as $ch) {
								if (in_array($ch->getName(), ['top','left','bottom','right','insideH','insideV']) ) {
									$line = $this->convertLine($ch['val']);
									if ($ch['color'] == 'auto'){
										$tbc = '000000';
									} else {
										$tbc = $ch['color'];
									}
									$zlinB = $ch['sz']/4;
									if ($zlinB >0 AND $zlinB <1){
										$zlinB = 1;
									} else{
										$zlinB = round($zlinB);
									}
									$Bname = "B".$ch->getName();
									$Rstyle[$Bname] = ":".$zlinB."px ".$line." #".$tbc.";";
								}
							}
						}
						if ($st1->nodeType == XMLREADER::ELEMENT && $st1->name === 'w:tblCellMar') { //Get table margin styles
							$tc3 = new SimpleXMLElement($st1->readOuterXML());
							foreach ($tc3->children('w',true) as $ch) {
								if (in_array($ch->getName(), ['top','left','bottom','right']) ) {
									$zlinM = round($ch['w']/$this->MTFactor);
									$Mname = "M".$ch->getName();
									$Rstyle[$Mname] = ":".$zlinM."px;";
								}
							}
						}
					}
				}
			}
		}
		if (! $Rstyle['Font']){
			if ($FontTheme AND $Rfont['major']){
				$Rstyle['Font'] = " font-family: ".$Rfont['major'].";";
			} else if (! $FontTheme AND $Rfont['minor']){
				$Rstyle['Font'] = " font-family: ".$Rfont['minor'].";";
			}
		}
		return $Rstyle;
	}
	
	

	/**
	 * CONVERTS A NUMBER TO ITS ROMAN PRESENTATION
	 * @param String/Integer $num - The number to be converted
	 * @return String - Roman number
	**/ 
	function numberToRoman($num)  
	{ 
		// Be sure to convert the given parameter into an integer
		$n = intval($num);
		$result = ''; 
 
		// Declare a lookup array that we will use to traverse the number: 
		$lookup = array(
			'm' => 1000, 'cm' => 900, 'd' => 500, 'cd' => 400, 
			'c' => 100, 'xc' => 90, 'l' => 50, 'xl' => 40, 
			'x' => 10, 'ix' => 9, 'v' => 5, 'iv' => 4, 'i' => 1
		); 
 
		foreach ($lookup as $roman => $value)  
		{
			// Look for number of matches
			$matches = intval($n / $value); 
 
			// Concatenate characters
			$result .= str_repeat($roman, $matches); 
 
			// Substract that from the number 
			$n = $n % $value; 
		} 

		return $result; 
	} 

	/**
	 * CHECKS THE FONT FORMATTING OF A GIVEN ELEMENT
	 * 
	 * @param XML $xml - The XML node
	 * @param String $Pstyle - The name of the paragraph style
	 * @param String $Tstyle - The name of the table style
	 * @param String $imgcss - The image and table option parameters
	 * @return Array - The elements styling and text
	 */
	private function checkFormating(&$xml,$Pstyle,$Tstyle,$imgcss)
	{	
		$Rstyle = $this->findstyles($Pstyle); // get the defined styles for this paragraph
		$node = trim($xml->readOuterXML());
		$t = '';
		// add <br> tags
		if (strstr($node,'<w:br ')) $t = '<br>';					 
		// look for formatting tags
		$f = "<span style='";
		$reader = new XMLReader();
		$reader->XML($node);
		$ret= array();
		$img = null;
		$Footref = '';
		$Ttmp = '';
		static $zimgcount = 1;
		$Lstyle = array();

		while ($reader->read()) {
			if ($reader->name === 'w:tab') {
				$Ttmp .= "&nbsp;&nbsp;&nbsp;";
			}

			if($reader->name == "w:b") {
				$Lstyle['Bold'] = "font-weight: bold;";
			}
			if($reader->name == "w:u") {
				$Lstyle['Under'] = "text-decoration: underline;";
			}
			if($reader->name == "w:i") {
				$Lstyle['Italic'] = " font-style: italic;";
			}
			if($reader->name == "w:color") {
				$Lstyle['Color'] = $reader->getAttribute("w:val");
			}
			if($reader->name == "w:rFonts") {
				 $FF = $reader->getAttribute("w:ascii");
				 if (substr($FF,0,9) == 'Helvetica'){
					 $FF = 'Helvetica';
				 }
				 $Lstyle['Font'] = "font-family: ".$FF.";";
			}
			if($reader->name == "w:shd" && $reader->getAttribute("w:val") != "clear") {
				$Lstyle['background-color'] = $reader->getAttribute("w:fill");
			}
			if($reader->name == "w:strike") {
				$f .=" text-decoration:line-through;";
			}
			if($reader->name == "w:vertAlign" && $reader->getAttribute("w:val") == "superscript") {
				$f .="position: relative; top: -0.6em;";
				$script = 'Y';
			}
			if($reader->name == "w:vertAlign" && $reader->getAttribute("w:val") == "subscript") {
				$f .="position: relative; bottom: -0.5em;";
				$script = 'Y';
			}
			if($reader->name == "w:sz") {
				$Lstyle['FontS'] = round($reader->getAttribute("w:val")/$this->FSFactor,2);
			}
			if($reader->name == "w:footnoteReference") {
				$Ftmp = $reader->getAttribute("w:id");
				$Footref = "<a id='FN".$Ftmp."R' href='#FN".$Ftmp."'>[".$Ftmp."]</a>";
				$f .="position: relative; top: -0.6em;font-weight: bold;";
				$script = 'Y';
			}
			if($reader->name == "w:endnoteReference") {
				$Ftmp = $reader->getAttribute("w:id");
				$Footref = "<a id='EN".$Ftmp."R' href='#EN".$Ftmp."'>[".$Ftmp."]</a>";
				$f .="position: relative; top: -0.6em;font-weight: bold;";
				$script = 'Y';
			}
			if($reader->name == 'w:drawing' && !empty($reader->readInnerXml())) {
				$r = $this->checkImageFormating($reader);
				if (substr($imgcss,0,1) == 'Y'){
					$img = $r !== null ? "<image class='Wimg".$zimgcount."' src='".$r['image']."' />" : null;
				} else {
					$img = $r !== null ? "<image src='".$r['image']."' ".$r['style']." />" : null;
				}
				$zimgcount++;
			}
			if($reader->name == "w:t") {
				$Ttmp .= htmlentities($reader->expand()->textContent);
			}
		}
		if ($Lstyle['Bold']){
			$f .= $Lstyle['Bold'];
		} else if($Rstyle['Bold']){
			$f .= $Rstyle['Bold'];
		}
		if ($Lstyle['Under']){
			$f .= $Lstyle['Under'];
		} else if($Rstyle['Under']){
			$f .= $Rstyle['Under'];
		}
		if ($Lstyle['Italic']){
			$f .= $Lstyle['Italic'];
		} else if($Rstyle['Italic']){
			$f .= $Rstyle['Italic'];
		}
		if ($Lstyle['Font']){
			$f .= $Lstyle['Font'];
		} else if ($Rstyle['Font']){
			$f .= $Rstyle['Font'];
		}
		if ($Lstyle['FontS']){
			$Fsize = $Lstyle['FontS'];
		} else if ($Rstyle['FontS']){
			$Fsize = $Rstyle['FontS'];
		} else if ($Rstyle['DFontS']){
			$Fsize = $Rstyle['DFontS'];
		}
		if ($script == 'Y'){
			$f .= "font-size: ".$Fsize * 0.65 ."rem;";
		} else {
			$f .= "font-size: ".$Fsize."rem;";
		}
		if ($Lstyle['Bcolor']){
			$Bcolor = $Lstyle['Bcolor'];
			$f .= "background-color: #".$Lstyle['Bcolor'].";";
		} else if($Rstyle['Bcolor']){
			$Bcolor = $Rstyle['Bcolor'];
			$f .= "background-color: #".$Rstyle['Bcolor'].";";
		}
		if ($Lstyle['Color']){
			$tcol = $Lstyle['Color'];
		} else if($Rstyle['Color']){
			$tcol = $Rstyle['Color'];
		} else if($Rstyle['DColor']){
			$tcol = $Rstyle['DColor'];
		}
		if ($tcol == 'auto'){
			if ($Bcolor){
				$red = hexdec(substr($Bcolor,0,2));
				$green = hexdec(substr($Bcolor,0,2));
				$blue = hexdec(substr($Bcolor,0,2));
				$color = (($red * 0.299) + ($green * 0.587) + ($blue * 0.114) > 186) ?  'FFFFFF' : '000000';
				$f .= "color: #".$color.";";
			} else {
				$f .= "color: #000000;";
			}
		} else if ($tcol){
			$f .= "color: #".$tcol.";";
		}
		if ($Rstyle['Caps']){
			$f .= $Rstyle['Caps'];
		}
		
		$f = rtrim($f, ',');
		$f .= "'>";
		if ($Footref <> ''){
			$t = $Footref;
		} else {
			$t .= ($img !== null ? $img : $Ttmp);
		}
		$ret['style'] = $f;
		$ret['text'] = $t;
		return $ret;
	}
	

	
	/**
	 * CHECKS THE ELEMENT FOR List ELEMENTS and their numbering and also PARAGRAGH FORMATTING
	 * 
	 * @param XML $xml - The XML node
	 * @param String $Tstyle - The name of the table style
	 * @return Array - Paragraph Styling and list details
	 */
	private function getListFormating(&$xml,$Tstyle)
	{
		$T1style = $T2style = array();
		static $Listcount = array();
		$T1style = $this->findstyles($Tstyle); // get the defined styles for this paragraph from the table style
		if ($Tstyle <> "TableNormal" AND $Tstyle <> ''){
			$T2style = $this->findstyles('TableNormal'); // get the basic styles for tables
		}
		$Dstyle = $this->findstyles(''); // get the default styles for the document
		$node = trim($xml->readOuterXML());

		$reader = new XMLReader();
		$reader->XML($node);
		$PSret= array();
		$LnumA = array();
		$Rstyle = array();
		$ListnumId = '';
		
		while ($reader->read()){
			if($reader->name == "w:pStyle" && $reader->hasAttributes ) {
				$Pstyle = $reader->getAttribute("w:val");
				$PSret['style'] = $Pstyle;
				$Rstyle = $this->findstyles($Pstyle); // get the defined styles for this paragraph
			}
			if($reader->name == "w:ilvl" && $reader->hasAttributes) { // List formating - list level
				$Listlevel = $reader->getAttribute("w:val");
			}
			if($reader->name == "w:numId" && $reader->hasAttributes) { // List formating - List cross reference
				$ListnumId = $reader->getAttribute("w:val");
			}

			if($reader->name == "w:rFonts") { // get font style for list numbering
				$font = " font-family: ".$reader->getAttribute("w:ascii").";";
			}
				
			if($reader->name == "w:sz") { // get font size for list numbering
				$Fsize = " font-size: ".round($reader->getAttribute("w:val")/$this->FSFactor,2)."rem;";
			}

			if($reader->name == "w:spacing") { // Checks for paragraph spacing
				if ($reader->getAttribute("w:before") <>''){
					$bmar =  "margin-top: ".round($reader->getAttribute("w:before")/$this->MTFactor)."px;";
				}
				if ($reader->getAttribute("w:after") <>''){
					$amar =  "margin-bottom: ".round($reader->getAttribute("w:after")/$this->MTFactor)."px;";
				}

			}
			if($reader->name == "w:ind") { // Checks for paragraph indent
				if ($reader->getAttribute("w:left") <>''){
					$aind =  "padding-left: ".round($reader->getAttribute("w:left")/$this->MTFactor)."px;";
				}
				if ($reader->getAttribute("w:right") <>''){
					$bind =  "padding-right: ".round($reader->getAttribute("w:right")/$this->MTFactor)."px;";
				}
				if ($reader->getAttribute("w:hanging") <>''){
					$cind =  "text-indent: -".round($reader->getAttribute("w:hanging")/$this->MTFactor)."px;";
				}
				if ($reader->getAttribute("w:firstLine") <>''){
					$dind =  "text-indent: ".round($reader->getAttribute("w:firstLine")/$this->MTFactor)."px;";
				}
			}
			if($reader->name == "w:jc") { // Checks for paragraph alignment
				switch($reader->getAttribute("w:val")) {
					case "left":
						$palign =  "text-align: left;";
						break;
					case "center":
						$palign =  "text-align: center;";
						break;
					case "right":
						$palign =  "text-align: right;";
						break;
					case "both":
						$palign =  "text-align: justify;";
						break;
				}
			} else if ($Rstyle['Align']){
				$palign =  $Rstyle['Align'];
			}
			if($reader->name == "w:pBdr") { // Add horizontal line
				$hr = "width:100%; height:1px; background: #000000";
			}

		}
		
		if ($ListnumId){
			// look for the List reference number of this element
			$reader1 = new XMLReader();
			$reader1->XML($this->numb_xml->saveXML());
			while ($reader1->read()) {
				if ($reader1->nodeType == XMLREADER::ELEMENT && $reader1->name == 'w:num' && $reader1->getAttribute("w:numId") == $ListnumId) {
					$st1 = new XMLReader;
					$st1->xml(trim($reader1->readOuterXML()));
					while ($st1->read()) {
						if($st1->name == 'w:abstractNumId') {
							$ListAbsNo = $st1->getAttribute("w:val");
						}
					}
				}
			}
			// look for the List details of this element
			$reader2 = new XMLReader();
			$reader2->XML($this->numb_xml->saveXML());
			while ($reader2->read()) {
				if ($reader2->nodeType == XMLREADER::ELEMENT && $reader2->name == 'w:abstractNum' && $reader2->getAttribute("w:abstractNumId") == $ListAbsNo) {
					$st2 = new XMLReader;
					$st2->xml(trim($reader2->readOuterXML()));
					while ($st2->read()) {
						if($st2->nodeType == XMLREADER::ELEMENT && $st2->name == 'w:lvl') {
							$Rlvl = $st2->getAttribute("w:ilvl");
						}
						if($st2->name == 'w:start') {
						$Rstart[$Rlvl] = $st2->getAttribute("w:val");
						}
						if($st2->name == 'w:numFmt') {
							$Rnumfmt[$Rlvl] = $st2->getAttribute("w:val");
						}
						if($st2->name == "w:ind") { // Gets the list hanging and indent
							if ($st2->getAttribute("w:left") <>''){
								$Rind[$Rlvl] =  "padding-left: ".round($st2->getAttribute("w:left")/$this->MTFactor)."px;";
							}
							if ($st2->getAttribute("w:hanging") <>''){
								$Rhang[$Rlvl] =  "text-indent: -".round($st2->getAttribute("w:hanging")/$this->MTFactor)."px;";
							}
						}
					}
				}
			}
		}
		
		if ($bmar == ''){ //set margin-top
			if ($Rstyle['MPtop']){
				$bmar =  $Rstyle['MPtop'];
			} else if ($T1style['MPtop']){
				$bmar =  $T1style['MPtop'];
			} else if ($T2style['MPtop']){ 
				$bmar =  ";".$T2style['MPtop'];
			} else if ($Dstyle['DMtop']){
				$bmar =  $Dstyle['DMtop'];
			} else {
				$bmar =  " margin-top:0px;";
			}	
		}
		if ($amar == ''){ // set margin-bottom
 			if ($Rstyle['MPbot']){
				$amar =  $Rstyle['MPbot'];
			} else if ($T1style['MPbot']){
				$amar =  $T1style['MPbot'];
			} else if ($T2style['MPbot']){
				$amar =  $T2style['MPbot'];
			} else if ($Dstyle['DMbot']){
				$amar =  $Dstyle['DMbot'];
			} else {
				$amar =  " margin-bottom:0px;";
			}
		}
		if ($T1style['Mleft']){ // set margin-left
			$cmar =  " margin-left".$T1style['Mleft'];
		} else if ($T2style['Mleft']){
			$cmar =  " margin-left".$T2style['Mleft'];
		}
		if ($T1style['Mright']){ //set margin-right
			$dmar =  " margin-right".$T1style['Mright'];
		} else if ($T2style['Mright']){
			$dmar =  " margin-right".$T2style['Mright'];
		}
		if ($Rstyle['Bcolor']){ // set text colour
			$bcol = " background-color:#".$Rstyle['Bcolor'].";";
		}
		if ($aind == ''){ // set left indent
			if ($Rind[$Listlevel]){
				$aind = $Rind[$Listlevel];
			} else if ($Rstyle['Ileft']){
				$aind =  $Rstyle['Ileft'];
			}
		}
		if ($bind == ''){ // set right indent
			if ($Rstyle['Iright']){
				$bind =  $Rstyle['Iright'];
			}
		}
		if ($cind == ''){ // set first line hanging indent
			if ($Rhang[$Listlevel]){
				$cind = $Rhang[$Listlevel];
			} else if ($Rstyle['Ihang']){
				$cind =  $Rstyle['Ihang'];
			}
		}
		if ($dind == ''){ // set first line indent
			if ($Rstyle['Ifirst']){
				$dind =  $Rstyle['Ifirst'];
			}
		}
		// return the paragraph styling
		$PSret['Pform'] = " style='".$bmar.$amar.$cmar.$dmar.$aind.$bind.$cind.$dind.$palign.$bcol.$hr."'";
		
		$alphabet = array( 'a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y'. 'z');
  
		if ($ListnumId){ // If the element is a list element get its number
			if ($Listcount[$ListnumId][$Listlevel] == ''){
				$Listcount[$ListnumId][$Listlevel] = $Rstart[$Listlevel];
				$Listcount[$ListnumId][$Listlevel + 1] = '';
			} else {
				$Listcount[$ListnumId][$Listlevel] = $Listcount[$ListnumId][$Listlevel] + 1;
				$Listcount[$ListnumId][$Listlevel + 1] = '';
			}
			$Lcount = 0;
			while ($Lcount <= $Listlevel){ // produce the list element number
				$LnumA[$Lcount] = $Listcount[$ListnumId][$Lcount];
				if ($Rnumfmt[$Lcount] == 'lowerLetter'){
					$LnumA[$Lcount] = $alphabet[$LnumA[$Lcount]-1].".";
				} else if ($Rnumfmt[$Lcount] == 'upperLetter'){
					$LnumA[$Lcount] = strtoupper($alphabet[$LnumA[$Lcount]-1].".");
				} else if ($Rnumfmt[$Lcount] == 'lowerRoman'){
					$LnumA[$Lcount] = $this->numberToRoman($LnumA[$Lcount]).".";
				} else if ($Rnumfmt[$Lcount] == 'upperRoman'){
					$LnumA[$Lcount] = strtoupper($this->numberToRoman($LnumA[$Lcount])).".";
				} else if ($Rnumfmt[$Lcount] == 'bullet'){
					$LnumA[$Lcount] = "•";
				} else {
					$LnumA[$Lcount] = $LnumA[$Lcount].".";
				}
				$Lnum .= $LnumA[$Lcount];
				$Lcount++;
			}
			$Lnum = $Lnum."&nbsp;&nbsp;&nbsp;";
		
		}
		
		$PSret['Lnum'] = $Lnum;  // return the element's list number
		$PSret['listnum'] = $ListnumId;
		return $PSret;

	}
	
	/**
	 * CHECKS IF THERE IS AN IMAGE PRESENT
	 * 
	 * @param XML $xml - The XML node
	 * @return String The location of the image
	 */
	private function checkImageFormating(&$xml)
	{
		$content = trim($xml->readInnerXml());

		if (!empty($content)) {

			$relId;
			$notfound = true;
			$reader = new XMLReader();
			$reader->XML($content);
			static $Icount = 1;
			$Inline = '';
			
			while ($reader->read() && $notfound) {
				if ($reader->name == "wp:inline") { // Checks if image is 'Inline'
					$Inline = 'Y';
				}
				if ($reader->name == "wp:posOffset" AND $offset == '') {
					$offset = (int)$xml->expand()->textContent;  // Checks if text flows round image and if so which side of the page to float the image
				}
				if ($reader->name == "wp:extent") { // Get image size
					$ImgW = round($reader->getAttribute("cx")/9000);
					$ImgH = round($reader->getAttribute("cy")/9000);
				}
				if ($reader->name == "a:blip") { // Get image name
					$relId = $reader->getAttribute("r:embed");
					$notfound = false;
				}
			}
			if ($Inline == ''){
				$Imgpos = ($offset < 10000) ? "float:left;" : "float:right;";
			}
			$image['style'] = "style='".$Imgpos."width:".$ImgW."px; height:".$ImgH."px; padding:10px 15px 10px 15px;'";

			// image id found, get the image location
			if (!$notfound && $relId) {
				$reader = new XMLReader();
				$reader->XML($this->rels_xml->saveXML());
				
				while ($reader->read()) {
					if ($reader->nodeType == XMLREADER::ELEMENT && $reader->name=='Relationship') {
						if($reader->getAttribute("Id") == $relId) {
							$link = "word/".$reader->getAttribute('Target');
							break;
						}
					}
				}

    			$zip = new ZipArchive();
    			$im = null;
    			if (true === $zip->open($this->file)) {
        			$image['image'] = $this->createImage($zip->getFromName($link), $relId, $link);
    			}
    			$zip->close();
    			return $image;
			}
		}

		return null;
	}

	/**
	 * Creates an image in the filesystem
	 *  
	 * @param objetc $image - The image object
	 * @param string $relId - The image relationship Id
	 * @param string $name - The image name
	 * @return Array - With image tag definition
	 */
	private function createImage($image, $relId, $name)
	{
		$arr = explode('.', $name);
		$l = count($arr);
		$ext = strtolower($arr[$l-1]);
		
		if (!is_dir($this->tmpDir)){
			mkdir($this->tmpDir, 0755, true);
		}

		$im = imagecreatefromstring($image);
		$fname = $this->tmpDir.'/'.$relId.'.'.$ext;

		switch ($ext) {
			case 'png':
				imagepng($im, $fname);
				break;
			case 'bmp':
				imagebmp($im, $fname);
				break;
			case 'gif':
				imagegif($im, $fname);
				break;
			case 'jpeg':
			case 'jpg':
				imagejpeg($im, $fname);
				break;
			default:
				return null;
		}
		imagedestroy($im);
		return $fname;
	}

	/**
	 * CHECKS IF ELEMENT IS AN HYPERLINK
	 *  
	 * @param XML $xml - The XML node
	 * @return Array - Hyperlink open and closing tag definition
	 */
	private function getHyperlink(&$xml)
	{
		$ret = array('open'=>'<ul>','close'=>'</ul>');
		$link ='';
		if($xml->hasAttributes) {
			$attribute = "";
			while($xml->moveToNextAttribute()) {
				if($xml->name == "r:id"){  // check for external hyperlinks
					$attribute = $xml->value;
				}
				if($xml->name == "w:anchor"){  // check for internal bookmark links
					$internal = substr($xml->value,1);
				}
			}
			
			if($attribute != "") {
				$reader = new XMLReader();
				$reader->XML($this->rels_xml->saveXML());
				
				while ($reader->read()) {
					if ($reader->nodeType == XMLREADER::ELEMENT && $reader->name=='Relationship') {
						if($reader->getAttribute("Id") == $attribute) {
							$link = $reader->getAttribute('Target');
							break;
						}
					}
				}
			}
		}
		
		if($link != "") { // external hyperlinks
			$ret['open'] = "<a href='".$link."' target='_blank'>";
			$ret['close'] = "</a>";
		} else { // internal bookmark links
			$ret['open'] = "<a id='R".$internal."' style='text-decoration:none;' href='#".$internal."'>";
			$ret['close'] = "</a>";
		}
		
		return $ret;
	}




	/**
	 * PROCESS PARAGRAPH CONTENT
	 *  
	 * @param XML $xml - The XML node
	 * @param String $Tstyle - The name of the table style
	 * @param Array $ListDet - An array with the details of any lists in the document
	 * @param String $imgcss - The image and table option parameters
	 * @return String - The HTML code of the paragraph
	 */
	private function getParagraph(&$paragraph,$Tstyle,$ListDet,$imgcss)
	{
		$text = '';
		$list_format=array();
		$zst = array();
		$zzz = $text;
		$zstc = 1;
		$Pformat = 'N';

		$Dstyle = $this->getListFormating($paragraph,$Tstyle); //default styles for the document
		// loop through paragraph dom
		while ($paragraph->read()) {
			// look for elements
			if ($paragraph->nodeType == XMLREADER::ELEMENT && $paragraph->name === 'w:r') {
				if ($Pformat == 'Y'){
					if ($list_format['Pform'] <> ''){
						$text = "<p".$BookMk.$list_format['Pform'].">"; // brings in paragraph formatting
					} else {
						$text .= "<p".$BookMk.$Dstyle['Pform'].">";
					}
					if($list_format['listnum']){
						$Pelement = $this->checkFormating($paragraph,$list_format['style'],$Tstyle,$imgcss); // Get inline style for list numbering
						$text .= $Pelement['style'].$list_format['Lnum']."</style>";
					}
					$Pformat = 'D';
				}
				if ($Pformat == 'N'){
					$text .= "<p".$BookMk.$Dstyle['Pform'].">";
					$Pformat = 'D';
				}
				$Pelement = $this->checkFormating($paragraph,$list_format['style'],$Tstyle,$imgcss); // Get inline style and associated text				
				$zst[$zstc] = $Pelement['style'];
				if ($zst[$zstc] != $zst[$zstc-1]){
					if ($zstc > 1){
						$text .= "</span>".$Pelement['style'];
					} else {
						$text .= $Pelement['style'];
					}
					$zstc++;
				}
				$text .= $Pelement['text']; 
//				$paragraph->next();
			} else if($paragraph->nodeType == XMLREADER::ELEMENT && $paragraph->name === 'w:pPr') { // Get list and paragraph formatting
				$list_format = $this->getListFormating($paragraph,$Tstyle);
				$Pformat = 'Y';
//				$paragraph->next();
			} 
			else if($paragraph->name == "w:bookmarkStart") { // check for internal bookmark link and its return
				$BM = $paragraph->getAttribute("w:name");
				$BookL = substr($BM,1);
				if ($BM  <> '_GoBack'){
					$BookMk = " id='".$BookL."'";
					$BookRet = "&nbsp;<a href='#R".$BookL."'><sup>[return]</a>";
				}
			}
			else if ($paragraph->nodeType == XMLREADER::ELEMENT && $paragraph->name === 'w:hyperlink') {
				if ($Pformat == 'Y'){ // Add in paragraph formatting for Bookmark links 
					if ($list_format['Pform'] <> ''){
						$text = "<p".$list_format['Pform'].">"; // brings in paragraph formatting
					} else {
						$text .= "<p".$Dstyle['Pform'].">";
					}
					if($list_format['listnum']){
						$Pelement = $this->checkFormating($paragraph,$list_format['style'],$Tstyle,$imgcss); // Get inline style for list numbering
						$text .= $Pelement['style'].$list_format['Lnum']."</style>";
					}
					$Pformat = 'D';
				}
				if ($Pformat == 'N'){
					$text .= "<p".$Dstyle['Pform'].">";
					$Pformat = 'D';
				}
				$hyperlink = $this->getHyperlink($paragraph); // Add in hyperlinks and bookmarks
				$text .= $hyperlink['open'];
				$Pelement2 = $this->checkFormating($paragraph,$Pstyle,$Tstyle,$imgcss);
				$zst[$zstc] = $Pelement2['style'];
				if ($zst[$zstc] != $zst[$zstc-1]){
					if ($zstc > 1){
						$text .= "</span>".$Pelement2['style'];
					} else {
						$text .= $Pelement2['style'];
					}
					$zstc++;
				}
				$text .= $Pelement2['text'];
				$text .= $hyperlink['close'];
				$paragraph->next();
			}
		}
		if ($zzz == $text){
			$text .= "<p".$Dstyle['Pform'].">&nbsp;</p> ";
		} else {
			$text .= "&nbsp;</span>".$BookRet."</p> ";
		}
		return $text;
	}
			

// ------------------- END OF PARAGRAPH PROCESSING ------------------------

// ------------------- START OF TABLE PROCESSING ------------------------


	/**
	 * FIND NUMBER OF ROWS IN THE TABLE
	 *  
	 * @param XML $content - The XML node
	 * @param Array $ListDet - An array with the details of any lists in the document
	 * @param String $imgcss - The image and table option parameters
	 * @return Array - The number of rows in the table and also details of which cells are part of a vertical merge
	 */
	private function getrows($content)
	{
		$ztext = new XMLReader;
		$ztext->xml($content);
		$Trow = 0;
		while ($ztext->read()) {
			if ($ztext->nodeType == XMLREADER::ELEMENT && $ztext->name === 'w:tr') { //find rows in the table
				$Trow++;
				$Tcol = 0;
			}
			if ($ztext->nodeType == XMLREADER::ELEMENT && $ztext->name === 'w:tc') { //find cells in the table
				$Tcol++;
				$cell[$Trow][$Tcol] = 1;
			}
			if ($ztext->name === 'w:vMerge') { //find merged cells in the table
				if ($ztext->getAttribute("w:val") == 'restart'){
					$mrow[$Tcol] = $Trow;
				} else {
					$cell[$mrow[$Tcol]][$Tcol]++;
					$cell[$Trow][$Tcol] = 0;
				}
			}
		}
		$ret['rows'] = $Trow;
		$ret['merge'] = $cell;

		return $ret;
	}



	/**
	 * PROCESS TABLE CONTENT
	 *  
	 * @param XML $xml - The XML node
	 * @param Array $ListDet - An array with the details of any lists in the document
	 * @param String $imgcss - The image and table option parameters
	 * @return String - The HTML code of the table
	 */
	private function checkTableFormating(&$xml,$ListDet,$imgcss)
	{

		if (substr($imgcss,1,1) == 'Y'){
			$table = "<table style='border-collapse:collapse; width:100%;'><tbody>";
		}  else {
			$table = "<table style='border-collapse:collapse;'><tbody>";
		}
		$Tstile = array();
		$TCcount = 0;
		$Twidth = 0;
		$Trow = 1;
		$T1style = $T2style = array();
		while ($xml->read()) {
			if ($xml->nodeType == XMLREADER::ELEMENT && $xml->name === 'w:tbl') { //Get number of rows in the table
				$Tinfo = $this->getrows(trim($xml->readOuterXML()));
				$Trows = $Tinfo['rows'];
				$Tmerge = $Tinfo['merge'];
			}
			if ($xml->name === 'w:tblStyle') { //Get table style
				$Tstyle = $xml->getAttribute("w:val");
				$T1style = $this->findstyles($Tstyle); // get the defined styles for this table
				if ($Tstyle <> "TableNormal"){
					$T2style = $this->findstyles('TableNormal'); // get the basic styles for tables
				}
			}

			if ($xml->nodeType == XMLREADER::ELEMENT && $xml->name === 'w:tblBorders') { //Get table default borders
				$tc2 = new SimpleXMLElement($xml->readOuterXML());
				foreach ($tc2->children('w',true) as $ch) {
					if (in_array($ch->getName(), ['top','left','bottom','right','insideH','insideV']) ) {
						$line = $this->convertLine($ch['val']);
						if ($ch['color'] == 'auto'){
							$tbc = '000000';
						} else {
							$tbc = $ch['color'];
						}
						$zlinT = $ch['sz']/4;
						if ($zlinT >0 AND $zlinT <1){
							$zlinT = 1;
						}
						$Tname = $ch->getName();
						$Tstile[$Tname] = ":".$zlinT."px ".$line." #".$tbc.";";
					}
				}
			}
			
			
			if ($xml->nodeType == XMLREADER::ELEMENT && $xml->name === 'w:tblGrid') { //Get number of columns and their widths in the table
				$tr9 = new XMLReader;
				$tr9->xml(trim($xml->readOuterXML()));
				while ($tr9->read()) {
					if($tr9->name === 'w:gridCol'){
						$TCcount++;
						$Cwidth[$TCcount] = round($tr9->getAttribute("w:w")/$this->MTFactor); // column width
						$Twidth = $Twidth + $Cwidth[$TCcount]; // get width of the table
					}
				}
			}
			
			
			if ($xml->nodeType == XMLREADER::ELEMENT && $xml->name === 'w:tr') { //find and process a table row
				if (! $Tstile['top']){
					if ($T1style['Btop']){
						$Tstile['top'] = $T1style['Btop'];
					} else if ($T2style['Btop']){
						$Tstile['top'] = $T2style['Btop'];
					}
				}
				if (! $Tstile['left']){
					if ($T1style['Bleft']){
						$Tstile['left'] = $T1style['Bleft'];
					} else if ($T2style['Bleft']){
						$Tstile['left'] = $T2style['Bleft'];
					}
				}
				if (! $Tstile['bottom']){
					if ($T1style['Bbottom']){
						$Tstile['bottom'] = $T1style['Bbottom'];
					} else if ($T2style['Bbottom']){
						$Tstile['bottom'] = $T2style['Bbottom'];
					}
				}
				if (! $Tstile['right']){
					if ($T1style['Bright']){
						$Tstile['right'] = $T1style['Bright'];
					} else if ($T2style['Bright']){
						$Tstile['right'] = $T2style['Bright'];
					}
				}
				if (! $Tstile['insideH']){
					if ($T1style['BinsideH']){
						$Tstile['insideH'] = $T1style['BinsideH'];
					} else if ($T2style['BinsideH']){
						$Tstile['insideH'] = $T2style['BinsideH'];
					}
				}
				if (! $Tstile['insideV']){
					if ($T1style['BinsideV']){
						$Tstile['insideV'] = $T1style['BinsideV'];
					} else if ($T2style['BinsideV']){
						$Tstile['insideV'] = $T2style['BinsideV'];
					}
				}
				$tc = "";
				$ts = array();
				$table .= "<tr>";

				$tr = new XMLReader;
				$tr->xml(trim($xml->readOuterXML()));
				$Tcol = 1;
				$TCoffset = 0;
				while ($tr->read()) {
					$Cstyle = '';
					if ($tr->nodeType == XMLREADER::ELEMENT && $tr->name === 'w:tc') { //get cell borders and cell text and its formatting
						$tc = $this->processTableRow(trim($tr->readOuterXML()),$Tstyle,$ListDet,$imgcss);
					}
					$style = '';
					if ($tr->nodeType == XMLREADER::ELEMENT && $tr->name === 'w:tcPr') { //get cell border formatting
						$ts = $this->processTableStyle(trim($tr->readOuterXML()));
						if ($ts['left']){  // set left border of a table cell
							$style .= " border-left".$ts['left'];
						} else{
						$style .= ($Tcol == 1) ? " border-left".$Tstile['left'] : " border-left".$Tstile['insideV'];
						}
						if ($ts['top']){  // set the top border of a table cell
							$style .= " border-top".$ts['top'];
						} else{
							$style .= ($Trow == 1) ? " border-top".$Tstile['top'] : " border-top".$Tstile['insideH'];
						}
						if ($ts['bottom']){  // set the bottom border of a table cell
							$style .= " border-bottom".$ts['bottom'];
						} else{
							if ($Tmerge[$Trow][$Tcol] > 1){  // set the bottom border of a table cell is a vertical cell merge
								$style .= ($Trow + $Tmerge[$Trow][$Tcol] - 1 == $Trows) ? " border-bottom".$Tstile['bottom'] : " border-bottom".$Tstile['insideH'];
							} else {  // set the bottom border of a table cell is not a vertical cell merge
								$style .= ($Trow == $Trows) ? " border-bottom".$Tstile['bottom'] : " border-bottom".$Tstile['insideH'];
							}
						}
						if ($ts['colspan'] <> ''){  // if a table cell is a horizontal cell merge, determine number of additional cells that make up the merge
							$TCoffset = $TCoffset + $ts['colspan'] - 1;
						}
						if ($ts['right']){  // set the right border of the table cell
							$style .= " border-right".$ts['right'];
						} else{
							$style .= ($Tcol + $TCoffset == $TCcount) ? " border-right".$Tstile['right'] : " border-right".$Tstile['insideV'];
						}


						
						if ($Tmerge[$Trow][$Tcol] > 0){
							if (substr($imgcss,1,1) == 'Y'){
								if ($Tmerge[$Trow][$Tcol] == 1){
									if ($ts['colspan'] <> ''){
										$table .= "<td colspan='".$ts['colspan']."'; style='width:".floor($Cwidth[$Tcol] / $Twidth *95)."%; ".$style;
									} else {
										$table .= "<td style='width:".floor($Cwidth[$Tcol] / $Twidth *95)."%; ".$style;
									}
								} else {
									if ($ts['colspan'] <> ''){
										$table .= "<td rowspan='".$Tmerge[$Trow][$Tcol]."'; colspan='".$ts['colspan']."'; style='width:".floor($Cwidth[$Tcol] / $Twidth *95)."%; ".$style;
									} else {
										$table .= "<td rowspan='".$Tmerge[$Trow][$Tcol]."'; style='width:".floor($Cwidth[$Tcol] / $Twidth *95)."%; ".$style;
									}
								}
							} else {
								if ($Tmerge[$Trow][$Tcol] == 1){
									if ($ts['colspan'] <> ''){
										$table .= "<td colspan='".$ts['colspan']."'; style='width:".$Cwidth[$Tcol]."px; ".$style;
									} else {
										$table .= "<td style='width:".$Cwidth[$Tcol]."px; ".$style;
									}
								} else {
									if ($ts['colspan'] <> ''){
										$table .= "<td rowspan='".$Tmerge[$Trow][$Tcol]."'; colspan='".$ts['colspan']."'; style='width:".$Cwidth[$Tcol]."px; ".$style;
									} else {
										$table .= "<td rowspan='".$Tmerge[$Trow][$Tcol]."'; style='width:".$Cwidth[$Tcol]."px; ".$style;
									}
								}
							}
							$table .= $tc['align']."'>".$tc['cell']."</td>";
						}
						$Tcol++;
					}
				}
				$table .= "</tr>";
				$Trow++;
			}
		}

		$table .= "</tbody></table>";
		return $table;
	}






	/**
	 * PROCESS THE TABLE CELL STYLE
	 *  
	 * @param string $content - The XML node content
	 * @return - The HTML code of the styling of a table cell
	 */
	private function processTableStyle($content)
	{

		$tc = new XMLReader;
		$tc->xml($content);
		$style = array();

		while ($tc->read()) {
			if ($tc->name === "w:tcBorders") {
				$tc2 = new SimpleXMLElement($tc->readOuterXML());
				foreach ($tc2->children('w',true) as $ch) {
					if (in_array($ch->getName(), ['left','top','bottom','right']) ) {
						$line = $this->convertLine($ch['val']);
						if ($ch['color'] == 'auto'){
							$tbc = '000000';
						} else {
							$tbc = $ch['color'];
						}
						$zlinT = $ch['sz']/4;
						if ($zlinT >0 AND $zlinT <1){
							$zlinT = 1;
						}
						$Tname = $ch->getName();
						$style[$Tname] = ":".$zlinT."px ".$line." #".$tbc.";";
					}
				}
				$tc->next();
			}
			if ($tc->name === "w:gridSpan") {
				$style['colspan'] = $tc->getAttribute("w:val");
			}

		}
		return $style;
	}

	/**
	 * PROCESS the BORDER LINE STYLE
	 *  
	 * @param string $in - The XML border line style
	 * @return - The HTML code of the border styling
	 */
	private function convertLine($in)
	{
		if (in_array($in, ['dotted']))
			return "dashed";

		if (in_array($in, ['dotDash','dotdotDash','dotted','dashDotStroked','dashed','dashSmallGap']))
			return "dashed";
		
		if (in_array($in, ['double','triple','threeDEmboss','threeDEngrave','thick']))
			return "double";

		if (in_array($in, ['nil','none']))
			return "none";

		return "solid";
	}

	/**
	 * PROCESS THE TABLE ROW
	 *  
	 * @param string $content - The XML node content
	 * @param String $Tstyle - The name of the table style
	 * @param Array $ListDet - An array with the details of any lists in the document
	 * @param String $imgcss - The image and table option parameters
	 * @return Array - The HTML code of the table row
	 */
	private function processTableRow($content,$Tstyle,$ListDet,$imgcss)
	{
		$tc = new XMLReader;
		$tc->xml($content);
		$ct = array();
		$count = 0;
		$valign = '';
		while ($tc->read()) {
			$ztpp = '';
			$ztp = '';
			$text = '';
			if ($tc->nodeType == XMLREADER::ELEMENT && $tc->name === "w:p") {  // get cell text and its formatting
				$paragraph = new XMLReader;
				$p = $tc->readOuterXML();
				$paragraph->xml($p);
				$ct['cell'] .= $this->getParagraph($paragraph,$Tstyle,$ListDet,$imgcss);
			}
			if ($tc->name === "w:jc") {  // cell text horizontal alignment
				switch($tc->getAttribute("w:val")) {
					case "left":
						$halign =  " text-align: left;";
						break;
					case "center":
						$halign =  " text-align: center;";
						break;
					case "right":
						$halign =  " text-align: right;";
						break;
					case "both":
						$halign =  " text-align: justify;";
						break;
				}
			}
			if ($tc->name === "w:vAlign") {  // cell text vertical alignment
				switch($tc->getAttribute("w:val")) {
					case "top":
						$valign =  " vertical-align: top;";
						break;
					case "center":
						$valign =  " vertical-align: center;";
						break;
					case "bottom":
						$valign =  " vertical-align: bottom;";
						break;
				}
			}
			if ($tc->name === "w:shd" AND $tc->getAttribute("w:fill")) {  // cell background color
				$BackCol = $tc->getAttribute("w:fill");
				$red = hexdec(substr($BackCol,0,2));
				$green = hexdec(substr($BackCol,0,2));
				$blue = hexdec(substr($BackCol,0,2));
				$color = (($red * 0.299) + ($green * 0.587) + ($blue * 0.114) > 186) ?  'FFFFFF' : '000000';
				$colours = "background-color: #".$BackCol."; color: #".$color.";";
			}

			
		}

		if ($valign == ''){
			$valign =  " vertical-align: top;";
		}
		$ct['align'] = $halign.$valign.$colours;
		if ($ct['cell'] == ''){
			$ct['cell'] = "&nbsp;";			
		}

		return $ct;
	}
	
	
// ------------------- END OF TABLE PROCESSING ------------------------


	/**
	 * FIND NUMBER OF LISTS AND THE NUMBER OF LIST PARAGRAPHS IN EACH ONE
	 *  
	 * @param XML $content - The XML node
	 * @return the DETAILS OF THE LISTS
	 */
	private function checkListNumbers($content)
	{
		$List = array();
		while ($content->read()) {
			if ($content->name === 'w:numId') { //find list paragraphs
				$Tlist = $content->getAttribute("w:val");
				if ($List[$Tlist] == ''){
					$List[$Tlist] = 1;
				} else {
					$List[$Tlist] = $List[$Tlist] + 1;
				}
			}
		}
		return $List;
	}
			


	/**
	 * READS THE GIVEN DOCX FILE INTO HTML FORMAT
	 *  
	 * @param String $filename - The DOCX file name
	 * @param String $imgcss - The image and table option parameters
	 * @return String - With HTML code of the DOCX file
	 */
	public function readDocument($filename,$imgcss)
	{
		
		$this->file = $filename;
		$this->readZipPart($filename);
		$reader = new XMLReader();
		$reader->XML($this->doc_xml->saveXML());
		$text = "<div style='position:fixed; bottom:50vh; right:10px; border:2px solid black; padding:2px; min-width:3%; text-align:center; background-color:#eeeeee';><a href='#top'>Top</a></div>";
		$text .= "<div style='margin:10px;'>"; // Provide a small margin around the html output
		while ($reader->read()) {
		// look for new paragraphs or table
			$paragraph = new XMLReader;
			$p = $reader->readOuterXML();
			if ($reader->nodeType == XMLREADER::ELEMENT && $reader->name === 'w:body') { //check for any lists in the document
				$paragraph->xml($p);
				$ListDet = $this->checkListNumbers($paragraph);
			}
			if ($reader->nodeType == XMLREADER::ELEMENT && $reader->name === 'w:tbl') { // finds and gets tables
				$paragraph->xml($p);
				$text .= $this->checkTableFormating($paragraph,$ListDet,$imgcss);
				$reader->next();
			}
			else if ($reader->nodeType == XMLREADER::ELEMENT && $reader->name === 'w:p') {
				// finds and gets paragraphs			
				$paragraph->xml($p); // set up new instance of XMLReader for parsing paragraph independantly	
				$text .= $this->getParagraph($paragraph,'',$ListDet,$imgcss);
				$reader->next();
			}
		}
		$Foot = $this->footnotes(); // Get any Footnotes in the document
		if ($Foot[1]) {
			$text .= "<br>&nbsp;";
			$text .= "<hr><p style='margin-top:6px;margin-bottom:6px;'><b>FOOTNOTES</b></p>";
			$Fcount = 1;
			while ($Foot[$Fcount]){
				$text .= "<p style='padding-left:50px;text-indent:-50px;margin-top:6px;margin-bottom:6px;'><sup><a id='FN".$Fcount."' href='#FN".$Fcount."R'>[".$Fcount."]</a></sup>&nbsp;&nbsp;&nbsp;".$Foot[$Fcount]."</p>";
				++$Fcount;
			}
		}
		
		$Endn = $this->endnotes(); //Get any Endnotes in the document
		if ($Endn[1]) {
			$text .= "<br>&nbsp;";
			$text .= "<hr><p style='margin-top:6px;margin-bottom:6px;'><b>ENDNOTES</b></p>";
			$Fcount = 1;
			while ($Endn[$Fcount]){
				$text .= "<p style='padding-left:50px;text-indent:-50px; margin-top:6px;margin-bottom:6px;'><sup><a id='EN".$Fcount."' href='#EN".$Fcount."R'>[".$this->numberToRoman($Fcount)."]</a></sup>&nbsp;&nbsp;&nbsp;".$Endn[$Fcount]."</p>";
				++$Fcount;
			}
		}
		$text .= "<br>&nbsp;";
		
		$text .= "</div>";
		$reader->close();
		if($this->debug) {  // if in DEBUG mode, display the generated HTML text of the DOCX document
			echo "<div style='width:100%;'>";
			echo mb_convert_encoding($text, $this->encoding);
			echo "</div>";
		}
		return mb_convert_encoding($text, $this->encoding); // Out put the generated HTML text of the DOCX document
	}
}






