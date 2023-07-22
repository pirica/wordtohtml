<?php
class WordPHP
{
	private $debug = false;
	private $file;
	private $rels_xml;
	private $doc_xml;
	private $Icss;
	private $Tcss;
	private $last = 'none';
	private $encoding = 'UTF-8';
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
	 * @param none
	 * @return void
	 */
	private function readZipPart()
	{
		$zip = new ZipArchive();
		$_xml = 'word/document.xml';
		$_xml_rels = 'word/_rels/document.xml.rels';
		
		if (true === $zip->open($this->file)) {
			//Get the main word document file
			if (($index = $zip->locateName($_xml)) !== false) {
				$xml = $zip->getFromIndex($index);
			}
			//Get the relationships
			if (($index = $zip->locateName($_xml_rels)) !== false) {
				$xml_rels = $zip->getFromIndex($index);
			}
			$zip->close();
		} else die('ERROR - non zip file');

		$enc = mb_detect_encoding($xml);
		$this->setXmlParts($this->doc_xml, $xml, $enc);
		$this->setXmlParts($this->rels_xml, $xml_rels, $enc);
		
		if($this->debug) {
			echo "XML File : word/document.xml<br>";
			echo "<textarea style='width:100%; height: 200px;'>";
			echo $this->doc_xml->saveXML();
			echo "</textarea>";
			echo "<br>XML File : word/_rels/document.xml.rels<br>";
			echo "<textarea style='width:100%; height: 200px;'>";
			echo $this->rels_xml->saveXML();
			echo "</textarea>";
		}
	}


	/**
	 * Looks up a font in the themes XML file and returns the various theme fonts
	 * 
	 * @param - None
	 * @returns Array - The major and minor font of the theme
	 */
	private function findfonts()
	{
		$zip = new ZipArchive();
		$_xml_theme = 'word/theme/theme1.xml';
		if (true === $zip->open($this->file)) {
			//Get the style references from the word themes file
			if (($index = $zip->locateName($_xml_theme)) !== false) {
				$xml_theme = $zip->getFromIndex($index);
			}
			$zip->close();
		}
		$enc = mb_detect_encoding($xml_theme);
		$this->setXmlParts($theme_xml, $xml_theme, $enc);
		if($this->debug) {
			echo "<br>XML File : word/theme/theme1.xml<br>";
			echo "<textarea style='width:100%; height: 200px;'>";
			echo $theme_xml->saveXML();
			echo "</textarea>";
		}

		$reader1 = new XMLReader();
		$reader1->XML($theme_xml->saveXML());
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
	 * @param - None
	 * @returns String - The footnote number and associated text
	 */
	private function footnotes()
	{
		$zip = new ZipArchive();
		$_xml_foot = 'word/footnotes.xml';
		if (true === $zip->open($this->file)) {
			//Get the footnotes from the word footnotes file file
			if (($index = $zip->locateName($_xml_foot)) !== false) {
				$xml_foot = $zip->getFromIndex($index);
			}
			$zip->close();
		}
		$Ftext = array();
		if (isset($xml_foot)){ // if the footnotes.xml file exists parse it
			$enc = mb_detect_encoding($_xml_foot);
			$this->setXmlParts($foot_xml, $xml_foot, $enc);
			if($this->debug) {
				echo "<br>XML File : word/footnotes.xml<br>";
				echo "<textarea style='width:100%; height: 200px;'>";
				echo $foot_xml->saveXML();
				echo "</textarea>";
			}
			$reader1 = new XMLReader();
			$reader1->XML($foot_xml->saveXML());
			while ($reader1->read()) {
			// look for required style
				$znum = 1;
				if ($reader1->nodeType == XMLREADER::ELEMENT && $reader1->name == 'w:footnote') { //Get footnote
					$zst = array();
					$zst[0] = '';
					$zstc = 1;
					$Footnum = $reader1->getAttribute("w:id");
					$Ftext[$Footnum] = '';
					$st2 = new XMLReader;
					$st2->xml(trim($reader1->readOuterXML()));
					while ($st2->read()) {
						if ($st2->nodeType == XMLREADER::ELEMENT && $st2->name == 'w:hyperlink') {
							$hyperlink = $this->getHyperlink($st2,'F'); // Add in hyperlinks in footnotes
							$Ftext[$Footnum] .= $hyperlink['open'];
							$Pelement2 = $this->checkFormating($st2,'','');
							$zst[$zstc] = $Pelement2['style'];
							if ($zst[$zstc] != $zst[$zstc-1]){
								if ($zstc > 1){
									$Ftext[$Footnum] .= "</span>".$Pelement2['style'];
								} else {
									$Ftext[$Footnum] .= $Pelement2['style'];
								}
								$zstc++;
							}
							$Ftext[$Footnum] .= $Pelement2['text'];
							$Ftext[$Footnum] .= $hyperlink['close'];
							$st2->next();
						}
						if ($st2->nodeType == XMLREADER::ELEMENT && $st2->name == 'w:br') {
							$Ftext[$Footnum] .= "<br>";
						}
						if ($st2->nodeType == XMLREADER::ELEMENT && $st2->name == 'w:r') {
			
							$Pelement = $this->checkFormating($st2,'',''); // Get inline style and associated text
							if (!isset($Pelement['style'])){
								$Pelement['style'] = '';
							}	
							$zst[$zstc] = $Pelement['style'];
							if ($zst[$zstc] != $zst[$zstc-1]){
								if ($zstc > 1){
									$Ftext[$Footnum] .= "</span>".$Pelement['style'];
								} else {
									$Ftext[$Footnum] .= $Pelement['style'];
								}
								$zstc++;
							}
							$Ftext[$Footnum] .= $Pelement['text'];
						}
					}
				}
			}
		}
		return $Ftext;
	}


	/**
	 * Looks up the endnotes XML file and returns the endnotes if any exist
	 * 
	 * @param - None
	 * @returns String - The endnote number and associated text
	 */
	private function endnotes()
	{
		$zip = new ZipArchive();
		$_xml_end = 'word/endnotes.xml';
		if (true === $zip->open($this->file)) {
			//Get the endnotes from the endnotes file
			if (($index = $zip->locateName($_xml_end)) !== false) {
				$xml_end = $zip->getFromIndex($index);
			}
			$zip->close();
		}
		$Etext = array();
		if (isset($xml_end)){ // if the endnotes.xml file exists parse it
			$enc = mb_detect_encoding($_xml_end);
			$this->setXmlParts($end_xml, $xml_end, $enc);
			if($this->debug) {
				echo "<br>XML File : word/endnotes.xml<br>";
				echo "<textarea style='width:100%; height: 200px;'>";
				echo $end_xml->saveXML();
				echo "</textarea>";
			}
		
			$reader1 = new XMLReader();
			$reader1->XML($end_xml->saveXML());
			while ($reader1->read()) {
			// look for required style
				$znum = 1;
				if ($reader1->nodeType == XMLREADER::ELEMENT && $reader1->name == 'w:endnote') { //Get endnote
					$zst = array();
					$zst[0] = '';
					$zstc = 1;
					$Endnum = $reader1->getAttribute("w:id");
					$Etext[$Endnum] = '';
					$st2 = new XMLReader;
					$st2->xml(trim($reader1->readOuterXML()));
					while ($st2->read()) {
						if ($st2->nodeType == XMLREADER::ELEMENT && $st2->name == 'w:hyperlink') {
							$hyperlink = $this->getHyperlink($st2,'E'); // Add in hyperlinks in endnotes
							$Etext[$Endnum] .= $hyperlink['open'];
							$Pelement2 = $this->checkFormating($st2,'','');
							$zst[$zstc] = $Pelement2['style'];
							if ($zst[$zstc] != $zst[$zstc-1]){
								if ($zstc > 1){
									$Etext[$Endnum] .= "</span>".$Pelement2['style'];
								} else {
									$Etext[$Endnum] .= $Pelement2['style'];
								}
								$zstc++;
							}
							$Etext[$Endnum] .= $Pelement2['text'];
							$Etext[$Endnum] .= $hyperlink['close'];
							$st2->next();
						}
						if ($st2->nodeType == XMLREADER::ELEMENT && $st2->name == 'w:br') {
							$Etext[$Endnum] .= "<br>";
						}
						if ($st2->nodeType == XMLREADER::ELEMENT && $st2->name == 'w:r') {
						
							$Pelement = $this->checkFormating($st2,'',''); // Get inline style and associated text				
							$zst[$zstc] = $Pelement['style'];
							if ($zst[$zstc] != $zst[$zstc-1]){
								if ($zstc > 1){
									$Etext[$Endnum] .= "</span>".$Pelement['style'];
								} else {
									$Etext[$Endnum] .= $Pelement['style'];
								}
								$zstc++;
							}
							$Etext[$Endnum] .= $Pelement['text'];
						}
					}
				}
			}
		}
		return $Etext;
	}


	/**
	 * Looks up the styles in the styles XML file and sets the parameters for all the styles
	 * 
	 * @param - None
	 * @return void
	 */
	private function findstyles()
	{
		$zip = new ZipArchive();
		$_xml_styles = 'word/styles.xml';
		if (true === $zip->open($this->file)) {
			//Get the style references from the word styles file
			if (($index = $zip->locateName($_xml_styles)) !== false) {
				$xml_styles = $zip->getFromIndex($index);
			}
			$zip->close();
		}

		$enc = mb_detect_encoding($xml_styles);
		$this->setXmlParts($styles_xml, $xml_styles, $enc);
		if($this->debug) {
			echo "<br>XML File : word/styles.xml<br>";
			echo "<textarea style='width:100%; height: 200px;'>";
			echo $styles_xml->saveXML();
			echo "</textarea>";
		}

		$Rfont = $this->findfonts();
		$reader1 = new XMLReader();
		$reader1->XML($styles_xml->saveXML());
		$FontTheme = '';
		$Rstyle = array();
		while ($reader1->read()) {
		// get all styles
			$znum = 1;
			if ($reader1->nodeType == XMLREADER::ELEMENT && $reader1->name == 'w:docDefaults') { //Get document default settings
				$st2 = new XMLReader;
				$st2->xml(trim($reader1->readOuterXML()));
				while ($st2->read()) {
					if($st2->name == "w:spacing") { // Checks for paragraph spacing
						if ($st2->getAttribute("w:before") <>''){
							$Rstyle['Default']['Mtop'] =  "margin-top: ".round($st2->getAttribute("w:before")/$this->MTFactor)."px;";
						}
						if ($st2->getAttribute("w:after") <>''){
							$Rstyle['Default']['Mbot'] =  "margin-bottom: ".round($st2->getAttribute("w:after")/$this->MTFactor)."px;";
						}
					}
					if($st2->name == "w:sz") { //Default font size
						$Rstyle['Default']['FontS'] = round($st2->getAttribute("w:val")/$this->FSFactor,2);
					}
					if($st2->name == "w:color") { // default font color
						$Rstyle['Default']['Color'] = $st2->getAttribute("w:val");
					}
					if($st2->name == "w:rFonts" and $st2->getAttribute("w:ascii")) { // Default font
						$DF = $st2->getAttribute("w:ascii");
						if (substr($DF,0,9) == 'Helvetica'){
							$DF = 'Helvetica';
						}
						$Rstyle['Default']['Font'] = " font-family: ".$DF.";";
					}
				}
			}
			if ($reader1->nodeType == XMLREADER::ELEMENT && $reader1->name == 'w:style') { //Get style settings
				$Fstyle = $reader1->getAttribute("w:styleId");
					$st1 = new XMLReader;
					$st1->xml(trim($reader1->readOuterXML()));
					while ($st1->read()) {

						if($st1->name == "w:rFonts" and $st1->getAttribute("w:ascii")) {
							$FF = $st1->getAttribute("w:ascii");
							if (substr($FF,0,9) == 'Helvetica'){
								$FF = 'Helvetica';
							}
							$Rstyle[$Fstyle]['Font'] = " font-family: ".$FF.";";
						}
						if($st1->name == "w:rFonts" and $st1->getAttribute("w:asciiTheme")) {
							$FontTheme = $st1->getAttribute("w:asciiTheme");
						}						
						if($st1->name == "w:sz") {
							$Rstyle[$Fstyle]['FontS'] = round($st1->getAttribute("w:val")/$this->FSFactor,2);
						}
						if($st1->name == "w:caps") {
							$Rstyle[$Fstyle]['Caps'] = " text-transform: uppercase;";
						}
						if($st1->name == "w:b") {
							$Rstyle[$Fstyle]['Bold'] = " font-weight: bold;";
						}
						if($st1->name == "w:u") {
							$Rstyle[$Fstyle]['Under'] = " text-decoration: underline;";
						}
						if($st1->name == "w:i") {
							$Rstyle[$Fstyle]['Italic'] = " font-style: italic;";
						}
						if($st1->name == "w:color") {
							$Rstyle[$Fstyle]['Color'] = $st1->getAttribute("w:val");
						}
						if($st1->name == "w:numId") { // Get paragraph numbering ref
							$Rstyle[$Fstyle]['parnum'] = $st1->getAttribute("w:val");
						}
						if($st1->name == "w:spacing") { // Checks for paragraph spacing
							if ($st1->getAttribute("w:before") <>''){
								$Rstyle[$Fstyle]['MPtop'] =  "margin-top: ".round($st1->getAttribute("w:before")/$this->MTFactor)."px;";
							}
							if ($st1->getAttribute("w:after") <>''){
								$Rstyle[$Fstyle]['MPbot'] =  "margin-bottom: ".round($st1->getAttribute("w:after")/$this->MTFactor)."px;";
							}
						}
						if($st1->name == "w:ind") { // Checks for paragragh indent
							if ($st1->getAttribute("w:left") <>''){
								$Rstyle[$Fstyle]['Ileft'] =  "padding-left: ".round($st1->getAttribute("w:left")/$this->MTFactor)."px;";
							}
							if ($st1->getAttribute("w:right") <>''){
								$Rstyle[$Fstyle]['Iright'] =  "padding-right: ".round($st1->getAttribute("w:right")/$this->MTFactor)."px;";
							}
							if ($st1->getAttribute("w:hanging") <>''){
								$Rstyle[$Fstyle]['Ihang'] =  "text-indent: -".round($st1->getAttribute("w:hanging")/$this->MTFactor)."px;";
							}
							if ($st1->getAttribute("w:firstLine") <>''){
								$Rstyle[$Fstyle]['Ifirst'] =  "text-indent: ".round($st1->getAttribute("w:firstLine")/$this->MTFactor)."px;";
							}
						}
						if($st1->name == "w:jc") { // Checks for paragragh alignment
							switch($st1->getAttribute("w:val")) {
								case "left":
									$Rstyle[$Fstyle]['Align'] =  "text-align: left;";
									break;
								case "center":
									$Rstyle[$Fstyle]['Align'] =  "text-align: center;";
									break;
								case "right":
									$Rstyle[$Fstyle]['Align'] =  "text-align: right;";
									break;
									case "both":
									$Rstyle[$Fstyle]['Align'] =  "text-align: justify;";
									break;
							}
						}
						if($st1->name == "w:shd" && $st1->getAttribute("w:fill") != "000000") { // get background colour
							$Rstyle[$Fstyle]['Bcolor'] = $st1->getAttribute("w:fill");
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
									$Rstyle[$Fstyle][$Bname] = ":".$zlinB."px ".$line." #".$tbc.";";
								}
							}
						}
						if ($st1->nodeType == XMLREADER::ELEMENT && $st1->name === 'w:tblCellMar') { //Get table margin styles
							$tc3 = new SimpleXMLElement($st1->readOuterXML());
							foreach ($tc3->children('w',true) as $ch) {
								if (in_array($ch->getName(), ['top','left','bottom','right']) ) {
									$zlinM = round($ch['w']/$this->MTFactor);
									$Mname = "M".$ch->getName();
									$Rstyle[$Fstyle][$Mname] = ":".$zlinM."px;";
								}
							}
						}
					}
					if ($FontTheme AND $Rfont['major']){ // get font defined by Font Theme or its absence
						$Rstyle[$Fstyle]['ThFont'] = " font-family: ".$Rfont['major'].";";
					} else if (! $FontTheme AND $Rfont['minor']){
						$Rstyle[$Fstyle]['ThFont'] = " font-family: ".$Rfont['minor'].";";
					}
					$FontTheme = '';
			}
		}
		$Rstyle['Theme']['minor'] = " font-family: ".$Rfont['minor'].";";
		$this->Rstyle = $Rstyle; // all the style parameters can now be accessed anywhere in the class
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
	 * @param String $Dcap - The type of drop capital if it exists
	 * @return Array - The elements styling and text
	 */
	private function checkFormating(&$xml,$Pstyle,$Dcap)
	{	
		$node = trim($xml->readOuterXML());
		$t = '';
		// add <br> tags
		if (strstr($node,'<w:br ')) $t = '<br>';					 
		// look for formatting tags
		$f = "<span style='";
		if ($Dcap == 'drop' OR $Dcap == 'margin'){
			$f .= "float:left; line-height: 80%; ";
		}
		$reader = new XMLReader();
		$reader->XML($node);
		$ret= array();
		$img = null;
		$Footref = '';
		$Ttmp = $script = $tcol = '';
		static $zimgcount = 1;
		$Lstyle = array();

		while ($reader->read()) {
			if ($reader->nodeType == XMLREADER::ELEMENT && $reader->name == 'w:instrText') {
				$Htext = htmlentities($reader->expand()->textContent);
				if (substr($Htext,0,5) == "HYPER"){
					$Htext = substr($Htext,16);
					$ret['Htext'] = substr($Htext,0,-6);
				}
				if (substr($Htext,0,4) == " REF"){
					$ret['CRtext'] = substr($Htext,6,11);
				}
			}
			if ($reader->name === 'w:rStyle' && $reader->getAttribute("w:val") == "Hyperlink") {
				$ret['Hyper'] = 'Y';
			}
			if ($reader->name === 'w:tab') {
				$Ttmp .= "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;";
			}
			if($reader->name == "w:br") { // Checks for page break
				if ($reader->getAttribute("w:type") == 'page'){
					$ret['Pbreak'] = 'Y';
				}
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
				 if ($FF == ''){
					 $FF = 'Times New Roman';
				 }
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
				$Footref = "<a id='EN".$Ftmp."R' href='#EN".$Ftmp."'>[".$this->numberToRoman($Ftmp)."]</a>";
				$f .="position: relative; top: -0.6em;font-weight: bold;";
				$script = 'Y';
			}
			if($reader->nodeType == XMLREADER::ELEMENT && $reader->name == 'w:drawing' ) { // Get a lower resolution image
				$r = $this->checkImageFormating($reader);
				if ($this->Icss == 'Y'){
					$img = $r !== null ? "<image class='Wimg".$zimgcount."' src='".$r['image']."' />" : null;
				} else {
					$img = $r !== null ? "<image src='".$r['image']."' ".$r['style']." />" : null;
				}
			}
			if($reader->nodeType == XMLREADER::ELEMENT && $reader->name == 'w:pict') { // Get a higher resolution image
				$r = $this->checkPictFormating($reader, '');
				if ($this->Icss == 'Y'){
					$img = $r !== null ? "<image class='Wimg".$zimgcount."' src='".$r['image']."' />" : null;
				} else {
					$img = $r !== null ? "<image src='".$r['image']."' ".$r['style']." />" : null;
				}
			} 
			if($reader->nodeType == XMLREADER::ELEMENT && $reader->name == 'v:shape' && $img == null) { // a fall back if the 'w:pict' does not get the higher resolution image
				$Psize = $reader->getAttribute("style");
				$r = $this->checkPictFormating($reader, $Psize);
				if ($this->Icss == 'Y'){
					$img = $r !== null ? "<image class='Wimg".$zimgcount."' src='".$r['image']."' />" : null;
				} else {
					$img = $r !== null ? "<image src='".$r['image']."' ".$r['style']." />" : null;
				}
			}
			if($reader->name == "w:t") {
				$Tmptext1 = htmlentities($reader->expand()->textContent);
				$Tmptext2 = preg_replace('~(?<=\s)\s~', '&nbsp;', $Tmptext1);
				$Ttmp .= $Tmptext2;
			}
		}
		if ($img !== null){
			$zimgcount++;
		}

		if (isset($Lstyle['Bold'])){
			$f .= $Lstyle['Bold'];
		} else if(isset($this->Rstyle[$Pstyle]['Bold'])){
			$f .= $this->Rstyle[$Pstyle]['Bold'];
		} else if(isset($this->Rstyle['Normal']['Bold'])){
			$f .= $this->Rstyle['Normal']['Bold'];
		}
		if (isset($Lstyle['Under'])){
			$f .= $Lstyle['Under'];
		} else if(isset($this->Rstyle[$Pstyle]['Under'])){
			$f .= $this->Rstyle[$Pstyle]['Under'];
		} else if(isset($this->Rstyle['Normal']['Under'])){
			$f .= $this->Rstyle['Normal']['Under'];
		}
		if (isset($Lstyle['Italic'])){
			$f .= $Lstyle['Italic'];
		} else if(isset($this->Rstyle[$Pstyle]['Italic'])){
			$f .= $this->Rstyle[$Pstyle]['Italic'];
		} else if(isset($this->Rstyle['Normal']['Italic'])){
			$f .= $this->Rstyle['Normal']['Italic'];
		}
		if (isset($Lstyle['Font'])){
			$f .= $Lstyle['Font'];
		} else if (isset($this->Rstyle[$Pstyle]['Font'])){
			$f .= $this->Rstyle[$Pstyle]['Font'];
		} else if (isset($this->Rstyle['Normal']['Font'])){
			$f .= $this->Rstyle['Normal']['Font'];
		} else if (isset($this->Rstyle[$Pstyle]['ThFont'])){
			$f .= $this->Rstyle[$Pstyle]['ThFont'];
		} else if (isset($this->Rstyle['Theme']['minor'])){
			$f .= $this->Rstyle['Theme']['minor'];
		} else if (isset($this->Rstyle['Default']['Font'])){
			$f .= $this->Rstyle['Default']['Font'];
		}else{
		}
		if (isset($Lstyle['FontS'])){
			$Fsize = $Lstyle['FontS'];
		} else if (isset($this->Rstyle[$Pstyle]['FontS'])){
			$Fsize = $this->Rstyle[$Pstyle]['FontS'];
		} else if (isset($this->Rstyle['Normal']['FontS'])){
			$Fsize = $this->Rstyle['Normal']['FontS'];
		} else if (isset($this->Rstyle['Default']['FontS'])){
			$Fsize = $this->Rstyle['Default']['FontS'];
		}
		if ($script == 'Y'){
			$f .= " font-size: ".$Fsize * 0.65 ."rem;";
		} else {
			$f .= " font-size: ".$Fsize."rem;";
		}
		if (isset($Lstyle['Bcolor'])){
			$Bcolor = $Lstyle['Bcolor'];
			$f .= " background-color: #".$Lstyle['Bcolor'].";";
		} else if(isset($this->Rstyle[$Pstyle]['Bcolor'])){
			$Bcolor = $this->Rstyle[$Pstyle]['Bcolor'];
			$f .= " background-color: #".$this->Rstyle[$Pstyle]['Bcolor'].";";
		}
		if (isset($Lstyle['Color'])){
			$tcol = $Lstyle['Color'];
		} else if(isset($this->Rstyle[$Pstyle]['Color'])){
			$tcol = $this->Rstyle[$Pstyle]['Color'];
		} else if(isset($this->Rstyle['Normal']['Color'])){
			$tcol = $this->Rstyle['Normal']['Color'];
		} else if(isset($this->Rstyle['Default']['Color'])){
			$tcol = $this->Rstyle['Default']['Color'];
		}
		if ($tcol == 'auto'){
			if (isset($Bcolor)){
				$red = hexdec(substr($Bcolor,0,2));
				$green = hexdec(substr($Bcolor,0,2));
				$blue = hexdec(substr($Bcolor,0,2));
				$color = (($red * 0.299) + ($green * 0.587) + ($blue * 0.114) > 186) ?  '000000' : 'FFFFFF';
				$f .= "color: #".$color.";";
			} else {
				$f .= "color: #000000;";
			}
		} else if ($tcol <> ''){
			$f .= "color: #".$tcol.";";
		}
		if (isset($this->Rstyle[$Pstyle]['Caps'])){
			$f .= $this->Rstyle[$Pstyle]['Caps'];
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
		
		$PSret= array();
		$DC = $Pstyle = $amar = $bmar = $cmar = $dmar = $PSret['Dcap'] = $aind = $bind = $cind = $dind = $Listlevel = $bcol = $hr = $Lnum = $palign = '';
		$node = trim($xml->readOuterXML());
		if ($node <>''){
			$reader = new XMLReader();
			$reader->XML($node);
			$LnumA = array();
			$ListnumId = '';
			static $Listcount = array();
			static $numb_xml = null;
		
			while ($reader->read()){
				if($reader->name == "w:framePr") { // get font style for list numbering
					$PSret['Dcap'] = $reader->getAttribute("w:dropCap");
				}
				if($reader->name == "w:pStyle" && $reader->hasAttributes ) {
					$Pstyle = $reader->getAttribute("w:val");
					$PSret['style'] = $Pstyle; //return the element's style
					if (substr($Pstyle,0,7) == 'numpara'){
						$ListnumId = $this->Rstyle[$Pstyle]['parnum'];
						$Listlevel = 0;
					}
				}
				if($reader->name == "w:ilvl" && $reader->hasAttributes) { // List formating - list level
					$Listlevel = $reader->getAttribute("w:val");
				}
				if($reader->name == "w:numId" && $reader->hasAttributes) { // List formating - List cross reference
					$ListnumId = $reader->getAttribute("w:val");
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
						$Thang = round($reader->getAttribute("w:hanging")/$this->MTFactor);
							if ($PSret['Dcap'] == 'margin'){
								$Thang = $Thang + 10;
							}
							$cind =  "text-indent: -".$Thang."px;";
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
				} else if (isset($this->Rstyle[$Pstyle]['Align'])){
					$palign =  $this->Rstyle[$Pstyle]['Align'];
				}
				if($reader->name == "w:pBdr") { // Add horizontal line
					$hr = "width:100%; height:1px; background: #000000";
				}
			}
		
			if ($ListnumId){
				if (!$numb_xml){
					$zip = new ZipArchive();
					$_xml_numb = 'word/numbering.xml';
		
					if (true === $zip->open($this->file)) {
						//Get the list references from the word numbering file
						if (($index = $zip->locateName($_xml_numb)) !== false) {
							$xml_numb = $zip->getFromIndex($index);
						}
						$zip->close();
					}

					$enc = mb_detect_encoding($xml_numb);
					$this->setXmlParts($numb_xml, $xml_numb, $enc);
		
					if($this->debug) {
						echo "<br>XML File : word/numbering.xml<br>";
						echo "<textarea style='width:100%; height: 200px;'>";
						echo $numb_xml->saveXML();
						echo "</textarea>";
					}
				}
			
			
			
				// look for the List reference number of this element
				$reader1 = new XMLReader();
				$reader1->XML($numb_xml->saveXML());
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
				$reader2->XML($numb_xml->saveXML());
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
							if($st2->name == 'w:lvlText') {
								$Rlvltxt[$Rlvl] = $st2->getAttribute("w:val");
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
		
			$alphabet = array( 'a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y'. 'z');
		
  
			if ($ListnumId){ // If the element is a list element get its number
				if (substr($Rlvltxt[$Listlevel],0,1) <> '%'){ // Check if there is a character before the list number and if so get it
					$LNfirst = substr($Rlvltxt[$Listlevel],0,1);
				} else {
					$LNfirst = '';
				}
				$LNlast = substr($Rlvltxt[$Listlevel],-1); // The last character of a list number
				if (!isset($Listcount[$ListnumId][$Listlevel])){
					$Listcount[$ListnumId][$Listlevel] = '';
				}
				if ($Listcount[$ListnumId][$Listlevel] == ''){ // Get the list number of the list element
					$Listcount[$ListnumId][$Listlevel] = $Rstart[$Listlevel];
					$Listcount[$ListnumId][$Listlevel + 1] = '';
				} else {
					$Listcount[$ListnumId][$Listlevel] = $Listcount[$ListnumId][$Listlevel] + 1;
					$Listcount[$ListnumId][$Listlevel + 1] = '';
				}
				if (strlen($Rlvltxt[$Listlevel]) > 4){
					$Lcount = 0;
				} else {
					$Lcount = $Listlevel;
				}
				while ($Lcount <= $Listlevel){ // produce the list element number
					$LnumA[$Lcount] = $Listcount[$ListnumId][$Lcount]; // The number of the list element
					if ($Rnumfmt[$Lcount] == 'lowerLetter'){
						$LnumA[$Lcount] = $LNfirst.$alphabet[$LnumA[$Lcount]-1].$LNlast;
					} else if ($Rnumfmt[$Lcount] == 'upperLetter'){
						$LnumA[$Lcount] = $LNfirst.strtoupper($alphabet[$LnumA[$Lcount]-1].$LNlast);
					} else if ($Rnumfmt[$Lcount] == 'lowerRoman'){
						$LnumA[$Lcount] = $LNfirst.$this->numberToRoman($LnumA[$Lcount]).$LNlast;
					} else if ($Rnumfmt[$Lcount] == 'upperRoman'){
						$LnumA[$Lcount] = $LNfirst.strtoupper($this->numberToRoman($LnumA[$Lcount])).$LNlast;
					} else if ($Rnumfmt[$Lcount] == 'bullet'){
						if (mb_ord($Rlvltxt[$Lcount]) == 61692){
							$LnumA[$Lcount] = "✓";
						} else if (mb_ord($Rlvltxt[$Lcount]) == 61623){
							$LnumA[$Lcount] = "•";
						} else if (mb_ord($Rlvltxt[$Lcount]) == 61607){
							$LnumA[$Lcount] = "◾";
						} else if (mb_ord($Rlvltxt[$Lcount]) == 61656){
							$LnumA[$Lcount] = "➢";
						} else if (mb_ord($Rlvltxt[$Lcount]) == 111){
							$LnumA[$Lcount] = $Rlvltxt[$Lcount];
						} else {
							$LnumA[$Lcount] = "◾";
						}
					} else {
						$LnumA[$Lcount] = $LNfirst.$LnumA[$Lcount].$LNlast;
					}
					$Lnum .= $LnumA[$Lcount];
					$Lcount++;
				}
				$Lnum = $Lnum."&nbsp;&nbsp;&nbsp;";
		
			}
		
			$PSret['Lnum'] = $Lnum;  // return the element's list number
			$PSret['listnum'] = $ListnumId;
		}
			if ($bmar == ''){ //set margin-top
				if (isset($this->Rstyle['Row']['MRtop'])){
					$bmar =  " margin-top".$this->Rstyle['Row']['MRtop'];
				} else if (isset($this->Rstyle[$Pstyle]['MPtop'])){
					$bmar =  $this->Rstyle[$Pstyle]['MPtop'];
				} else if (isset($this->Rstyle[$Tstyle]['MPtop'])){
					$bmar =  $this->Rstyle[$Tstyle]['MPtop'];
				} else if (isset($this->Rstyle['TableNormal']['MPtop'])){ 
					$bmar =  ";".$this->Rstyle['TableNormal']['MPtop'];
				} else if (isset($this->Rstyle['Default']['Mtop'])){
					$bmar =  " margin-top".$this->Rstyle['Default']['Mtop'];
				} else {
					$bmar =  " margin-top:0px;";
				}	
			}
			if ($amar == ''){ // set margin-bottom
				if (isset($this->Rstyle['Row']['MRbot'])){
					$bmar =  " margin-bottom".$this->Rstyle['Row']['MRbot'];
				} elseif (isset($this->Rstyle[$Pstyle]['MPbot'])){
					$amar =  $this->Rstyle[$Pstyle]['MPbot'];
				} else if (isset($this->Rstyle[$Tstyle]['MPbot'])){
					$amar =  $this->Rstyle[$Tstyle]['MPbot'];
				} else if (isset($this->Rstyle['TableNormal']['MPbot'])){
					$amar =  $this->Rstyle['TableNormal']['MPbot'];
				} else if (isset($this->Rstyle['Default']['Mbot'])){
					$amar =  " margin-bottom".$this->Rstyle['Default']['Mbot'];
				} else {
					$amar =  " margin-bottom:0px;";
				}
			}
			if ($PSret['Dcap'] == 'margin'){
				$cmar =  " margin-left:0px";
			} else if (isset($this->Rstyle[$Tstyle]['Mleft'])){ // set margin-left
				$cmar =  " margin-left".$this->Rstyle[$Tstyle]['Mleft'];
			} else if (isset($this->Rstyle['TableNormal']['Mleft'])){
				$cmar =  " margin-left".$this->Rstyle['TableNormal']['Mleft'];
			}
			if (isset($this->Rstyle[$Tstyle]['Mright'])){ //set margin-right
				$dmar =  " margin-right".$this->Rstyle[$Tstyle]['Mright'];
			} else if (isset($this->Rstyle['TableNormal']['Mright'])){
				$dmar =  " margin-right".$this->Rstyle['TableNormal']['Mright'];
			}
			if (isset($this->Rstyle[$Pstyle]['Bcolor'])){ // set text colour
				$bcol = " background-color:#".$this->Rstyle[$Pstyle]['Bcolor'].";";
			}
			if ($aind == ''){ // set left indent
				if (isset($Rind[$Listlevel])){
					$aind = $Rind[$Listlevel];
				} else if (isset($this->Rstyle[$Pstyle]['Ileft'])){
					$aind =  $this->Rstyle[$Pstyle]['Ileft'];
				}
			}
			if ($bind == ''){ // set right indent
				if (isset($this->Rstyle[$Pstyle]['Iright'])){
					$bind =  $this->Rstyle[$Pstyle]['Iright'];
				}
			}
			if ($cind == ''){ // set first line hanging indent
				if (isset($Rhang[$Listlevel])){
					$cind = $Rhang[$Listlevel];
				} else if (isset($this->Rstyle[$Pstyle]['Ihang'])){
					$cind =  $this->Rstyle[$Pstyle]['Ihang'];
				}
			}
			if ($dind == ''){ // set first line indent
				if (isset($this->Rstyle[$Pstyle]['Ifirst'])){
					$dind =  $this->Rstyle[$Pstyle]['Ifirst'];
				}
			}
			// return the paragraph styling
			$PSret['Pform'] = " style='".$bmar.$amar.$cmar.$dmar.$aind.$bind.$cind.$dind.$palign.$bcol.$hr."'";
			return $PSret;
		
		$PSret['Dcap'] = '';
		$PSret['Style'] = '';
		$PSret['Pform'] = '';
		$PSret['Lnum'] = '';
		$PSret['listnum'] = '';
	}
	
	/**
	 * CHECKS IF THERE IS A PICTURE PRESENT (used for higher resolution images)
	 * 
	 * @param XML $xml - The XML node
	 * @param $Psize - The image size sent in some modes
	 * @return String The location of the image
	 */
	private function checkPictFormating(&$xml, $Psize)
	{
		$content = trim($xml->readInnerXml());

		if (!empty($content)) {
			$ImgL = $Wtmp = $Htmp = $Ltmp = $Imgpos = '';
			$arr = explode(';', $Psize); // Get image size if $Psize was supplied to this function
			$l = count($arr);
			$d = 0;
			while ($d < $l){
				if (substr($arr[$d],0,5) == 'width' ) {
					$Wtmp = substr($arr[$d],6);
				}
				if (substr($arr[$d],0,6) == 'height' ) {
					$Htmp = substr($arr[$d],7);
				}
				if (substr($arr[$d],0,11) == 'margin-left' ) {
					$Ltmp = substr($arr[$d],12);
				}
				$d++;
			}
			if (isset($Wtmp)){	
				$ImgW = (float)substr($Wtmp,0,-2) * 1.4;
				$ImgH = (float)substr($Htmp,0,-2) * 1.4;
			}
			
			$relId;
			$notfound = true;
			$reader = new XMLReader();
			$reader->XML($content);
			static $Icount = 1;
			while ($reader->read()) {
				if ($reader->nodeType == XMLREADER::ELEMENT && $reader->name == "v:shape") { // Get image size if $Psize was not supplied to this function
					$Psize = $reader->getAttribute("style");
					$brr = explode(';', $Psize);
					$l = count($brr);
					$c = 0;
					while ($c < $l){
						if (substr($brr[$c],0,5) == 'width' ) {
							$Wtmp = substr($brr[$c],6);
						}
						if (substr($brr[$c],0,6) == 'height' ) {
							$Htmp = substr($brr[$c],7);
						}
						if (substr($brr[$c],0,11) == 'margin-left' ) {
							$Ltmp = substr($brr[$c],12);
						}
						$c++;
					}					
					$ImgW = (float)substr($Wtmp,0,-2) * 1.4;
					$ImgH = (float)substr($Htmp,0,-2) * 1.4;
					$ImgL = substr($Ltmp,0,-2);
				}
				if ($reader->name == "v:imagedata") { // Get image name
					$relId = $reader->getAttribute("r:id");
					$notfound = false;
				}
			}
			if ($ImgL <> ''){
				$Imgpos = ($ImgL < 50) ? "float:left;" : "float:right;";
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

	}






	/**
	 * CHECKS IF THERE IS AN IMAGE PRESENT (Used for lower resolution images)
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
			$Inline = $offset = $Imgpos = '';
			$Pcount = 0;
			
			while ($reader->read()) {
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
					if ($Pcount == 0){
						$relId = $reader->getAttribute("r:embed");
						$Pcount = 1;
						$notfound = false;
					} else {
						$relId = $reader->getAttribute("r:embed");  //In order to get the alternative .png image when the original is a .pdf image
						$notfound = false;
					}
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
	private function getHyperlink(&$xml,$RelF)
	{
		$ret = array('open'=>'<ul>','close'=>'</ul>');
		$link ='';
		static $rels_end = null ;
		static $rels_foot = null ;
		if($xml->hasAttributes) {
			$attribute = "";
			while($xml->moveToNextAttribute()) {
				if($xml->name == "r:id"){  // check for external hyperlinks
					$attribute = $xml->value;
				}
				if($xml->name == "w:anchor"){  // check for internal bookmark links
					$internalT = $xml->value;
					if (substr($internalT,0,1) == '_'){
					$internal = substr($internalT,1);
				} else {
					$internal = $internalT;
				}

					$this->anchor[$internal] = 'Y';
				}
			}
			
			if($attribute != "") {
				$reader = new XMLReader();
				if ($RelF == 'P'){
					$reader->XML($this->rels_xml->saveXML());
				} else if($RelF == 'F'){
					if(!$rels_foot){
						$zip = new ZipArchive();
						$_foot_rels = 'word/_rels/footnotes.xml.rels';
						if (true === $zip->open($this->file)) {
							//Get the footnotes relationships file
							if (($index = $zip->locateName($_foot_rels)) !== false) {
								$foot_rels = $zip->getFromIndex($index);
							}
							$zip->close();
						}
						$enc = mb_detect_encoding($foot_rels);
						$this->setXmlParts($rels_foot, $foot_rels, $enc);
						if($this->debug) {
							echo "<br>XML File : word/_rels/footnotes.xml.rels<br>";
							echo "<textarea style='width:100%; height: 200px;'>";
							echo $rels_foot->saveXML();
							echo "</textarea>";
						}
					}
					$reader->XML($rels_foot->saveXML()); //Get the footnote hyperlinks from the footnotes relationships file

				} else if($RelF == 'E'){
					if(!$rels_end){
						$zip = new ZipArchive();
						$_end_rels = 'word/_rels/endnotes.xml.rels';
						if (true === $zip->open($this->file)) {
							//Get the endnotes relationships file
							if (($index = $zip->locateName($_end_rels)) !== false) {
								$end_rels = $zip->getFromIndex($index);
							}
							$zip->close();
						}
						$enc = mb_detect_encoding($end_rels);
						$this->setXmlParts($rels_end, $end_rels, $enc);
						if($this->debug) {
							echo "<br>XML File : word/_rels/endnotes.xml.rels<br>";
							echo "<textarea style='width:100%; height: 200px;'>";
							echo $rels_end->saveXML();
							echo "</textarea>";
						}
					}
					$reader->XML($rels_end->saveXML()); //Get the endnote hyperlinks from the endnotes relationships file

				}
				
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
			$ret['open'] = "<a href='".$link."' target='_blank' rel='noopener noreferrer'>";
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
	 * @return String - The HTML code of the paragraph
	 */
	private function getParagraph(&$paragraph,$Tstyle)
	{
		$zst = array();
		$text = $BookMk = $BookRet = $Pstyle = $zst[0] = $CBdef = $CBdef = '';
		$list_format=array();
		$zzz = $text;
		$zstc = 1;
		$Pformat = 'N';

		$Dstyle = $this->getListFormating($paragraph,$Tstyle); //default styles for the document
		if (!isset($Dstyle['Pform'])){
			$Dstyle['Pform'] = '';
		}
		// loop through paragraph dom
		while ($paragraph->read()) {
			// look for elements
			if ($paragraph->nodeType == XMLREADER::ELEMENT && $paragraph->name === 'w:r') {
				if (!isset($list_format['style'])){
					$list_format['style'] = '';
				}
				if (!isset($list_format['Dcap'])){
					$list_format['Dcap'] = '';
				}
				$Pelement = $this->checkFormating($paragraph,$list_format['style'],$list_format['Dcap']); //check if this element is a page break and if so ignore this element
				if (!isset($Pelement['Pbreak'])){
					if ($Pformat == 'Y'){
						if ($list_format['Pform'] <> ''){
							$text = "<p".$BookMk.$list_format['Pform'].">"; // brings in paragraph formatting
						} else {
							$text .= "<p".$BookMk.$Dstyle['Pform'].">";
						}
						if(isset($list_format['listnum'])){
							$text .= $Pelement['style'].$list_format['Lnum']."</style>";
						}
						$Pformat = 'D';
					}
					if ($Pformat == 'N'){
						$text .= "<p".$BookMk.$Dstyle['Pform'].">";
						$Pformat = 'D';
					}
					$zst[$zstc] = $Pelement['style'];
					if ($zst[$zstc] != $zst[$zstc-1]){
						if ($zstc > 1){
							$text .= "</span>".$Pelement['style'];
						} else {
							$text .= $Pelement['style'];
						}
						$zstc++;
					}
					if (isset($Pelement['CRtext'])){
						$CRlink = $Pelement['CRtext'];
					}
					if (isset($Pelement['text'])){
						if (isset($CRlink)){
							$text .= "<a id='R".$CRlink."' href='#".$CRlink."'>".$Pelement['text']."</a>";
							$CRlink = '';
						} else {
							$text .= $Pelement['text']; 
						}
					}
				}
			} else if($paragraph->nodeType == XMLREADER::ELEMENT && $paragraph->name === 'w:pPr') { // Get list and paragraph formatting
				$list_format = $this->getListFormating($paragraph,$Tstyle);
				$Pformat = 'Y';
			} 
			else if($paragraph->name == "w:bookmarkStart") { // check for internal bookmark link and its return
				$BM = $paragraph->getAttribute("w:name");
				if (substr($BM,0,1) == '_'){
					$BookL = substr($BM,1);
				} else {
					$BookL = $BM;
				}
				if (isset($this->anchor[$BookL])){
					if ($BM  <> '_GoBack'){
						$BookMk = " id='".$BookL."'";
						$BookRet = "&nbsp;<a href='#R".$BookL."'><sup>[return]</a>";
					}
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
						$Pelement = $this->checkFormating($paragraph,$list_format['style'],$list_format['Dcap']); // Get inline style for list numbering
						$text .= $Pelement['style'].$list_format['Lnum']."</style>";
					}
					$Pformat = 'D';
				}
				if ($Pformat == 'N'){
					$text .= "<p".$Dstyle['Pform'].">";
					$Pformat = 'D';
				}
				$hyperlink = $this->getHyperlink($paragraph,'P'); // Add in hyperlinks and bookmarks
				$text .= $hyperlink['open'];
				$Pelement2 = $this->checkFormating($paragraph,$Pstyle,$list_format['Dcap']);
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
			else if ($paragraph->nodeType == XMLREADER::ELEMENT && $paragraph->name === 'w:checkBox') {
				$cb2 = new SimpleXMLElement($paragraph->readOuterXML());
				foreach ($cb2->children('w',true) as $ch) {
					if (in_array($ch->getName(), ['default']) ) {
						$CBdef = $ch['val'];
					}
				}
				if ($CBdef == '0'){
					$text .= "<input type=\"checkbox\">";
				} else {
					$text .= "<input type=\"checkbox\" checked>";
				}
			}
		}
		if ($zzz == $text){
			$text .= "<p".$Dstyle['Pform'].">&nbsp;</p> ";
		} else {
			if (substr($text,-1) == '>'){
				$text .= "&nbsp;</span>".$BookRet."</p> ";
			} else {
				$text .= "</span>".$BookRet."</p> ";
			}
		}
		return $text;
	}
			

// ------------------- END OF PARAGRAPH PROCESSING ------------------------

// ------------------- START OF TABLE PROCESSING ------------------------


	/**
	 * FIND NUMBER OF ROWS IN THE TABLE
	 *  
	 * @param XML $content - The XML node
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
	 * @return String - The HTML code of the table
	 */
	private function checkTableFormating(&$xml)
	{

		if ($this->Tcss == 'Y'){
			$table = "<table style='border-collapse:collapse; width:100%;'><tbody>";
		}  else {
			$table = "<table style='border-collapse:collapse;margin-left:auto;margin-right:auto;'><tbody>";
		}
		$Tstile = array();
		$TCcount = 0;
		$Twidth = 0;
		$Trow = 1;
		$Tstyle = '';
		$Colnum = 1;
		while ($xml->read()) {
			if ($xml->nodeType == XMLREADER::ELEMENT && $xml->name === 'w:tbl') { //Get number of rows in the table
				$Tinfo = $this->getrows(trim($xml->readOuterXML()));
				$Trows = $Tinfo['rows'];
				$Tmerge = $Tinfo['merge'];
			}
			if ($xml->name === 'w:tblStyle') { //Get table style
				$Tstyle = $xml->getAttribute("w:val");
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
						$Cwidth[$TCcount] = $tr9->getAttribute("w:w"); // column width
						$Twidth = $Twidth + $Cwidth[$TCcount]; // get width of the table
					}
				}
			}
			
			if ($xml->nodeType == XMLREADER::ELEMENT && $xml->name === 'w:tr') { //find and process a table row
				if (!isset($Tstile['top'])){
					if (isset($this->Rstyle[$Tstyle]['Btop'])){
						$Tstile['top'] = $this->Rstyle[$Tstyle]['Btop'];
					} else if (isset($this->Rstyle['TableNormal']['Btop'])){
						$Tstile['top'] = $this->Rstyle['TableNormal']['Btop'];
					} else {
						$Tstile['top'] = '';
					}
				}
				if (!isset($Tstile['left'])){
					if (isset($this->Rstyle[$Tstyle]['Bleft'])){
						$Tstile['left'] = $this->Rstyle[$Tstyle]['Bleft'];
					} else if (isset($this->Rstyle['TableNormal']['Bleft'])){
						$Tstile['left'] = $this->Rstyle['TableNormal']['Bleft'];
					} else {
						$Tstile['left'] = '';
					}
				}
				if (!isset($Tstile['bottom'])){
					if (isset($this->Rstyle[$Tstyle]['Bbottom'])){
						$Tstile['bottom'] = $this->Rstyle[$Tstyle]['Bbottom'];
					} else if (isset($this->Rstyle['TableNormal']['Bbottom'])){
						$Tstile['bottom'] = $this->Rstyle['TableNormal']['Bbottom'];
					} else {
						$Tstile['bottom'] = '';
					}
				}
				if (!isset($Tstile['right'])){
					if (isset($this->Rstyle[$Tstyle]['Bright'])){
						$Tstile['right'] = $this->Rstyle[$Tstyle]['Bright'];
					} else if (isset($this->Rstyle['TableNormal']['Bright'])){
						$Tstile['right'] = $this->Rstyle['TableNormal']['Bright'];
					} else {
						$Tstile['right'] = '';
					}
				}
				if (!isset($Tstile['insideH'])){
					if (isset($this->Rstyle[$Tstyle]['BinsideH'])){
						$Tstile['insideH'] = $this->Rstyle[$Tstyle]['BinsideH'];
					} else if (isset($this->Rstyle['TableNormal']['BinsideH'])){
						$Tstile['insideH'] = $this->Rstyle['TableNormal']['BinsideH'];
					} else {
						$Tstile['insideH'] = '';
					}
				}
				if (!isset($Tstile['insideV'])){
					if (isset($this->Rstyle[$Tstyle]['BinsideV'])){
						$Tstile['insideV'] = $this->Rstyle[$Tstyle]['BinsideV'];
					} else if (isset($this->Rstyle['TableNormal']['BinsideV'])){
						$Tstile['insideV'] = $this->Rstyle['TableNormal']['BinsideV'];
					} else {
						$Tstile['insideV'] = '';
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
					if ($tr->nodeType == XMLREADER::ELEMENT && $tr->name === 'w:tblCellMar') { //Get table row margin styles
						$tr3 = new SimpleXMLElement($tr->readOuterXML());
						foreach ($tr3->children('w',true) as $ch) {
							if (in_array($ch->getName(), ['top','left','bottom','right']) ) {
								$zlinM = round($ch['w']/$this->MTFactor);
								$Mname = "MR".$ch->getName();
								$this->Rstyle['Row'][$Mname] = ":".$zlinM."px;";
							}
						}
					}

					if ($tr->nodeType == XMLREADER::ELEMENT && $tr->name === 'w:tc') { //get cell borders and cell text and its formatting
						$tc = $this->processTableRow(trim($tr->readOuterXML()),$Tstyle);
					}
					$style = '';
					if ($tr->nodeType == XMLREADER::ELEMENT && $tr->name === 'w:tcPr') { //get cell border formatting
						$ts = $this->processTableStyle(trim($tr->readOuterXML()));
						if (isset($ts['left'])){  // set left border of a table cell
							$style .= " border-left".$ts['left'];
						} else{
							$style .= ($Tcol == 1) ? " border-left".$Tstile['left'] : " border-left".$Tstile['insideV'];
						}
						if (isset($ts['top'])){  // set the top border of a table cell
							$style .= " border-top".$ts['top'];
						} else{
							$style .= ($Trow == 1) ? " border-top".$Tstile['top'] : " border-top".$Tstile['insideH'];
						}
						if (isset($ts['bottom'])){  // set the bottom border of a table cell
							$style .= " border-bottom".$ts['bottom'];
						} else{
							if ($Tmerge[$Trow][$Tcol] > 1){  // set the bottom border of a table cell is a vertical cell merge
								$style .= ($Trow + $Tmerge[$Trow][$Tcol] - 1 == $Trows) ? " border-bottom".$Tstile['bottom'] : " border-bottom".$Tstile['insideH'];
							} else {  // set the bottom border of a table cell is not a vertical cell merge
								$style .= ($Trow == $Trows) ? " border-bottom".$Tstile['bottom'] : " border-bottom".$Tstile['insideH'];
							}
						}
						if (isset($ts['colspan'])){  // if a table cell is a horizontal cell merge, determine number of additional cells that make up the merge
							$TCoffset = $TCoffset + $ts['colspan'] - 1;
						}
						if (isset($ts['right'])){  // set the right border of the table cell
							$style .= " border-right".$ts['right'];
						} else{
							$style .= ($Tcol + $TCoffset == $TCcount) ? " border-right".$Tstile['right'] : " border-right".$Tstile['insideV'];
						}


						
						if ($Tmerge[$Trow][$Tcol] > 0){
							if ($this->Tcss == 'Y'){
								if ($Tmerge[$Trow][$Tcol] == 1){
									if (isset($ts['colspan'])){
										$table .= "<td colspan='".$ts['colspan']."'; style='width:".floor($ts['cellwidth'] / $Twidth *95)."%; ".$style;
									} else {
										$table .= "<td style='width:".floor($ts['cellwidth'] / $Twidth *95)."%; ".$style;
									}
								} else {
									if (isset($ts['colspan'])){
										$table .= "<td rowspan='".$Tmerge[$Trow][$Tcol]."'; colspan='".$ts['colspan']."'; style='width:".floor($ts['cellwidth'] / $Twidth *95)."%; ".$style;
									} else {
										$table .= "<td rowspan='".$Tmerge[$Trow][$Tcol]."'; style='width:".floor($ts['cellwidth'] / $Twidth *95)."%; ".$style;
									}
								}
							} else {
								if ($Tmerge[$Trow][$Tcol] == 1){
									if (isset($ts['colspan'])){
										$table .= "<td colspan='".$ts['colspan']."'; style='width:".$ts['cellwidth']/$this->MTFactor."px; ".$style;
									} else {
										$table .= "<td style='width:".$ts['cellwidth']/$this->MTFactor."px; ".$style;
									}
								} else {
									if (isset($ts['colspan'])){
										$table .= "<td rowspan='".$Tmerge[$Trow][$Tcol]."'; colspan='".$ts['colspan']."'; style='width:".$ts['cellwidth']/$this->MTFactor."px; ".$style;
									} else {
										$table .= "<td rowspan='".$Tmerge[$Trow][$Tcol]."'; style='width:".$ts['cellwidth']/$this->MTFactor."px; ".$style;
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
			if ($tc->name === "w:tcW") {
				$style['cellwidth'] = $tc->getAttribute("w:w");
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
	 * @return Array - The HTML code of the table row
	 */
	private function processTableRow($content,$Tstyle)
	{
		$tc = new XMLReader;
		$tc->xml($content);
		$ct = array();
		$count = 0;
		$valign = $halign = $ct['cell'] = $colours = $BackCol = $TextCol = '';
		$Cpos = $Cwidth = 0;
		while ($tc->read()) {
			$ztpp = '';
			$ztp = '';
			$text = '';
			if ($tc->nodeType == XMLREADER::ELEMENT && $tc->name === "w:p") {  // get cell text and its formatting
				$paragraph = new XMLReader;
				$p = $tc->readOuterXML();
				$paragraph->xml($p);
				$ct['cell'] .= $this->getParagraph($paragraph,$Tstyle);
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
			if ($tc->name === "w:tcW") {  // cell width
				$Cwidth = $tc->getAttribute("w:w");
			}
			if ($tc->name === "w:tab") {  // cell text position
				$Cpos = $tc->getAttribute("w:pos");
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
				if ($BackCol == 'auto'){
					$BackCol = 'FFFFFF';
				}
			}
			if ($tc->name === "w:shd" AND $tc->getAttribute("w:color")) {  // cell background color
				$TextCol = $tc->getAttribute("w:color");
			}
			if ($BackCol <> '' OR $TextCol <> ''){
				if ($TextCol == ''){
					$red = hexdec(substr($BackCol,0,2));
					$green = hexdec(substr($BackCol,2,2));
					$blue = hexdec(substr($BackCol,4,2));
					$color = (($red * 0.299) + ($green * 0.587) + ($blue * 0.114) > 186) ?  '000000' : 'FFFFFF';
				} else {
					if ($BackCol == ''){
						$BackCol = 'FFFFFF';
					}
					$color = $TextCol;
				}
				$colours = "background-color: #".$BackCol."; color: #".$color.";";
			}

			
		}

		if ($valign == ''){
			$valign =  " vertical-align: top;";
		}
		if ($halign == ''){
			if ($Cwidth > 0){
				if ($Cpos / $Cwidth < .3){
					$halign =  " text-align: left;";
				} else if ($Cpos / $Cwidth < .7){
					$halign =  " text-align: center;";
				} else {
					$halign =  " text-align: right;";
				}
			}
		}
		$ct['align'] = $halign.$valign.$colours;
		if ($ct['cell'] == ''){
			$ct['cell'] = "&nbsp;";			
		}

		return $ct;
	}
	
	
// ------------------- END OF TABLE PROCESSING ------------------------



	/**
	 * READS THE GIVEN DOCX FILE INTO HTML FORMAT
	 *  
	 * @param String $filename - The DOCX file name
	 * @param String $imgcss - The image and table option parameters
	 * @return String - With HTML code of the DOCX file
	 */
	public function readDocument($filename,$imgcss)
	{
		if (!file_exists($filename)){
			exit("Cannot find file : ".$filename." ");
		}
		$this->file = $filename; // makes the filename available throughout the class
		$this->Icss = substr($imgcss,0,1); // A 'Y' enables images to be styled by external CSS
		$this->Tcss = substr($imgcss,1,1); //A 'Y' makes tables to be 100% width
		$this->readZipPart(); // Makes the document and relationships file available throughout the class
		$this->findstyles(); // Makes the style parameters from the styles XML file available throughout the class
		$anchor = array();
		$this->anchor = $anchor;
		
		$reader = new XMLReader();
		$reader->XML($this->doc_xml->saveXML());
		$text = "<div style='position:fixed; bottom:50vh; right:10px; border:2px solid black; padding:2px; min-width:3%; text-align:center; background-color:#eeeeee';><a href='#top'>Top</a></div>";
		$text .= "<div style='margin:10px;'>"; // Provide a small margin around the html output
		while ($reader->read()) {
		// look for new paragraphs or table
			$paragraph = new XMLReader;
			$p = $reader->readOuterXML();
			if ($reader->nodeType == XMLREADER::ELEMENT && $reader->name === 'w:tbl') { // finds and gets tables
				$paragraph->xml($p);
				$text .= $this->checkTableFormating($paragraph);
				$reader->next();
			}
			else if ($reader->nodeType == XMLREADER::ELEMENT && $reader->name === 'w:p') {
				// finds and gets paragraphs			
				$paragraph->xml($p); // set up new instance of XMLReader for parsing paragraph independantly	
				$text .= $this->getParagraph($paragraph,'');
				$reader->next();
			}
		}
		$Foot = $this->footnotes(); // Get any Footnotes in the document
		if (isset($Foot[1])) {
			$text .= "<br>&nbsp;";
			$text .= "<hr><p style='margin-top:6px;margin-bottom:6px;'><b>FOOTNOTES</b></p>";
			$Fcount = 1;
			while (isset($Foot[$Fcount])){
				$text .= "<p style='padding-left:50px;text-indent:-50px;margin-top:6px;margin-bottom:6px;'><sup><a id='FN".$Fcount."' href='#FN".$Fcount."R'>[".$Fcount."]</a></sup>&nbsp;&nbsp;&nbsp;".$Foot[$Fcount]."</p>";
				++$Fcount;
			}
		}
		
		$Endn = $this->endnotes(); //Get any Endnotes in the document
		if (isset($Endn[1])) {
			$text .= "<br>&nbsp;";
			$text .= "<hr><p style='margin-top:6px;margin-bottom:6px;'><b>ENDNOTES</b></p>";
			$Fcount = 1;
			while (isset($Endn[$Fcount])){
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






