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
	private $tmpDir = 'tmp';
	
	/**
	 * CONSTRUCTOR
	 * 
	 * @param Boolean $debug Debug mode or not
	 * @return void
	 */
	public function __construct($debug_=null, $encoding=null)
	{
		if($debug_ != null) {
			$this->debug = $debug_;
		}
		if ($encoding != null) {
			$this->encoding = $encoding;
		}
		$this->tmpDir = dirname(__FILE__);
	}

	/**
	 * Sets the tmp directory where images will be stored
	 * 
	 * @param string $tmp The location 
	 * @return void
	 */
	private function setTmpDir($tmp)
	{
		$this->tmpDir = $tmp;
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
		
		if($this->debug) {
			echo "XML File : word/document.xml<br>";
			echo "<textarea style='width:100%; height: 200px;'>";
			echo $this->doc_xml->saveXML();
			echo "</textarea>";
			echo "XML File : word/_rels/document.xml.rels<br>";
			echo "<textarea style='width:100%; height: 200px;'>";
			echo $this->rels_xml->saveXML();
			echo "</textarea>";
			echo "XML File : word/numbering.xml<br>";
			echo "<textarea style='width:100%; height: 200px;'>";
			echo $this->numb_xml->saveXML();
			echo "</textarea>";
			echo "XML File : word/styles.xml<br>";
			echo "<textarea style='width:100%; height: 200px;'>";
			echo $this->styles_xml->saveXML();
			echo "</textarea>";
			echo "XML File : word/fontTable.xml<br>";
			echo "<textarea style='width:100%; height: 200px;'>";
			echo $this->fonts_xml->saveXML();
			echo "</textarea>";
			echo "XML File : word/theme/theme1.xml<br>";
			echo "<textarea style='width:100%; height: 200px;'>";
			echo $this->theme_xml->saveXML();
			echo "</textarea>";
			echo "XML File : word/settings.xml<br>";
			echo "<textarea style='width:100%; height: 200px;'>";
			echo $this->settings_xml->saveXML();
			echo "</textarea>";
		}
	}


	/**
	 * Looks up a font in the themes XML file and returns the various fonts
	 * 
	 * @param String $style - The name of the style
	 * @returns the major and minor font of the theme
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
					}
				}
			}
			if ($reader1->nodeType == XMLREADER::ELEMENT && $reader1->name == 'a:minorFont') {
				$st2 = new XMLReader;
				$st2->xml(trim($reader1->readOuterXML()));
				while ($st2->read()) {
					if ($st2->name == 'a:latin') {
						$Mfont['minor'] = $st2->getAttribute("typeface");
					}
				}
			}
		}
		return $Mfont;
	}





	/**
	 * Looks up a style in the styles XML file and returns the various style parameters
	 * 
	 * @param String $style - The name of the style
	 * @returns the various parameters of the style
	 */
	private function findstyles($style)
	{
		$Rfont = $this->findfonts();
		$reader1 = new XMLReader();
		$reader1->XML($this->styles_xml->saveXML());
		while ($reader1->read()) {
		// look for required style
			$znum = 1;
			if ($reader1->nodeType == XMLREADER::ELEMENT && $reader1->name == 'w:style') {
				if ($reader1->getAttribute("w:styleId") == $style){
					$st1 = new XMLReader;
					$st1->xml(trim($reader1->readOuterXML()));
					while ($st1->read()) {

						if($st1->name == "w:rFonts" and $st1->getAttribute("w:ascii")) {
							$Rstyle['Font'] = " font-family: ".$st1->getAttribute("w:ascii").";";
						}
						if($st1->name == "w:rFonts" and $st1->getAttribute("w:asciiTheme")) {
							$FontTheme = $st1->getAttribute("w:asciiTheme");
						}						
						if($st1->name == "w:sz") {
							$Rstyle['FontS'] = round($st1->getAttribute("w:val")/22*100);
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
							if ($st1->getAttribute("w:before")){
								$Rstyle['Mtop'] =  "margin-top: ".round($st1->getAttribute("w:before")/20)."px;";
							}
							if ($st1->getAttribute("w:after")){
								$Rstyle['Mbot'] =  "margin-bottom: ".round($st1->getAttribute("w:after")/20)."px;";
							}
						}
						if($st1->name == "w:ind") { // Checks for paragragh indent
							if ($st1->getAttribute("w:left")){
								$Rstyle['Ileft'] =  "padding-left: ".round($st1->getAttribute("w:left")/20)."px;";
							}
							if ($st1->getAttribute("w:right")){
								$Rstyle['Iright'] =  "padding-right: ".round($st1->getAttribute("w:right")/20)."px;";
							}
							if ($st1->getAttribute("w:hanging")){
								$Rstyle['Ihang'] =  "text-indent: -".round($st1->getAttribute("w:hanging")/20)."px;";
							}
							if ($st1->getAttribute("w:firstLine")){
								$Rstyle['Ifirst'] =  "text-indent: ".round($st1->getAttribute("w:firstLine")/20)."px;";
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
									$Rstyle[$Bname] = ":".$zlinB."px $line #".$tbc.";";
								}
							}
						}
						if ($st1->nodeType == XMLREADER::ELEMENT && $st1->name === 'w:tblCellMar') { //Get table margin styles
							$tc3 = new SimpleXMLElement($st1->readOuterXML());
							foreach ($tc3->children('w',true) as $ch) {
								if (in_array($ch->getName(), ['top','left','bottom','right']) ) {
									$zlinM = round($ch['w']/13);
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
	 * CHECKS THE FONT FORMATTING OF A GIVEN ELEMENT
	 * 
	 * @param XML $xml The XML node
	 * @return String HTML formatted code
	 */
	private function checkFormating(&$xml,$Pstyle,$Tstyle)
	{	
//		if ($Pstyle <> ''){
			$Rstyle = $this->findstyles($Pstyle); // get the defined styles for this paragraph
//		}
		
		$node = trim($xml->readOuterXML());
		$t = '';
		// add <br> tags
		if (strstr($node,'<w:br ')) $t = '<br>';					 
		// look for formatting tags
		$f = "<span style='";
		$reader = new XMLReader();
		$reader->XML($node);
		$ret="";
		$img = null;

		while ($reader->read()) {
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
				$Lstyle['Font'] ="font-family: ".$reader->getAttribute("w:ascii").";";
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
				$Lstyle['FontS'] = round($reader->getAttribute("w:val")/22*100);
			}
			if($reader->name == 'w:drawing' && !empty($reader->readInnerXml())) {
				$r = $this->checkImageFormating($reader);
				$img = $r !== null ? "<image src='".$r['image']."' ".$r['style']." />" : null;
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
		}
		if ($script == 'Y' AND $Fsize){
			$Fsize = round($Fsize * 0.65);
		}
		if ($Fsize){
			$f .= "font-size: ".$Fsize."%;";
		} else if ($script == 'Y'){
			$f .= "font-size: 65%;";
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
		$t .= ($img !== null ? $img : htmlentities($xml->expand()->textContent));

		$ret['style'] = $f;
		$ret['text'] = $t;
		return $ret;
	}
	

	
	/**
	 * CHECKS THE ELEMENT FOR List ELEMENTS and their numbering and PARAGRAGH FORMATTING
	 * 
	 * @param XML $xml The XML node
	 * @return String HTML formatted code
	 */
	private function getListFormating(&$xml,$Tstyle)
	{
		$T1style = $this->findstyles($Tstyle); // get the defined styles for this paragraph from the table style
		if ($Tstyle <> "TableNormal"){
			$T2style = $this->findstyles('TableNormal'); // get the basic styles for tables
		}
		$zzlistilvl = '';
		$node = trim($xml->readOuterXML());

		$reader = new XMLReader();
		$reader->XML($node);
		$ret="";
		$close = "";
		while ($reader->read()){
			if($reader->name == "w:pStyle" && $reader->hasAttributes ) {
				$Pstyle = $reader->getAttribute("w:val");
				$ret['style'] = $Pstyle;
				$Rstyle = $this->findstyles($Pstyle); // get the defined styles for this paragraph

			}
			if($reader->name == "w:ilvl" && $reader->hasAttributes) { // List formating - list level
				$zzlistilvl = $reader->getAttribute("w:val");
			}
			if($reader->name == "w:numId" && $reader->hasAttributes) { // List formating - List cross reference
				$zzlistnumId = $reader->getAttribute("w:val");
			}

			if($reader->name == "w:rFonts") { // get font style for list numbering
				$font =" style = 'font-family: ".$reader->getAttribute("w:ascii")."';";
			}
				
			if($reader->name == "w:spacing") { // Checks for paragragh spacing
				if ($reader->getAttribute("w:before")){
					$bmar =  "margin-top: ".round($reader->getAttribute("w:before")/13)."px;";
				}
				if ($reader->getAttribute("w:after")){
					$amar =  "margin-bottom: ".round($reader->getAttribute("w:after")/13)."px;";
				}

			}
			if($reader->name == "w:ind") { // Checks for paragragh indent
				if ($reader->getAttribute("w:left")){
					$aind =  "padding-left: ".round($reader->getAttribute("w:left")/13)."px;";
				}
				if ($reader->getAttribute("w:right")){
					$bind =  "padding-right: ".round($reader->getAttribute("w:right")/13)."px;";
				}
				if ($reader->getAttribute("w:hanging")){
					$cind =  "text-indent: -".round($reader->getAttribute("w:hanging")/13)."px;";
				}
				if ($reader->getAttribute("w:firstLine")){
					$dind =  "text-indent: ".round($reader->getAttribute("w:firstLine")/13)."px;";
				}
			}
			if($reader->name == "w:jc") { // Checks for paragragh alignment
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
		
		if ($bmar == ''){
			if ($Rstyle['Mtop']){
				$bmar =  $Rstyle['Mtop'];
			} else if ($T1style['Mtop']){
				$bmar =  " margin-top".$T1style['Mtop'];
			} else if ($T2style['Mtop']){
				$bmar =  " margin-top".$T2style['Mtop'];
			}	
		}
		if ($amar == ''){
			if ($Rstyle['Mbot']){
				$amar =  $Rstyle['Mbot'];
			} else if ($T1style['Mbottom']){
				$amar =  " margin-bottom".$T1style['Mbottom'];
			} else if ($T2style['Mbottom']){
				$amar =  " margin-bottom".$T2style['Mbottom'];
			}	
		}
		if ($T1style['Mleft']){
			$cmar =  " margin-left".$T1style['Mleft'];
		} else if ($T2style['Mleft']){
			$cmar =  " margin-left".$T2style['Mleft'];
		}
		if ($T1style['Mright']){
			$dmar =  " margin-right".$T1style['Mright'];
		} else if ($T2style['Mright']){
			$dmar =  " margin-right".$T2style['Mright'];
		}
		if ($Rstyle['Bcolor']){
			$bcol = " background-color:#".$Rstyle['Bcolor'].";";
		}
		if ($aind == ''){
			if ($Rstyle['Ileft']){
				$aind =  $Rstyle['Ileft'];
			}
		}
		if ($bind == ''){
			if ($Rstyle['Iright']){
				$bind =  $Rstyle['Iright'];
			}
		}
		if ($cind == ''){
			if ($Rstyle['Ihang']){
				$bind =  $Rstyle['Ihang'];
			}
		}
		if ($dind == ''){
			if ($Rstyle['Ifirst']){
				$bind =  $Rstyle['Ifirst'];
			}
		}
		$ret['Pform'] = " style='".$bmar.$amar.$cmar.$dmar.$aind.$bind.$cind.$dind.$palign.$bcol.$hr."'";
		
		
		//CHECKS the List Numbering file for details of selected list type and start numbering
		$reader1 = new XMLReader();
		$reader1->XML($this->numb_xml->saveXML());
		static $zlistonce = 0;
		static $zcref = '';
		static $zlstart = '';
		static $zltype = '';
		if ($zlistonce == 0){
			while ($reader1->read()) {
			// look for new List definitions
				$znum = 1;
				if ($reader1->nodeType == XMLREADER::ELEMENT && $reader1->name == 'w:num') {
					$znumid = $reader1->getAttribute("w:numId");
				}
				if($reader1->nodeType == XMLREADER::ELEMENT && $reader1->name == 'w:abstractNumId') {
					$zabsno = $reader1->getAttribute("w:val");
					$zcref[$znumid] = $zabsno;
					$znum++;
				}
				
				if($reader1->nodeType == XMLREADER::ELEMENT && $reader1->name == 'w:abstractNum') {
					$zabsnum = $reader1->getAttribute("w:abstractNumId");
				}
				
				if($reader1->nodeType == XMLREADER::ELEMENT && $reader1->name == 'w:lvl') {
					$zlvl = $reader1->getAttribute("w:ilvl");
				}
				if($reader1->nodeType == XMLREADER::ELEMENT && $reader1->name == 'w:start') {
					$zstart = $reader1->getAttribute("w:val");
				$zlstart[$zabsnum][$zlvl] = $zstart;
				}
				if($reader1->nodeType == XMLREADER::ELEMENT && $reader1->name == 'w:numFmt') {
					$znumf= $reader1->getAttribute("w:val");
					$zltype[$zabsnum][$zlvl] = $znumf;
				}
			$zlistonce = 1;
			}
		}
		$zzcrossref = $zcref[$zzlistnumId];
		if ($zlstart[$zzcrossref][$zzlistilvl] == 1){
			$zstnum = '';
		} else {
			$zstnum = "start = '".$zlstart[$zzcrossref][$zzlistilvl]."' ";
		}

		switch($zltype[$zzcrossref][$zzlistilvl]) {
			case 'decimal':
				$ret['openF'] = "<ol ".$zstnum."><li".$font.">";
				$ret['openL'] = "<li".$font.">";
				$ret['closeF'] = "</span></li>";
				$ret['closeL'] = "</span></li></ol>";
				break;
			case 'lowerLetter':
				$ret['openF'] = "<ol type = 'a' ".$zstnum."><li".$font.">";
				$ret['openL'] = "<li".$font.">";
				$ret['closeF'] = "</span></li>";
				$ret['closeL'] = "</span></li></ol>";
				break;
			case 'upperLetter':
				$ret['openF'] = "<ol type = 'A' ".$zstnum."><li".$font.">";
				$ret['openL'] = "<li".$font.">";
				$ret['closeF'] = "</span></li>";
				$ret['closeL'] = "</span></li></ol>";
				break;
			case 'lowerRoman':
				$ret['openF'] = "<ol type = 'i' ".$zstnum."><li".$font.">";
				$ret['openL'] = "<li".$font.">";
				$ret['closeF'] = "</span></li>";
				$ret['closeL'] = "</span></li></ol>";
				break;
			case 'upperRoman':
				$ret['openF'] = "<ol type = 'I' ".$zstnum."><li".$font.">";
				$ret['openL'] = "<li".$font.">";
				$ret['closeF'] = "</span></li>";
				$ret['closeL'] = "</span></li></ol>";
				break;
			case 'bullet':
				$ret['openF'] = "<ul><li>";
				$ret['openL'] = "<li>";
				$ret['closeF'] = "</span></li>";
				$ret['closeL'] = "</span></li></ul>";
				break;
		}

		$ret['level'] = $zzlistilvl;
		
		return $ret;

	}
	
	/**
	 * CHECKS IF THERE IS AN IMAGE PRESENT
	 * 
	 * @param XML $xml The XML node
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
			
			while ($reader->read() && $notfound) {
				if ($reader->name == "wp:inline") {
					$Inline = 'Y';
				}
				if ($reader->name == "wp:posOffset") {
					$offset = (int)$xml->expand()->textContent;
				}
				if ($reader->name == "wp:extent") {
					$ImgW = round($reader->getAttribute("cx")/9000);
					$ImgH = round($reader->getAttribute("cy")/9000);
				}
				if ($reader->name == "a:blip") {
					$relId = $reader->getAttribute("r:embed");
					$notfound = false;
				}
			}
			if ($offset){
				$Imgpos = ($offset < 1000) ? "float:left;" : "float:right;";
			}
			$image['style'] = "style='".$Imgpos."width:".$ImgW."px;height:".$ImgH."px;padding:10px 15px 10px 15px;'";

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
	 * @param objetc $image The image object
	 * @param string $relId The image relationship Id
	 * @param string $name The image name
	 * @return Array With HTML open and closing tag definition
	 */
	private function createImage($image, $relId, $name)
	{
		$arr = explode('.', $name);
		$l = count($arr);
		$ext = strtolower($arr[$l-1]);
		
		if (!is_dir('images')){
			mkdir('images', 0755, true);
		}

		$im = imagecreatefromstring($image);
		$fname = 'images/'.$relId.'.'.$ext;

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
	 * @param XML $xml The XML node
	 * @return Array With HTML open and closing tag definition
	 */
	private function getHyperlink(&$xml)
	{
		$ret = array('open'=>'<ul>','close'=>'</ul>');
		$link ='';
		if($xml->hasAttributes) {
			$attribute = "";
			while($xml->moveToNextAttribute()) {
				if($xml->name == "r:id")
					$attribute = $xml->value;
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
		
		if($link != "") {
			$ret['open'] = "<a href='".$link."' target='_blank'>";
			$ret['close'] = "</a>";
		}
		
		return $ret;
	}




	/**
	 * PROCESS PARAGRAPH CONTENT
	 *  
	 * @param XML $xml The XML node
	 * @return The HTML code of the paragraph
	 */
	private function getParagraph(&$paragraph,$Tstyle)
	{
		static $Zlist = 0;
		$zzliF = '';
		$zzliL = '';
		$text = '';
		$pcount = 1;
		static $zzLlevel = -1;
		static $zzLlevelP = -1;
		static $zzliCF = '';
		static $zzliCL = '';
		static $zlistE = '';
		$list_format="";
		if ($Zlist == 0){
			$text .= '<p';
		} else {
			$zlistS = '<p';
		}
		$zzz = $text;
		$zct = 0;
		$zst = '';
		$zstc = 1;
		// loop through paragraph dom
		while ($paragraph->read()) {
			// look for elements
			if ($paragraph->nodeType == XMLREADER::ELEMENT && $paragraph->name === 'w:r') {
				$wr = 'Y';
				if($para_format == ""){
					if ($zzPE == '' AND $Zlist > 0){
						$text .= $zzliCL.$zlistE.$zlistS.$zlistPF;
						$Zlist = 0;
						$zct = 1;
					} else
					if ($zct == 0){
						$text .= '>';
						$zct = 1;
					}
					$zfor = $this->checkFormating($paragraph,$list_format['style'],$Tstyle); // Get inline style and associated text
					$zst[$zstc] = $zfor['style'];
					if ($zst[$zstc] != $zst[$zstc-1]){
						if ($zstc > 1){
							$text .= "</span>".$zfor['style'];
						} else if ($list_format['style'] == 'ListParagraph' AND $list_format['level'] == ''){
							$text .= "<br></span>".$zfor['style'];
						} else {
							$text .= $zfor['style'];
						}
					}
					$zstc++;
					$ztpp = $zfor['text'];  // Processing the text to remove unnecessary spaces
					if (substr($ztpp,1,5) == 'image'){
						$text .= $ztpp;
					} else {
						if (strlen($ztpp) - strlen(rtrim($ztpp)) > 12){
							$ztp = substr($ztpp,0,-13);
						} else {
							$ztp = substr($ztpp,0,-7);
						}
						if (strlen($ztp) - strlen(ltrim($ztp)) > 129){
							$zzs = 130;						
						} else if (strlen($ztp) - strlen(ltrim($ztp)) > 115){
							$zzs = 116;						
						} else if (strlen($ztp) - strlen(ltrim($ztp)) > 112){
							$zzs = 113;						
						} else if (strlen($ztp) - strlen(ltrim($ztp)) > 95){
							$zzs = 96;						
						} else if (strlen($ztp) - strlen(ltrim($ztp)) > 78){
							$zzs = 79;						
						} else if (strlen($ztp) - strlen(ltrim($ztp)) > 70){
							$zzs = 71;						
						} else if (strlen($ztp) - strlen(ltrim($ztp)) > 61){
							$zzs = 62;						
						} else if (strlen($ztp) - strlen(ltrim($ztp)) > 59){
							$zzs = 60;						
						} else if (strlen($ztp) - strlen(ltrim($ztp)) > 48){
							$zzs = 49;						
						} else if (strlen($ztp) - strlen(ltrim($ztp)) > 44){
							$zzs = 45;						
						} else if (strlen($ztp) - strlen(ltrim($ztp)) > 37){
							$zzs = 38;
						} else if (strlen($ztp) - strlen(ltrim($ztp)) > 27){
							$zzs = 28;
						} else if (strlen($ztp) - strlen(ltrim($ztp)) > 14){
							$zzs = 15;
						} else {
							$zzs = 9;
						}
						$ztest = substr($ztp,$zzs);
						if (ctype_space($ztest)){
							$text .= "&nbsp;";
						}
						$text .= substr($ztp,$zzs);
					}
				} else {
					$zzLlevel = $list_format['level']; // level of this list element
					if ($list_format['style'] = "ListParagraph"){
						$zzList = ' Y';
					}
					if ($zzList = 'Y' AND $zzLlevel <> ''){
						$zzliOF = $list_format['openF']; // add in lists
						$zzliOL = $list_format['openL'];
						if ($zzliOF <> '') {
							if ($Zlist == 0) {
								$text .= $zzliOF;
							} else if ($zzLlevel > $zzLlevelP){ // if this list level is greater than the previous one
								$text .= $zlistE.$zlistS.$zlistPF;
								$text .= $zzliOF;
							} else if ($zzLlevel < $zzLlevelP){ // if this list level is less than the previous one
								$text .= $zzliCL.$zlistE.$zlistS.$zlistPF;
								$text .= $zzliCF.$zzliOL;
							} else {
								$text .= $zzliCF.$zlistE.$zlistS.$zlistPF;
								$text .= $zzliOL;
							}
							$Zlist++;
							$zzLlevelP = $zzLlevel;
						} else if ($Zlist > 0){
							$text .= $zzliCL.$zlistE.$zlistS.$zlistPF;
							$Zlist = 0;
							$zzLlevel = '';
							$zzLlevelP = '';
						}
					} else if ($Zlist > 0){
						$text .= $zzliCL.$zlistE.$zlistS.$zlistPF;
						$Zlist = 0;
						$zzLlevel = '';
						$zzLlevelP = '';
					}

							
					$zfor = $this->checkFormating($paragraph,$list_format['style'],$Tstyle); //add in initial paragraph style and text
					$zst[$zstc] = $zfor['style'];
					if ($zst[$zstc] != $zst[$zstc-1]){
						if ($zstc > 1){
							$text .= "</span>".$zfor['style'];
						} else {
							$text .= $zfor['style'];
						}
					}
					$zstc++;
					$ztpp = $zfor['text'];
					if (substr($ztpp,1,5) == 'image'){
						$text .= $ztpp;
					} else {
						if (strlen($ztp) - strlen(rtrim($ztp)) > 12){
							$ztp = substr($ztpp,0,-13);
						} else {
							$ztp = substr($ztpp,0,-7);
						}
						if (strlen($ztp) - strlen(ltrim($ztp)) > 129){
							$zzs = 130;						
						} else if (strlen($ztp) - strlen(ltrim($ztp)) > 115){
							$zzs = 116;						
						} else if (strlen($ztp) - strlen(ltrim($ztp)) > 112){
							$zzs = 113;						
						} else if (strlen($ztp) - strlen(ltrim($ztp)) > 95){
							$zzs = 96;						
						} else if (strlen($ztp) - strlen(ltrim($ztp)) > 78){
							$zzs = 79;						
						} else if (strlen($ztp) - strlen(ltrim($ztp)) > 70){
							$zzs = 71;						
						} else if (strlen($ztp) - strlen(ltrim($ztp)) > 61){
							$zzs = 62;						
						} else if (strlen($ztp) - strlen(ltrim($ztp)) > 59){
							$zzs = 60;						
						} else if (strlen($ztp) - strlen(ltrim($ztp)) > 48){
							$zzs = 49;						
						} else if (strlen($ztp) - strlen(ltrim($ztp)) > 44){
							$zzs = 45;						
						} else if (strlen($ztp) - strlen(ltrim($ztp)) > 37){
							$zzs = 38;
						} else if (strlen($ztp) - strlen(ltrim($ztp)) > 27){
							$zzs = 28;
						} else if (strlen($ztp) - strlen(ltrim($ztp)) > 14){
							$zzs = 15;
						} else {
							$zzs = 9;
						}
						$ztest = substr($ztp,$zzs);
						if (ctype_space($ztest)){
							$text .= "&nbsp;";
						}
						$text .= substr($ztp,$zzs);
					}
					$para_format = '';	
					$paragraph->next();
				}
			} else if($paragraph->nodeType == XMLREADER::ELEMENT && $paragraph->name === 'w:pPr') { //lists and paragraph formatting
				$list_format = $this->getListFormating($paragraph,$Tstyle);
				$zta = $list_format['Pform']; // brings in paragraph formatting
				if ($pcount > 1){
					$text .= "</p><p";
				}
				$zct = 1;
				if ($list_format['style'] == 'ListParagraph' AND $list_format['level'] <> ''){
					$para_format = 'Y';
				} 
				if ($list_format['style'] == 'ListParagraph'){
					$zzPE = 'Y';
				} 
				if ($Zlist == 0){
					$text .= ($zta == '' && $zct == 0) ? '>' : $zta.'>';
				} else {
					$zlistPF = ($zta == '' && $zct == 0) ? '>' : $zta.'>';
				}
				$paragraph->next();
				$pcount++;
			}
			else if($paragraph->nodeType == XMLREADER::ELEMENT && $paragraph->name === 'w:drawing') { //images
				if ($zct == 0){
					$text .= '>';
					$zct = 1;
				}
				$paragraph->next();
			}
			else if ($paragraph->nodeType == XMLREADER::ELEMENT && $paragraph->name === 'w:hyperlink') {
				if ($zct == 0){
					$text .= '>';
					$zct = 1;
				}
				$hyperlink = $this->getHyperlink($paragraph);
				$text .= $hyperlink['open'];
				$zfor = $this->checkFormating($paragraph,$Pstyle,$Tstyle);
				if ($zstc > 1){
					$text .= "</span>".$zfor['style'];
				} else {
					$text .= $zfor['style'];
				}
				$zstc++;
				$ztp = substr($zfor['text'],0,-16);
				$text .= substr($ztp,68);
				$text .= $hyperlink['close'];
				$paragraph->next();
			}
		}
		if ($wr <> 'Y' AND $Zlist > 0){
			$text .= $zzliCL.$zlistE.$zlistS.$zlistPF;
			$Zlist = 0;
			$zzLlevel = '';
			$zzLlevelP = '';
		} 
		$wr = 'N';
		$zzliCF = $list_format['closeF']; // list closing
		$zzliCL = $list_format['closeL'];
		$list_format ="";
		$zzPE = '';
		if ($Zlist == 0 ){
			if ($zzz == $text){
				$text .= '>&nbsp;</p> ';
			} else {
				$text .= '&nbsp;</p> ';
			}
		} else {
			if ($zzz == $text){
				$zlistE = '></p> ';
			} else {
				$zlistE = '</p> ';
			}
		}
		return $text;
	}
			

// ------------------- END OF PARAGRAPH PROCESSING ------------------------

// ------------------- START OF TABLE PROCESSING ------------------------


	/**
	 * FIND NUMBER OF ROWS IN THE TABLE
	 *  
	 * @param XML $xml The XML node
	 * @return the number of rows in the table
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
//					$mcol[$Trow] = $Tcol;
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
	 * @param XML $xml The XML node
	 * @return THe HTML code of the table
	 */
	private function checkTableFormating(&$xml)
	{

		$table = "<table style='border-collapse:collapse;'><tbody>";
		$Tstile = '';
		$TCcount = 0;
		$Trow = 1;
		$T1style = $T2style = '';
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
						$Tstile[$Tname] = ":".$zlinT."px $line #".$tbc.";";
					}
				}
			}
			
			
			if ($xml->nodeType == XMLREADER::ELEMENT && $xml->name === 'w:tblGrid') { //Get number of columns and their widths in the table
				$tr9 = new XMLReader;
				$tr9->xml(trim($xml->readOuterXML()));
				while ($tr9->read()) {
					if($tr9->name === 'w:gridCol'){
						$TCcount++;
						$Cwidth[$TCcount] = round($tr9->getAttribute("w:w")/13); // column width
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
				$tc = $ts = "";
				$table .= '<tr>';

				$tr = new XMLReader;
				$tr->xml(trim($xml->readOuterXML()));
				$Tcol = 1;
				$TCoffset = 0;
				while ($tr->read()) {
					$Cstyle = '';
					if ($tr->nodeType == XMLREADER::ELEMENT && $tr->name === 'w:tc') { //get cell borders and cell text and its formatting
						$tc = $this->processTableRow(trim($tr->readOuterXML()),$Tstyle);
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
						if ($ts['colspan'] <> ''){  // if a table cell is a horizonatal cell merge, determine its width
							$TCoffset = $TCoffset + $ts['colspan'] - 1;
						}
						if ($ts['right']){  // set the right border of the table cell
							$style .= " border-right".$ts['right'];
						} else{
							$style .= ($Tcol + $TCoffset == $TCcount) ? " border-right".$Tstile['right'] : " border-right".$Tstile['insideV'];
						}


						
						if ($Tmerge[$Trow][$Tcol] > 0){
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
	 * @param string $content The XML node content
	 * @return THe HTML code of the table
	 */
	private function processTableStyle($content)
	{

		$tc = new XMLReader;
		$tc->xml($content);
		$style = '';

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
						$style[$Tname] = ":".$zlinT."px $line #".$tbc.";";
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
	 * @param string $content The XML node content
	 * @return The HTML code of the table
	 */
	private function processTableRow($content,$Tstyle)
	{
		$tc = new XMLReader;
		$tc->xml($content);
		$ct = "";
		$count = 0;
		$text = '';
		$valign = '';
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
	 * READS THE GIVEN DOCX FILE INTO HTML FORMAT
	 *  
	 * @param String $filename The DOCX file name
	 * @return String With HTML code
	 */
	public function readDocument($filename)
	{
		
		$this->file = $filename;
		$this->readZipPart($filename);
		$reader = new XMLReader();
		$reader->XML($this->doc_xml->saveXML());
		$text = "<div style='margin-left:10px; margin-right:10px;margin-top:10px; margin-bottom:10px;'>"; // Provide a small margin around the html output
		while ($reader->read()) {
		// look for new paragraphs or table
			$paragraph = new XMLReader;
			$p = $reader->readOuterXML();
			if ($reader->nodeType == XMLREADER::ELEMENT && $reader->name === 'w:tbl') { //tables
				$paragraph->xml($p);
				$text .= $this->checkTableFormating($paragraph);
				$reader->next();
			}
			else if ($reader->nodeType == XMLREADER::ELEMENT && $reader->name === 'w:p') {
				// finds paragraphs			
				$paragraph->xml($p); // set up new instance of XMLReader for parsing paragraph independantly	
				$text .= $this->getParagraph($paragraph,'');
				$reader->next();
			}
		}
		$text .= "</div>";
		$reader->close();
		if($this->debug) {
			echo "<div style='width:100%;'>";
			echo mb_convert_encoding($text, $this->encoding);
			echo "</div>";
		}
		return mb_convert_encoding($text, $this->encoding);
	}
}






