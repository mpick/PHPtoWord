<?php
/*******************************************************************************
* PHPtoWord                                                                    *
*                                                                              *
* Version: 0.1                                                                 *
* Date:    2011-08-19                                                          *
* Author:  Miles Pickens                                                       *
*******************************************************************************/
class PhpToWord
{
	var $font =			'Times New Roman'; //font name
	var $fontsize =			'12.0pt'; //font size
	var $fontstyle =		'normal'; //italic, obliquie or normal
	var $fontweight =		'normal'; //bold or normal
	
	function __construct(array $fontvars=null)
	{
		if($font['font'])		$this->font = $font['font'];
		if($fontvars['fontstyle'])	$this->fontstyle = $fontvars['fontstyle'];
		if($font['fontsize'])		$this->fontsize = $font['fontsize'].'.0pt';
		if($font['fontweight'])		$this->fontweight = $font['fontweight'];
		
	
	}
	
	
	function setFont($font=null, $fontstyle=null, $fontsize=null,$fontweight=null)
	{
		if($font)			$this->font =	$font;
		if($fontstyle)			$this->fontstyle = $fontstyle;
		if($fontsize)			$this->fontsize = $fontsize.'.0pt';
		if($fontweight)			$this->fontweight = $fontweight;
				
	}
	public function paragraph($txt,$align = null)
	{
		if(is_null($align)) $align = 'left';
		$string = "<p class=MsoNormal align={$align} style='text-align:{$align} "
			. "font-family:{$this->font};font-size:{$this->fontsize};font-weight:{$this->fontweight};font-style:{$this->fontstyle};'>"
			. $txt
			. '</p>';
		echo $string;
	}
	
	public function writeFile($filename)
	{
		ini_set('auto_detect_line_endings', true);
		$contents = file($filename);
		foreach ($contents as $row) {
			if (is_null($row)) {
				$this->linebreak();
			} else {
				$this->paragraph($row);
				$this->linebreak();
			}
		}
		
	}
	
	public function doubleSpace()
	{
		echo "<span style='mso-spacerun:yes'> </span>";
	}
	
	public function headerText($txt, $align=null)
	{
		if(is_null($align)) $align = 'center';
		$string = "<p class=MsoHeader align=center style='text-align:{$align}
			font-family:{$this->font};font-size:{$this->fontsize};font-weight:{$this->fontweight};font-style:{$this->fontstyle};'>"
			."<span style='font-size:{$this->fontsize};font-family:'{$this->font}','serif';color:#0033CC'>"
			.$txt
			.'</span></p>';
		echo $string;
	}
	
	public function image($w=null, $h=null, $src, $align=null)
	{
		//Expected src should not include the domain name, it will be added by this script
		$constrain = '';
		if(is_null($align)) $align = 'center';
		if($w && $h) $constrain = " width={$w} height={$h}"; 
		$url = $_SERVER['SERVER_NAME'];
		$string = "<img {$constrain} src='"
			.'http://'
			.$url
			.$src
			."'>";
		$string = $this->paragraph($string,$align);
		echo $string;
	}
	
	public function pageBreak()
	{
		//$this->footer();
		$string = "<br clear=all style='page-break-before:always'>";
		echo $string;
	}
	
	public function columns(array $content, $int)
	{
		$push =				1;
		$string =			"<p class=MsoHeader style='font-family:{$this->font};font-size:{$this->fontsize};font-weight:{$this->fontweight};font-style:{$this->fontstyle};'>";			
		$count =			count($content);
		if ($count > 3) {
			echo			'Error, more than 3 columns!';
		} else {
			foreach ($content as $row){
				$string .=	$row
					.	"<span style='mso-tab-count:{$push}'></span>";				
			}
		}
		$string .=			'</p>';
		echo $string;
		
		
	}
	
	public function footer($txt = null, $align = null, $showdate = false, $showpage = false)
	{	
		if(!is_array($txt)) $txt = array($txt);
		if(!$align) $align = 'center';
		$count = 1;
		if (is_array($txt)) {
			$count = count($txt);
		}
		for ($i=0;$i<$count;$i++)
		{
			
			$string = "<p class='MsoNormal' align='{$align}' style='text-align:{$align};font-family:{$this->font};font-size:{$this->fontsize};font-weight:{$this->fontweight};font-style:{$this->fontstyle};'>"
				. ($showdate ? 'Printed: '.date("m/d/Y")."<span style='mso-tab-count:1'></span>" : '')
				. $txt[$i]
				. ($showpage ? "<span style='mso-tab-count:2'></span>Page <span style='mso-field-code:\" PAGE \"'></span> of <span style='mso-field-code:\" NUMPAGES \"'></span>" : '')
				. '</p>';
			echo $string;
		}
	}
	
	public function lineBreak($x=null)
	{
		if(!$x) $x = 1;
		for($i=0;$i<$x;$i++){
			$string = "<p class=MsoNormal><o:p>&nbsp;</o:p></p>";
			echo $string;
		}

	}
	
	public function table($rows,$columns,array $content)
	{
		$string = "<table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0
		 style='font-family:{$this->font};font-size:{$this->fontsize};font-weight:{$this->fontweight};font-style:{$this->fontstyle};
		 border-collapse:collapse;border:none;mso-border-alt:solid black .5pt;
		 mso-border-themecolor:text1;mso-yfti-tbllook:1184;mso-padding-alt:0in 5.4pt 0in 5.4pt'>";
		for($i=0;$i<$rows;$i++){
			 $string .= "<tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes'>";
			 for($n=0;$n<$columns;$n++){
				$string .= "<td width=148 valign=top style='width:110.7pt;border:solid black 1.0pt;
				  mso-border-themecolor:text1;border-left:none;mso-border-left-alt:solid black .5pt;
				  mso-border-left-themecolor:text1;mso-border-alt:solid black .5pt;mso-border-themecolor:
				  text1;padding:0in 5.4pt 0in 5.4pt'>"
					. "<p class=MsoNormal style='text-align:justify'>"
					. $content[$i][$n]
					. "</p>"
					. "</td>";
			 }
			 $string .= "</tr>";
		}
		$string .= "</table>";
		
		echo $string;
		
	}
	
	public function bulletedList(array $content)
	{
		$string = "<ul style='margin-top:0in' type=disc>";
		foreach($content as $row){
			$string .= "<li class=MsoNormal style='text-align:justify;mso-list:l0 level1 lfo1'>"
				. $row
				. "</li>";
		}
		$string .= "</ul>";
		echo $string;

	}
	
	
	public function close()
	{
		$string = '</div>'
			. '</body>'
			. '</html>';
		echo $string;
	}
	
	public function open($filename=null)
	{
		if(!$filename) $filename = 'Document';
		/* if web output is needed, but it will look off
		 * $string = "<html xmlns:v=\"urn:schemas-microsoft-com:vml\""
			. "xmlns:o=\"urn:schemas-microsoft-com:office:office\""
			. "xmlns:w=\"urn:schemas-microsoft-com:office:word\">"
		 * 
		 */
		$string = "<html>"
			. '<head>';
		echo $string;
		header("Content-type: application/vnd.ms-word");
		header("Content-Disposition: attachment;Filename={$filename}.doc");
		$string = "<meta http-equiv='Content-Type' content='text/html; charset=UTF-8'>"
			."<style>"
			."@page Section1 {mso-footer:f1;}"
			."div.Section1{page:Section1;}"
			."p.MsoFooter, li.MsoFooter, div.MsoFooter{"
			."mso-pagination:widow-orphan;"
			."tab-stops:center 216.0pt right 432.0pt;}"
			."p.MsoNormal, li.MsoNormal, div.MsoNormal"
			."{mso-style-unhide:no;"
			."mso-style-qformat:yes;"
			."mso-style-parent:'';"
			."margin:0in;"
			."margin-bottom:.0001pt;"
			."mso-pagination:widow-orphan;"
			."font-size:{$this->fontsize};"
			."font-family:'{$this->font}','serif';"
			."mso-fareast-font-family:'{$this->font}';}"
			."p.MsoHeader, li.MsoHeader, div.MsoHeader"
			."{mso-style-unhide:no;"
			."margin:0in;"
			."margin-bottom:.0001pt;"
			."mso-pagination:widow-orphan;"
			."tab-stops:center 3.0in right 6.0in;"
			."font-size:{$this->fontsize};"
			."mso-bidi-font-size:{$this->fontsize};"
			."font-family:'{$this->font}','sans-serif';"
			."mso-fareast-font-family:'{$this->font}';"
			."mso-bidi-font-family:'{$this->font}';}"
			."a:link, span.MsoHyperlink"
			."{mso-style-unhide:no;"
			."color:blue;"
			."text-decoration:underline;"
			."text-underline:single;}"
			."a:visited, span.MsoHyperlinkFollowed"
			."{mso-style-unhide:no;"
			."color:purple;"
			."mso-themecolor:followedhyperlink;"
			."text-decoration:underline;"
			."text-underline:single;}"
			."p.MsoAcetate, li.MsoAcetate, div.MsoAcetate"
			."{mso-style-noshow:yes;"
			."mso-style-unhide:no;"
			."margin:0in;"
			."margin-bottom:.0001pt;"
			."mso-pagination:widow-orphan;"
			."font-size:{$this->fontsize};"
			."font-family:'{$this->font}','sans-serif';"
			."mso-fareast-font-family:'{$this->font}';}"
			."</style>"
			."</head>"
			."<body>"
			."<div class=Section1>";
		echo $string;
	}
}
