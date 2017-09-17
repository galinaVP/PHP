<?php

use PhpOffice\PhpWord;

require 'vendor/autoload.php';

// Creating the new document...
$phpWord = new \PhpOffice\PhpWord\PhpWord();
/* Note: any element you append to a document must reside inside of a Section. */

                               //////STYLES HEADER&FOOTER////
$SectionStyle = array('orientation' => 'portrait',
               'marginLeft' => \PhpOffice\PhpWord\Shared\Converter::cmToPixel(30), 
               'marginRight' => \PhpOffice\PhpWord\Shared\Converter::cmToPixel(15),
               'marginTop' => \PhpOffice\PhpWord\Shared\Converter::cmToPixel(10),
               'borderTopColor' => '000000'
         );
		 
$lineStyle = array(
        'width'       => \PhpOffice\PhpWord\Shared\Converter::cmToPixel(18),
        'height'      => \PhpOffice\PhpWord\Shared\Converter::cmToPixel(0),
        'positioning' => 'absolute',
        'weight'      => 1,
    );
$lineStyleBold = array(
        'width'       => \PhpOffice\PhpWord\Shared\Converter::cmToPixel(18),
        'height'      => \PhpOffice\PhpWord\Shared\Converter::cmToPixel(0),
        'positioning' => 'absolute',
        'weight'      => 2,
    );
$titleStyle = array('name' => 'Arial Black', 'color' => '191970','bold' => true,'size' => '12');
$headerTextBoldV11Style = array('name' => 'Verdana', 'color' => '191970','bold' => true,'size' => '11');
$headerTextBoldV9Style = array('name' => 'Verdana', 'color' => '191970','bold' => true,'size' => '9');
$headerTextV9Style = array('name' => 'Verdana', 'color' => '191970','bold' => false,'size' => '9');
$headerRightAlignStyle = array('alignment' => 'right');
$headerLeftAlignStyle = array('alignment' => 'left');
$footerTextStyle = array('name' => 'Tahoma', 'color' => '191970','bold' => false,'size' => '8');
$tableStyle = array('cellMarginTop'=>'0', 'cellMarginBottom'=>'0');


// Adding an empty Section to the document...
$section = $phpWord -> createSection($SectionStyle);

                                /////HEADER&FOOTER/////
								
//Adding header to the list
$header = $section->createHeader();

//$header->addText('Nick Halle				', $titleStyle, $headerRightAlignStyle); 
//Adding line1
//$header->addLine($lineStyle);

//Adding main header text
$table=$header->addTable($tableStyle);

$table->addRow();
$table->addCell(4000);
$table->addCell(4500);
$table->addCell(3600)->addText('Nick Halle', $titleStyle, array('underline'=>'single'));

$table->addRow();
$table->addCell(4000)->addText('RECHNUNG № 0137', $headerTextBoldV11Style, array ('spaceAfter' => '0','spaceAfter' => '0'));
$table->addCell(4500);
$table->addCell(3600)->addText('IMPORT EXPORT <w:br/>Inh Halle Nick <w:br/>Volkmar Str 1-7 Halle Nr.9.c <w:br/>DE – 12099 Berlin', $headerTextBoldV9Style);

$table->addRow();
$table->addCell(1000)->addText('RECHNUNGSANSCHRIFT: <w:br/>UAB RIMELANA <w:br/>V.a graiciuno 6 <w:br/>LT02241 Vilnius <w:br/>LITAUEN <w:br/>LT253729811', $headerTextV9Style);
$table->addCell(4500);
$table->addCell(2000)->addPreserveText('<w:br/><w:br/><w:br/><w:br/>Seite {PAGE} von {NUMPAGES}', $headerTextBoldV9Style, $headerRightAlignStyle);

$table->addRow();
$table->addCell(1000);
$table->addCell(4500);
$table->addCell(2000)->addText('date:17/09/2017',$headerTextV9Style, $headerRightAlignStyle);
//$table->addCell(2000)->addField('DATE', array('dateformat' => 'dd.MM.yyyy'), array('PreserveFormat'),$headerTextV9Style);

$header->addLine($lineStyleBold);

//Create footer
$footer = $section->createFooter();
$footer->addLine($lineStyleBold);

//create table with text
$table = $footer->addTable($tableStyle);

$table->addRow();
$table->addCell(4500)->addText('Bankverbindung: Inhaber: Nick Halle <w:br/>Bank: Netbank <w:br/>IBAN:DE46 2009 0500 0002 5600 11 <w:br/>BIC: GENODEF1S15', $footerTextStyle);
$table->addCell(4500)->addText('Bankverbindung: Inhaber: Nick Halle <w:br/>Bank: HypoVereinsbank <w:br/>IBAN: DE83 2003 0000 0015 9936 37 <w:br/>HYVEDEMM300', $footerTextStyle);
$table->addCell(2500)->addText('Steuer-Nr. 21/328/00469 <w:br/>Ust-IdNr. DE298265166 <w:br/>E-Mail:nick.halle@gmx.de', $footerTextStyle);

                           ////FILE BODY - TABLE/////
//Table styles
$TableStyle = array('cellMargin' => 0, 'alignment' => \PhpOffice\PhpWord\SimpleType\JcTable::CENTER);
$TableCellStyle = array('valign' => 'center');
$TableFontStyle = array('bold' => false, 'alignment' => 'center', 'name'=>'Calibri', 'size'=>'12');
//Read from file
$lines = file ('load.txt');
    
//Add item while i<amount of items			   
$arrayLenght=count($lines);					   
$table = $section->addTable($TableStyle);
for ($i=0; $i<=$arrayLenght; $i++){
	//devide line to cell (devider is ",")
     list($cell1, $cell2, $cell3, $cell4, $cell5) = explode (",", $lines[$i]);
	 $table->addRow(300);
     $table->addCell(600)->addText("$cell1", $TableFontStyle);
     $table->addCell(3000)->addText("$cell2", $TableFontStyle);
     $table->addCell(5000)->addText("$cell3", $TableFontStyle);
	 $table->addCell(1000)->addText("$cell4", $TableFontStyle);
	 $table->addCell(2000)->addText("$cell5", $TableFontStyle);
};
						   
						   ////SUMMARY   TABLE/////
$TableStyleName = 'Summary Table';	
$TableStyle = array('borderSize' => 6, 'borderColor' => '000000', 'cellMargin' => 0, 'alignment' => \PhpOffice\PhpWord\SimpleType\JcTable::CENTER);
$TableFirstRowStyle = array('borderBottomSize' => 18, 'borderBottomColor' => '000000');
$TableCellStyle = array('valign' => 'center');
$TableFontStyle = array('bold' => true, 'alignment' => 'right', 'name'=>'Arial', 'size'=>'10');
$phpWord->addTableStyle($TableStyleName, $TableStyle, $TableFirstRowStyle);
$table = $section->addTable($TableStyleName);
    //Head of a summary table - constant
$table->addRow(200);
$table->addCell(2000, $TableCellStyle)->addText('NETTO, €', $TableFontStyle);
$table->addCell(2000, $TableCellStyle)->addText('MwSt, 00', $TableFontStyle);
$table->addCell(2000, $TableCellStyle)->addText('BRUTTO, €',$TableFontStyle);
    //Body of a summary table - various:
$table->addRow(300);
$table->addCell(2000, $TableCellStyle)->addText('Value1', $TableFontStyle);
$table->addCell(2000, $TableCellStyle)->addText('Value2', $TableFontStyle);
$table->addCell(2000, $TableCellStyle)->addText('Value3', $TableFontStyle);



// Saving the document as OOXML file...
$objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
$objWriter->save('Template.docx');

?>