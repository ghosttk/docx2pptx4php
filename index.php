<?php

require './vendor/autoload.php';

function docx2html($source)
{
    $targetDocx = new \PhpOffice\PhpWord\PhpWord();
    $targetSection = $targetDocx->addSection();
    $phpWord = \PhpOffice\PhpWord\IOFactory::load($source);
    $html = '';
    foreach ($phpWord->getSections() as $section) {
		    $tarTextRun = "";
        foreach ($section->getElements() as $ele1) {
            $paragraphStyle = $ele1->getParagraphStyle();
            if ($paragraphStyle) {
                $html .= '<p style="text-align:'. $paragraphStyle->getAlignment() .';text-indent:20px;">';
            } else {
                $html .= '<p>';
            }
            if ($ele1 instanceof \PhpOffice\PhpWord\Element\TextRun) {
		    $fBreakLine = false;
		    $fBreakPage = false;
		    $numMatched = $charMatched = array();
		    $oldNum = $oldChar = "";
                foreach ($ele1->getElements() as $ele2) {
                    if ($ele2 instanceof \PhpOffice\PhpWord\Element\Text) {
			$tarTextRun.=$strResult;
			if(preg_match("/^\d./", $tarTextRun, $numMatched)){
			        //$tarTextRun = "";
				if($oldNum === $numMatched[0]){
					$fBreakPage = false;
				}
				else{
					$fBreakPage = true;
					$fBreakLine = true;
					$oldNum = $numMatched[0];
					$targetSection->addPageBreak();
				}
			}
			else{
				if(preg_match("/[ABCD]\./", $tarTextRun, $charMatched)){
					if($oldChar === $charMatched[0]){
						$fBreakLine = false;
					}
					else{
						$oldChar = $charMatched[0];
						$fBreakLine = true;
				$targetSection->addText("oldChar=".$oldChar);
				$targetSection->addText("charMatched=".$charMatched[0]);
					}
				}
				else{
				}
			}
		    	$strResult = $ele2->getText();
			if($fBreakLine===true || $fBreakPage===true ){
				$targetSection->addText($tarTextRun);
				$tarTextRun = "";
			}
			var_dump($tarTextRun);
		    	$nPosResult = strpos($strResult, "B");
			if($nPosResult===0){
				var_dump($strResult);
				$subStrA=substr($strResult, 0, $nPosResult);
			}
                        $style = $ele2->getFontStyle();
                        $fontFamily = mb_convert_encoding($style->getName(), 'GBK', 'UTF-8');
                        $fontSize = $style->getSize();
                        $isBold = $style->isBold();
                        $styleString = '';
                        $fontFamily && $styleString .= "font-family:{$fontFamily};";
                        $fontSize && $styleString .= "font-size:{$fontSize}px;";
                        $isBold && $styleString .= "font-weight:bold;";
                        $html .= sprintf('<span style="%s">%s</span>',
                            $styleString,
                            mb_convert_encoding($ele2->getText(), 'GBK', 'UTF-8')
                        );
                    } elseif ($ele2 instanceof \PhpOffice\PhpWord\Element\Image) {
                        $imageSrc = 'images/' . md5($ele2->getSource()) . '.' . $ele2->getImageExtension();
                        $imageData = $ele2->getImageStringData(true);
                        // $imageData = 'data:' . $ele2->getImageType() . ';base64,' . $imageData;
                        file_put_contents($imageSrc, base64_decode($imageData));
                        $html .= '<img src="'. $imageSrc .'" style="width:100%;height:auto">';
                    }
                }
            }
            $html .= '</p>';
        }
    }
    $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($targetDocx, 'Word2007');
    $objWriter->save('target.docx');
    return mb_convert_encoding($html, 'UTF-8', 'GBK');
}




$dir = str_replace('\\', '/', __DIR__) . '/';
$source = $dir . '0306.docx';
echo docx2html($source);


?>
