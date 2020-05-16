<?php
require_once 'vendor/autoload.php';

use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Alignment;

class pptxXmlAnalysis{

    /*
     * read ppth presentation
     * @params: $path - path to presentation
     * @return -- pptx instance or null, if presentation is not serializable
     */
    private function readPptx($path){
        $samples = array('PowerPoint97','Serialized','PowerPoint2007', 'ODPresentation');
        foreach ($samples as $str){
            $oReader = IOFactory::createReader($str);
            if($oReader->canRead($path)) {
                return $oReader->load($path);
            }
        }
        return null;
    }

    /*
     * Main function to analyse other pptx presentation
     * @params: $pathToPPTX - path to testing pptx
     *
     * @return -- mark from 0 to 1, whitch takes from formula : 0.2*Slides + 0.6*analysText + 0.2*count_other_shapes.
     * Slides is a mark with compare slides. It is 0 or 1.
     * analysText it is analys text. It double from 0 to 1, which takes Font and Text analyses in same parts(0.5 0.5)
     * count_other_shapes is a mark which compare all shapes from slide. It is the mark from 0 to 1.
     */
    public function analysis($pathToSamplePPTX, $pathToTestedPPTX){
        $samplePptx = $this->readPptx($pathToSamplePPTX);
        $analysPptx = $this->readPptx($pathToTestedPPTX);

        if($analysPptx == null || $samplePptx == null){
            //Error PPTX
            return 0;
        }

        $cmpSlides = $this->compareSlidesCount($samplePptx, $analysPptx);
        $analysText = $this->testFontsText($samplePptx, $analysPptx);
        $cmpShapes = $this->compareShapes($samplePptx, $analysPptx);
        $cmpLayouts = $this->compareLayout($samplePptx, $analysPptx);
        system('python test.py '.$pathToSamplePPTX.' '.$pathToTestedPPTX, $scored);
        system('python test.py '.$pathToSamplePPTX.' '.$pathToSamplePPTX, $max);
        $scored += $cmpShapes[0] + $cmpSlides[0] + $analysText[0] + $cmpLayouts[0];
        $max += $cmpShapes[1] + $cmpSlides[1] + $analysText[1] + $cmpLayouts[1];
        if($max == 0){
            return 1;
        }
        return $scored/$max;
    }

    private function getSlidesArray($pptx){
        $slides = array();
        foreach($pptx->getAllSlides() as $slide){
            array_push($slides, $slide);
        }
        return $slides;
    }

    private function compareLayout($sample_pptx, $tested_pptx){
        $sample_slides = $this->getSlidesArray($sample_pptx);
        $tested_slides = $this->getSlidesArray($tested_pptx);
        $scored = 0;
        $max = 0;
        for ($i = 0; $i < min(count($sample_slides), count($tested_slides)); $i++){
            $scored += $this->compareSlideLayouts($sample_slides[$i], $tested_slides[$i]);
            $max++;
        }
        return array($scored, $max);

    }

    //Подсчет слайдов. Возврящается количество слайдов
    private function compareSlidesCount($sample_pptx , $tested_pptx){
        return array(min(count($tested_pptx->getAllSlides()),count($sample_pptx->getAllSlides())),max(count($tested_pptx->getAllSlides()),count($sample_pptx->getAllSlides())));
    }

    //Получение текста с презентации
    private function getText($PPTX){
        $slidesArray = $PPTX->getAllSlides();
        $result = '';
        foreach ($slidesArray as $slide){
            $result = $result . $this->getTextFromSlide($slide);
        }
        return $result;
    }

    /*
     * Comparing text from original presentation and tested presentation
     * @params $analysPptx -- Analys presentation
     *
     * @return percentage of comparing presentations text. When text is big it can work incorrectly
     */
    private function compareText($sample_pptx, $tested_pptx){
        $sampleText = $this->getText($sample_pptx);
        $testText = $this->getText($tested_pptx);
        $diff = levenshtein($sampleText, $testText);
        return 1 - abs($diff/max(strlen($sampleText),strlen($testText)));
    }

    //Получение текста со слайда
    private function getTextFromSlide($slide){
        $res = "";
        $shapes = $slide->getShapeCollection();
        foreach($shapes as $shape){
            if($shape instanceof PhpOffice\PhpPresentation\Shape\RichText){
                $paragraphs = $shape->getParagraphs();
                foreach($paragraphs as $paragraph){
                    $res = $res . $paragraph->getPlainText();
                }
            }
        }
        return $res."\n";
    }

    /*
     * Comparing shapes without RichText shapes
     * @params: $analysPptx -- analys presentation
     *
     * @return mark from 0 to 1 with comparing.
     */
    private function compareShapes($sample_pptx, $tested_pptx){
        return array(min($this->countNotRichTextShapes($sample_pptx), $this->countNotRichTextShapes($tested_pptx)),max($this->countNotRichTextShapes($tested_pptx),$this->countNotRichTextShapes($sample_pptx)));
    }

    /*
     * Counting Shapes from pptx presentation
     * @params: $pptx -- presentation
     * @return -- numbers of shapes in this presentation
     */
    private function countNotRichTextShapes($pptx){
        $slides = $pptx->getAllSlides();
        $cnt = 0;
        foreach($slides as $slide){
            foreach ($slide->getShapeCollection() as $shape){
                if(!$shape instanceof PhpOffice\PhpPresentation\Shape\RichText){
                    $cnt++;
                }
            }
        }
        return $cnt;
    }

    /*
     * Analyses text from all slides from testPptx with samplePptx
     * @params: $testing -- testing pptx
     * @return -- mark from 0 to 1
     */
    public function testFontsText($sample_pptx, $tested_pptx){
        $smplSlides = $this->getSlidesArray($sample_pptx);
        $tstSlides = $this->getSlidesArray($tested_pptx);
        $max = 0;
        $scored = 0;
        for ($i = 0; $i < min(count($smplSlides),count($tstSlides)); $i++){
            $temp = $this->analysTextFromSlide($smplSlides[$i], $tstSlides[$i]);
            $scored += $temp[0];
            $max += $temp[1];
        }
        return array($scored,$max);
    }

    /*
     * Font analyse with 5 parametrs: Size, Bolt, Italic, Underline, Strikethrough
     * @params $sampleSlide -- slide with perfect task
     * @params $testSlide -- testing slide
     * @return -- mark of analyse from 0 to 1
     */
    private function analysTextFromSlide($sampleSlide, $testSlide){
        $sample_shapes = $sampleSlide->getShapeCollection();
        $test_shapes = $testSlide->getShapeCollection();
        $sample_rtb_elements = array();
        $test_rtb_elements = array();

        $max = 0;
        $scored = 0;

        foreach($sample_shapes as $shape){
            $tmp = $this->getCollectionRichTextElements($shape);
            foreach ($tmp as $elem)
                array_push($sample_rtb_elements, $elem);
        }
        foreach($test_shapes as $shape){
            $tmp = $this->getCollectionRichTextElements($shape);
            foreach ($tmp as $elem)
                array_push($test_rtb_elements, $elem);
        }

        for ($i = 0; $i < min(count($sample_rtb_elements),count($test_rtb_elements));$i++){
            $sampleFont = $sample_rtb_elements[$i]->getFont();
            $testFont = $test_rtb_elements[$i]->getFont();
            if($sampleFont == null&& $testFont == null)
                continue;
            $max+=10;
            if($sampleFont== null){
                $scored+=10;
                continue;
            }
            if($testFont == null){
                continue;
            }
            $scored += 5 - 5*levenshtein($sample_rtb_elements[$i]->getText(),$test_rtb_elements[$i]->getText())/max(strlen($sample_rtb_elements[$i]->getText()),strlen($test_rtb_elements[$i]->getText()));
            if($sampleFont->isItalic() == $testFont->isItalic())
            {
                $scored+=1;
            }
            if($sampleFont->getUnderline() == $testFont->getUnderline())
            {
                $scored+=1;
            }
            if($sampleFont->getSize() == $testFont->getSize()){
                $scored += 1;
            }
            if($sampleFont->isBold() == $testFont->isBold()){
                $scored+=1;
            }
            if($sampleFont->isStrikethrough() == $testFont->isStrikethrough()){
                $scored+=1;
            }
        }
        return array($scored,$max);
    }

    /*
     * Getting all RichTextElements from RichTextShape
     * @param: $shape -- @shape from presentation slide
     * @return -- array of RichTextElements from this shape
     */
    private function getCollectionRichTextElements($shape){
        $arrayElements = array();
        if($shape instanceof PhpOffice\PhpPresentation\Shape\RichText){
            $paragraphs = $shape->getParagraphs();
            foreach($paragraphs as $paragraph){
                foreach ($paragraph->getRichTextElements() as $richTextElement) {
                    array_push($arrayElements, $richTextElement);
                }
            }
        }
        return $arrayElements;
    }

    private function compareSlideLayouts($sample_slide, $tested_slide){
        if($sample_slide == null || $tested_slide == null)
            return 0;
        if($sample_slide->getSlideLayout()->getLayoutName() == $tested_slide->getSlideLayout()->getLayoutName())
            return 1;
        return 0;
    }

}
system('pip install python-pptx', $output);
$program = new pptxXmlAnalysis();
echo $program->analysis("sample.pptx", "pres.pptx");


