<?php

set_time_limit(10);

include_once 'Sample_Header.php';

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\DocumentLayout;
use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Slide;
use PhpOffice\PhpPresentation\AbstractShape;
use PhpOffice\PhpPresentation\Shape\MemoryDrawing;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Shape\Group;

class PhpPptTree {
    protected $oPhpPresentation;
    protected $htmlOutput;
    
    public function __construct(PhpPresentation $oPHPPpt)
    {
        $this->oPhpPresentation = $oPHPPpt;
    }
    
    public function display()
    {
        $this->append('<div class="container-fluid pptTree">');
        $this->append('<div class="row">');
        $this->append('<div class="collapse in col-md-6">');
        $this->append('<div class="tree">');
        $this->append('<ul>');
        $this->displayPhpPresentation($this->oPhpPresentation);
        $this->append('</ul>');
        $this->append('</div>');
        $this->append('</div>');
        $this->append('<div class="col-md-6">');
        $this->displayPhpPresentationInfo($this->oPhpPresentation);
        $this->append('</div>');
        $this->append('</div>');
        $this->append('</div>');
        
        return $this->htmlOutput;
    }
    
    protected function append($sHTML)
    {
        $this->htmlOutput .= $sHTML;
    }
    
    protected function displayPhpPresentation(PhpPresentation $oPHPPpt)
    {
        $this->append('<li><span><i class="fa fa-folder-open"></i> PhpPresentation</span>');
        $this->append('<ul>');
        $this->append('<li><span class="shape" id="divPhpPresentation"><i class="fa fa-info-circle"></i> Info "PhpPresentation"</span></li>');
        foreach ($oPHPPpt->getAllSlides() as $oSlide) {
            $this->append('<li><span><i class="fa fa-minus-square"></i> Slide</span>');
            $this->append('<ul>');
            $this->append('<li><span class="shape" id="div'.$oSlide->getHashCode().'"><i class="fa fa-info-circle"></i> Info "Slide"</span></li>');
            foreach ($oSlide->getShapeCollection() as $oShape) {
                if($oShape instanceof Group) {
                    $this->append('<li><span><i class="fa fa-minus-square"></i> Shape "Group"</span>');
                    $this->append('<ul>');
                    // $this->append('<li><span class="shape" id="div'.$oShape->getHashCode().'"><i class="fa fa-info-circle"></i> Info "Group"</span></li>');
                    foreach ($oShape->getShapeCollection() as $oShapeChild) {
                        $this->displayShape($oShapeChild);
                    }
                    $this->append('</ul>');
                    $this->append('</li>');
                } else {
                    $this->displayShape($oShape);
                }
            }
            $this->append('</ul>');
            $this->append('</li>');
        }
        $this->append('</ul>');
        $this->append('</li>');
    }
    
    protected function displayShape(AbstractShape $shape)
    {
        if($shape instanceof MemoryDrawing) {
            $this->append('<li><span class="shape" id="div'.$shape->getHashCode().'">Shape "MemoryDrawing"</span></li>');
        } elseif($shape instanceof RichText) {
            $this->append('<li><span class="shape" id="div'.$shape->getHashCode().'">Shape "RichText"</span></li>');
        } else {
            var_export($shape);
        }
    }
    
    protected function displayPhpPresentationInfo(PhpPresentation $oPHPPpt)
    {
        $this->append('<div class="infoBlk" id="divPhpPresentationInfo">');
        $this->append('<dl>');
        $this->append('<dt>Number of slides</dt><dd>'.$oPHPPpt->getSlideCount().'</dd>');
        $this->append('<dt>Document Layout Height</dt><dd>'.$oPHPPpt->getLayout()->getCY(DocumentLayout::UNIT_MILLIMETER).' mm</dd>');
        $this->append('<dt>Document Layout Width</dt><dd>'.$oPHPPpt->getLayout()->getCX(DocumentLayout::UNIT_MILLIMETER).' mm</dd>');
        $this->append('<dt>Properties : Category</dt><dd>'.$oPHPPpt->getProperties()->getCategory().'</dd>');
        $this->append('<dt>Properties : Company</dt><dd>'.$oPHPPpt->getProperties()->getCompany().'</dd>');
        $this->append('<dt>Properties : Created</dt><dd>'.$oPHPPpt->getProperties()->getCreated().'</dd>');
        $this->append('<dt>Properties : Creator</dt><dd>'.$oPHPPpt->getProperties()->getCreator().'</dd>');
        $this->append('<dt>Properties : Description</dt><dd>'.$oPHPPpt->getProperties()->getDescription().'</dd>');
        $this->append('<dt>Properties : Keywords</dt><dd>'.$oPHPPpt->getProperties()->getKeywords().'</dd>');
        $this->append('<dt>Properties : Last Modified By</dt><dd>'.$oPHPPpt->getProperties()->getLastModifiedBy().'</dd>');
        $this->append('<dt>Properties : Modified</dt><dd>'.$oPHPPpt->getProperties()->getModified().'</dd>');
        $this->append('<dt>Properties : Subject</dt><dd>'.$oPHPPpt->getProperties()->getSubject().'</dd>');
        $this->append('<dt>Properties : Title</dt><dd>'.$oPHPPpt->getProperties()->getTitle().'</dd>');
        $this->append('</dl>');
        $this->append('</div>');
        
        foreach ($oPHPPpt->getAllSlides() as $oSlide) {
            $this->append('<div class="infoBlk" id="div'.$oSlide->getHashCode().'Info">');
            $this->append('<dl>');
            $this->append('<dt>HashCode</dt><dd>'.$oSlide->getHashCode().'</dd>');
            $this->append('<dt>Slide Layout</dt><dd>'.$oSlide->getSlideLayout().'</dd>');
            $this->append('<dt>Offset X</dt><dd>'.$oSlide->getOffsetX().'</dd>');
            $this->append('<dt>Offset Y</dt><dd>'.$oSlide->getOffsetY().'</dd>');
            $this->append('<dt>Extent X</dt><dd>'.$oSlide->getExtentX().'</dd>');
            $this->append('<dt>Extent Y</dt><dd>'.$oSlide->getExtentY().'</dd>');
            $this->append('</dl>');
            $this->append('</div>');
            
            foreach ($oSlide->getShapeCollection() as $oShape) {
                if($oShape instanceof Group) {
                    foreach ($oShape->getShapeCollection() as $oShapeChild) {
                        $this->displayShapeInfo($oShapeChild);
                    }
                } else {
                    $this->displayShapeInfo($oShape);
                }
            }
        }
    }
    
    protected function displayShapeInfo(AbstractShape $oShape)
    {
        $this->append('<div class="infoBlk" id="div'.$oShape->getHashCode().'Info">');
        $this->append('<dl>');
        $this->append('<dt>HashCode</dt><dd>'.$oShape->getHashCode().'</dd>');
        $this->append('<dt>Offset X</dt><dd>'.$oShape->getOffsetX().'</dd>');
        $this->append('<dt>Offset Y</dt><dd>'.$oShape->getOffsetY().'</dd>');
        $this->append('<dt>Height</dt><dd>'.$oShape->getHeight().'</dd>');
        $this->append('<dt>Width</dt><dd>'.$oShape->getWidth().'</dd>');
        $this->append('<dt>Rotation</dt><dd>'.$oShape->getRotation().'°</dd>');
        $this->append('<dt>Hyperlink</dt><dd>'.ucfirst(var_export($oShape->hasHyperlink(), true)).'</dd>');
        $this->append('<dt>Fill</dt><dd>@Todo</dd>');
        $this->append('<dt>Border</dt><dd>@Todo</dd>');
        if($oShape instanceof MemoryDrawing) {
            ob_start();
            call_user_func($oShape->getRenderingFunction(), $oShape->getImageResource());
            $sShapeImgContents = ob_get_contents();
            ob_end_clean();
            $this->append('<dt>Mime-Type</dt><dd>'.$oShape->getMimeType().'</dd>');
            $this->append('<dt>Image</dt><dd><img src="data:'.$oShape->getMimeType().';base64,'.base64_encode($sShapeImgContents).'"></dd>');
        } else {
            // Add another shape
        }
        $this->append('</dl>');
        $this->append('</div>');
    }
}

$pptReader = PhpOffice\PhpPresentation\IOFactory::createReader('PowerPoint97');
$oPHPPresentation = $pptReader->load('resources/Sample_12.ppt');

$oTree = new PhpPptTree($oPHPPresentation);
echo $oTree->display();
if (!CLI) {
	include_once 'Sample_Footer.php';
}
