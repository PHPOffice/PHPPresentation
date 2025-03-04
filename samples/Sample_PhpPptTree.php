<?php

use PhpOffice\PhpPresentation\AbstractShape;
use PhpOffice\PhpPresentation\DocumentLayout;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\Drawing;
use PhpOffice\PhpPresentation\Shape\Group;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Shape\RichText\BreakElement;
use PhpOffice\PhpPresentation\Style\Bullet;

class Sample_PhpPptTree
{
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

    protected function append($sHTML): void
    {
        $this->htmlOutput .= $sHTML;
    }

    protected function displayPhpPresentation(PhpPresentation $oPHPPpt): void
    {
        $this->append('<li><span><i class="fa fa-folder-open"></i> PhpPresentation</span>');
        $this->append('<ul>');
        $this->append('<li><span class="shape" id="divPhpPresentation"><i class="fa fa-info-circle"></i> Info "PhpPresentation"</span></li>');
        foreach ($oPHPPpt->getAllSlides() as $oSlide) {
            $this->append('<li><span><i class="fa fa-minus-square"></i> Slide</span>');
            $this->append('<ul>');
            $this->append('<li><span class="shape" id="div' . $oSlide->getHashCode() . '"><i class="fa fa-info-circle"></i> Info "Slide"</span></li>');
            foreach ($oSlide->getShapeCollection() as $oShape) {
                if ($oShape instanceof Group) {
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

    protected function displayShape(AbstractShape $shape): void
    {
        if ($shape instanceof Drawing\Gd) {
            $this->append('<li><span class="shape" id="div' . $shape->getHashCode() . '">Shape "Drawing\Gd"</span></li>');
        } elseif ($shape instanceof Drawing\File) {
            $this->append('<li><span class="shape" id="div' . $shape->getHashCode() . '">Shape "Drawing\File"</span></li>');
        } elseif ($shape instanceof Drawing\Base64) {
            $this->append('<li><span class="shape" id="div' . $shape->getHashCode() . '">Shape "Drawing\Base64"</span></li>');
        } elseif ($shape instanceof Drawing\ZipFile) {
            $this->append('<li><span class="shape" id="div' . $shape->getHashCode() . '">Shape "Drawing\Zip"</span></li>');
        } elseif ($shape instanceof RichText) {
            $this->append('<li><span class="shape" id="div' . $shape->getHashCode() . '">Shape "RichText"</span></li>');
        } elseif ($shape instanceof Table) {
            $this->append('<li><span class="shape" id="div' . $shape->getHashCode() . '">Shape "Table"</span></li>');
        } else {
            var_dump($shape);
        }
    }

    protected function displayPhpPresentationInfo(PhpPresentation $oPHPPpt): void
    {
        $this->append('<div class="infoBlk" id="divPhpPresentationInfo">');
        $this->append('<dl>');
        $this->append('<dt>Number of slides</dt><dd>' . $oPHPPpt->getSlideCount() . '</dd>');
        $this->append('<dt>Document Layout Name</dt><dd>' . (empty($oPHPPpt->getLayout()->getDocumentLayout()) ? 'Custom' : $oPHPPpt->getLayout()->getDocumentLayout()) . '</dd>');
        $this->append('<dt>Document Layout Height</dt><dd>' . $oPHPPpt->getLayout()->getCY(DocumentLayout::UNIT_MILLIMETER) . ' mm</dd>');
        $this->append('<dt>Document Layout Width</dt><dd>' . $oPHPPpt->getLayout()->getCX(DocumentLayout::UNIT_MILLIMETER) . ' mm</dd>');
        $this->append('<dt>Properties : Category</dt><dd>' . $oPHPPpt->getDocumentProperties()->getCategory() . '</dd>');
        $this->append('<dt>Properties : Company</dt><dd>' . $oPHPPpt->getDocumentProperties()->getCompany() . '</dd>');
        $this->append('<dt>Properties : Created</dt><dd>' . $oPHPPpt->getDocumentProperties()->getCreated() . '</dd>');
        $this->append('<dt>Properties : Creator</dt><dd>' . $oPHPPpt->getDocumentProperties()->getCreator() . '</dd>');
        $this->append('<dt>Properties : Description</dt><dd>' . $oPHPPpt->getDocumentProperties()->getDescription() . '</dd>');
        $this->append('<dt>Properties : Keywords</dt><dd>' . $oPHPPpt->getDocumentProperties()->getKeywords() . '</dd>');
        $this->append('<dt>Properties : Last Modified By</dt><dd>' . $oPHPPpt->getDocumentProperties()->getLastModifiedBy() . '</dd>');
        $this->append('<dt>Properties : Modified</dt><dd>' . $oPHPPpt->getDocumentProperties()->getModified() . '</dd>');
        $this->append('<dt>Properties : Subject</dt><dd>' . $oPHPPpt->getDocumentProperties()->getSubject() . '</dd>');
        $this->append('<dt>Properties : Title</dt><dd>' . $oPHPPpt->getDocumentProperties()->getTitle() . '</dd>');
        $this->append('</dl>');
        $this->append('</div>');

        foreach ($oPHPPpt->getAllSlides() as $oSlide) {
            $this->append('<div class="infoBlk" id="div' . $oSlide->getHashCode() . 'Info">');
            $this->append('<dl>');
            $this->append('<dt>HashCode</dt><dd>' . $oSlide->getHashCode() . '</dd>');
            $this->append('<dt>Slide Layout</dt><dd>Layout::' . $this->getConstantName('\PhpOffice\PhpPresentation\Slide\Layout', $oSlide->getSlideLayout()) . '</dd>');

            $this->append('<dt>Offset X</dt><dd>' . $oSlide->getOffsetX() . '</dd>');
            $this->append('<dt>Offset Y</dt><dd>' . $oSlide->getOffsetY() . '</dd>');
            $this->append('<dt>Extent X</dt><dd>' . $oSlide->getExtentX() . '</dd>');
            $this->append('<dt>Extent Y</dt><dd>' . $oSlide->getExtentY() . '</dd>');
            $oBkg = $oSlide->getBackground();
            if ($oBkg instanceof Slide\AbstractBackground) {
                if ($oBkg instanceof Slide\Background\Color) {
                    $this->append('<dt>Background Color</dt><dd>#' . $oBkg->getColor()->getRGB() . '</dd>');
                }
                if ($oBkg instanceof Slide\Background\Image) {
                    $sBkgImgContents = file_get_contents($oBkg->getPath());
                    $this->append('<dt>Background Image</dt><dd><img src="data:image/png;base64,' . base64_encode($sBkgImgContents) . '"></dd>');
                }
            }
            $oNote = $oSlide->getNote();
            if (count($oNote->getShapeCollection()) > 0) {
                $this->append('<dt>Notes</dt>');
                foreach ($oNote->getShapeCollection() as $oShape) {
                    if ($oShape instanceof RichText) {
                        $this->append('<dd>' . $oShape->getPlainText() . '</dd>');
                    }
                }
            }

            $this->append('</dl>');
            $this->append('</div>');

            foreach ($oSlide->getShapeCollection() as $oShape) {
                if ($oShape instanceof Group) {
                    foreach ($oShape->getShapeCollection() as $oShapeChild) {
                        $this->displayShapeInfo($oShapeChild);
                    }
                } else {
                    $this->displayShapeInfo($oShape);
                }
            }
        }
    }

    protected function displayShapeInfo(AbstractShape $oShape): void
    {
        $this->append('<div class="infoBlk" id="div' . $oShape->getHashCode() . 'Info">');
        $this->append('<dl>');
        $this->append('<dt>HashCode</dt><dd>' . $oShape->getHashCode() . '</dd>');
        $this->append('<dt>Offset X</dt><dd>' . $oShape->getOffsetX() . '</dd>');
        $this->append('<dt>Offset Y</dt><dd>' . $oShape->getOffsetY() . '</dd>');
        $this->append('<dt>Height</dt><dd>' . $oShape->getHeight() . '</dd>');
        $this->append('<dt>Width</dt><dd>' . $oShape->getWidth() . '</dd>');
        $this->append('<dt>Rotation</dt><dd>' . $oShape->getRotation() . '°</dd>');
        $this->append('<dt>Hyperlink</dt><dd>' . ucfirst(var_export($oShape->hasHyperlink(), true)) . '</dd>');
        $this->append('<dt>Fill</dt>');
        if (null === $oShape->getFill()) {
            $this->append('<dd>None</dd>');
        } else {
            switch ($oShape->getFill()->getFillType()) {
                case PhpOffice\PhpPresentation\Style\Fill::FILL_NONE:
                    $this->append('<dd>None</dd>');

                    break;
                case PhpOffice\PhpPresentation\Style\Fill::FILL_SOLID:
                    $this->append('<dd>Solid (');
                    $this->append('Color : #' . $oShape->getFill()->getStartColor()->getRGB());
                    $this->append(' - Alpha : ' . $oShape->getFill()->getStartColor()->getAlpha() . '%');
                    $this->append(')</dd>');

                    break;
            }
        }
        $this->append('<dt>Border</dt><dd>@Todo</dd>');
        $this->append('<dt>IsPlaceholder</dt><dd>' . ($oShape->isPlaceholder() ? 'true' : 'false') . '</dd>');
        if ($oShape instanceof Drawing\Gd) {
            $this->append('<dt>Name</dt><dd>' . $oShape->getName() . '</dd>');
            $this->append('<dt>Description</dt><dd>' . $oShape->getDescription() . '</dd>');
            ob_start();
            call_user_func($oShape->getRenderingFunction(), $oShape->getImageResource());
            $sShapeImgContents = ob_get_contents();
            ob_end_clean();
            $this->append('<dt>Mime-Type</dt><dd>' . $oShape->getMimeType() . '</dd>');
            $this->append('<dt>Image</dt><dd><img src="data:' . $oShape->getMimeType() . ';base64,' . base64_encode($sShapeImgContents) . '"></dd>');
            if ($oShape->hasHyperlink()) {
                $this->append('<dt>Hyperlink URL</dt><dd>' . $oShape->getHyperlink()->getUrl() . '</dd>');
                $this->append('<dt>Hyperlink Tooltip</dt><dd>' . $oShape->getHyperlink()->getTooltip() . '</dd>');
            }
        } elseif ($oShape instanceof Drawing\AbstractDrawingAdapter) {
            $this->append('<dt>Name</dt><dd>' . $oShape->getName() . '</dd>');
            $this->append('<dt>Description</dt><dd>' . $oShape->getDescription() . '</dd>');
        } elseif ($oShape instanceof RichText) {
            $this->append('<dt># of paragraphs</dt><dd>' . count($oShape->getParagraphs()) . '</dd>');
            $this->append('<dt>Inset (T / R / B / L)</dt><dd>' . $oShape->getInsetTop() . 'px / ' . $oShape->getInsetRight() . 'px / ' . $oShape->getInsetBottom() . 'px / ' . $oShape->getInsetLeft() . 'px</dd>');
            $this->append('<dt>Text</dt>');
            $this->append('<dd>');
            foreach ($oShape->getParagraphs() as $oParagraph) {
                $this->append('Paragraph<dl>');
                $this->append('<dt>Alignment Horizontal</dt><dd> Alignment::' . $this->getConstantName('\PhpOffice\PhpPresentation\Style\Alignment', $oParagraph->getAlignment()->getHorizontal()) . '</dd>');
                $this->append('<dt>Alignment Vertical</dt><dd> Alignment::' . $this->getConstantName('\PhpOffice\PhpPresentation\Style\Alignment', $oParagraph->getAlignment()->getVertical()) . '</dd>');
                $this->append('<dt>Alignment Margin (L / R)</dt><dd>' . $oParagraph->getAlignment()->getMarginLeft() . ' px / ' . $oParagraph->getAlignment()->getMarginRight() . 'px</dd>');
                $this->append('<dt>Alignment Indent</dt><dd>' . $oParagraph->getAlignment()->getIndent() . ' px</dd>');
                $this->append('<dt>Alignment Level</dt><dd>' . $oParagraph->getAlignment()->getLevel() . '</dd>');
                $this->append('<dt>Bullet Style</dt><dd> Bullet::' . $this->getConstantName('\PhpOffice\PhpPresentation\Style\Bullet', $oParagraph->getBulletStyle()->getBulletType()) . '</dd>');
                if (Bullet::TYPE_NONE != $oParagraph->getBulletStyle()->getBulletType()) {
                    $this->append('<dt>Bullet Font</dt><dd>' . $oParagraph->getBulletStyle()->getBulletFont() . '</dd>');
                    $this->append('<dt>Bullet Color</dt><dd>' . $oParagraph->getBulletStyle()->getBulletColor()->getARGB() . '</dd>');
                }
                if (Bullet::TYPE_BULLET == $oParagraph->getBulletStyle()->getBulletType()) {
                    $this->append('<dt>Bullet Char</dt><dd>' . $oParagraph->getBulletStyle()->getBulletChar() . '</dd>');
                }
                if (Bullet::TYPE_NUMERIC == $oParagraph->getBulletStyle()->getBulletType()) {
                    $this->append('<dt>Bullet Start At</dt><dd>' . $oParagraph->getBulletStyle()->getBulletNumericStartAt() . '</dd>');
                    $this->append('<dt>Bullet Style</dt><dd>' . $oParagraph->getBulletStyle()->getBulletNumericStyle() . '</dd>');
                }
                $this->append('<dt>Line Spacing</dt><dd>' . $oParagraph->getLineSpacing() . '</dd>');
                $this->append('<dt>RichText</dt><dd><dl>');
                foreach ($oParagraph->getRichTextElements() as $oRichText) {
                    if ($oRichText instanceof BreakElement) {
                        $this->append('<dt><i>Break</i></dt>');
                    } else {
                        if ($oRichText instanceof TextElement) {
                            $this->append('<dt><i>TextElement</i></dt>');
                        } else {
                            $this->append('<dt><i>Run</i></dt>');
                        }
                        $this->append('<dd>' . $oRichText->getText());
                        $this->append('<dl>');
                        $this->append('<dt>Font Name</dt><dd>' . $oRichText->getFont()->getName() . '</dd>');
                        $this->append('<dt>Font Size</dt><dd>' . $oRichText->getFont()->getSize() . '</dd>');
                        $this->append('<dt>Font Color</dt><dd>#' . $oRichText->getFont()->getColor()->getARGB() . '</dd>');
                        $this->append('<dt>Font Transform</dt><dd>');
                        $this->append('<abbr title="Bold">Bold</abbr> : ' . ($oRichText->getFont()->isBold() ? 'Y' : 'N') . ' - ');
                        $this->append('<abbr title="Italic">Italic</abbr> : ' . ($oRichText->getFont()->isItalic() ? 'Y' : 'N') . ' - ');
                        $this->append('<abbr title="Underline">Underline</abbr> : Underline::' . $this->getConstantName('\PhpOffice\PhpPresentation\Style\Font', $oRichText->getFont()->getUnderline()) . ' - ');
                        $this->append('<abbr title="Strikethrough">Strikethrough</abbr> : ' . ($oRichText->getFont()->isStrikethrough() ? 'Y' : 'N') . ' - ');
                        $this->append('<abbr title="Baseline">Baseline</abbr> : ' . $oRichText->getFont()->getBaseline() . ' - ');
                        $this->append('<abbr title="SubScript">SubScript</abbr> : ' . ($oRichText->getFont()->isSubScript() ? 'Y' : 'N') . ' - ');
                        $this->append('<abbr title="SuperScript">SuperScript</abbr> : ' . ($oRichText->getFont()->isSuperScript() ? 'Y' : 'N'));
                        $this->append('</dd>');
                        if ($oRichText instanceof TextElement) {
                            if ($oRichText->hasHyperlink()) {
                                $this->append('<dt>Hyperlink URL</dt><dd>' . $oRichText->getHyperlink()->getUrl() . '</dd>');
                                $this->append('<dt>Hyperlink Tooltip</dt><dd>' . $oRichText->getHyperlink()->getTooltip() . '</dd>');
                            }
                        }
                        $this->append('</dl>');
                        $this->append('</dd>');
                    }
                }
                $this->append('</dl></dd></dl>');
            }
            $this->append('</dd>');
        }
        // Add another shape

        $this->append('</dl>');
        $this->append('</div>');
    }

    protected function getConstantName($class, $search, $startWith = '')
    {
        $fooClass = new ReflectionClass($class);
        $constants = $fooClass->getConstants();
        $constName = null;
        foreach ($constants as $key => $value) {
            if ($value == $search) {
                if (empty($startWith) || (!empty($startWith) && 0 === strpos($key, $startWith))) {
                    $constName = $key;
                }

                break;
            }
        }

        return $constName;
    }
}
