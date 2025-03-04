<?php

/**
 * This file is part of PHPPresentation - A pure PHP library for reading and writing
 * presentations documents.
 *
 * PHPPresentation is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPPresentation/contributors.
 *
 * @see        https://github.com/PHPOffice/PHPPresentation
 *
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

declare(strict_types=1);

namespace PhpOffice\PhpPresentation\Writer;

use PhpOffice\PhpPresentation\DocumentLayout;
use PhpOffice\PhpPresentation\Exception\InvalidParameterException;
use PhpOffice\PhpPresentation\HashTable;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\Drawing;
use PhpOffice\PhpPresentation\Shape\Media;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Shape\Table;
use PhpOffice\PhpPresentation\Slide;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Font;
use PhpOffice\PhpPresentation\Style\Shadow;

/**
 * HTML writer.
 */
class HTML extends AbstractWriter implements WriterInterface
{
    /**
     * @var string
     */
    protected $style = '';

    /**
     * @var string
     */
    protected $bodyList = '';

    /**
     * @var string
     */
    protected $bodySlides = '';

    /**
     * @var float
     */
    protected $ratioX;

    /**
     * @var float
     */
    protected $ratioY;

    /**
     * @var bool
     */
    protected $isPDF = false;

    /**
     * Create a new \PhpOffice\PhpPresentation\Writer\ODPresentation.
     */
    public function __construct(?PhpPresentation $pPhpPresentation = null)
    {
        // Assign PhpPresentation
        $this->setPhpPresentation($pPhpPresentation ?? new PhpPresentation());

        // Set HashTable variables
        $this->oDrawingHashTable = new HashTable();
    }

    /**
     * Save PhpPresentation to file.
     */
    public function save(string $pFilename): void
    {
        if (empty($pFilename)) {
            throw new InvalidParameterException('pFilename', '');
        }

        file_put_contents($pFilename, $this->getHtmlContent());
    }

    protected function getHtmlContent(): string
    {
        $presentation = $this->getPhpPresentation();
        $this->ratioX = $presentation->getLayout()->getCX(DocumentLayout::UNIT_PIXEL);
        $this->ratioY = $presentation->getLayout()->getCY(DocumentLayout::UNIT_PIXEL);

        // Style
        $this->writeCSS();

        // Slides
        $slides = $presentation->getAllSlides();
        $numSlides = count($slides);
        foreach ($slides as $key => $slide) {
            $this->writeSlide($slide, $key, $numSlides);
        }

        return  '<html>'
            . '<head>'
            . '<meta charset="utf-8">'
            . '<style>' . $this->style . '</style>'
            . '</head>'
            . '<body>' . $this->bodyList . $this->bodySlides . '</body>'
            . '</html>';
    }

    protected function writeCSS(): void
    {
        $this->style = '* {box-sizing: border-box;}
      body {width: 100%;height: 100%;margin: 0;}
      .slideItem {width: 100%;height: 100%;}
      .slideItem .content {width: 100%;height: 100%;position: relative;top: 0;left: 0;overflow: hidden;}

      .slideItem#slidePage0 .navigation {display: block;}
      .slideItem#slidePage0 .content {opacity: 1;}

      .slideList:target ~ .slideItem .navigation {display: none !important;}
      .slideList:target ~ .slideItem .content {opacity: 0 !important;}
      .slideList[id="slide0"]:target ~ .slideItem#slidePage0 .navigation,
      .slideList[id="slide1"]:target ~ .slideItem#slidePage1 .navigation,
      .slideList[id="slide2"]:target ~ .slideItem#slidePage2 .navigation,
      .slideList[id="slide3"]:target ~ .slideItem#slidePage3 .navigation,
      .slideList[id="slide4"]:target ~ .slideItem#slidePage4 .navigation {
        display: block !important;
      }';

        if ($this->isPDF) {
            // Style PDF Output
            $this->style .= PHP_EOL
            . '.slideItem { page-break-after: always; position: relative;}' . PHP_EOL
            . '.slideItem .content { opacity: 1; }' . PHP_EOL;
        } else {
            // Style HTML Output
            $this->style .= PHP_EOL
            . 'body { overflow: hidden; }' . PHP_EOL
            . '.slideItem { position: absolute;}' . PHP_EOL
            . '.slideItem .content { opacity: 0; }' . PHP_EOL
            // Navigation
            . '.slideItem .navigation {display: none;}
                .slideItem .navigation a {text-decoration: none}
                .navigation {position: absolute;z-index: 5;bottom: 30px;right: 30px;font-size: 60px;}
                .navigation .next, .navigation .prev {color: #7dbeeb;}
                .navigation .disabled {color: #0a629f;}
                .navigation .next::after {content: "⮞";}
                .navigation .prev {margin-right: 0.3em;}
                .navigation .prev::after {content: "⮜";}' . PHP_EOL
            // Animations
            . '.slideList[id="slide0"]:target ~ .slideItem#slidePage0 .content,
                .slideList[id="slide1"]:target ~ .slideItem#slidePage1 .content,
                .slideList[id="slide2"]:target ~ .slideItem#slidePage2 .content,
                .slideList[id="slide3"]:target ~ .slideItem#slidePage3 .content,
                .slideList[id="slide4"]:target ~ .slideItem#slidePage4 .content {
                  animation-name: fade_in;
                  animation-duration: 0.5s;
                  opacity: 1 !important;
                }' . PHP_EOL
            . '@keyframes fade_in {from {opacity: 0;} to {opacity: 1;}}' . PHP_EOL;
        }
    }

    protected function writeSlide(Slide $slide, int $numSlideCurrent, int $numSlides): void
    {
        // PDF : No need navigation
        if (!$this->isPDF) {
            $this->bodyList .= '<div id="slide' . $numSlideCurrent . '" class="slideList"></div>';
        }
        $this->bodySlides .= '<div id="slidePage' . $numSlideCurrent . '" class="slideItem">';

        // PDF : No need navigation
        if (!$this->isPDF) {
            if ($numSlides >= 1) {
                $this->bodySlides .= '<div class="navigation">';
                if (($numSlideCurrent - 1) >= 0) {
                    $this->bodySlides .= '<a href="#slide' . ($numSlideCurrent - 1) . '" class="prev a' . $numSlideCurrent . '"></a>';
                }
                if (($numSlideCurrent + 1) < $numSlides) {
                    $this->bodySlides .= '<a href="#slide' . ($numSlideCurrent + 1) . '" class="next a' . $numSlideCurrent . '"></a>';
                }
                $this->bodySlides .= '</div>';
            }
        }

        $this->bodySlides .= '<div class="content">';
        foreach ($slide->getShapeCollection() as $shape) {
            if ($shape instanceof Media) {
                $this->writeVideo($shape);

                continue;
            }
            if ($shape instanceof Drawing\AbstractDrawingAdapter) {
                $this->writeImage($shape);

                continue;
            }
            if ($shape instanceof RichText) {
                $this->writeRichText($shape);

                continue;
            }
            if ($shape instanceof Table) {
                $this->writeTable($shape);

                continue;
            }
        }
        $this->bodySlides .= '</div></div>';
    }

    protected function writeImage(Drawing\AbstractDrawingAdapter $shape): void
    {
        $imageData = 'data:' . $shape->getMimeType() . ';base64,' . base64_encode($shape->getContents());

        $styles = [];
        $styles[] = 'position: absolute';
        $styles[] = 'width: ' . $shape->getWidth() . 'px';
        $styles[] = 'height: ' . $shape->getHeight() . 'px';
        $styles[] = 'top: ' . $shape->getOffsetY() . 'px';
        $styles[] = 'left: ' . $shape->getOffsetX() . 'px';
        $styles = array_merge($styles, $this->getStyleShadow($shape->getShadow()));

        $this->bodySlides .= '<img src="' . $imageData . '" style="' . implode(';', $styles) . '" alt="' . $shape->getDescription() . '" title="' . $shape->getDescription() . '" />';
    }

    protected function writeRichText(RichText $shape): void
    {
        $styles = [];
        $styles[] = 'position: absolute';
        $styles[] = 'width: ' . $shape->getWidth() . 'px';
        $styles[] = 'height: ' . $shape->getHeight() . 'px';
        $styles[] = 'top: ' . $shape->getOffsetY() . 'px';
        $styles[] = 'left: ' . $shape->getOffsetX() . 'px';

        $this->bodySlides .= '<div style="' . implode(';', $styles) . '">';
        foreach ($shape->getParagraphs() as $paragraph) {
            if ($paragraph instanceof RichText\Paragraph) {
                $this->writeRichTextParagraph($paragraph);

                continue;
            }
        }
        $this->bodySlides .= '</div>';
    }

    protected function writeTable(Table $shape): void
    {
        $styles = [];
        $styles[] = 'position: absolute';
        $styles[] = 'width: ' . $shape->getWidth() . 'px';
        $styles[] = 'height: ' . $shape->getHeight() . 'px';
        $styles[] = 'top: ' . $shape->getOffsetY() . 'px';
        $styles[] = 'left: ' . $shape->getOffsetX() . 'px';

        $this->bodySlides .= '<div style="' . implode(';', $styles) . '">';
        $this->bodySlides .= '<table style="width:100%"><tbody>';
        foreach ($shape->getRows() as $row) {
            $this->bodySlides .= '<tr>';
            foreach ($row->getCells() as $cell) {
                $this->bodySlides .= '<td';
                if ($cell->getColspan() > 1) {
                    $this->bodySlides .= ' colspan="' . $cell->getColspan() . '"';
                }
                $this->bodySlides .= '>';
                foreach ($cell->getParagraphs() as $paragraph) {
                    if ($paragraph instanceof RichText\Paragraph) {
                        $this->writeRichTextParagraph($paragraph);

                        continue;
                    }
                }
                $this->bodySlides .= '</td>';
            }
            $this->bodySlides .= '</tr>';
        }
        $this->bodySlides .= '</tbody></table>';
        $this->bodySlides .= '</div>';
    }

    protected function writeVideo(Media $shape): void
    {
        $imageData = 'data:' . $shape->getMimeType() . ';base64,' . base64_encode($shape->getContents());

        $styles = [];
        $styles[] = 'position: absolute';
        $styles[] = 'width: ' . $shape->getWidth() . 'px';
        $styles[] = 'height: ' . $shape->getHeight() . 'px';
        $styles[] = 'top: ' . $shape->getOffsetY() . 'px';
        $styles[] = 'left: ' . $shape->getOffsetX() . 'px';
        $styles = array_merge($styles, $this->getStyleShadow($shape->getShadow()));

        $this->bodySlides .= '<video controls style="' . implode(';', $styles) . '">'
          . '<source type ="' . $shape->getMimeType() . '" src="' . $imageData . '" />'
          . '</video>';
    }

    protected function writeRichTextParagraph(RichText\Paragraph $paragraph): void
    {
        $styles = array_merge([], $this->getStyleAlignment($paragraph->getAlignment()));
        $this->bodySlides .= '<div style="' . implode(';', $styles) . '">';
        foreach ($paragraph->getRichTextElements() as $richTextElement) {
            if ($richTextElement instanceof RichText\Run) {
                if ($richTextElement->hasHyperlink() && '' != $richTextElement->getHyperlink()->getUrl()) {
                    $this->bodySlides .= '<a href="' . $richTextElement->getHyperlink()->getUrl() . '">';
                }
                $styles = array_merge([], $this->getStyleFont($richTextElement->getFont()));
                $this->bodySlides .= '<span style="' . implode(';', $styles) . '">';
                $this->bodySlides .= $richTextElement->getText();
                $this->bodySlides .= '</span>';
                if ($richTextElement->hasHyperlink() && '' != $richTextElement->getHyperlink()->getUrl()) {
                    $this->bodySlides .= '</a>';
                }

                continue;
            }
            if ($richTextElement instanceof RichText\BreakElement) {
                $this->bodySlides .= '<br />';

                continue;
            }
        }
        $this->bodySlides .= '</div>';
    }

    /**
     * @return array<string>
     */
    protected function getStyleAlignment(Alignment $style): array
    {
        $styles = [];
        $styles[] = 'text-align: ' . ($style->getHorizontal() === Alignment::HORIZONTAL_CENTER ? 'center' : ($style->getHorizontal() === Alignment::HORIZONTAL_RIGHT ? 'right' : 'left'));

        return $styles;
    }

    /**
     * @return array<string>
     */
    protected function getStyleFont(Font $style): array
    {
        $styles = [];
        if ($style->isBold()) {
            $styles[] = 'font-weight: bold';
        }
        $styles[] = 'font-family: ' . $style->getName();
        $styles[] = 'font-size: ' . $style->getSize() . 'px';
        $styles[] = 'color: #' . $style->getColor()->getRGB();

        return $styles;
    }

    /**
     * @return array<string>
     */
    protected function getStyleShadow(Shadow $style): array
    {
        if (!$style->isVisible()) {
            return [];
        }
        $boxShadow = '';
        switch($style->getDirection()) {
            case 45:
                $boxShadow = $style->getDistance() . 'px ' . $style->getDistance() . 'px';

                break;
            case 90:
                $boxShadow = $style->getDistance() . 'px 0px';

                break;
            case 135:
                $boxShadow = '-' . $style->getDistance() . 'px ' . $style->getDistance() . 'px';

                break;
            case 180:
                $boxShadow = '-' . $style->getDistance() . 'px 0px';

                break;
            case 225:
                $boxShadow = '-' . $style->getDistance() . 'px -' . $style->getDistance() . 'px';

                break;
            case 270:
                $boxShadow = '-' . $style->getDistance() . 'px 0px';

                break;
            case 315:
                $boxShadow = $style->getDistance() . 'px -' . $style->getDistance() . 'px';

                break;
        }

        $styles = [];
        if ($boxShadow) {
            $styles[] = 'box-shadow: ' . $boxShadow . ' ' . $style->getBlurRadius() . 'px  #' . $style->getColor()->getRGB();
        }
        $styles[] = 'opacity: ' . ($style->getColor()->getAlpha() / 100);

        return $styles;
    }
}
