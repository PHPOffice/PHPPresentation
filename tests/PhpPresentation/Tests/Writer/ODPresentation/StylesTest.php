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

namespace PhpOffice\PhpPresentation\Tests\Writer\ODPresentation;

use PhpOffice\PhpPresentation\DocumentLayout;
use PhpOffice\PhpPresentation\Slide\Background\Color as BackgroundColor;
use PhpOffice\PhpPresentation\Slide\Background\Image as BackgroundImage;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Tests\PhpPresentationTestCase;

/**
 * Test class for PhpOffice\PhpPresentation\Writer\ODPresentation.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Writer\ODPresentation
 */
class StylesTest extends PhpPresentationTestCase
{
    protected $writerName = 'ODPresentation';

    public function testDocumentLayout(): void
    {
        $element = '/office:document-styles/office:automatic-styles/style:page-layout/style:page-layout-properties';

        $oDocumentLayout = new DocumentLayout();
        $oDocumentLayout->setDocumentLayout(DocumentLayout::LAYOUT_A4, true);
        $this->oPresentation->setLayout($oDocumentLayout);

        $this->assertZipXmlElementExists('styles.xml', $element);
        $this->assertZipXmlAttributeEquals('styles.xml', $element, 'style:print-orientation', 'landscape');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oDocumentLayout->setDocumentLayout(DocumentLayout::LAYOUT_A4, false);
        $this->oPresentation->setLayout($oDocumentLayout);
        $this->resetPresentationFile();

        $this->assertZipXmlElementExists('styles.xml', $element);
        $this->assertZipXmlAttributeEquals('styles.xml', $element, 'style:print-orientation', 'portrait');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testCustomDocumentLayout(): void
    {
        $oDocumentLayout = new DocumentLayout();
        $oDocumentLayout->setDocumentLayout(['cx' => mt_rand(1, 100), 'cy' => mt_rand(1, 100)]);
        $this->oPresentation->setLayout($oDocumentLayout);

        $element = '/office:document-styles/office:master-styles/style:master-page';
        $this->assertZipXmlElementExists('styles.xml', $element);

        // the page layout the master page actually names, whatever it was called
        $pageLayoutName = $this->getZipXmlAttributeValue('styles.xml', $element, 'style:page-layout-name');
        $this->assertZipXmlElementExists(
            'styles.xml',
            '/office:document-styles/office:automatic-styles/style:page-layout[@style:name=\'' . $pageLayoutName . '\']'
        );

        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testMasterSlideBackgroundColor(): void
    {
        $element = $this->getMasterPageStyleXPath() . '/style:drawing-page-properties';

        // The style the master page is drawn with carries no fill until one is asked for
        $this->assertZipXmlElementExists('styles.xml', $element);
        $this->assertZipXmlAttributeNotExists('styles.xml', $element, 'draw:fill');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oBackground = new BackgroundColor();
        $oBackground->setColor(new Color('FFCC00'));
        $this->oPresentation->getAllMasterSlides()[0]->setBackground($oBackground);
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeEquals('styles.xml', $element, 'draw:fill', 'solid');
        $this->assertZipXmlAttributeEquals('styles.xml', $element, 'draw:fill-color', '#FFCC00');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testMasterSlideBackgroundImage(): void
    {
        $imagePath = PHPPRESENTATION_TESTS_BASE_DIR . DIRECTORY_SEPARATOR . 'resources' . DIRECTORY_SEPARATOR . 'images' . DIRECTORY_SEPARATOR . 'PhpPresentationLogo.png';

        $oBackground = new BackgroundImage();
        $oBackground->setPath($imagePath);
        $this->oPresentation->getAllMasterSlides()[0]->setBackground($oBackground);

        $element = $this->getMasterPageStyleXPath() . '/style:drawing-page-properties';
        $this->assertZipXmlAttributeEquals('styles.xml', $element, 'draw:fill', 'bitmap');

        // The image the style names has to be in the package, and in the manifest with it
        $imageName = $this->getZipXmlAttributeValue('styles.xml', $element, 'draw:fill-image-name');
        $element = '/office:document-styles/office:styles/draw:fill-image[@draw:name=\'' . $imageName . '\']';
        $imageHref = $this->getZipXmlAttributeValue('styles.xml', $element, 'xlink:href');
        $this->assertZipFileExists($imageHref);
        $this->assertZipXmlElementExists('META-INF/manifest.xml', '/manifest:manifest/manifest:file-entry[@manifest:full-path=\'' . $imageHref . '\']');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testPatternFillHatch(): void
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createRichTextShape();
        $oShape->getFill()->setFillType(Fill::FILL_PATTERN_WDDNDIAG)
            ->setStartColor(new Color('FF4472C4'))
            ->setEndColor(new Color('FFFFFFFF'));

        // the hatch defined in `styles.xml` is the one `content.xml` names
        $element = '//style:graphic-properties[@draw:fill-hatch-name]';
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:fill', 'hatch');
        // The ground the lines are drawn on, which is painted only because this says so
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:fill-hatch-solid', 'true');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:fill-color', '#FFFFFF');
        $hatchName = $this->getZipXmlAttributeValue('content.xml', $element, 'draw:fill-hatch-name');

        $element = '/office:document-styles/office:styles/draw:hatch';
        $this->assertZipXmlAttributeEquals('styles.xml', $element, 'draw:name', $hatchName);
        $this->assertZipXmlAttributeEquals('styles.xml', $element, 'draw:style', 'single');
        $this->assertZipXmlAttributeEquals('styles.xml', $element, 'draw:color', '#4472C4');
        $this->assertZipXmlAttributeEquals('styles.xml', $element, 'draw:rotation', '315');
        $this->assertZipXmlAttributeEquals('styles.xml', $element, 'draw:distance', '0.2cm');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testPatternFillWithoutHatch(): void
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createRichTextShape();
        // Confetti is not a family of lines, and a hatch is all ODF has: the shape is painted in
        // the colour of the pattern rather than left with no fill at all.
        $oShape->getFill()->setFillType(Fill::FILL_PATTERN_SMCONFETTI)
            ->setStartColor(new Color('FFFF0000'))
            ->setEndColor(new Color('FF00FF00'));

        $this->assertZipXmlElementNotExists('styles.xml', '/office:document-styles/office:styles/draw:hatch');
        $element = '//style:style[@style:family=\'graphic\']/style:graphic-properties';
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:fill', 'solid');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:fill-color', '#FF0000');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testPatternFillTable(): void
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createTableShape();
        $oRow = $oShape->createRow();
        $oCell = $oRow->getCell();
        $oCell->getFill()->setFillType(Fill::FILL_PATTERN_LGGRID)
            ->setStartColor(new Color('FFFF7700'))
            ->setEndColor(new Color('FFFFFFFF'));

        // the hatch defined in `styles.xml` is the one `content.xml` names
        $hatchName = $this->getZipXmlAttributeValue(
            'content.xml',
            '//style:graphic-properties[@draw:fill-hatch-name]',
            'draw:fill-hatch-name'
        );
        $element = '/office:document-styles/office:styles/draw:hatch';
        $this->assertZipXmlAttributeEquals('styles.xml', $element, 'draw:name', $hatchName);
        $this->assertZipXmlAttributeEquals('styles.xml', $element, 'draw:style', 'double');

        $this->assertIsSchemaOpenDocumentNotValid('1.2');
    }

    public function testGradientTable(): void
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createTableShape();
        $oRow = $oShape->createRow();
        $oCell = $oRow->getCell();
        $oCell->getFill()->setFillType(Fill::FILL_GRADIENT_LINEAR)->setStartColor(new Color('FFFF7700'))->setEndColor(new Color('FFFFFFFF'));

        // the gradient defined in `styles.xml` is the one `content.xml` names
        $gradientName = $this->getZipXmlAttributeValue(
            'content.xml',
            '//style:graphic-properties[@draw:fill-gradient-name]',
            'draw:fill-gradient-name'
        );
        $element = '/office:document-styles/office:styles/draw:gradient';
        $this->assertZipXmlAttributeEquals('styles.xml', $element, 'draw:name', $gradientName);

        $this->assertIsSchemaOpenDocumentNotValid('1.2');
    }

    public function testGradientTableRow(): void
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oShape = $oSlide->createTableShape();
        $oRow = $oShape->createRow();

        // Neither the cell nor its row asks for a fill, so there is no gradient to define
        $this->assertZipXmlElementNotExists('styles.xml', '/office:document-styles/office:styles/draw:gradient');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        // The gradient of a row is defined once the row asks for one, exactly as a cell's is
        $oRow->getFill()->setFillType(Fill::FILL_GRADIENT_LINEAR)->setStartColor(new Color('FFFF7700'))->setEndColor(new Color('FFFFFFFF'));
        $this->resetPresentationFile();

        // the gradient defined in `styles.xml` is the one `content.xml` names
        $gradientName = $this->getZipXmlAttributeValue(
            'content.xml',
            '//style:graphic-properties[@draw:fill-gradient-name]',
            'draw:fill-gradient-name'
        );
        $element = '/office:document-styles/office:styles/draw:gradient';
        $this->assertZipXmlAttributeEquals('styles.xml', $element, 'draw:name', $gradientName);
        $this->assertIsSchemaOpenDocumentNotValid('1.2');
    }

    public function testStrokeDash(): void
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oRichText1 = $oSlide->createRichTextShape();
        $oRichText1->getBorder()->setColor(new Color('FF4672A8'))->setLineStyle(Border::LINE_SINGLE);
        $arrayDashStyle = [
            Border::DASH_DASH,
            Border::DASH_DASHDOT,
            Border::DASH_DOT,
            Border::DASH_LARGEDASH,
            Border::DASH_LARGEDASHDOT,
            Border::DASH_LARGEDASHDOTDOT,
            Border::DASH_SYSDASH,
            Border::DASH_SYSDASHDOT,
            Border::DASH_SYSDASHDOTDOT,
            Border::DASH_SYSDOT,
        ];

        foreach ($arrayDashStyle as $style) {
            $oRichText1->getBorder()->setDashStyle($style);

            $element = '/office:document-styles/office:styles/draw:stroke-dash[@draw:name=\'strokeDash_' . $style . '\']';
            $this->assertZipXmlElementExists('styles.xml', $element);
            $this->assertZipXmlAttributeEquals('styles.xml', $element, 'draw:style', 'rect');
            $this->assertZipXmlAttributeExists('styles.xml', $element, 'draw:distance');

            switch ($style) {
                case Border::DASH_DOT:
                case Border::DASH_SYSDOT:
                    $this->assertZipXmlAttributeExists('styles.xml', $element, 'draw:dots1');
                    $this->assertZipXmlAttributeExists('styles.xml', $element, 'draw:dots1-length');

                    break;
                case Border::DASH_DASH:
                case Border::DASH_LARGEDASH:
                case Border::DASH_SYSDASH:
                    $this->assertZipXmlAttributeExists('styles.xml', $element, 'draw:dots2');
                    $this->assertZipXmlAttributeExists('styles.xml', $element, 'draw:dots2-length');

                    break;
                case Border::DASH_DASHDOT:
                case Border::DASH_LARGEDASHDOT:
                case Border::DASH_LARGEDASHDOTDOT:
                case Border::DASH_SYSDASHDOT:
                case Border::DASH_SYSDASHDOTDOT:
                    $this->assertZipXmlAttributeExists('styles.xml', $element, 'draw:dots1');
                    $this->assertZipXmlAttributeExists('styles.xml', $element, 'draw:dots1-length');
                    $this->assertZipXmlAttributeExists('styles.xml', $element, 'draw:dots2');
                    $this->assertZipXmlAttributeExists('styles.xml', $element, 'draw:dots2-length');

                    break;
            }
            $this->assertIsSchemaOpenDocumentValid('1.2');
            $this->resetPresentationFile();
        }
    }

    /**
     * The `style:style` of the master page, addressed by the name its `style:master-page` carries.
     *
     * The name is a constant in the writer rather than a generated one, but the hole is the same:
     * with `draw:style-name` broken and every definition still correctly named, a spelled-out name
     * passes and a resolved one does not.
     */
    private function getMasterPageStyleXPath(): string
    {
        $styleName = $this->getZipXmlAttributeValue(
            'styles.xml',
            '/office:document-styles/office:master-styles/style:master-page',
            'draw:style-name'
        );

        return '/office:document-styles/office:automatic-styles/style:style[@style:name=\'' . $styleName . '\']';
    }
}
