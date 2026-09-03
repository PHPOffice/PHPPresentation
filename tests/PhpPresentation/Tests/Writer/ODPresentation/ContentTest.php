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

use PhpOffice\Common\Drawing;
use PhpOffice\Common\Drawing as CommonDrawing;
use PhpOffice\Common\Text;
use PhpOffice\PhpPresentation\PresentationProperties;
use PhpOffice\PhpPresentation\Shape\Chart\Series;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Line as ChartTypeLine;
use PhpOffice\PhpPresentation\Shape\Comment;
use PhpOffice\PhpPresentation\Shape\Group;
use PhpOffice\PhpPresentation\Shape\Media;
use PhpOffice\PhpPresentation\Shape\Placeholder;
use PhpOffice\PhpPresentation\Shape\RichText\Field;
use PhpOffice\PhpPresentation\Shape\RichText\Paragraph;
use PhpOffice\PhpPresentation\Shape\RichText\Run;
use PhpOffice\PhpPresentation\Slide\Transition;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Bullet;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Font;
use PhpOffice\PhpPresentation\Tests\PhpPresentationTestCase;
use PHPUnit\Framework\Attributes\DataProvider;
use ReflectionClass;

/**
 * Test class for PhpOffice\PhpPresentation\Writer\ODPresentation\Manifest.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Writer\ODPresentation\Manifest
 */
class ContentTest extends PhpPresentationTestCase
{
    protected $writerName = 'ODPresentation';

    public function testComment(): void
    {
        $expectedName = 'Name';
        $expectedText = 'Text';

        $oAuthor = new Comment\Author();
        $oAuthor->setName($expectedName);
        $oComment = new Comment();
        $oComment->setAuthor($oAuthor);
        $oComment->setText($expectedText);
        $this->oPresentation->getActiveSlide()->addShape($oComment);

        $element = '/office:document-content';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeExists('content.xml', $element, 'xmlns:officeooo');
        $element = '/office:document-content/office:body/office:presentation/draw:page/officeooo:annotation';
        $this->assertZipXmlElementExists('content.xml', $element);
        $element = '/office:document-content/office:body/office:presentation/draw:page/officeooo:annotation/dc:creator';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlElementEquals('content.xml', $element, $expectedName);
        $element = '/office:document-content/office:body/office:presentation/draw:page/officeooo:annotation/text:p';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlElementEquals('content.xml', $element, $expectedText);
        $this->assertIsSchemaOpenDocumentNotValid('1.2');
    }

    public function testCommentWithoutAuthor(): void
    {
        $oComment = new Comment();
        $this->oPresentation->getActiveSlide()->addShape($oComment);

        $element = '/office:document-content/office:body/office:presentation/draw:page/officeooo:annotation';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'dc:creator');
        $this->assertIsSchemaOpenDocumentNotValid('1.2');
    }

    public function testShapeDecorative(): void
    {
        $oSlide = $this->oPresentation->getActiveSlide();

        $oRichText = $oSlide->createRichTextShape();
        $oRichText->createTextRun('AAA');

        $oLine = $oSlide->createLineShape(10, 10, 100, 100);
        $oLine->setDecorative();

        $basePath = '/office:document-content/office:body/office:presentation/draw:page';
        $this->assertZipXmlAttributeNotExists('content.xml', $basePath . '/draw:frame', 'loext:decorative');
        $this->assertZipXmlAttributeEquals('content.xml', $basePath . '/draw:line', 'loext:decorative', 'true');
        // Invalid because `loext:decorative` is a LibreOffice extension, standardized in ODF 1.4
        $this->assertIsSchemaOpenDocumentNotValid('1.2');
    }

    public function testShapeDescription(): void
    {
        $oSlide = $this->oPresentation->getActiveSlide();

        $oRichText = $oSlide->createRichTextShape();
        $oRichText->createTextRun('AAA');
        $oRichText->setDescription('RichText Alternative Text');

        $oLine = $oSlide->createLineShape(10, 10, 100, 100);
        $oLine->setDescription('Line Alternative Text');

        $oGroup = new Group();
        $oGroup->setDescription('Group Alternative Text');
        $oGroup->createRichTextShape()->createTextRun('BBB');
        $oSlide->addShape($oGroup);

        $basePath = '/office:document-content/office:body/office:presentation/draw:page';
        $this->assertZipXmlElementEquals('content.xml', $basePath . '/draw:frame/svg:desc', 'RichText Alternative Text');
        $this->assertZipXmlElementEquals('content.xml', $basePath . '/draw:line/svg:desc', 'Line Alternative Text');
        $this->assertZipXmlElementEquals('content.xml', $basePath . '/draw:g/svg:desc', 'Group Alternative Text');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testShapeDescriptionOmittedWhenEmpty(): void
    {
        $oRichText = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oRichText->createTextRun('AAA');

        $this->assertZipXmlElementNotExists(
            'content.xml',
            '/office:document-content/office:body/office:presentation/draw:page/draw:frame/svg:desc'
        );
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    /**
     * A hyperlink given to a shape as a whole, which every kind of shape accepts and only a
     * picture used to write. Each shape here carries a description as well, so that the schema
     * check below sees both children and pins their order: `office:event-listeners` precedes
     * `svg:desc` inside a `draw:frame` and follows it inside `draw:line` and `draw:g`.
     */
    public function testShapeHyperlink(): void
    {
        $expectedUrl = 'https://github.com/PHPOffice/PHPPresentation/';
        $oSlide = $this->oPresentation->getActiveSlide();

        $oRichText = $oSlide->createRichTextShape();
        $oRichText->createTextRun('AAA');

        $oTable = $oSlide->createTableShape();
        $oTable->createRow();

        $oChart = $oSlide->createChartShape();
        $oChart->getPlotArea()->setType((new ChartTypeLine())->addSeries(new Series('Serie', ['A' => '1'])));

        $oLine = $oSlide->createLineShape(10, 10, 100, 100);

        $oGroup = new Group();
        $oGroup->createRichTextShape()->createTextRun('BBB');
        $oSlide->addShape($oGroup);

        $oMedia = new Media();
        $oMedia->setPath(PHPPRESENTATION_TESTS_BASE_DIR . '/resources/videos/sintel_trailer-480p.ogv');
        $oSlide->addShape($oMedia);

        $basePath = '/office:document-content/office:body/office:presentation/draw:page';
        $expectedShapes = [
            $basePath . '/draw:frame[1]' => $oRichText,
            $basePath . '/draw:frame[2]' => $oTable,
            $basePath . '/draw:frame[3]' => $oChart,
            $basePath . '/draw:line' => $oLine,
            $basePath . '/draw:g' => $oGroup,
            $basePath . '/draw:frame[4]' => $oMedia,
        ];
        foreach ($expectedShapes as $shape) {
            $shape->setDescription('Alternative Text');
            $shape->getHyperlink()->setUrl($expectedUrl);
        }

        foreach ($expectedShapes as $shapePath => $shape) {
            $element = $shapePath . '/office:event-listeners/presentation:event-listener';
            $this->assertZipXmlElementExists('content.xml', $element);
            $this->assertZipXmlAttributeEquals('content.xml', $element, 'script:event-name', 'dom:click');
            $this->assertZipXmlAttributeEquals('content.xml', $element, 'presentation:action', 'show');
            $this->assertZipXmlAttributeEquals('content.xml', $element, 'xlink:href', $expectedUrl);
        }
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    /**
     * A shape given no hyperlink writes no listener, which is what keeps the element optional
     * rather than empty.
     */
    public function testShapeHyperlinkOmittedWhenAbsent(): void
    {
        $oRichText = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oRichText->createTextRun('AAA');

        $this->assertZipXmlElementNotExists(
            'content.xml',
            '/office:document-content/office:body/office:presentation/draw:page/draw:frame/office:event-listeners'
        );
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    /**
     * A hyperlink to another slide, which ODF addresses by the name of the page rather than by
     * the PowerPoint action string the model stores.
     */
    public function testShapeHyperlinkToASlide(): void
    {
        $oRichText = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oRichText->createTextRun('AAA');
        $oRichText->getHyperlink()->setSlideNumber(1);

        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/office:event-listeners/presentation:event-listener';
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'xlink:href', '#Slide 1');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testDrawingMimetype(): void
    {
        $oShape = $this->oPresentation->getActiveSlide()->createDrawingShape();
        $oShape->setPath(PHPPRESENTATION_TESTS_BASE_DIR . '/resources/images/PhpPresentationLogo.png');

        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:image';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeExists('content.xml', $element, 'loext:mime-type');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'loext:mime-type', 'image/png');
        // Invalid because `draw:image` has attribute `loext:mime-type`
        $this->assertIsSchemaOpenDocumentNotValid('1.2');

        $this->resetPresentationFile();
        $oShape->setPath(PHPPRESENTATION_TESTS_BASE_DIR . '/resources/images/tiger.svg');

        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:image';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeExists('content.xml', $element, 'loext:mime-type');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'loext:mime-type', 'image/svg+xml');
        // Invalid because `draw:image` has attribute `loext:mime-type`
        $this->assertIsSchemaOpenDocumentNotValid('1.2');
    }

    public function testDrawingShapeFill(): void
    {
        $oShape = $this->oPresentation->getActiveSlide()->createDrawingShape();
        $oShape->setPath(PHPPRESENTATION_TESTS_BASE_DIR . '/resources/images/PhpPresentationLogo.png');

        $element = '/office:document-content/office:automatic-styles/style:style/style:graphic-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:fill', 'none');
        // Invalid because `draw:image` has attribute `loext:mime-type`
        $this->assertIsSchemaOpenDocumentNotValid('1.2');

        $oColor = new Color(Color::COLOR_DARKRED);
        $oColor->setAlpha(mt_rand(0, 100));
        $oShape->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor($oColor);
        $this->resetPresentationFile();

        $element = '/office:document-content/office:automatic-styles/style:style/style:graphic-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:fill', 'solid');
        $this->assertZipXmlAttributeStartsWith('content.xml', $element, 'draw:fill-color', '#');
        $this->assertZipXmlAttributeEndsWith('content.xml', $element, 'draw:fill-color', $oColor->getRGB());
        // Invalid because `draw:image` has attribute `loext:mime-type`
        $this->assertIsSchemaOpenDocumentNotValid('1.2');
    }

    public function testDrawingWithHyperlink(): void
    {
        $oShape = $this->oPresentation->getActiveSlide()->createDrawingShape();
        $oShape->setPath(PHPPRESENTATION_TESTS_BASE_DIR . '/resources/images/PhpPresentationLogo.png');
        $oShape->getHyperlink()->setUrl('https://github.com/PHPOffice/PHPPresentation/');

        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/office:event-listeners/presentation:event-listener';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'xlink:href', 'https://github.com/PHPOffice/PHPPresentation/');
        // Invalid because `draw:image` has attribute `loext:mime-type`
        $this->assertIsSchemaOpenDocumentNotValid('1.2');
    }

    public function testFillSetBackToNull(): void
    {
        $oShape = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oShape->createTextRun('Nothing was said about my fill');
        $oShape->setFill(null);

        // Saving at all is the point: `getFill()` used to hand back the null it was given, and the
        // switch below dereferenced it. The style now simply names no `draw:fill`, so the shape
        // takes the one its parent style gives it.
        $element = '/office:document-content/office:automatic-styles/style:style/style:graphic-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'draw:fill');
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'draw:fill-color');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testFillGradientLinearRichText(): void
    {
        $oShape = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oShape->getFill()->setFillType(Fill::FILL_GRADIENT_LINEAR)->setStartColor(new Color('FFFF7700'))->setEndColor(new Color('FFFFFFFF'));

        $element = $this->getShapeStyleXPath() . '/style:graphic-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:fill', 'gradient');

        // the gradient the shape names is the gradient that `styles.xml` defines
        $gradientName = $this->getZipXmlAttributeValue('content.xml', $element, 'draw:fill-gradient-name');
        $this->assertZipXmlAttributeEquals(
            'styles.xml',
            '/office:document-styles/office:styles/draw:gradient',
            'draw:name',
            $gradientName
        );
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testFillSolidRichText(): void
    {
        $oShape = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oShape->getFill()->setFillType(Fill::FILL_SOLID)->setRotation(90)->setStartColor(new Color('FF4672A8'))->setEndColor(new Color('FF4672A8'));

        $element = $this->getShapeStyleXPath() . '/style:graphic-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:fill', 'solid');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:fill-color', '#' . $oShape->getFill()->getStartColor()->getRGB());
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:fill-color', '#' . $oShape->getFill()->getEndColor()->getRGB());
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testGroup(): void
    {
        $oShapeGroup = $this->oPresentation->getActiveSlide()->createGroup();
        $oShape = $oShapeGroup->createDrawingShape();
        $oShape->setPath(PHPPRESENTATION_TESTS_BASE_DIR . '/resources/images/PhpPresentationLogo.png');
        $oShape->getHyperlink()->setUrl('https://github.com/PHPOffice/PHPPresentation/');

        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:g';
        $this->assertZipXmlElementExists('content.xml', $element);
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:g/draw:frame/office:event-listeners/presentation:event-listener';
        $this->assertZipXmlElementExists('content.xml', $element);
        // Invalid because `draw:image` has attribute `loext:mime-type`
        $this->assertIsSchemaOpenDocumentNotValid('1.2');
    }

    public function testList(): void
    {
        $oRichText = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oRichText->getActiveParagraph()->getBulletStyle()->setBulletType(Bullet::TYPE_BULLET);
        $oRichText->createTextRun('Alpha');
        $oRichText->createParagraph()->createTextRun('Beta');
        $oRichText->createParagraph()->createTextRun('Delta');
        $oRichText->createParagraph()->createTextRun('Epsilon');

        $element = '/office:document-content/office:automatic-styles/text:list-style/text:list-level-style-bullet';
        $this->assertZipXmlElementExists('content.xml', $element);
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:text-box';
        $this->assertZipXmlElementExists('content.xml', $element);
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:text-box/text:list/text:list-item/text:p/text:span';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    /**
     * A paragraph whose bullet is numeric, which used to match neither branch of the writer and
     * so lost its text along with its marker.
     */
    public function testNumericBullet(): void
    {
        $oRichText = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oRichText->getActiveParagraph()->getBulletStyle()->setBulletType(Bullet::TYPE_NUMERIC);
        $oRichText->createTextRun('Alpha');

        $textBox = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:text-box';
        $this->assertZipXmlElementEquals('content.xml', $textBox . '/text:list/text:list-item/text:p/text:span', 'Alpha');
        $this->assertZipXmlElementExists(
            'content.xml',
            '/office:document-content/office:automatic-styles/text:list-style/text:list-level-style-number'
        );
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    /**
     * @return array<array<int, string>>
     */
    public static function dataProviderNumericBulletFormat(): array
    {
        return [
            // The five formats ODF names outright, and the punctuation it keeps apart from them
            [Bullet::NUMERIC_ARABICPERIOD, '1', '', '.'],
            [Bullet::NUMERIC_ARABICPLAIN, '1', '', ''],
            [Bullet::NUMERIC_ARABICPARENBOTH, '1', '(', ')'],
            [Bullet::NUMERIC_ALPHALCPARENR, 'a', '', ')'],
            [Bullet::NUMERIC_ALPHAUCPERIOD, 'A', '', '.'],
            [Bullet::NUMERIC_ROMANLCPERIOD, 'i', '', '.'],
            [Bullet::NUMERIC_ROMANUCPARENBOTH, 'I', '(', ')'],
            // The sequences LibreOffice names, measured against its own conversion of the same deck
            [Bullet::NUMERIC_CIRCLENUMDBPLAIN, "\u{2460}, \u{2461}, \u{2462}, ...", '', ''],
            [Bullet::NUMERIC_EA1CHSPERIOD, "\u{58F9}, \u{8D30}, \u{53C1}, ...", '', '.'],
            [Bullet::NUMERIC_HEBREW2MINUS, "\u{05D0}, \u{05D1}, \u{05D2}, ...", '', '-'],
            [Bullet::NUMERIC_THAINUMPARENR, "\u{0E01}, \u{0E02}, \u{0E04}, ...", '', ')'],
            // The schemes LibreOffice drops the numbering of: arabic, and the punctuation kept
            [Bullet::NUMERIC_HINDINUMPARENR, '1', '', ')'],
            [Bullet::NUMERIC_ARABICDBPLAIN, '1', '', ''],
            [Bullet::NUMERIC_ARABIC1MINUS, '1', '', '-'],
        ];
    }

    /**
     * @dataProvider dataProviderNumericBulletFormat
     */
    #[DataProvider('dataProviderNumericBulletFormat')]
    public function testNumericBulletFormat(string $scheme, string $expectedFormat, string $expectedPrefix, string $expectedSuffix): void
    {
        $oRichText = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oRichText->getActiveParagraph()->getBulletStyle()
            ->setBulletType(Bullet::TYPE_NUMERIC)
            ->setBulletNumericStyle($scheme);
        $oRichText->createTextRun('Alpha');

        $element = '/office:document-content/office:automatic-styles/text:list-style/text:list-level-style-number';
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'style:num-format', $expectedFormat);
        if ('' === $expectedPrefix) {
            $this->assertZipXmlAttributeNotExists('content.xml', $element, 'style:num-prefix');
        } else {
            $this->assertZipXmlAttributeEquals('content.xml', $element, 'style:num-prefix', $expectedPrefix);
        }
        if ('' === $expectedSuffix) {
            $this->assertZipXmlAttributeNotExists('content.xml', $element, 'style:num-suffix');
        } else {
            $this->assertZipXmlAttributeEquals('content.xml', $element, 'style:num-suffix', $expectedSuffix);
        }
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testNumericBulletStartAt(): void
    {
        $oRichText = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oRichText->getActiveParagraph()->getBulletStyle()
            ->setBulletType(Bullet::TYPE_NUMERIC)
            ->setBulletNumericStartAt(5);
        $oRichText->createTextRun('Alpha');

        $element = '/office:document-content/office:automatic-styles/text:list-style/text:list-level-style-number';
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'text:start-value', 5);
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testNumericBulletStartAtOmittedWhenFirst(): void
    {
        $oRichText = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oRichText->getActiveParagraph()->getBulletStyle()->setBulletType(Bullet::TYPE_NUMERIC);
        $oRichText->createTextRun('Alpha');

        $element = '/office:document-content/office:automatic-styles/text:list-style/text:list-level-style-number';
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'text:start-value');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    /**
     * A bullet list and a numbered list are two lists, not one: they name two styles, and a
     * `text:list` names exactly one.
     */
    public function testABulletListAndANumberedListAreTwoLists(): void
    {
        $oRichText = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oRichText->getActiveParagraph()->getBulletStyle()->setBulletType(Bullet::TYPE_BULLET);
        $oRichText->createTextRun('Alpha');
        $oParagraph = $oRichText->createParagraph();
        $oParagraph->getBulletStyle()->setBulletType(Bullet::TYPE_NUMERIC);
        $oParagraph->createTextRun('Beta');

        $textBox = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:text-box';
        $this->assertZipXmlElementEquals('content.xml', $textBox . '/text:list[1]/text:list-item/text:p/text:span', 'Alpha');
        $this->assertZipXmlElementEquals('content.xml', $textBox . '/text:list[2]/text:list-item/text:p/text:span', 'Beta');
        $this->assertZipXmlElementExists(
            'content.xml',
            '/office:document-content/office:automatic-styles/text:list-style/text:list-level-style-bullet'
        );
        $this->assertZipXmlElementExists(
            'content.xml',
            '/office:document-content/office:automatic-styles/text:list-style/text:list-level-style-number'
        );
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    /**
     * The list ends where the paragraphs that asked for a marker end. The paragraph that follows
     * one used to be written inside the last item of the list it did not ask to join, and only
     * the second plain paragraph in a row escaped it.
     */
    public function testListEndsBeforeAParagraphThatAsksForNoBullet(): void
    {
        $oRichText = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oRichText->getActiveParagraph()->getBulletStyle()->setBulletType(Bullet::TYPE_BULLET);
        $oRichText->createTextRun('Alpha');
        // `createParagraph()` clones the bullet style of the paragraph before it, so asking for
        // no bullet has to be said out loud
        $oParagraph = $oRichText->createParagraph();
        $oParagraph->getBulletStyle()->setBulletType(Bullet::TYPE_NONE);
        $oParagraph->createTextRun('Beta');

        $textBox = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:text-box';
        $this->assertZipXmlElementEquals('content.xml', $textBox . '/text:list/text:list-item/text:p/text:span', 'Alpha');
        $this->assertZipXmlElementEquals('content.xml', $textBox . '/text:p/text:span', 'Beta');
        $this->assertZipXmlElementNotExists('content.xml', $textBox . '/text:list/text:list-item/text:p[2]');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testListStyleIsNotWrittenWithoutAList(): void
    {
        $oRichText = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oRichText->createTextRun('Alpha');

        // A paragraph left at Bullet::TYPE_NONE is written as a plain `text:p`, so nothing in the
        // document names a list style and none is written for it
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:text-box/text:p';
        $this->assertZipXmlElementExists('content.xml', $element);
        $element = '/office:document-content/office:automatic-styles/text:list-style';
        $this->assertZipXmlElementNotExists('content.xml', $element);
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testListStyleIsWrittenOnlyForTheParagraphsThatUseOne(): void
    {
        $oRichText = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oRichText->getActiveParagraph()->getBulletStyle()->setBulletType(Bullet::TYPE_BULLET);
        $oRichText->createTextRun('Alpha');
        $oRichText->createParagraph()->createTextRun('Beta');

        $oPlain = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oPlain->createTextRun('Delta');

        // The bulleted shape keeps its style and the plain one adds none of its own
        $element = '/office:document-content/office:automatic-styles/text:list-style';
        $this->assertZipXmlElementCount('content.xml', $element, 1);

        // and the one that is written is the one a list names
        $element = '/office:document-content/office:automatic-styles/text:list-style[@style:name = //text:list/@text:style-name]';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testStyleIsSharedByEverythingThatWritesIt(): void
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        foreach (['Alpha', 'Beta', 'Delta'] as $text) {
            $oSlide->createRichTextShape()->createTextRun($text);
        }

        // Three shapes formatted alike hold different text, so they hash differently -- and write
        // the same paragraph style and the same text style, which is what names them
        $this->assertZipXmlElementCount(
            'content.xml',
            '/office:document-content/office:automatic-styles/style:style[@style:family=\'paragraph\']',
            1
        );
        $this->assertZipXmlElementCount(
            'content.xml',
            '/office:document-content/office:automatic-styles/style:style[@style:family=\'text\']',
            1
        );

        // and all three name the one that was written
        $expected = $this->getRunStyleXPath(1);
        self::assertSame($expected, $this->getRunStyleXPath(2));
        self::assertSame($expected, $this->getRunStyleXPath(3));
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testStyleIsNotSharedBetweenRunsThatDiffer(): void
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oSlide->createRichTextShape()->createTextRun('Alpha');
        $oSlide->createRichTextShape()->createTextRun('Beta')->getFont()->setBold(true);

        $this->assertZipXmlElementCount(
            'content.xml',
            '/office:document-content/office:automatic-styles/style:style[@style:family=\'text\']',
            2
        );
        self::assertNotSame($this->getRunStyleXPath(1), $this->getRunStyleXPath(2));
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testNestedGroupNamesStylesThatAreWritten(): void
    {
        $oInner = new Group();
        $oInner->createRichTextShape()->createTextRun('Nested');
        $oOuter = new Group();
        $oOuter->addShape($oInner);
        $this->oPresentation->getActiveSlide()->addShape($oOuter);

        // The shapes of a group inside a group are written by a recursive pass, and have to be
        // collected by one too, or they name styles that nothing defines
        $frame = '/office:document-content/office:body/office:presentation/draw:page/draw:g/draw:g/draw:frame';
        foreach ([
            $frame => 'draw:style-name',
            $frame . '/draw:text-box/text:p' => 'text:style-name',
            $frame . '/draw:text-box/text:p/text:span' => 'text:style-name',
        ] as $xPath => $attribute) {
            $styleName = $this->getZipXmlAttributeValue('content.xml', $xPath, $attribute);
            $this->assertZipXmlElementExists(
                'content.xml',
                '/office:document-content/office:automatic-styles/style:style[@style:name=\'' . $styleName . '\']'
            );
        }
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testSlideNoteNamesStylesThatAreWritten(): void
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oSlide->createRichTextShape()->createTextRun('Body');
        $oSlide->getNote()->createRichTextShape()->createTextRun('Note');

        // The note is written by a pass of its own, and its shapes have to be collected with the
        // rest or it names a paragraph style, a text style and a graphic style that nothing defines
        $noteText = '/office:document-content/office:body/office:presentation/draw:page'
            . '/presentation:notes/draw:frame/draw:text-box/text:p';
        foreach ([$noteText => 'text:style-name', $noteText . '/text:span' => 'text:style-name'] as $xPath => $attribute) {
            $styleName = $this->getZipXmlAttributeValue('content.xml', $xPath, $attribute);
            $this->assertZipXmlElementExists(
                'content.xml',
                '/office:document-content/office:automatic-styles/style:style[@style:name=\'' . $styleName . '\']'
            );
        }

        $frameStyle = $this->getZipXmlAttributeValue(
            'content.xml',
            '/office:document-content/office:body/office:presentation/draw:page/presentation:notes/draw:frame',
            'draw:style-name'
        );
        $this->assertZipXmlElementExists(
            'content.xml',
            '/office:document-content/office:automatic-styles/style:style[@style:name=\'' . $frameStyle . '\']'
        );
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testInnerList(): void
    {
        $oRichText = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oRichText->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT)->setMarginLeft(25)->setIndent(-25);
        $oRichText->getActiveParagraph()->getBulletStyle()->setBulletType(Bullet::TYPE_BULLET);

        $oRichText->createTextRun('Alpha');
        $oRichText->createParagraph()->getAlignment()->setLevel(1)->setMarginLeft(75)->setIndent(-25);
        $oRichText->createTextRun('Alpha.Alpha');
        $oRichText->createParagraph()->createTextRun('Alpha.Beta');
        $oRichText->createParagraph()->createTextRun('Alpha.Delta');

        $oRichText->createParagraph()->getAlignment()->setLevel(0)->setMarginLeft(25)->setIndent(-25);
        $oRichText->createTextRun('Beta');
        $oRichText->createParagraph()->getAlignment()->setLevel(1)->setMarginLeft(75)->setIndent(-25);
        $oRichText->createTextRun('Beta.Alpha');
        $oRichText->createParagraph()->createTextRun('Beta.Beta');
        $oRichText->createParagraph()->createTextRun('Beta.Delta');

        $element = '/office:document-content/office:automatic-styles/text:list-style/text:list-level-style-bullet';
        $this->assertZipXmlElementExists('content.xml', $element);
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:text-box';
        $this->assertZipXmlElementExists('content.xml', $element);
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:text-box/text:list/text:list-item/text:list/text:list-item/text:p/text:span';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testParagraphRichText(): void
    {
        $oRichText = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oRichText->createTextRun('Alpha');
        $oRichText->createBreak();
        $oRichText->createText('Beta');
        $oRichText->createBreak();
        $oRun = $oRichText->createTextRun('Delta');
        $oRun->getHyperlink()->setUrl('http://www.google.fr');

        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:text-box/text:p/text:span/text:line-break';
        $this->assertZipXmlElementExists('content.xml', $element);

        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:text-box/text:p/text:span/text:a';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'xlink:href', 'http://www.google.fr');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testListWithRichText(): void
    {
        $oRichText = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oRichText->getActiveParagraph()->getBulletStyle()->setBulletType(Bullet::TYPE_BULLET);
        $oRun = $oRichText->createTextRun('Alpha');
        $oRun->getHyperlink()->setUrl('http://www.google.fr');
        $oRichText->createBreak();
        $oRichText->createTextRun('Beta');

        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:text-box/text:list/text:list-item/text:p/text:span/text:a';
        $this->assertZipXmlElementExists('content.xml', $element);
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:text-box/text:list/text:list-item/text:p/text:span/text:line-break';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testLineShadow(): void
    {
        $this->oPresentation->getActiveSlide()->createLineShape(10, 10, 100, 100)
            ->getShadow()->setVisible(true)->setAlpha(75)->setDirection(45)->setDistance(10);

        $element = $this->getLineStyleXPath() . '/style:graphic-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:shadow', 'visible');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testShadowSetBackToNull(): void
    {
        // one of the two shapes this writer reads a shadow from: a text frame, by writeTxtStyle()
        $oSlide = $this->oPresentation->getActiveSlide();
        $oSlide->createRichTextShape()->setShadow(null);

        $this->assertZipFileExists('content.xml');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        // the other: a drawing, by writeDrawingStyle() -- which is what a video is written as too.
        // No schema assertion here: a `draw:image` this writer produces is not ODF 1.2 valid for a
        // reason of its own, which testDrawingMimetype() pins down.
        $oSlide->createDrawingShape()
            ->setPath(PHPPRESENTATION_TESTS_BASE_DIR . '/resources/images/PhpPresentationLogo.png')
            ->setShadow(null);
        $oMedia = new Media();
        $oMedia->setPath(PHPPRESENTATION_TESTS_BASE_DIR . '/resources/videos/sintel_trailer-480p.ogv')
            ->setShadow(null);
        $oSlide->addShape($oMedia);
        $this->resetPresentationFile();

        $this->assertZipFileExists('content.xml');
        $this->assertZipXmlElementExists('content.xml', '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:image');
        $this->assertZipXmlElementExists('content.xml', '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:plugin');
    }

    public function testMedia(): void
    {
        $expectedName = 'MyName';
        $expectedWidth = mt_rand(1, 100);
        $expectedHeight = mt_rand(1, 100);
        $expectedX = mt_rand(1, 100);
        $expectedY = mt_rand(1, 100);

        $oMedia = new Media();
        $oMedia->setPath(PHPPRESENTATION_TESTS_BASE_DIR . '/resources/videos/sintel_trailer-480p.ogv')
            ->setName($expectedName)
            ->setResizeProportional(false)
            ->setHeight($expectedHeight)
            ->setWidth($expectedWidth)
            ->setOffsetX($expectedX)
            ->setOffsetY($expectedY);
        $this->oPresentation->getActiveSlide()->addShape($oMedia);

        $expectedWidth = Text::numberFormat(CommonDrawing::pixelsToCentimeters((int) $expectedWidth), 3) . 'cm';
        $expectedHeight = Text::numberFormat(CommonDrawing::pixelsToCentimeters((int) $expectedHeight), 3) . 'cm';
        $expectedX = Text::numberFormat(CommonDrawing::pixelsToCentimeters((int) $expectedX), 3) . 'cm';
        $expectedY = Text::numberFormat(CommonDrawing::pixelsToCentimeters((int) $expectedY), 3) . 'cm';

        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:name', $expectedName);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'svg:width', $expectedWidth);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'svg:height', $expectedHeight);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'svg:x', $expectedX);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'svg:y', $expectedY);
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:plugin';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:mime-type', 'application/vnd.sun.star.media');
        $this->assertZipXmlAttributeStartsWith('content.xml', $element, 'draw:mime-type', 'application/vnd.sun.star.media');
        $this->assertZipXmlAttributeStartsWith('content.xml', $element, 'xlink:href', 'Pictures/');
        $this->assertZipXmlAttributeEndsWith('content.xml', $element, 'xlink:href', 'ogv');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testNote(): void
    {
        $oNote = $this->oPresentation->getActiveSlide()->getNote();
        $oRichText = $oNote->createRichTextShape()->setHeight(300)->setWidth(600);
        $oRichText->createTextRun('testNote');

        $element = '/office:document-content/office:body/office:presentation/draw:page/presentation:notes';
        $this->assertZipXmlElementExists('content.xml', $element);
        $element = '/office:document-content/office:body/office:presentation/draw:page/presentation:notes/draw:frame/draw:text-box/text:p/text:span';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testParagraphLineSpacing(): void
    {
        $richText = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $richText->getActiveParagraph()->setLineSpacingMode(Paragraph::LINE_SPACING_MODE_PERCENT);
        $richText->getActiveParagraph()->setLineSpacing(200);

        $element = $this->getParagraphStyleXPath() . '/style:paragraph-properties';
        $this->assertZipXmlAttributeExists('content.xml', $element, 'fo:line-height');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'fo:line-height', '200%');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $richText->getActiveParagraph()->setLineSpacingMode(Paragraph::LINE_SPACING_MODE_POINT);
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeExists('content.xml', $element, 'fo:line-height');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'fo:line-height', '200pt');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testParagraphSpacingBefore(): void
    {
        $richText = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $richText->getActiveParagraph()->setSpacingBefore(123);

        $element = $this->getParagraphStyleXPath() . '/style:paragraph-properties';
        $this->assertZipXmlAttributeExists('content.xml', $element, 'fo:margin-top');
        // six decimals of a centimetre, so that 123 points come back as 123 rather than as 122.99
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'fo:margin-top', '4.339167cm');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testParagraphSpacingAfter(): void
    {
        $richText = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $richText->getActiveParagraph()->setSpacingAfter(123);

        $element = $this->getParagraphStyleXPath() . '/style:paragraph-properties';
        $this->assertZipXmlAttributeExists('content.xml', $element, 'fo:margin-bottom');
        // six decimals of a centimetre, so that 123 points come back as 123 rather than as 122.99
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'fo:margin-bottom', '4.339167cm');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testRichTextAutoShrink(): void
    {
        $oRichText1 = $this->oPresentation->getActiveSlide()->createRichTextShape();

        $element = $this->getShapeStyleXPath() . '/style:graphic-properties';
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'draw:auto-grow-height');
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'draw:auto-grow-width');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oRichText1->setAutoShrinkHorizontal(false);
        $oRichText1->setAutoShrinkVertical(true);
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeExists('content.xml', $element, 'draw:auto-grow-height');
        $this->assertZipXmlAttributeExists('content.xml', $element, 'draw:auto-grow-width');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:auto-grow-height', 'true');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:auto-grow-width', 'false');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oRichText1->setAutoShrinkHorizontal(true);
        $oRichText1->setAutoShrinkVertical(false);
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeExists('content.xml', $element, 'draw:auto-grow-height');
        $this->assertZipXmlAttributeExists('content.xml', $element, 'draw:auto-grow-width');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:auto-grow-height', 'false');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:auto-grow-width', 'true');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testRichTextRunFontState(): void
    {
        $oRichText = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oRun = $oRichText->createTextRun('Run1');

        $element = $this->getRunStyleXPath() . '/style:text-properties';
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'fo:font-style');
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'style:text-underline-style');
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'style:text-line-through-style');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oRun->getFont()->setItalic(true);
        $oRun->getFont()->setUnderline(Font::UNDERLINE_WAVYHEAVY);
        $oRun->getFont()->setStrikethrough(Font::STRIKE_DOUBLE);
        $this->resetPresentationFile();

        $element = $this->getRunStyleXPath() . '/style:text-properties';
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'fo:font-style', 'italic');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'style:text-underline-style', 'wave');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'style:text-underline-type', 'single');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'style:text-underline-width', 'bold');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'style:text-line-through-style', 'solid');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'style:text-line-through-type', 'double');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        // the italic is spelled once per script, the way the family, the size and the weight are
        $oRun->getFont()->setFormat(Font::FORMAT_EAST_ASIAN);
        $this->resetPresentationFile();

        $element = $this->getRunStyleXPath() . '/style:text-properties';
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'fo:font-style');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'style:font-style-asian', 'italic');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'style:text-underline-style', 'wave');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'style:text-line-through-type', 'double');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oRun->getFont()->setFormat(Font::FORMAT_COMPLEX_SCRIPT);
        $this->resetPresentationFile();

        $element = $this->getRunStyleXPath() . '/style:text-properties';
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'fo:font-style');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'style:font-style-complex', 'italic');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testRichTextRunLanguage(): void
    {
        $oRichText = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oRun = $oRichText->createTextRun('Run1');
        $oRun->getFont()->setFormat(Font::FORMAT_LATIN);

        $element = $this->getRunStyleXPath() . '/style:text-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'style:language-asian');
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'style:language-complex');
        $this->assertZipXmlAttributeExists('content.xml', $element, 'fo:language');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'fo:language', 'en');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oRun->getFont()->setFormat(Font::FORMAT_EAST_ASIAN);
        $this->resetPresentationFile();

        $element = $this->getRunStyleXPath() . '/style:text-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'fo:language');
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'style:language-complex');
        $this->assertZipXmlAttributeExists('content.xml', $element, 'style:language-asian');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'style:language-asian', 'en');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oRun->getFont()->setFormat(Font::FORMAT_COMPLEX_SCRIPT);
        $this->resetPresentationFile();

        $element = $this->getRunStyleXPath() . '/style:text-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'fo:language');
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'style:language-asian');
        $this->assertZipXmlAttributeExists('content.xml', $element, 'style:language-complex');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'style:language-complex', 'en');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oRun->setLanguage('de');
        $oRun->getFont()->setFormat(Font::FORMAT_LATIN);
        $this->resetPresentationFile();

        $element = $this->getRunStyleXPath() . '/style:text-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'style:language-asian');
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'style:language-complex');
        $this->assertZipXmlAttributeExists('content.xml', $element, 'fo:language');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'fo:language', 'de');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oRun->getFont()->setFormat(Font::FORMAT_EAST_ASIAN);
        $this->resetPresentationFile();

        $element = $this->getRunStyleXPath() . '/style:text-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'fo:language');
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'style:language-complex');
        $this->assertZipXmlAttributeExists('content.xml', $element, 'style:language-asian');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'style:language-asian', 'de');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oRun->getFont()->setFormat(Font::FORMAT_COMPLEX_SCRIPT);
        $this->resetPresentationFile();

        $element = $this->getRunStyleXPath() . '/style:text-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'fo:language');
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'style:language-asian');
        $this->assertZipXmlAttributeExists('content.xml', $element, 'style:language-complex');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'style:language-complex', 'de');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testRichTextBorder(): void
    {
        $oRichText1 = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oRichText1->getBorder()->setColor(new Color('FF4672A8'))->setDashStyle(Border::DASH_SOLID)->setLineStyle(Border::LINE_NONE);

        $element = $this->getShapeStyleXPath() . '/style:graphic-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'svg:stroke-color');
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'svg:stroke-width');
        $this->assertZipXmlAttributeExists('content.xml', $element, 'draw:stroke');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:stroke', 'none');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oRichText1->getBorder()->setColor(new Color('FF4672A8'))->setDashStyle(Border::DASH_SOLID)->setLineStyle(Border::LINE_SINGLE);
        $this->resetPresentationFile();

        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeExists('content.xml', $element, 'svg:stroke-color');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'svg:stroke-color', '#' . $oRichText1->getBorder()->getColor()->getRGB());
        $this->assertZipXmlAttributeExists('content.xml', $element, 'svg:stroke-width');
        $this->assertZipXmlAttributeStartsWith('content.xml', $element, 'svg:stroke-width', (string) number_format(CommonDrawing::pointsToCentimeters($oRichText1->getBorder()->getLineWidth()), 3, '.', ''));
        $this->assertZipXmlAttributeEndsWith('content.xml', $element, 'svg:stroke-width', 'cm');
        $this->assertZipXmlAttributeExists('content.xml', $element, 'draw:stroke');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:stroke', 'solid');
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'draw:stroke-dash');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oRichText1->getBorder()->setColor(new Color('FF4672A8'))->setDashStyle(Border::DASH_DASH);
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:stroke', 'dash');
        $this->assertZipXmlAttributeExists('content.xml', $element, 'draw:stroke-dash');
        $this->assertZipXmlAttributeStartsWith('content.xml', $element, 'draw:stroke-dash', 'strokeDash_');
        $this->assertZipXmlAttributeEndsWith('content.xml', $element, 'draw:stroke-dash', $oRichText1->getBorder()->getDashStyle());
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    /**
     * @return array<array<string>>
     */
    public static function dataProviderField(): array
    {
        return [
            [Field::TYPE_SLIDENUM, 'text:page-number'],
            [Field::TYPE_DATETIME, 'text:date'],
            // OpenDocument names a field by what it is, so the fourteen dated formats OOXML
            // numbers come down to the two it has
            ['datetime3', 'text:date'],
            ['datetime11', 'text:time'],
            ['author', 'text:author-name'],
            ['file2', 'text:file-name'],
        ];
    }

    /**
     * @dataProvider dataProviderField
     */
    #[DataProvider('dataProviderField')]
    public function testRichTextField(string $type, string $expectedElement): void
    {
        $oParagraph = $this->oPresentation->getActiveSlide()->createRichTextShape()->getActiveParagraph();
        $oParagraph->createTextRun('page ');
        $oParagraph->createField($type, '7');

        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:text-box/text:p/text:span';
        // the field is written where it stands, and the run beside it stays a run
        $this->assertZipXmlElementEquals('content.xml', $element . '[1]', 'page ');
        $this->assertZipXmlElementNotExists('content.xml', $element . '[1]/' . $expectedElement);
        $this->assertZipXmlElementEquals('content.xml', $element . '[2]/' . $expectedElement, '7');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testRichTextFieldSlideCount(): void
    {
        $oParagraph = $this->oPresentation->getActiveSlide()->createRichTextShape()->getActiveParagraph();
        $oParagraph->createField(Field::TYPE_SLIDECOUNT, '12');

        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:text-box/text:p/text:span/text:page-count';
        $this->assertZipXmlElementEquals('content.xml', $element, '12');
        // No schema assertion here, and the file is not the reason. libxml rejects
        // `text:page-count` inside a `text:span` -- along with every other element of the two
        // groups the schema spells as one `<element>` holding a choice of names -- while accepting
        // the same element as a direct child of `text:p`, which has the identical content model
        // (`zeroOrMore paragraph-content-or-hyperlink`). Jing, the reference RELAX NG
        // implementation, validates this file. LibreOffice writes the count inside a span too.
    }

    public function testRichTextFieldOfNoKnownKind(): void
    {
        $oParagraph = $this->oPresentation->getActiveSlide()->createRichTextShape()->getActiveParagraph();
        // OpenDocument has no element for this one, so what is left is the text it stands in for
        $oParagraph->createField('slidename', 'Slide 7');

        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:text-box/text:p/text:span';
        $this->assertZipXmlElementEquals('content.xml', $element, 'Slide 7');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    /**
     * @return array<array<string>>
     */
    public static function dataProviderPlaceholderField(): array
    {
        return [
            [Placeholder::PH_TYPE_SLIDENUM, 'text:page-number'],
            [Placeholder::PH_TYPE_DATETIME, 'text:date'],
        ];
    }

    /**
     * @dataProvider dataProviderPlaceholderField
     */
    #[DataProvider('dataProviderPlaceholderField')]
    public function testRichTextPlaceholderField(string $type, string $expectedElement): void
    {
        $oRichText = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oRichText->createTextRun('7');
        $oRichText->setPlaceHolder(new Placeholder($type));

        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:text-box/text:p/text:span/' . $expectedElement;
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlElementEquals('content.xml', $element, '7');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testRichTextPlaceholderTitle(): void
    {
        $oRichText = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oRichText->createTextRun('Title');
        $oRichText->setPlaceHolder(new Placeholder(Placeholder::PH_TYPE_TITLE));

        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:text-box/text:p/text:span';
        $this->assertZipXmlElementEquals('content.xml', $element, 'Title');
        $this->assertZipXmlElementNotExists('content.xml', $element . '/text:page-number');
        $this->assertZipXmlElementNotExists('content.xml', $element . '/text:date');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testRichTextRotation(): void
    {
        $expectedValue = mt_rand(1, 360);
        $oRichText1 = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oRichText1->setRotation($expectedValue);

        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeExists('content.xml', $element, 'draw:transform');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:transform', 'rotate (-' . deg2rad($expectedValue) . ') translate (0.000cm 0.000cm)');
    }

    public function testRichTextRotationTranslatesTheTurnedCorner(): void
    {
        $oRichText = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oRichText->setOffsetX(400)->setOffsetY(100)->setWidth(200)->setHeight(100)->setRotation(30);

        // `translate` moves the frame after `rotate` has turned it about the origin, so the point
        // written is where the top left corner lands, not where it started. These are the values
        // LibreOffice itself writes for a shape of this size, at this offset, turned this far.
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame';
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:transform', 'rotate (-' . deg2rad(30) . ') translate (11.599cm 1.500cm)');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testRichTextShadow(): void
    {
        $randAlpha = mt_rand(0, 100);
        $oRichText = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oRichText->createTextRun('AAA');
        $oRichText->getShadow()->setVisible(true)->setAlpha($randAlpha)->setBlurRadius(2);

        for ($inc = 0; $inc <= 360; $inc += 45) {
            $randDistance = mt_rand(0, 100);
            $oRichText->getShadow()->setDirection($inc)->setDistance($randDistance);

            // resolved inside the loop: the shape is mutated on every pass and the file is reset
            // at the end of it, so the name has to be read from what this pass writes
            $element = $this->getShapeStyleXPath() . '/style:graphic-properties';
            $this->assertZipXmlElementExists('content.xml', $element);
            $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:shadow', 'visible');
            $this->assertZipXmlAttributeEquals('content.xml', $element, 'style:mirror', 'none');
            // Opacity
            $this->assertZipXmlAttributeStartsWith('content.xml', $element, 'draw:shadow-opacity', (string) (100 - $randAlpha));
            $this->assertZipXmlAttributeEndsWith('content.xml', $element, 'draw:shadow-opacity', '%');
            // Color
            $this->assertZipXmlAttributeStartsWith('content.xml', $element, 'draw:shadow-color', '#');
            // X
            if (90 == $inc || 270 == $inc) {
                $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:shadow-offset-x', '0cm');
            } else {
                if ($inc > 90 && $inc < 270) {
                    $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:shadow-offset-x', '-' . Drawing::pixelsToCentimeters((int) $randDistance) . 'cm');
                } else {
                    $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:shadow-offset-x', Drawing::pixelsToCentimeters((int) $randDistance) . 'cm');
                }
            }
            // Y
            if (0 == $inc || 180 == $inc || 360 == $inc) {
                $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:shadow-offset-y', '0cm');
            } else {
                if ($inc < 180) {
                    $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:shadow-offset-y', Drawing::pixelsToCentimeters((int) $randDistance) . 'cm');
                } else {
                    $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:shadow-offset-y', '-' . Drawing::pixelsToCentimeters((int) $randDistance) . 'cm');
                }
            }
            $this->assertIsSchemaOpenDocumentValid('1.2');
            $this->resetPresentationFile();
        }
    }

    public function testSlideName(): void
    {
        $element = '/office:document-content/office:body/office:presentation/draw:page';

        // Every page is named, so that a link can address it; a slide with no name of its own is
        // named after its position
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:name', 'Slide 1');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $this->oPresentation->getActiveSlide()->setName('AAAA');
        $this->resetPresentationFile();

        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeExists('content.xml', $element, 'draw:name');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:name', 'AAAA');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $this->oPresentation->getActiveSlide()->setName();
        $this->resetPresentationFile();

        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:name', 'Slide 1');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testSlideLink(): void
    {
        $oRun = $this->oPresentation->getActiveSlide()->createRichTextShape()->createTextRun('AAAA');
        $oRun->getHyperlink()->setSlideNumber(3);
        $this->oPresentation->createSlide();
        $this->oPresentation->createSlide()->setName('Milestone Overview');

        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:text-box/text:p/text:span/text:a';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'xlink:href', '#Milestone Overview');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        // A slide with no name of its own is addressed by the name its page carries
        $this->oPresentation->getSlide(2)->setName();
        $this->resetPresentationFile();

        $this->assertZipXmlAttributeEquals('content.xml', $element, 'xlink:href', '#Slide 3');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testSlideLinkOutOfRange(): void
    {
        $oRun = $this->oPresentation->getActiveSlide()->createRichTextShape()->createTextRun('AAAA');
        $oRun->getHyperlink()->setSlideNumber(9);

        // Nothing to address, so the action string is left as it was rather than pointing nowhere
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/draw:text-box/text:p/text:span/text:a';
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'xlink:href', 'ppaction://hlinksldjump');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testStyleAlignment(): void
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oRichText1 = $oSlide->createRichTextShape();
        $oRichText1->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $oRichText1->createTextRun('Run1');
        $oRichText2 = $oSlide->createRichTextShape();
        $oRichText2->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_DISTRIBUTED);
        $oRichText2->createTextRun('Run2');
        $oRichText3 = $oSlide->createRichTextShape();
        $oRichText3->getActiveParagraph()->getAlignment()->setHorizontal('AAAAA');
        $oRichText3->createTextRun('Run3');
        $oRichText4 = $oSlide->createRichTextShape();
        $oRichText4->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_JUSTIFY);
        $oRichText4->createTextRun('Run4');
        $oRichText5 = $oSlide->createRichTextShape();
        $oRichText5->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);
        $oRichText5->createTextRun('Run5');
        $oRichText6 = $oSlide->createRichTextShape();
        $oRichText6->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
        $oRichText6->createTextRun('Run6');

        $element = $this->getParagraphStyleXPath(1) . '/style:paragraph-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'fo:text-align', 'center');

        $element = $this->getParagraphStyleXPath(2) . '/style:paragraph-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'fo:text-align', 'justify');

        // an unknown alignment falls back to left, which is what shape 5 asks for outright -- the
        // two resolve to the same style the moment the writer shares one
        $element = $this->getParagraphStyleXPath(3) . '/style:paragraph-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'fo:text-align', 'left');

        $element = $this->getParagraphStyleXPath(4) . '/style:paragraph-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'fo:text-align', 'justify');

        $element = $this->getParagraphStyleXPath(5) . '/style:paragraph-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'fo:text-align', 'left');

        $element = $this->getParagraphStyleXPath(6) . '/style:paragraph-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'fo:text-align', 'right');

        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testStyleAlignmentRTL(): void
    {
        $oSlide = $this->oPresentation->getActiveSlide();
        $oRichText1 = $oSlide->createRichTextShape();
        $oRichText1->getActiveParagraph()->getAlignment()->setIsRTL(true);
        $oRichText1->createTextRun('Run1');
        $oRichText2 = $oSlide->createRichTextShape();
        $oRichText2->getActiveParagraph()->getAlignment()->setIsRTL(false);
        $oRichText2->createTextRun('Run2');

        $element = $this->getParagraphStyleXPath(1) . '/style:paragraph-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'style:writing-mode', 'rl-tb');

        $element = $this->getParagraphStyleXPath(2) . '/style:paragraph-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'style:writing-mode', 'lr-tb');

        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testStyleFontBold(): void
    {
        $oRichText = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oRun = $oRichText->createTextRun('Run1');
        $oRun->getFont()->setBold(true);
        $oRun->getFont()->setFormat(Font::FORMAT_LATIN);

        $element = $this->getRunStyleXPath() . '/style:text-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'style:font-weight-asian');
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'style:font-weight-complex');
        $this->assertZipXmlAttributeExists('content.xml', $element, 'fo:font-weight');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'fo:font-weight', 'bold');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oRun->getFont()->setFormat(Font::FORMAT_EAST_ASIAN);
        $this->resetPresentationFile();

        $element = $this->getRunStyleXPath() . '/style:text-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'fo:font-weight');
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'style:font-weight-complex');
        $this->assertZipXmlAttributeExists('content.xml', $element, 'style:font-weight-asian');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'style:font-weight-asian', 'bold');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oRun->getFont()->setFormat(Font::FORMAT_COMPLEX_SCRIPT);
        $this->resetPresentationFile();

        $element = $this->getRunStyleXPath() . '/style:text-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'fo:font-weight');
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'style:font-weight-asian');
        $this->assertZipXmlAttributeExists('content.xml', $element, 'style:font-weight-complex');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'style:font-weight-complex', 'bold');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oRun->getFont()->setBold(false);
        $oRun->getFont()->setFormat(Font::FORMAT_LATIN);
        $this->resetPresentationFile();

        $element = $this->getRunStyleXPath() . '/style:text-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'fo:font-weight');
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'style:font-weight-asian');
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'style:font-weight-complex');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oRun->getFont()->setFormat(Font::FORMAT_EAST_ASIAN);
        $this->resetPresentationFile();

        $element = $this->getRunStyleXPath() . '/style:text-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'fo:font-weight');
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'style:font-weight-asian');
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'style:font-weight-complex');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oRun->getFont()->setFormat(Font::FORMAT_COMPLEX_SCRIPT);
        $this->resetPresentationFile();

        $element = $this->getRunStyleXPath() . '/style:text-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'fo:font-weight');
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'style:font-weight-asian');
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'style:font-weight-complex');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testStyleFontFormat(): void
    {
        $oRichText = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oRun = $oRichText->createTextRun('Run1');
        $oRun->getFont()->setFormat(Font::FORMAT_LATIN);

        $element = $this->getRunStyleXPath() . '/style:text-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeExists('content.xml', $element, 'style:script-type');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'style:script-type', 'latin');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oRun->getFont()->setFormat(Font::FORMAT_EAST_ASIAN);
        $this->resetPresentationFile();

        $element = $this->getRunStyleXPath() . '/style:text-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeExists('content.xml', $element, 'style:script-type');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'style:script-type', 'asian');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oRun->getFont()->setFormat(Font::FORMAT_COMPLEX_SCRIPT);
        $this->resetPresentationFile();

        $element = $this->getRunStyleXPath() . '/style:text-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeExists('content.xml', $element, 'style:script-type');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'style:script-type', 'complex');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testStyleFontCapitalization(): void
    {
        $oRichText = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oRun = $oRichText->createTextRun('Run1');
        $oRun->getFont()->setCapitalization(Font::CAPITALIZATION_ALL);

        $element = $this->getRunStyleXPath() . '/style:text-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeExists('content.xml', $element, 'fo:text-transform');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'fo:text-transform', 'uppercase');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oRun->getFont()->setCapitalization(Font::CAPITALIZATION_SMALL);
        $this->resetPresentationFile();

        $element = $this->getRunStyleXPath() . '/style:text-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeExists('content.xml', $element, 'fo:text-transform');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'fo:text-transform', 'lowercase');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oRun->getFont()->setCapitalization(Font::CAPITALIZATION_NONE);
        $this->resetPresentationFile();

        $element = $this->getRunStyleXPath() . '/style:text-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeExists('content.xml', $element, 'fo:text-transform');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'fo:text-transform', 'none');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testStyleFontName(): void
    {
        $oRichText = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oRun = $oRichText->createTextRun('Run1');
        $oRun->getFont()->setName('Calibri');
        $oRun->getFont()->setFormat(Font::FORMAT_LATIN);

        $element = $this->getRunStyleXPath() . '/style:text-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'style:font-family-asian');
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'style:font-family-complex');
        $this->assertZipXmlAttributeExists('content.xml', $element, 'fo:font-family');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'fo:font-family', 'Calibri');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oRun->getFont()->setFormat(Font::FORMAT_EAST_ASIAN);
        $this->resetPresentationFile();

        $element = $this->getRunStyleXPath() . '/style:text-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'fo:font-family');
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'style:font-family-complex');
        $this->assertZipXmlAttributeExists('content.xml', $element, 'style:font-family-asian');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'style:font-family-asian', 'Calibri');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oRun->getFont()->setFormat(Font::FORMAT_COMPLEX_SCRIPT);
        $this->resetPresentationFile();

        $element = $this->getRunStyleXPath() . '/style:text-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'fo:font-family');
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'style:font-family-asian');
        $this->assertZipXmlAttributeExists('content.xml', $element, 'style:font-family-complex');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'style:font-family-complex', 'Calibri');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testStyleFontSize(): void
    {
        $oRichText = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oRun = $oRichText->createTextRun('Run1');
        $oRun->getFont()->setSize(12);
        $oRun->getFont()->setFormat(Font::FORMAT_LATIN);

        $element = $this->getRunStyleXPath() . '/style:text-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'style:font-size-asian');
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'style:font-size-complex');
        $this->assertZipXmlAttributeExists('content.xml', $element, 'fo:font-size');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'fo:font-size', '12pt');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oRun->getFont()->setFormat(Font::FORMAT_EAST_ASIAN);
        $this->resetPresentationFile();

        $element = $this->getRunStyleXPath() . '/style:text-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'fo:font-size');
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'style:font-size-complex');
        $this->assertZipXmlAttributeExists('content.xml', $element, 'style:font-size-asian');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'style:font-size-asian', '12pt');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oRun->getFont()->setFormat(Font::FORMAT_COMPLEX_SCRIPT);
        $this->resetPresentationFile();

        $element = $this->getRunStyleXPath() . '/style:text-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'fo:font-size');
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'style:font-size-asian');
        $this->assertZipXmlAttributeExists('content.xml', $element, 'style:font-size-complex');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'style:font-size-complex', '12pt');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testTable(): void
    {
        $oShape = $this->oPresentation->getActiveSlide()->createTableShape();
        $oShape->createRow();

        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/table:table';
        $this->assertZipXmlElementExists('content.xml', $element);

        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testTableFirstRow(): void
    {
        $oShape = $this->oPresentation->getActiveSlide()->createTableShape();
        $oShape->createRow();
        $oShape->createRow();

        $table = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/table:table';

        // The first row is a header row by default, and every row stays a row of the table
        $this->assertZipXmlAttributeEquals('content.xml', $table, 'table:use-first-row-styles', 'true');
        $this->assertZipXmlElementCount('content.xml', $table . '/table:table-row', 2);
        // and none of them is wrapped: that is the other ODF table model, and a consumer reading a
        // drawing table drops the rows it wraps
        $this->assertZipXmlElementNotExists('content.xml', $table . '/table:table-header-rows');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        // A table whose first row holds data, not column labels
        $this->resetPresentationFile();
        $oShape->setFirstRow(false);

        $this->assertZipXmlAttributeNotExists('content.xml', $table, 'table:use-first-row-styles');
        $this->assertZipXmlElementCount('content.xml', $table . '/table:table-row', 2);
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testTableBandRow(): void
    {
        $oShape = $this->oPresentation->getActiveSlide()->createTableShape();
        $oShape->createRow();

        $table = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/table:table';

        // Banded rows are on by default, the same as `bandRow` on `a:tblPr`
        $this->assertZipXmlAttributeEquals('content.xml', $table, 'table:use-banding-rows-styles', 'true');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $this->resetPresentationFile();
        $oShape->setBandRow(false);

        $this->assertZipXmlAttributeNotExists('content.xml', $table, 'table:use-banding-rows-styles');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testTableEmpty(): void
    {
        $this->oPresentation->getActiveSlide()->createTableShape();

        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/table:table';
        $this->assertZipXmlElementNotExists('content.xml', $element);

        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    /**
     * @return array<array{0: string, 1: string, 2: string}>
     */
    public static function dataProviderTableCellBorder(): array
    {
        return [
            // A line style the ODF border styles can name outright
            [Border::LINE_NONE, Border::DASH_SOLID, 'none'],
            [Border::LINE_DOUBLE, Border::DASH_SOLID, 'double'],
            // The compound lines have no CSS2 equivalent, double is the closest
            [Border::LINE_THICKTHIN, Border::DASH_SOLID, 'double'],
            [Border::LINE_THINTHICK, Border::DASH_SOLID, 'double'],
            [Border::LINE_TRI, Border::DASH_SOLID, 'double'],
            // A dash pattern decides the style, and CSS2 knows two of them
            [Border::LINE_SINGLE, Border::DASH_DASH, 'dashed'],
            [Border::LINE_SINGLE, Border::DASH_LARGEDASHDOT, 'dashed'],
            [Border::LINE_SINGLE, Border::DASH_DOT, 'dotted'],
            [Border::LINE_SINGLE, Border::DASH_SYSDOT, 'dotted'],
            // No line at all outranks any dash pattern
            [Border::LINE_NONE, Border::DASH_DASH, 'none'],
        ];
    }

    /**
     * @dataProvider dataProviderTableCellBorder
     */
    #[DataProvider('dataProviderTableCellBorder')]
    public function testTableCellBorder(string $lineStyle, string $dashStyle, string $expected): void
    {
        $oShape = $this->oPresentation->getActiveSlide()->createTableShape();
        $oCell = $oShape->createRow()->getCell();
        $oCell->getBorders()->getBottom()->setLineStyle($lineStyle)->setDashStyle($dashStyle);

        // Only the bottom differs from the default, so each side is written on its own
        $element = $this->getTableCellStyleXPath(1, 1) . '/style:paragraph-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEndsWith('content.xml', $element, 'fo:border-bottom', $expected . ' #000000');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    /**
     * A border given no colour is written without `svg:stroke-color`: the stroke and its width stay,
     * and the parent style answers for the colour.
     */
    public function testRichTextBorderWithoutAColor(): void
    {
        $oRichText = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oRichText->getBorder()->setLineStyle(Border::LINE_SINGLE)->setColor();

        $element = $this->getShapeStyleXPath() . '/style:graphic-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'svg:stroke-color');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:stroke', 'solid');
        $this->assertZipXmlAttributeExists('content.xml', $element, 'svg:stroke-width');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    /**
     * `fo:border` takes the CSS2 shorthand, where the colour is the optional third part. A cell
     * border given no colour is written as a width and a style alone.
     */
    public function testTableCellBorderWithoutAColor(): void
    {
        $oShape = $this->oPresentation->getActiveSlide()->createTableShape();
        $oCell = $oShape->createRow()->getCell();
        foreach (['getBottom', 'getTop', 'getLeft', 'getRight'] as $side) {
            $oCell->getBorders()->{$side}()->setLineStyle(Border::LINE_SINGLE)->setColor();
        }

        $element = $this->getTableCellStyleXPath(1, 1) . '/style:paragraph-properties';
        $this->assertZipXmlAttributeEndsWith('content.xml', $element, 'fo:border', 'pt solid');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testTableCellBorderOnEverySide(): void
    {
        $oShape = $this->oPresentation->getActiveSlide()->createTableShape();
        $oCell = $oShape->createRow()->getCell();
        foreach (['getBottom', 'getTop', 'getLeft', 'getRight'] as $side) {
            $oCell->getBorders()->{$side}()->setLineStyle(Border::LINE_DOUBLE);
        }

        // The four sides agree, so they collapse into the fo:border shorthand
        $element = $this->getTableCellStyleXPath(1, 1) . '/style:paragraph-properties';
        $this->assertZipXmlAttributeEndsWith('content.xml', $element, 'fo:border', 'double #000000');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testTableCellFill(): void
    {
        $oColor = new Color();
        $oColor->setRGB(Color::COLOR_BLUE);

        $oFill = new Fill();
        $oFill->setFillType(Fill::FILL_SOLID)->setStartColor($oColor);

        $oShape = $this->oPresentation->getActiveSlide()->createTableShape();
        $oRow = $oShape->createRow();
        $oCell = $oRow->getCell();
        $oCell->setFill($oFill);

        $element = $this->getTableCellStyleXPath(1, 1) . '';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'style:family', 'table-cell');

        $element = $this->getTableCellStyleXPath(1, 1) . '/style:graphic-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:fill', 'solid');
        $this->assertZipXmlAttributeStartsWith('content.xml', $element, 'draw:fill-color', '#');
        $this->assertZipXmlAttributeEndsWith('content.xml', $element, 'draw:fill-color', $oColor->getRGB());

        $this->assertIsSchemaOpenDocumentNotValid('1.2');
    }

    public function testTableRowFill(): void
    {
        $oFill = new Fill();
        $oFill->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFE06B20'));

        $oShape = $this->oPresentation->getActiveSlide()->createTableShape(2);
        $oRow = $oShape->createRow();
        $oRow->setFill($oFill);

        // ODF puts no fill on a table-row style, so the fill of the row lands on each of its cells
        foreach ([1, 2] as $cell) {
            $element = $this->getTableCellStyleXPath(1, $cell) . '/style:graphic-properties';
            $this->assertZipXmlElementExists('content.xml', $element);
            $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:fill', 'solid');
            $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:fill-color', '#E06B20');
        }
    }

    public function testTableRowFillUntouched(): void
    {
        $oShape = $this->oPresentation->getActiveSlide()->createTableShape(2);
        $oShape->createRow();

        // Nobody asked for a fill, on the cell or on its row, which is how a table starts out:
        // the cell style is still written, and carries no graphic properties at all
        foreach ([1, 2] as $cell) {
            $element = $this->getTableCellStyleXPath(1, $cell) . '';
            $this->assertZipXmlElementExists('content.xml', $element);
            $this->assertZipXmlElementNotExists('content.xml', $element . '/style:graphic-properties');
        }

        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testTableCellFillNoneInFilledRow(): void
    {
        $oFill = new Fill();
        $oFill->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFE06B20'));

        $oShape = $this->oPresentation->getActiveSlide()->createTableShape(2);
        $oRow = $oShape->createRow();
        $oRow->setFill($oFill);

        // The second cell refuses a fill outright, which is not the same as never asking for one:
        // it stays transparent while its neighbour takes the colour of the row
        $oRow->getCell(1)->getFill()->setFillType(Fill::FILL_NONE);

        $element = $this->getTableCellStyleXPath(1, 1) . '/style:graphic-properties';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'draw:fill-color', '#E06B20');

        $element = $this->getTableCellStyleXPath(1, 2) . '';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlElementNotExists('content.xml', $element . '/style:graphic-properties');
    }

    public function testTableWithColspan(): void
    {
        $value = mt_rand(2, 100);

        $oShape = $this->oPresentation->getActiveSlide()->createTableShape($value);
        $oRow = $oShape->createRow();
        $oCell = $oRow->getCell();
        $oCell->setColSpan($value);

        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/table:table/table:table-row/table:table-cell';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'table:number-columns-spanned', $value);

        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    /**
     * @see : https://github.com/PHPOffice/PHPPresentation/issues/70
     */
    public function testTableWithHyperlink(): void
    {
        $oShape = $this->oPresentation->getActiveSlide()->createTableShape(4);
        $oShape->setHeight(200)->setWidth(600)->setOffsetX(150)->setOffsetY(300);
        $oRow = $oShape->createRow();
        $oCell = $oRow->getCell();
        $oTextRun = $oCell->createTextRun('AAA');
        $oHyperlink = $oTextRun->getHyperlink();
        $oHyperlink->setUrl('https://github.com/PHPOffice/PHPPresentation/');

        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/table:table/table:table-row/table:table-cell/text:p/text:span/text:a';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'xlink:href', 'https://github.com/PHPOffice/PHPPresentation/');

        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testTableWithText(): void
    {
        $oRun = new Run();
        $oRun->setText('Test');

        $oShape = $this->oPresentation->getActiveSlide()->createTableShape();
        $oRow = $oShape->createRow();
        $oCell = $oRow->getCell();
        $oCell->addText($oRun);
        $oCell->createBreak();

        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/table:table/table:table-row/table:table-cell/text:p/text:span';
        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlElementEquals('content.xml', $element, 'Test');
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame/table:table/table:table-row/table:table-cell/text:p/text:span/text:line-break';
        $this->assertZipXmlElementExists('content.xml', $element);

        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testTransition(): void
    {
        $value = mt_rand(1000, 5000);

        $element = $this->getSlideStyleXPath() . '/style:drawing-page-properties';

        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'presentation:duration');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oTransition = new Transition();
        $oTransition->setTimeTrigger(true, $value);
        $this->oPresentation->getActiveSlide()->setTransition($oTransition);
        $this->resetPresentationFile();
        $element = $this->getSlideStyleXPath() . '/style:drawing-page-properties';

        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeExists('content.xml', $element, 'presentation:duration');
        $this->assertZipXmlAttributeStartsWith('content.xml', $element, 'presentation:duration', 'PT');
        $this->assertZipXmlAttributeEndsWith('content.xml', $element, 'presentation:duration', 'S');
        $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:duration', number_format($value / 1000, 6, '.', ''));
        $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-type', 'automatic');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oTransition->setSpeed(Transition::SPEED_FAST);
        $this->resetPresentationFile();
        $element = $this->getSlideStyleXPath() . '/style:drawing-page-properties';

        $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-speed', 'fast');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oTransition->setSpeed(Transition::SPEED_MEDIUM);
        $this->resetPresentationFile();
        $element = $this->getSlideStyleXPath() . '/style:drawing-page-properties';

        $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-speed', 'medium');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oTransition->setSpeed(Transition::SPEED_SLOW);
        $this->resetPresentationFile();
        $element = $this->getSlideStyleXPath() . '/style:drawing-page-properties';

        $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-speed', 'slow');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $rcTransition = new ReflectionClass('PhpOffice\PhpPresentation\Slide\Transition');
        $arrayConstants = $rcTransition->getConstants();
        foreach ($arrayConstants as $key => $value) {
            if (0 !== strpos($key, 'TRANSITION_')) {
                continue;
            }
            $this->resetPresentationFile();
            $oTransition->setTransitionType($rcTransition->getConstant($key));
            $this->oPresentation->getActiveSlide()->setTransition($oTransition);
            $element = $this->getSlideStyleXPath() . '/style:drawing-page-properties';
            switch ($key) {
                case 'TRANSITION_BLINDS_HORIZONTAL':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'horizontal-stripes');

                    break;
                case 'TRANSITION_BLINDS_VERTICAL':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'vertical-stripes');

                    break;
                case 'TRANSITION_CHECKER_HORIZONTAL':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'horizontal-checkerboard');

                    break;
                case 'TRANSITION_CHECKER_VERTICAL':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'vertical-checkerboard');

                    break;
                case 'TRANSITION_CIRCLE_HORIZONTAL':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'none');

                    break;
                case 'TRANSITION_CIRCLE_VERTICAL':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'none');

                    break;
                case 'TRANSITION_COMB_HORIZONTAL':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'none');

                    break;
                case 'TRANSITION_COMB_VERTICAL':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'none');

                    break;
                case 'TRANSITION_COVER_DOWN':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'uncover-to-bottom');

                    break;
                case 'TRANSITION_COVER_LEFT':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'uncover-to-left');

                    break;
                case 'TRANSITION_COVER_LEFT_DOWN':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'uncover-to-lowerleft');

                    break;
                case 'TRANSITION_COVER_LEFT_UP':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'uncover-to-upperleft');

                    break;
                case 'TRANSITION_COVER_RIGHT':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'uncover-to-right');

                    break;
                case 'TRANSITION_COVER_RIGHT_DOWN':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'uncover-to-lowerright');

                    break;
                case 'TRANSITION_COVER_RIGHT_UP':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'uncover-to-upperright');

                    break;
                case 'TRANSITION_COVER_UP':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'uncover-to-top');

                    break;
                case 'TRANSITION_CUT':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'none');

                    break;
                case 'TRANSITION_DIAMOND':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'none');

                    break;
                case 'TRANSITION_DISSOLVE':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'dissolve');

                    break;
                case 'TRANSITION_FADE':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'fade-from-center');

                    break;
                case 'TRANSITION_NEWSFLASH':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'none');

                    break;
                case 'TRANSITION_PLUS':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'close');

                    break;
                case 'TRANSITION_PULL_DOWN':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'stretch-from-bottom');

                    break;
                case 'TRANSITION_PULL_LEFT':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'stretch-from-left');

                    break;
                case 'TRANSITION_PULL_RIGHT':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'stretch-from-right');

                    break;
                case 'TRANSITION_PULL_UP':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'stretch-from-top');

                    break;
                case 'TRANSITION_PUSH_DOWN':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'roll-from-bottom');

                    break;
                case 'TRANSITION_PUSH_LEFT':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'roll-from-left');

                    break;
                case 'TRANSITION_PUSH_RIGHT':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'roll-from-right');

                    break;
                case 'TRANSITION_PUSH_UP':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'roll-from-top');

                    break;
                case 'TRANSITION_RANDOM':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'random');

                    break;
                case 'TRANSITION_RANDOMBAR_HORIZONTAL':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'horizontal-lines');

                    break;
                case 'TRANSITION_RANDOMBAR_VERTICAL':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'vertical-lines');

                    break;
                case 'TRANSITION_SPLIT_IN_HORIZONTAL':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'close-horizontal');

                    break;
                case 'TRANSITION_SPLIT_OUT_HORIZONTAL':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'open-horizontal');

                    break;
                case 'TRANSITION_SPLIT_IN_VERTICAL':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'close-vertical');

                    break;
                case 'TRANSITION_SPLIT_OUT_VERTICAL':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'open-vertical');

                    break;
                case 'TRANSITION_STRIPS_LEFT_DOWN':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'none');

                    break;
                case 'TRANSITION_STRIPS_LEFT_UP':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'none');

                    break;
                case 'TRANSITION_STRIPS_RIGHT_DOWN':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'none');

                    break;
                case 'TRANSITION_STRIPS_RIGHT_UP':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'none');

                    break;
                case 'TRANSITION_WEDGE':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'none');

                    break;
                case 'TRANSITION_WIPE_DOWN':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'fade-from-bottom');

                    break;
                case 'TRANSITION_WIPE_LEFT':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'fade-from-left');

                    break;
                case 'TRANSITION_WIPE_RIGHT':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'fade-from-right');

                    break;
                case 'TRANSITION_WIPE_UP':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'fade-from-top');

                    break;
                case 'TRANSITION_ZOOM_IN':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'none');

                    break;
                case 'TRANSITION_ZOOM_OUT':
                    $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-style', 'none');

                    break;
            }
            $this->assertIsSchemaOpenDocumentValid('1.2');
        }

        $oTransition->setTimeTrigger(false);
        $oTransition->setManualTrigger(true);
        $this->resetPresentationFile();
        $element = $this->getSlideStyleXPath() . '/style:drawing-page-properties';
        $this->assertZipXmlAttributeContains('content.xml', $element, 'presentation:transition-type', 'manual');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testVisibility(): void
    {
        $element = $this->getSlideStyleXPath() . '/style:drawing-page-properties';

        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeNotExists('content.xml', $element, 'presentation:visibility');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $oSlide = $this->oPresentation->getActiveSlide();
        $oSlide->setIsVisible(false);
        $this->resetPresentationFile();
        $element = $this->getSlideStyleXPath() . '/style:drawing-page-properties';

        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeExists('content.xml', $element, 'presentation:visibility');
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'presentation:visibility', 'hidden');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    /**
     * @dataProvider dataProviderShowType
     */
    #[DataProvider('dataProviderShowType')]
    public function testShowType(string $slideshowType, bool $withAttribute): void
    {
        $this->oPresentation->getPresentationProperties()->setSlideshowType($slideshowType);

        $this->assertZipFileExists('content.xml');
        $element = '/office:document-content/office:body/office:presentation/presentation:settings';
        $this->assertZipXmlElementExists('content.xml', $element);
        if ($withAttribute) {
            $this->assertZipXmlAttributeExists('content.xml', $element, 'presentation:full-screen');
            $this->assertZipXmlAttributeEquals('content.xml', $element, 'presentation:full-screen', 'false');
        } else {
            $this->assertZipXmlAttributeNotExists('content.xml', $element, 'presentation:full-screen');
        }
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    /**
     * @return array<array<bool|string>>
     */
    public static function dataProviderShowType(): array
    {
        return [
            [
                PresentationProperties::SLIDESHOW_TYPE_PRESENT,
                false,
            ],
            [
                PresentationProperties::SLIDESHOW_TYPE_BROWSE,
                true,
            ],
            [
                PresentationProperties::SLIDESHOW_TYPE_KIOSK,
                false,
            ],
        ];
    }

    /**
     * @dataProvider dataProviderLoopContinuouslyUntilEsc
     */
    #[DataProvider('dataProviderLoopContinuouslyUntilEsc')]
    public function testLoopContinuouslyUntilEsc(bool $isLoopContinuouslyUntilEsc): void
    {
        $this->oPresentation->getPresentationProperties()->setLoopContinuouslyUntilEsc($isLoopContinuouslyUntilEsc);

        $this->assertZipFileExists('content.xml');
        $element = '/office:document-content/office:body/office:presentation/presentation:settings';
        $this->assertZipXmlElementExists('content.xml', $element);
        if ($isLoopContinuouslyUntilEsc) {
            $this->assertZipXmlAttributeExists('content.xml', $element, 'presentation:endless');
            $this->assertZipXmlAttributeEquals('content.xml', $element, 'presentation:endless', 'true');
            $this->assertZipXmlAttributeExists('content.xml', $element, 'presentation:pause');
            $this->assertZipXmlAttributeEquals('content.xml', $element, 'presentation:pause', 'PT0S');
            $this->assertZipXmlAttributeExists('content.xml', $element, 'presentation:mouse-visible');
            $this->assertZipXmlAttributeEquals('content.xml', $element, 'presentation:mouse-visible', 'false');
        } else {
            $this->assertZipXmlAttributeNotExists('content.xml', $element, 'presentation:endless');
            $this->assertZipXmlAttributeNotExists('content.xml', $element, 'presentation:pause');
            $this->assertZipXmlAttributeNotExists('content.xml', $element, 'presentation:mouse-visible');
        }
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    /**
     * @return array<array<bool>>
     */
    public static function dataProviderLoopContinuouslyUntilEsc(): array
    {
        return [
            [
                true,
            ],
            [
                false,
            ],
        ];
    }

    public function testRichTextColumnsRTL(): void
    {
        $oRichText = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oRichText->createTextRun('AAA');

        $element = $this->getShapeStyleXPath() . '/style:graphic-properties';

        $this->assertZipXmlAttributeEquals('content.xml', $element, 'style:writing-mode', 'lr-tb');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        // a shape whose every paragraph reads right to left orders its columns the same way
        $this->resetPresentationFile();
        $oRichText->setColumns(2);
        $oRichText->getActiveParagraph()->getAlignment()->setIsRTL(true);

        $this->assertZipXmlAttributeEquals('content.xml', $element, 'style:writing-mode', 'rl-tb');
        $this->assertIsSchemaOpenDocumentValid('1.2');

        // and the order can be said outright, which is the only way to right to left text in
        // left to right columns
        $this->resetPresentationFile();
        $oRichText->setColumnsRTL(false);

        $this->assertZipXmlAttributeEquals('content.xml', $element, 'style:writing-mode', 'lr-tb');
        // the paragraph keeps saying what it says for itself
        $this->assertZipXmlAttributeEquals(
            'content.xml',
            $this->getParagraphStyleXPath() . '/style:paragraph-properties',
            'style:writing-mode',
            'rl-tb'
        );
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    public function testRichTextColumns(): void
    {
        $oRichText = $this->oPresentation->getActiveSlide()->createRichTextShape();
        $oRichText->createTextRun('AAA');

        $element = $this->getShapeStyleXPath() . '/style:graphic-properties/style:columns';

        // a single column is the default, and says nothing
        $this->assertZipXmlElementNotExists('content.xml', $element);
        $this->assertIsSchemaOpenDocumentValid('1.2');

        $this->resetPresentationFile();
        $oRichText->setColumns(3)->setColumnSpacing(20);

        $this->assertZipXmlElementExists('content.xml', $element);
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'fo:column-count', '3');
        // 20 pixels, as the gap is measured everywhere else in the model
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'fo:column-gap', '0.529cm');
        $this->assertIsSchemaOpenDocumentValid('1.2');
    }

    /**
     * The `style:style` of the paragraph of a shape, addressed by the name its `text:p` carries.
     *
     * Spelling the generated name out, or recomputing it the way the writer does, asserts on a
     * definition without asking whether anything points at it -- and a dangling reference then
     * passes. Going through the reference is what a consumer does.
     */
    private function getParagraphStyleXPath(int $frame = 1): string
    {
        return $this->getAutomaticStyleXPath(
            sprintf(
                '/office:document-content/office:body/office:presentation/draw:page/draw:frame[%d]/draw:text-box/text:p',
                $frame
            ),
            'text:style-name'
        );
    }

    /**
     * The `style:style` of the first run of a shape, addressed by the name its `text:span` carries.
     */
    private function getRunStyleXPath(int $frame = 1): string
    {
        return $this->getAutomaticStyleXPath(
            sprintf(
                '/office:document-content/office:body/office:presentation/draw:page/draw:frame[%d]/draw:text-box/text:p/text:span',
                $frame
            ),
            'text:style-name'
        );
    }

    /**
     * The `style:style` of a shape, addressed by the name its `draw:frame` carries.
     */
    private function getSlideStyleXPath(int $page = 1): string
    {
        return $this->getAutomaticStyleXPath(
            sprintf(
                '/office:document-content/office:body/office:presentation/draw:page[%d]',
                $page
            ),
            'draw:style-name'
        );
    }

    private function getShapeStyleXPath(int $frame = 1): string
    {
        return $this->getAutomaticStyleXPath(
            sprintf(
                '/office:document-content/office:body/office:presentation/draw:page/draw:frame[%d]',
                $frame
            ),
            'draw:style-name'
        );
    }

    private function getLineStyleXPath(int $line = 1): string
    {
        return $this->getAutomaticStyleXPath(
            sprintf(
                '/office:document-content/office:body/office:presentation/draw:page/draw:line[%d]',
                $line
            ),
            'draw:style-name'
        );
    }

    /**
     * The `style:style` of a table cell, addressed by the name its `table:table-cell` carries.
     */
    private function getTableCellStyleXPath(int $row = 1, int $cell = 1): string
    {
        return $this->getAutomaticStyleXPath(
            sprintf('//table:table-row[%d]/table:table-cell[%d]', $row, $cell),
            'table:style-name'
        );
    }

    private function getAutomaticStyleXPath(string $referenceXPath, string $referenceAttribute): string
    {
        $styleName = $this->getZipXmlAttributeValue('content.xml', $referenceXPath, $referenceAttribute);

        return '/office:document-content/office:automatic-styles/style:style[@style:name=\'' . $styleName . '\']';
    }

    /**
     * The hash table leaves one `Pictures/` entry for two shapes holding the same image, and the
     * manifest declares that one entry. `draw:image` was written from `getIndexedFilename()`
     * instead, so the second shape named a picture the archive does not carry.
     */
    public function testTwoIdenticalDrawingsShareOneWrittenPart(): void
    {
        $oSlide = $this->oPresentation->getActiveSlide();

        $oShape1 = $oSlide->createDrawingShape();
        $oShape1->setPath(PHPPRESENTATION_TESTS_BASE_DIR . '/resources/images/PhpPresentationLogo.png');
        $oShape2 = $oSlide->createDrawingShape();
        $oShape2->setPath(PHPPRESENTATION_TESTS_BASE_DIR . '/resources/images/PhpPresentationLogo.png');

        $this->assertZipFileExists('Pictures/' . $oShape1->getIndexedFilename());
        $this->assertZipFileNotExists('Pictures/' . $oShape2->getIndexedFilename());
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame[1]/draw:image';
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'xlink:href', 'Pictures/' . $oShape1->getIndexedFilename());
        $element = '/office:document-content/office:body/office:presentation/draw:page/draw:frame[2]/draw:image';
        $this->assertZipXmlAttributeEquals('content.xml', $element, 'xlink:href', 'Pictures/' . $oShape1->getIndexedFilename());
        // Invalid because `draw:image` has attribute `loext:mime-type`
        $this->assertIsSchemaOpenDocumentNotValid('1.2');
    }
}
