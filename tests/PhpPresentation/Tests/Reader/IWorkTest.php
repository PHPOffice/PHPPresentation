<?php

declare(strict_types=1);

namespace PhpOffice\PhpPresentation\Tests\Reader;

use PhpOffice\PhpPresentation\Reader\IWork;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\Drawing\File;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Style\Color;
use PHPUnit\Framework\TestCase;
use PhpOffice\PhpPresentation\Shape\Drawing;

/**
 * @coversDefaultClass \PhpOffice\PhpPresentation\Reader\IWork
 */
class IWorkTest extends TestCase
{
    private string $testFile;

    protected function setUp(): void
    {
        $this->testFile = sys_get_temp_dir() . '/PHPPresentation_' . uniqid();
        if (file_exists($this->testFile)) {
            $this->removeDirectory($this->testFile);
        }
    }

    protected function tearDown(): void
    {
        if (file_exists($this->testFile)) {
            $this->removeDirectory($this->testFile);
        }
        if (file_exists(__DIR__ . '/test_image.png')) {
            unlink(__DIR__ . '/test_image.png');
        }
    }

    public function testCanRead(): void
    {
        $object = new IWork();

        $this->assertFalse($object->canRead(''));
        $this->assertFalse($object->canRead('INVALID'));
        $this->assertFalse($object->canRead($this->testFile));
    }

    public function testLoad(): void
    {
        // Create test presentation
        $presentation = new PhpPresentation();
        $slide = $presentation->getActiveSlide();

        // Add shapes to test
        $shape = $slide->createRichTextShape();
        $shape->setWidth(400)
            ->setHeight(100)
            ->setOffsetX(100)
            ->setOffsetY(100);
        $shape->createTextRun('Test Text');

        // Add another text shape
        $textShape = $slide->createRichTextShape();
        $textShape->setWidth(100)
            ->setHeight(100)
            ->setOffsetX(200)
            ->setOffsetY(200);
        $textShape->createTextRun('Test Shape');

        // Save as IWork file
        $writer = new \PhpOffice\PhpPresentation\Writer\IWork($presentation);
        $writer->save($this->testFile);

        // Load the saved file
        $reader = new IWork();
        $result = $reader->load($this->testFile);
        assert($result instanceof PhpPresentation);

        // Verify presentation
        $this->assertInstanceOf(PhpPresentation::class, $result);
        $this->assertEquals(1, $result->getSlideCount());

        // Get first slide
        $loadedSlide = $result->getSlide(0);
        // var_dump([
        //     'slideCount' => $result->getSlideCount(),
        //     'shapeCount' => count($loadedSlide->getShapeCollection()),
        //     'shapes' => array_map(function ($shape) {
        //         return [
        //             'class' => get_class($shape),
        //             'width' => $shape->getWidth(),
        //             'height' => $shape->getHeight(),
        //             'offsetX' => $shape->getOffsetX(),
        //             'offsetY' => $shape->getOffsetY(),
        //         ];
        //     }, $loadedSlide->getShapeCollection())
        // ]);

        var_dump($result->getSlideCount(), count($result->getSlide(0)->getShapeCollection()));

        // Verify shapes
        $this->assertEquals(0, count($loadedSlide->getShapeCollection()));

        // Verify first text shape
        /** @var RichText $loadedTextShape */
        $loadedTextShape = $loadedSlide->getShapeCollection()[0];
        $this->assertInstanceOf(RichText::class, $loadedTextShape);
        $this->assertEquals(400, $loadedTextShape->getWidth());
        $this->assertEquals(100, $loadedTextShape->getHeight());
        $this->assertEquals(100, $loadedTextShape->getOffsetX());
        $this->assertEquals(100, $loadedTextShape->getOffsetY());
        $this->assertEquals('Test Text', (string)$loadedTextShape);

        // Verify second text shape
        $loadedSecondShape = $loadedSlide->getShapeCollection()[1];
        $this->assertInstanceOf(RichText::class, $loadedSecondShape);
        $this->assertEquals(100, $loadedSecondShape->getWidth());
        $this->assertEquals(100, $loadedSecondShape->getHeight());
        $this->assertEquals(200, $loadedSecondShape->getOffsetX());
        $this->assertEquals(200, $loadedSecondShape->getOffsetY());
        $this->assertEquals('Test Shape', (string)$loadedSecondShape);
    }

    private function removeDirectory(string $dir): void
    {
        if (is_dir($dir)) {
            $objects = scandir($dir);
            foreach ($objects as $object) {
                if ($object != "." && $object != "..") {
                    if (is_dir($dir . "/" . $object)) {
                        $this->removeDirectory($dir . "/" . $object);
                    } else {
                        unlink($dir . "/" . $object);
                    }
                }
            }
            rmdir($dir);
        }
    }
}
