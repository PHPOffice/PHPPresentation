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

namespace PhpOffice\PhpPresentation\Tests\Reader;

use ArrayObject;
use PhpOffice\PhpPresentation\Exception\FileNotFoundException;
use PhpOffice\PhpPresentation\Exception\InvalidFileFormatException;
use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Reader\Serialized;
use PhpOffice\PhpPresentation\Shape;
use PhpOffice\PhpPresentation\Writer\Serialized as SerializedWriter;
use PHPUnit\Framework\Attributes\DataProvider;
use PHPUnit\Framework\TestCase;
use ReflectionClass;
use ReflectionProperty;
use ZipArchive;

/**
 * Test class for serialized reader.
 *
 * @coversDefaultClass \PhpOffice\PhpPresentation\Reader\Serialized
 */
class SerializedTest extends TestCase
{
    /**
     * Test can read.
     */
    public function testCanRead(): void
    {
        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/serialized.phppt';
        $object = new Serialized();

        self::assertTrue($object->canRead($file));
    }

    public function testLoadFileNotExists(): void
    {
        $this->expectException(FileNotFoundException::class);
        $this->expectExceptionMessage('The file "" doesn\'t exist');

        $object = new Serialized();
        $object->load('');
    }

    public function testLoadFileBadFormat(): void
    {
        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Sample_01_Simple.pptx';
        $this->expectException(InvalidFileFormatException::class);
        $this->expectExceptionMessage(sprintf(
            'The file %s is not in the format supported by class PhpOffice\PhpPresentation\Reader\Serialized',
            $file
        ));

        $object = new Serialized();
        $object->load($file);
    }

    public function testFileSupportsNotExists(): void
    {
        $this->expectException(FileNotFoundException::class);
        $this->expectExceptionMessage('The file "" doesn\'t exist');

        $object = new Serialized();
        $object->fileSupportsUnserializePhpPresentation('');
    }

    public function testLoadSerializedFileNotExists(): void
    {
        $file = tempnam(sys_get_temp_dir(), 'PhpPresentation_Serialized');

        $this->expectException(InvalidFileFormatException::class);
        $this->expectExceptionMessage(sprintf(
            'The file %s is not in the format supported by class PhpOffice\PhpPresentation\Reader\Serialized (The file PhpPresentation.xml is malformed)',
            $file
        ));

        $oArchive = new ZipArchive();
        $oArchive->open($file, ZipArchive::OVERWRITE);
        $oArchive->addFromString('PhpPresentation.xml', '');
        $oArchive->close();

        $object = new Serialized();
        $object->load($file);
    }

    public function testLoadExistingSerializedFixture(): void
    {
        $loaded = (new Serialized())->load(PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/serialized.phppt');
        self::assertGreaterThan(0, $loaded->getSlideCount());
    }

    public function testWriterRoundTrip(): void
    {
        $presentation = new PhpPresentation();
        $presentation->getActiveSlide()->createRichTextShape()->createTextRun('Round trip');
        $row = $presentation->getActiveSlide()->createTableShape(2)->createRow();
        $row->getCell(0)->createTextRun('Product');
        $row->getCell(1)->createTextRun('Price');
        $loaded = $this->roundTrip($presentation);
        $shape = $loaded->getActiveSlide()->getShapeCollection()[0];
        self::assertInstanceOf(Shape\RichText::class, $shape);
        self::assertSame('Round trip', $shape->getPlainText());
        $table = $loaded->getActiveSlide()->getShapeCollection()[1];
        self::assertInstanceOf(Shape\Table::class, $table);
        self::assertSame(2, $table->getNumColumns());
        self::assertSame('Product', $table->getRow(0)->getCell(0)->getPlainText());
        self::assertSame('Price', $table->getRow(0)->getCell(1)->getPlainText());
    }

    public function testImageRoundTrip(): void
    {
        $presentation = new PhpPresentation();
        $path = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/images/PhpPresentationLogo.png';
        $presentation->getActiveSlide()->createDrawingShape()->setPath($path);
        $filename = tempnam(sys_get_temp_dir(), 'SerializedImage');

        try {
            (new SerializedWriter($presentation))->save($filename);
            $loaded = (new Serialized())->load($filename)->getActiveSlide()->getShapeCollection()[0];
            self::assertInstanceOf(Shape\Drawing\File::class, $loaded);
            self::assertStringStartsWith('zip://' . $filename . '#media/', $loaded->getPath());
            self::assertSame(file_get_contents($path), file_get_contents($loaded->getPath()));
        } finally {
            unlink($filename);
        }
    }

    public function testBase64ImageRoundTrip(): void
    {
        $path = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/images/PhpPresentationLogo.png';
        $contents = file_get_contents($path);
        $drawing = new Shape\Drawing\Base64();
        $drawing->setPath($path);
        $drawing->setData('data:image/png;base64,' . base64_encode($contents));
        $presentation = new PhpPresentation();
        $presentation->getActiveSlide()->addShape($drawing);
        $loaded = $this->roundTrip($presentation)->getActiveSlide()->getShapeCollection()[0];
        self::assertInstanceOf(Shape\Drawing\Base64::class, $loaded);
        self::assertSame($contents, $loaded->getContents());
    }

    public function testZipImageRoundTrip(): void
    {
        $path = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/images/PhpPresentationLogo.png';
        $contents = file_get_contents($path);
        $source = tempnam(sys_get_temp_dir(), 'SerializedSourceImage');
        $filename = tempnam(sys_get_temp_dir(), 'SerializedZipImage');
        $archive = new ZipArchive();
        $archive->open($source, ZipArchive::OVERWRITE);
        $archive->addFromString('media/image.png', $contents);
        $archive->close();

        try {
            $drawing = new Shape\Drawing\ZipFile();
            $drawing->setPath('zip://' . $source . '#media/image.png');
            $presentation = new PhpPresentation();
            $presentation->getActiveSlide()->addShape($drawing);
            (new SerializedWriter($presentation))->save($filename);
            $loaded = (new Serialized())->load($filename)->getActiveSlide()->getShapeCollection()[0];
            self::assertInstanceOf(Shape\Drawing\ZipFile::class, $loaded);
            self::assertStringStartsWith('zip://' . $filename . '#media/', $loaded->getPath());
            // Verify the imported image does not depend on the source archive.
            unlink($source);
            self::assertSame($contents, $loaded->getContents());
        } finally {
            if (file_exists($source)) {
                unlink($source);
            }
            unlink($filename);
        }
    }

    public function testSlideMasterRoundTrip(): void
    {
        $presentation = new PhpPresentation();
        $presentation->setAllMasterSlides([]);
        $master = $presentation->createMasterSlide();
        $master->createRichTextShape()->createTextRun('Master text');
        $layout = $master->createSlideLayout();
        $presentation->getActiveSlide()->setSlideLayout($layout);
        $presentation->setAllMasterSlides(new ArrayObject($presentation->getAllMasterSlides()));
        $loaded = $this->roundTrip($presentation);
        $loadedLayout = $loaded->getActiveSlide()->getSlideLayout();
        self::assertNotNull($loadedLayout);
        $masters = $loaded->getAllMasterSlides();
        self::assertInstanceOf(ArrayObject::class, $masters);
        self::assertCount(1, $masters);
        self::assertSame($masters[0]->getAllSlideLayouts()[0], $loadedLayout);
        $parent = $masters[0]->getParent();
        self::assertNotNull($parent);
        self::assertSame($masters[0], $parent->getAllMasterSlides()[0]);
        $shape = $masters[0]->getShapeCollection()[0];
        self::assertInstanceOf(Shape\RichText::class, $shape);
        self::assertSame('Master text', $shape->getPlainText());
    }

    public function testRejectsMalformedXml(): void
    {
        $this->expectException(InvalidFileFormatException::class);
        $this->loadXml('<PhpPresentation><data>unterminated');
    }

    public function testRejectsMalformedBase64(): void
    {
        $this->expectException(InvalidFileFormatException::class);
        $this->loadXml('<PhpPresentation><data>%%%invalid%%%</data></PhpPresentation>');
    }

    public function testRejectsMissingData(): void
    {
        $this->expectException(InvalidFileFormatException::class);
        $this->loadXml('<PhpPresentation/>');
    }

    public function testRejectsRecursiveArray(): void
    {
        $recursive = [];
        $recursive['self'] = &$recursive;
        $presentation = new PhpPresentation();
        $presentation->getDocumentProperties()->setCustomProperty('recursive', $recursive);
        $this->expectException(InvalidFileFormatException::class);
        $this->loadPayload(serialize($presentation));
    }

    public function testRejectsObjectInjection(): void
    {
        SerializedWakeupProbe::$awakened = false;

        try {
            $this->loadPayload(serialize(new SerializedWakeupProbe()));
            self::fail('An unexpected object must be rejected');
        } catch (InvalidFileFormatException $exception) {
            self::assertFalse(SerializedWakeupProbe::$awakened);
        }
    }

    public function testRejectsNestedObjectInjection(): void
    {
        SerializedWakeupProbe::$awakened = false;
        $presentation = new PhpPresentation();
        $presentation->getDocumentProperties()->setCustomProperty('payload', new SerializedWakeupProbe());

        try {
            $this->loadPayload(serialize($presentation));
            self::fail('A nested unexpected object must be rejected');
        } catch (InvalidFileFormatException $exception) {
            self::assertFalse(SerializedWakeupProbe::$awakened);
        }
    }

    /** @return array<string, array{bool, bool}> */
    public static function deletionGadgetLocations(): array
    {
        return [
            'automatic, root' => [false, false],
            'automatic, nested' => [false, true],
            'direct, root' => [true, false],
            'direct, nested' => [true, true],
        ];
    }

    /** @dataProvider deletionGadgetLocations */
    #[DataProvider('deletionGadgetLocations')]
    public function testRejectsFileDeletionGadget(bool $direct, bool $nested): void
    {
        $target = tempnam(sys_get_temp_dir(), 'SerializedVictim');
        $class = 'PhpOffice\\Common\\XMLWriter';
        $property = "\0" . $class . "\0tempFileName";
        $payload = 'O:' . strlen($class) . ':"' . $class . '":1:{s:' . strlen($property)
            . ':"' . $property . '";s:' . strlen($target) . ':"' . $target . '";}';
        if ($nested) {
            $presentation = new PhpPresentation();
            $presentation->getDocumentProperties()->setCustomProperty('gadget', 'deletion-placeholder');
            $payload = str_replace(serialize('deletion-placeholder'), $payload, serialize($presentation));
        }

        try {
            try {
                $this->loadXml('<p><data>' . base64_encode($payload) . '</data></p>', $direct);
                self::fail('The deletion gadget must be rejected');
            } catch (InvalidFileFormatException $exception) {
                clearstatcache();
                self::assertFileExists($target);
            }
        } finally {
            if (file_exists($target)) {
                unlink($target);
            }
        }
    }

    public function testAllowedClassesHaveNoAutomaticUserlandHooks(): void
    {
        $reader = new ReflectionClass(Serialized::class);
        $classes = $reader->getReflectionConstant('ALLOWED_CLASSES')->getValue();
        foreach ($classes as $class) {
            $reflection = new ReflectionClass($class);
            self::assertTrue($reflection->isInstantiable(), $class);
            // ArrayObject's internal restoration handler is checked by DataValidator.
            if ($class === ArrayObject::class) {
                continue;
            }
            foreach (['__wakeup', '__unserialize', '__destruct'] as $hook) {
                self::assertFalse($reflection->hasMethod($hook), $class . '::' . $hook);
            }
        }
    }

    public function testRejectsGdRenderingCallback(): void
    {
        $presentation = new PhpPresentation();
        $drawing = new Shape\Drawing\Gd();
        $drawing->setPath('image.png');
        // A scalar argument and arbitrary callable survive PHP serialization.
        $drawing->setRenderingFunction('strlen');
        $property = new ReflectionProperty($drawing, 'imageResource');
        if (PHP_VERSION_ID < 80100) {
            $property->setAccessible(true);
        }
        $property->setValue($drawing, 'attacker-controlled argument');
        $presentation->getActiveSlide()->addShape($drawing);
        $this->expectException(InvalidFileFormatException::class);
        $this->loadPayload(serialize($presentation));
    }

    public function testRejectsMalformedAllowedObject(): void
    {
        $class = PhpPresentation::class;
        $property = "\0" . $class . "\0slideCollection";
        $payload = 'O:' . strlen($class) . ':"' . $class . '":1:{s:' . strlen($property)
            . ':"' . $property . '";s:3:"bad";}';
        $this->expectException(InvalidFileFormatException::class);
        $this->loadPayload($payload);
    }

    public function testRejectsRepeatedArrayExpansion(): void
    {
        $branches = [['leaf']];
        for ($i = 1; $i <= 20; ++$i) {
            $branches[$i] = [&$branches[$i - 1], &$branches[$i - 1]];
        }
        $presentation = new PhpPresentation();
        $presentation->getDocumentProperties()->setCustomProperty('branches', $branches[20]);
        $this->expectException(InvalidFileFormatException::class);
        $this->loadPayload(serialize($presentation));
    }

    public function testForbiddenClassIsNotAutoloaded(): void
    {
        // Sharing the library namespace must not grant permission to instantiate.
        $class = 'PhpOffice\\PhpPresentation\\UntrustedSerializedShape';
        $autoloaded = [];
        $loader = static function (string $name) use (&$autoloaded, $class): void {
            if ($name === $class) {
                $autoloaded[] = $name;
            }
        };
        spl_autoload_register($loader, true, true);

        try {
            try {
                $this->loadPayload('O:' . strlen($class) . ':"' . $class . '":0:{}');
                self::fail('An unknown class must be rejected');
            } catch (InvalidFileFormatException $exception) {
                self::assertSame([], $autoloaded);
            }
        } finally {
            spl_autoload_unregister($loader);
        }
    }

    public function testRejectsPresentationSubclass(): void
    {
        SerializedPresentationProbe::$awakened = false;

        try {
            $this->loadPayload(serialize(new SerializedPresentationProbe()));
            self::fail('A subclass must not inherit deserialization permission');
        } catch (InvalidFileFormatException $exception) {
            self::assertFalse(SerializedPresentationProbe::$awakened);
        }
    }

    public function testRejectsNonPresentationRoots(): void
    {
        foreach ([false, null, 42, 'text', [], new Shape\RichText()] as $value) {
            try {
                $this->loadPayload(serialize($value));
                self::fail('Only a presentation may be the root value');
            } catch (InvalidFileFormatException $exception) {
                self::assertStringContainsString('Invalid serialized presentation data', $exception->getMessage());
            }
        }
    }

    public function testAcceptsNestingAtLimit(): void
    {
        $value = 'leaf';
        // Four property/array steps lead from the root to the custom property's value.
        for ($i = 0; $i < 508; ++$i) {
            $value = [$value];
        }
        $presentation = new PhpPresentation();
        $presentation->getDocumentProperties()->setCustomProperty('nested', $value);
        $loaded = $this->loadPayload(serialize($presentation));
        $loadedValue = $loaded->getDocumentProperties()->getCustomPropertyValue('nested');
        self::assertSame($value, $loadedValue);
    }

    public function testRejectsNestingOnePastLimit(): void
    {
        $value = 'leaf';
        for ($i = 0; $i < 509; ++$i) {
            $value = [$value];
        }
        $presentation = new PhpPresentation();
        $presentation->getDocumentProperties()->setCustomProperty('nested', $value);
        $this->expectException(InvalidFileFormatException::class);
        $this->loadPayload(serialize($presentation));
    }

    public function testAcceptsSharedArrayReferences(): void
    {
        $shared = ['text', 42, true, null];
        $presentation = new PhpPresentation();
        $presentation->getDocumentProperties()->setCustomProperty('shared', [&$shared, &$shared]);
        $loaded = $this->loadPayload(serialize($presentation));
        self::assertSame([$shared, $shared], $loaded->getDocumentProperties()->getCustomPropertyValue('shared'));
    }

    public function testRejectsTruncatedPresentation(): void
    {
        $this->expectException(InvalidFileFormatException::class);
        $this->loadPayload(substr(serialize(new PhpPresentation()), 0, -1));
    }

    public function testRejectsUnserializeAndDestructorHooks(): void
    {
        $payload = serialize(new SerializedMagicMethodProbe());
        SerializedMagicMethodProbe::$called = false;

        try {
            $this->loadPayload($payload);
            self::fail('Forbidden magic methods must never execute');
        } catch (InvalidFileFormatException $exception) {
            self::assertFalse(SerializedMagicMethodProbe::$called);
        }
    }

    public function testDoesNotExpandExternalXmlEntities(): void
    {
        $filename = tempnam(sys_get_temp_dir(), 'SerializedEntity');
        file_put_contents($filename, base64_encode(serialize(new PhpPresentation())));

        try {
            $this->expectException(InvalidFileFormatException::class);
            $this->loadXml('<!DOCTYPE p [<!ENTITY payload SYSTEM "file://' . $filename
                . '">]><p><data>&payload;</data></p>');
        } finally {
            unlink($filename);
        }
    }

    public function testRejectsUnsafeArrayObjectIteratorWithoutAutoload(): void
    {
        $class = 'UntrustedSerializedIterator';
        $presentation = new PhpPresentation();
        $presentation->setAllMasterSlides(new ArrayObject($presentation->getAllMasterSlides()));
        $payload = str_replace('i:3;N;', 'i:3;s:' . strlen($class) . ':"' . $class . '";', serialize($presentation));
        $autoloaded = false;
        $loader = static function (string $name) use (&$autoloaded, $class): void {
            if ($name === $class) {
                $autoloaded = true;
            }
        };
        spl_autoload_register($loader, true, true);

        try {
            try {
                $this->loadPayload($payload);
                self::fail('An unsafe iterator must be rejected before autoloading');
            } catch (InvalidFileFormatException $exception) {
                self::assertFalse($autoloaded);
            }
        } finally {
            spl_autoload_unregister($loader);
        }
    }

    public function testRejectsForbiddenObjectInArrayObjectStorage(): void
    {
        foreach ([0, ArrayObject::STD_PROP_LIST, ArrayObject::ARRAY_AS_PROPS] as $flags) {
            SerializedWakeupProbe::$awakened = false;
            $presentation = new PhpPresentation();
            $presentation->getDocumentProperties()->setCustomProperty('storage', new ArrayObject([new SerializedWakeupProbe()], $flags));

            try {
                $this->loadPayload(serialize($presentation));
                self::fail('ArrayObject storage must be checked regardless of flags');
            } catch (InvalidFileFormatException $exception) {
                self::assertFalse(SerializedWakeupProbe::$awakened);
            }
        }
    }

    public function testAcceptsSafeGdDrawing(): void
    {
        $presentation = new PhpPresentation();
        $drawing = new Shape\Drawing\Gd();
        $drawing->setPath('image.png');
        $presentation->getActiveSlide()->addShape($drawing);
        $loaded = $this->loadPayload(serialize($presentation));
        $image = $loaded->getActiveSlide()->getShapeCollection()[0];
        self::assertInstanceOf(Shape\Drawing\Gd::class, $image);
        self::assertSame('imagepng', $image->getRenderingFunction());
    }

    public function testRejectsAllocationBombBeforeUnserialize(): void
    {
        // A tiny payload declaring billions of entries must never allocate a PHP array.
        $this->expectException(InvalidFileFormatException::class);
        $this->loadPayload('a:999999999:{}');
    }

    public function testRejectsTooManySerializedValues(): void
    {
        $this->expectException(InvalidFileFormatException::class);
        $presentation = new PhpPresentation();
        $presentation->getDocumentProperties()->setCustomProperty('manyValues', array_fill(0, 50001, null));
        $this->loadPayload(serialize($presentation));
    }

    public function testRejectsTrailingPayloadData(): void
    {
        $this->expectException(InvalidFileFormatException::class);
        $this->loadPayload(serialize(new PhpPresentation()) . 'trailing');
    }

    public function testScalarAndBinaryDataRoundTrip(): void
    {
        $values = [true, false, null, -42, 1.5, -0.0, 1.0e100, INF, -INF, "\0binary\xff", '多字节字符串'];
        $presentation = new PhpPresentation();
        $presentation->getDocumentProperties()->setCustomProperty('values', $values);
        $loaded = $this->loadPayload(serialize($presentation));
        self::assertSame($values, $loaded->getDocumentProperties()->getCustomPropertyValue('values'));
    }

    public function testAllChartTypesRoundTrip(): void
    {
        $classes = [Shape\Chart\Type\Area::class, Shape\Chart\Type\Bar::class,
            Shape\Chart\Type\Bar3D::class, Shape\Chart\Type\Doughnut::class,
            Shape\Chart\Type\Line::class, Shape\Chart\Type\Pie::class,
            Shape\Chart\Type\Pie3D::class, Shape\Chart\Type\Radar::class,
            Shape\Chart\Type\Scatter::class];
        foreach ($classes as $class) {
            $presentation = new PhpPresentation();
            $chart = $presentation->getActiveSlide()->createChartShape();
            $type = new $class();
            $type->addSeries(new Shape\Chart\Series('Series', ['A' => '1', 'B' => '2']));
            $chart->getPlotArea()->setType($type);
            $loaded = $this->roundTrip($presentation)->getActiveSlide()->getShapeCollection()[0];
            self::assertInstanceOf(Shape\Chart::class, $loaded);
            self::assertInstanceOf($class, $loaded->getPlotArea()->getType());
            self::assertSame(['A' => '1', 'B' => '2'], $loaded->getPlotArea()->getType()->getSeries()[0]->getValues());
        }
    }

    public function testRejectsCustomSerializationHandler(): void
    {
        $this->expectException(InvalidFileFormatException::class);
        $this->loadPayload('C:11:"ArrayObject":21:{x:i:0;a:0:{};m:a:0:{}}');
    }

    public function testRejectsEnumWithoutAutoload(): void
    {
        $autoloaded = false;
        $loader = static function (string $class) use (&$autoloaded): void {
            if ($class === 'UnknownEnum') {
                $autoloaded = true;
            }
        };
        spl_autoload_register($loader, true, true);

        try {
            try {
                $this->loadPayload('E:17:"UnknownEnum:Value";');
                self::fail('Enums must be rejected before autoloading');
            } catch (InvalidFileFormatException $exception) {
                self::assertFalse($autoloaded);
            }
        } finally {
            spl_autoload_unregister($loader);
        }
    }

    private function roundTrip(PhpPresentation $presentation): PhpPresentation
    {
        $filename = tempnam(sys_get_temp_dir(), 'SerializedRoundTrip');

        try {
            (new SerializedWriter($presentation))->save($filename);

            return (new Serialized())->load($filename);
        } finally {
            unlink($filename);
        }
    }

    private function loadPayload(string $payload): PhpPresentation
    {
        return $this->loadXml('<PhpPresentation><data>' . base64_encode($payload) . '</data></PhpPresentation>');
    }

    private function loadXml(string $xml, bool $direct = false): PhpPresentation
    {
        $filename = tempnam(sys_get_temp_dir(), 'SerializedPayload');
        $archive = new ZipArchive();
        $archive->open($filename, ZipArchive::OVERWRITE);
        $archive->addFromString('PhpPresentation.xml', $xml);
        $archive->close();

        try {
            return $direct ? (new Serialized())->load($filename) : IOFactory::load($filename);
        } finally {
            unlink($filename);
        }
    }
}

class SerializedWakeupProbe
{
    public static bool $awakened = false;

    public function __wakeup(): void
    {
        self::$awakened = true;
    }
}

class SerializedPresentationProbe extends PhpPresentation
{
    public static bool $awakened = false;

    public function __wakeup(): void
    {
        self::$awakened = true;
    }
}

class SerializedMagicMethodProbe
{
    public static bool $called = false;

    /** @param array<mixed> $data */
    public function __unserialize(array $data): void
    {
        // This probe is serialized without any instance properties.
        self::$called = $data === [];
    }

    public function __destruct()
    {
        self::$called = true;
    }
}
