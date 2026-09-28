<?php

declare(strict_types=1);

namespace PhpPresentation\Tests\Writer\PowerPoint2007;

use DK\OpenXml\Exception\InvalidPasswordException;
use DK\OpenXml\Exception\OpenXmlException;
use DK\OpenXml\OfficeFileDetector;
use DK\OpenXml\OfficeFileFormat;
use PhpOffice\PhpPresentation\Exception\InvalidParameterException;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Reader\PowerPoint2007 as PowerPoint2007Reader;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Writer\PowerPoint2007 as PowerPoint2007Writer;
use PHPUnit\Framework\TestCase;

/**
 * A protected presentation is the package wrapped in a compound file, which is a thing neither
 * Writer could produce and neither Reader could open.
 */
class EncryptionTest extends TestCase
{
    /**
     * @var string
     */
    private $file = '';

    protected function tearDown(): void
    {
        if ('' !== $this->file) {
            @unlink($this->file);
            $this->file = '';
        }
    }

    private function write(?string $password): string
    {
        $oPhpPresentation = new PhpPresentation();
        $oPhpPresentation->getActiveSlide()->createRichTextShape()->createTextRun('Delta');

        $this->file = (string) tempnam(sys_get_temp_dir(), 'PhpPresentation');
        $oWriter = new PowerPoint2007Writer($oPhpPresentation);
        if (null !== $password) {
            $oWriter->setEncryptionPassword($password);
        }
        $oWriter->save($this->file);

        return $this->file;
    }

    public function testAPresentationWithNoPasswordIsAPlainPackage(): void
    {
        self::assertEquals(OfficeFileFormat::OpcPackage, OfficeFileDetector::detect($this->write(null)));
    }

    public function testAProtectedPresentationSurvivesTheRoundTrip(): void
    {
        $file = $this->write('a strong password');

        self::assertEquals(OfficeFileFormat::EncryptedOpcPackage, OfficeFileDetector::detect($file));

        $oReader = new PowerPoint2007Reader();
        $oReader->setEncryptionPassword('a strong password');
        $oPhpPresentation = $oReader->load($file);

        self::assertEquals(1, $oPhpPresentation->getSlideCount());
        $arrayShape = array_values((array) $oPhpPresentation->getActiveSlide()->getShapeCollection());
        self::assertCount(1, $arrayShape);
        self::assertInstanceOf(RichText::class, $arrayShape[0]);
        self::assertEquals('Delta', $arrayShape[0]->getParagraph()->getRichTextElements()[0]->getText());
    }

    public function testAProfileThatCannotBeWrittenIsRefused(): void
    {
        $oWriter = new PowerPoint2007Writer(new PhpPresentation());

        $this->expectException(InvalidParameterException::class);
        $oWriter->setEncryptionProfile(128, 'SHA1');
    }

    public function testTheSpinCountIsHonoured(): void
    {
        $oPhpPresentation = new PhpPresentation();
        $oPhpPresentation->getActiveSlide()->createRichTextShape()->createTextRun('Delta');

        $this->file = (string) tempnam(sys_get_temp_dir(), 'PhpPresentation');
        $oWriter = new PowerPoint2007Writer($oPhpPresentation);
        $oWriter->setEncryptionPassword('a strong password')->setEncryptionProfile(256, 'SHA512', 1000);
        $oWriter->save($this->file);

        // a reader that allows less work than the file asks for refuses it
        $oReader = new PowerPoint2007Reader();
        $oReader->setEncryptionPassword('a strong password')->setMaxEncryptionSpinCount(999);

        $this->expectException(OpenXmlException::class);
        $oReader->load($this->file);
    }

    public function testTheWrongPasswordIsRefused(): void
    {
        $file = $this->write('a strong password');

        $oReader = new PowerPoint2007Reader();
        $oReader->setEncryptionPassword('the wrong password');

        $this->expectException(InvalidPasswordException::class);
        $oReader->load($file);
    }
}
