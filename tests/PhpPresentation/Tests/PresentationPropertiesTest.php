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
 * @copyright   2009-2015 PHPPresentation contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 * @link        https://github.com/PHPOffice/PHPPresentation
 */

namespace PhpOffice\PhpPresentation\Tests;

use PhpOffice\PhpPresentation\PresentationProperties;

/**
 * Test class for DocumentProperties
 *
 * @coversDefaultClass PhpOffice\PhpPresentation\PresentationProperties
 */
class PresentationPropertiesTest extends \PHPUnit_Framework_TestCase
{
    public function testCommentVisible()
    {
        $object = new PresentationProperties();
        $this->assertFalse($object->isCommentVisible());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PresentationProperties', $object->setCommentVisible('AAAA'));
        $this->assertFalse($object->isCommentVisible());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PresentationProperties', $object->setCommentVisible(true));
        $this->assertTrue($object->isCommentVisible());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PresentationProperties', $object->setCommentVisible(false));
        $this->assertFalse($object->isCommentVisible());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PresentationProperties', $object->setCommentVisible());
        $this->assertFalse($object->isCommentVisible());
    }

    public function testLoopUntilEsc()
    {
        $object = new PresentationProperties();
        $this->assertFalse($object->isLoopContinuouslyUntilEsc());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PresentationProperties', $object->setLoopContinuouslyUntilEsc('AAAA'));
        $this->assertFalse($object->isLoopContinuouslyUntilEsc());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PresentationProperties', $object->setLoopContinuouslyUntilEsc(true));
        $this->assertTrue($object->isLoopContinuouslyUntilEsc());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PresentationProperties', $object->setLoopContinuouslyUntilEsc(false));
        $this->assertFalse($object->isLoopContinuouslyUntilEsc());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PresentationProperties', $object->setLoopContinuouslyUntilEsc());
        $this->assertFalse($object->isLoopContinuouslyUntilEsc());
    }

    public function testLastView()
    {
        $object = new PresentationProperties();
        $this->assertEquals(PresentationProperties::VIEW_SLIDE, $object->getLastView());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PresentationProperties', $object->setLastView('AAAA'));
        $this->assertEquals(PresentationProperties::VIEW_SLIDE, $object->getLastView());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PresentationProperties', $object->setLastView(PresentationProperties::VIEW_OUTLINE));
        $this->assertEquals(PresentationProperties::VIEW_OUTLINE, $object->getLastView());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PresentationProperties', $object->setLastView());
        $this->assertEquals(PresentationProperties::VIEW_SLIDE, $object->getLastView());
    }

    public function testMarkAsFinal()
    {
        $object = new PresentationProperties();
        $this->assertFalse($object->isMarkedAsFinal());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PresentationProperties', $object->markAsFinal('AAAA'));
        $this->assertFalse($object->isMarkedAsFinal());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PresentationProperties', $object->markAsFinal(true));
        $this->assertTrue($object->isMarkedAsFinal());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PresentationProperties', $object->markAsFinal(false));
        $this->assertFalse($object->isMarkedAsFinal());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PresentationProperties', $object->markAsFinal());
        $this->assertTrue($object->isMarkedAsFinal());
    }

    public function testThumbnail()
    {
        $imagePath = PHPPRESENTATION_TESTS_BASE_DIR.DIRECTORY_SEPARATOR.'resources'.DIRECTORY_SEPARATOR.'images'.DIRECTORY_SEPARATOR.'PhpPresentationLogo.png';

        $object = new PresentationProperties();
        $this->assertNull($object->getThumbnailPath());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PresentationProperties', $object->setThumbnailPath('AAAA'));
        $this->assertNull($object->getThumbnailPath());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PresentationProperties', $object->setThumbnailPath());
        $this->assertNull($object->getThumbnailPath());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PresentationProperties', $object->setThumbnailPath($imagePath));
        $this->assertEquals($imagePath, $object->getThumbnailPath());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PresentationProperties', $object->setThumbnailPath());
        $this->assertEquals($imagePath, $object->getThumbnailPath());
    }

    public function testZoom()
    {
        $object = new PresentationProperties();
        $this->assertEquals(1, $object->getZoom());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PresentationProperties', $object->setZoom('AAAA'));
        $this->assertEquals(1, $object->getZoom());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PresentationProperties', $object->setZoom(2.3));
        $this->assertEquals(2.3, $object->getZoom());
        $this->assertInstanceOf('PhpOffice\\PhpPresentation\\PresentationProperties', $object->setZoom());
        $this->assertEquals(1, $object->getZoom());
    }
}
