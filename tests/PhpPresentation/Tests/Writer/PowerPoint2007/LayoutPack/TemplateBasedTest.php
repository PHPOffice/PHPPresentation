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
 * @link        https://github.com/PHPOffice/PHPPresentation
 * @copyright   2009-2015 PHPPresentation contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

namespace PhpOffice\PhpPresentation\Tests\Writer\PowerPoint2007\LayoutPack;

use PhpOffice\PhpPresentation\Writer\PowerPoint2007\LayoutPack\PackDefault;
use PhpOffice\PhpPresentation\Writer\PowerPoint2007\LayoutPack\TemplateBased;

/**
 * Test class for TemplateBased
 *
 * @deprecated 0.7
 * @coversDefaultClass TemplateBased
 */
class TemplateBasedTest extends \PHPUnit_Framework_TestCase
{
    public function testFindLayout()
    {
        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Sample_00_01.pptx';
        $templateBased = new TemplateBased($file);
        $layouts = $templateBased->getLayouts();

        foreach ($layouts as $layout) {
            $foundLayout = $templateBased->findLayout($layout['name']);
            $this->assertEquals($layout, $foundLayout);
        }
    }

    /**
     * @expectedException \Exception
     */
    public function testFindLayoutException()
    {
        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Sample_00_01.pptx';
        $templateBased = new TemplateBased($file);
        $name = 'Invalid';
        $templateBased->findLayout($name);
    }

    public function testFindLayoutId()
    {
        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Sample_00_01.pptx';
        $templateBased = new TemplateBased($file);
        $layouts = $templateBased->getLayouts();

        foreach ($layouts as $layout) {
            $foundLayoutId = $templateBased->findLayoutId($layout['name']);
            $this->assertEquals($layout['id'], $foundLayoutId);
        }
    }

    /**
     * @expectedException \Exception
     */
    public function testFindLayoutIdException()
    {
        $file = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Sample_00_01.pptx';
        $templateBased = new TemplateBased($file);
        $name = 'Invalid';
        $templateBased->findLayoutId($name);
    }

    public function testFindLayoutName()
    {
        $oLayout = new PackDefault();
        foreach ($oLayout->getLayouts() as $keyLayout => $layout) {
            $foundLayoutName = $oLayout->findLayoutName($keyLayout);
            $this->assertEquals($layout['name'], $foundLayoutName);
        }
    }

    /**
     * @expectedException \Exception
     */
    public function testFindLayoutNameException()
    {
        $oLayout = new PackDefault();
        $oLayout->findLayoutName(1000);
    }
}
