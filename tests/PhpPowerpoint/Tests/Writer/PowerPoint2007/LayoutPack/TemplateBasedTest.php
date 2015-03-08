<?php
/**
 * This file is part of PHPPowerPoint - A pure PHP library for reading and writing
 * presentations documents.
 *
 * PHPPowerPoint is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPWord/contributors.
 *
 * @link        https://github.com/PHPOffice/PHPPowerPoint
 * @copyright   2009-2014 PHPPowerPoint contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

namespace PhpOffice\PhpPowerpoint\Tests\Writer\PowerPoint2007\LayoutPack;

use PhpOffice\PhpPowerpoint\Writer\PowerPoint2007\LayoutPack\TemplateBased;

/**
 * Test class for TemplateBased
 *
 * @coversDefaultClass TemplateBased
 */
class TemplateBasedTest extends \PHPUnit_Framework_TestCase
{
    /**
     * Test create new instance
     */
    public function testConstruct()
    {
        $file = PHPPOWERPOINT_TESTS_BASE_DIR . '/resources/files/Sample_00_01.pptx';
        $templateBased = new TemplateBased($file);
        $layouts = $templateBased->getLayouts();

        $expectedSlideLayoutId = 1;
        foreach ($layouts as $layoutId => $layout) {
            $this->assertEquals($expectedSlideLayoutId, $layoutId);
            ++$expectedSlideLayoutId;
        }

        $this->markTestIncomplete(
          'This test has not been fully implemented yet.'
        );
    }

    public function testFindLayout()
    {
        $file = PHPPOWERPOINT_TESTS_BASE_DIR . '/resources/files/Sample_00_01.pptx';
        $templateBased = new TemplateBased($file);
        $layouts = $templateBased->getLayouts();

        foreach ($layouts as $layoutId => $layout) {
            $foundLayout = $templateBased->findLayout($layout['name']);
            $this->assertEquals($layout, $foundLayout);
        }
    }

    /**
     * @expectedException \Exception
     */
    public function testFindLayoutException()
    {
        $file = PHPPOWERPOINT_TESTS_BASE_DIR . '/resources/files/Sample_00_01.pptx';
        $templateBased = new TemplateBased($file);
        $name = 'Invalid';
        $templateBased->findLayout($name);
    }

    public function testFindLayoutIndex()
    {
        $file = PHPPOWERPOINT_TESTS_BASE_DIR . '/resources/files/Sample_00_01.pptx';
        $templateBased = new TemplateBased($file);
        $layouts = $templateBased->getLayouts();

        $expectedSlideLayoutId = 1;
        foreach ($layouts as $layoutId => $layout) {
            $layoutIndex = $templateBased->findLayoutIndex($layout['name']);
            $this->assertEquals($expectedSlideLayoutId, ($layoutIndex + 1));
            ++$expectedSlideLayoutId;
        }
    }

    /**
     * @expectedException \Exception
     */
    public function testFindLayoutIndexException()
    {
        $file = PHPPOWERPOINT_TESTS_BASE_DIR . '/resources/files/Sample_00_01.pptx';
        $templateBased = new TemplateBased($file);
        $name = 'Invalid';
        $templateBased->findLayoutIndex($name);
    }
}