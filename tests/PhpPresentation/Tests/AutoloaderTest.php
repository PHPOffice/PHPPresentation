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
 * @copyright   2009-2015 PHPPresentation contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

declare(strict_types=1);

namespace PhpOffice\PhpPresentation\Tests;

use PhpOffice\PhpPresentation\Autoloader;
use PHPUnit\Framework\TestCase;

/**
 * Test class for Autoloader.
 */
class AutoloaderTest extends TestCase
{
    /**
     * Register.
     */
    public function testRegister(): void
    {
        Autoloader::register();
        $this->assertContains(
            ['PhpOffice\\PhpPresentation\\Autoloader', 'autoload'],
            spl_autoload_functions()
        );
    }

    /**
     * Autoload.
     */
    public function testAutoload(): void
    {
        $declared = get_declared_classes();
        $declaredCount = count($declared);
        Autoloader::autoload('Foo');
        $this->assertEquals(
            $declaredCount,
            count(get_declared_classes()),
            'PhpOffice\\PhpPresentation\\Autoloader::autoload() is trying to load ' .
            'classes outside of the PhpOffice\\PhpPresentation namespace'
        );
    }
}
