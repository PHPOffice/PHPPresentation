<?php

/**
 * This file is part of PHPPresentation - A pure PHP library for reading and writing
 * presentations documents.
 *
 * PHPPresentation is free software distributed under the terms of the GNU Lesser
 * General Public License as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PhpPresentation/contributors.
 *
 * @see        https://github.com/PHPOffice/PhpPresentation
 *
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL v3
 */

declare(strict_types=1);

namespace PhpOffice\PhpPresentation\Tests\Reader\PowerPoint2007;

use PhpOffice\PhpPresentation\Reader\PowerPoint2007;
use PHPUnit\Framework\TestCase;

require 'PackagePathReader.php';

/**
 * The names of the parts of a package are separated by forward slashes, and they are not paths of
 * a file system.
 *
 * `dirname()` is a function of the file system, and on Windows it answers with backslashes: the
 * root of a package is named `\` there. A reader that asks the file system for the folder of a
 * part therefore looks for `_rels/.rels` under a name the package does not have, never finds the
 * part the presentation is in, and says that the file it was given is not a presentation.
 */
class PackagePathTest extends TestCase
{
    /**
     * Test the folder of a part is found with the separator of the package.
     */
    public function testTheFolderOfAPartIsFoundWithTheSeparatorOfThePackage(): void
    {
        $reader = new PackagePathReader();

        self::assertSame('', $reader->folderOf(''));
        self::assertSame('', $reader->folderOf('presentation.xml'));
        self::assertSame('ppt', $reader->folderOf('ppt/presentation.xml'));
        self::assertSame('ppt/slides', $reader->folderOf('ppt/slides/slide1.xml'));
    }

    /**
     * Test the relationships of the root of the package are named as the package names them.
     */
    public function testTheRelationshipsOfTheRootAreNamedAsThePackageNamesThem(): void
    {
        self::assertSame('_rels/.rels', (new PackagePathReader())->relsOf(''));
    }

    /**
     * Test the relationships of a part are named beside it.
     */
    public function testTheRelationshipsOfAPartAreNamedBesideIt(): void
    {
        $reader = new PackagePathReader();

        self::assertSame('_rels/presentation.xml.rels', $reader->relsOf('presentation.xml'));
        self::assertSame('ppt/_rels/presentation.xml.rels', $reader->relsOf('ppt/presentation.xml'));
        self::assertSame('ppt/slides/_rels/slide1.xml.rels', $reader->relsOf('ppt/slides/slide1.xml'));
    }

    /**
     * Test a target is resolved against the part that declared it.
     */
    public function testATargetIsResolvedAgainstThePartThatDeclaredIt(): void
    {
        $reader = new PackagePathReader();

        self::assertSame('ppt/presentation.xml', $reader->resolve('_rels/.rels', 'ppt/presentation.xml'));
        self::assertSame('ppt/media/image1.png', $reader->resolve('ppt/slides/_rels/slide1.xml.rels', '../media/image1.png'));
        self::assertSame('ppt/slides/media/image1.png', $reader->resolve('ppt/slides/_rels/slide1.xml.rels', 'media/image1.png'));
    }

    /**
     * Test a presentation is read through the relationships of the root of its package.
     */
    public function testAPresentationIsReadThroughTheRelationshipsOfTheRoot(): void
    {
        $presentation = (new PowerPoint2007())->load(
            PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/Sample_12.pptx'
        );

        self::assertGreaterThan(0, $presentation->getSlideCount());
    }
}
