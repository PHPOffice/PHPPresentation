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

namespace PhpOffice\PhpPresentation\Tests\Reader\PowerPoint2007;

use PhpOffice\PhpPresentation\Reader\PowerPoint2007;
use PHPUnit\Framework\Attributes\DataProvider;
use PHPUnit\Framework\TestCase;
use ZipArchive;

/**
 * A presentation is read through the part its package points at, whatever that part is named.
 *
 * `/ppt/presentation.xml`, and the folders beside it, are the names Office happens to write.
 * The specification asks only that the package relate to a part of the presentation content
 * type, and that each relationship name its target relative to the part that declares it.
 */
class MainDocumentPartTest extends TestCase
{
    /**
     * @var list<string>
     */
    private $files = [];

    protected function tearDown(): void
    {
        foreach ($this->files as $file) {
            @unlink($file);
        }
        $this->files = [];
    }

    /**
     * @return array<string, array{string}>
     */
    public static function dataProviderSamples(): array
    {
        return [
            'chart' => ['PPTX_ChartBar.pptx'],
            'group' => ['PPTX_Group.pptx'],
            'image' => ['PPTX_InvalidImage.pptx'],
            'note' => ['PPTX_SlideNoteWithRichText.pptx'],
            'several masters' => ['Issue_00150.pptx'],
            'several slides' => ['Sample_12.pptx'],
        ];
    }

    /**
     * @dataProvider dataProviderSamples
     */
    #[DataProvider('dataProviderSamples')]
    public function testReadsAPresentationWhoseMainPartIsElsewhere(string $sample): void
    {
        $original = PHPPRESENTATION_TESTS_BASE_DIR . '/resources/files/' . $sample;
        $moved = $this->moveMainDocumentPart($original);

        self::assertSame($this->describe($original), $this->describe($moved));
    }

    /**
     * Move the main document part to `/deck/main.xml`, and say so everywhere it is named.
     *
     * Its relationships are declared one folder further from the parts they point at, so each
     * relative target grows the step back out of `/deck` that it now needs.
     */
    private function moveMainDocumentPart(string $original): string
    {
        $moved = (string) tempnam(sys_get_temp_dir(), 'PhpPresentationMainPart');
        $this->files[] = $moved;
        copy($original, $moved);

        $zip = new ZipArchive();
        $zip->open($moved);
        $relationships = (string) preg_replace_callback(
            '/Target="([^"]+)"/',
            function (array $matches): string {
                return 0 === strncmp($matches[1], '/', 1) || false !== strpos($matches[1], '://')
                    ? $matches[0]
                    : 'Target="../ppt/' . $matches[1] . '"';
            },
            (string) $zip->getFromName('ppt/_rels/presentation.xml.rels')
        );
        $zip->addFromString('deck/main.xml', (string) $zip->getFromName('ppt/presentation.xml'));
        $zip->addFromString('deck/_rels/main.xml.rels', $relationships);
        $zip->addFromString('[Content_Types].xml', str_replace(
            '"/ppt/presentation.xml"',
            '"/deck/main.xml"',
            (string) $zip->getFromName('[Content_Types].xml')
        ));
        $zip->addFromString('_rels/.rels', str_replace(
            '"ppt/presentation.xml"',
            '"deck/main.xml"',
            (string) $zip->getFromName('_rels/.rels')
        ));
        $zip->deleteName('ppt/presentation.xml');
        $zip->deleteName('ppt/_rels/presentation.xml.rels');
        $zip->close();

        return $moved;
    }

    /**
     * Everything the reader found, as a string two presentations can be compared by.
     */
    private function describe(string $file): string
    {
        $presentation = (new PowerPoint2007())->load($file);
        $read = [
            'slides' => $presentation->getSlideCount(),
            'masters' => count($presentation->getAllMasterSlides()),
            'title' => $presentation->getDocumentProperties()->getTitle(),
            'zoom' => $presentation->getPresentationProperties()->getZoom(),
        ];
        foreach ($presentation->getAllSlides() as $index => $slide) {
            $shapes = [];
            foreach ($slide->getShapeCollection() as $shape) {
                $shapes[] = get_class($shape) . ':' . $shape->getName();
            }
            $read['slide' . $index] = [
                'name' => $slide->getName(),
                'layout' => null === $slide->getSlideLayout() ? null : $slide->getSlideLayout()->getLayoutName(),
                'background' => null === $slide->getBackground() ? null : get_class($slide->getBackground()),
                'notes' => count($slide->getNote()->getShapeCollection()),
                'shapes' => $shapes,
            ];
        }
        foreach ($presentation->getAllMasterSlides() as $index => $master) {
            $read['master' . $index] = [
                'layouts' => count($master->getAllSlideLayouts()),
                'shapes' => count($master->getShapeCollection()),
            ];
        }

        return (string) json_encode($read);
    }
}
