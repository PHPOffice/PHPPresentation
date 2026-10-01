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

namespace PhpOffice\PhpPresentation\Tests\Shape\RichText;

use PhpOffice\PhpPresentation\Shape\RichText\Numbering;
use PhpOffice\PhpPresentation\Shape\RichText\Paragraph;
use PhpOffice\PhpPresentation\Style\Bullet;
use PHPUnit\Framework\TestCase;

/**
 * @coversDefaultClass \PhpOffice\PhpPresentation\Shape\RichText\Numbering
 */
class NumberingTest extends TestCase
{
    public function testNumberingIsContinuedOverDeeperParagraphsAndEndedByShallowerOnesAndGaps(): void
    {
        // level, marker, start: a numbering from 2 with a deeper one inside it, a new one
        // from 3, a paragraph with no marker, and one more after it
        $paragraphs = [];
        foreach ([[0, 2], [0, null], [1, null], [1, null], [0, 3], [0, null], [0, false], [0, null]] as [$level, $startAt]) {
            $paragraph = new Paragraph();
            $paragraph->getAlignment()->setLevel($level);
            $paragraph->getBulletStyle()
                ->setBulletType(false === $startAt ? Bullet::TYPE_NONE : Bullet::TYPE_NUMERIC)
                ->setBulletNumericStartAt(false === $startAt ? null : $startAt);
            $paragraphs[] = $paragraph;
        }

        self::assertSame(
            [0 => null, 1 => 2, 2 => null, 3 => 1, 4 => 2, 5 => 3, 7 => null],
            (new Numbering($paragraphs))->getContinuedStartAts()
        );
    }

    public function testNumberingIsEndedByABulletOrAnotherSchemeAtItsLevel(): void
    {
        $paragraphs = [];
        foreach ([[Bullet::TYPE_NUMERIC, Bullet::NUMERIC_ARABICPERIOD], [Bullet::TYPE_NUMERIC, Bullet::NUMERIC_ALPHALCPERIOD], [Bullet::TYPE_BULLET, Bullet::NUMERIC_ALPHALCPERIOD], [Bullet::TYPE_NUMERIC, Bullet::NUMERIC_ALPHALCPERIOD]] as [$type, $scheme]) {
            $paragraph = new Paragraph();
            $paragraph->getBulletStyle()->setBulletType($type)->setBulletNumericStyle($scheme);
            $paragraphs[] = $paragraph;
        }
        $paragraphs[0]->getBulletStyle()->setBulletNumericStartAt(2);

        self::assertSame([0 => null, 1 => null, 3 => null], (new Numbering($paragraphs))->getContinuedStartAts());
    }

    public function testNumberingAskedToContinueTakesTheNumberAfterTheLastOne(): void
    {
        // 1, 2, a paragraph with no marker, 3, 4 asked to continue, another, then 1
        $paragraphs = [];
        foreach ([[true, false], [true, false], [false, false], [true, true], [true, false], [false, false], [true, false]] as [$numbered, $continue]) {
            $paragraph = new Paragraph();
            $paragraph->getBulletStyle()
                ->setBulletType($numbered ? Bullet::TYPE_NUMERIC : Bullet::TYPE_NONE)
                ->setBulletNumericContinue($continue);
            $paragraphs[] = $paragraph;
        }

        self::assertSame([0 => null, 1 => 1, 3 => 3, 4 => 3, 6 => null], (new Numbering($paragraphs))->getContinuedStartAts());
        self::assertSame([0 => null, 3 => 3, 6 => 5], (new Numbering($paragraphs))->getResumedStartAts());
    }
}
