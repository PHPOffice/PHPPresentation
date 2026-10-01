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

namespace PhpOffice\PhpPresentation\Shape\RichText;

use PhpOffice\PhpPresentation\Style\Bullet;

/**
 * How the numbered paragraphs of one text number.
 *
 * A numbering goes on over the paragraphs of its level that share its scheme, and over any
 * deeper paragraph between them. A shallower paragraph, one with no marker, or one at its
 * level with a bullet or another scheme ends it -- what PowerPoint shows. A paragraph asked
 * to continue the numbering before it (Bullet::setBulletNumericContinue()) takes, when it
 * begins one, the number after the last one of its level and scheme.
 */
class Numbering
{
    /**
     * @var array<int, null|int>
     */
    private $continuedStartAts = [];

    /**
     * @var array<int, null|int>
     */
    private $resumedStartAts = [];

    /**
     * @param array<int, Paragraph> $paragraphs the paragraphs of one text, in order
     */
    public function __construct(array $paragraphs)
    {
        // the scheme, the start and the last number of the numbering open at each level
        $open = [];
        // the last number of the last numbering ended at each level, by scheme
        $ended = [];
        foreach ($paragraphs as $key => $paragraph) {
            $bullet = $paragraph->getBulletStyle();
            $type = null === $bullet ? Bullet::TYPE_NONE : $bullet->getBulletType();
            $level = $paragraph->getAlignment()->getLevel();
            $scheme = null === $bullet ? '' : $bullet->getBulletNumericStyle();
            foreach ($open as $openLevel => [$openScheme, , $openNumber]) {
                if (Bullet::TYPE_NONE == $type || $openLevel > $level
                    || ($openLevel === $level && (Bullet::TYPE_NUMERIC != $type || $openScheme !== $scheme))) {
                    $ended[$openLevel][$openScheme] = $openNumber;
                    unset($open[$openLevel]);
                }
            }
            if (null === $bullet || Bullet::TYPE_NUMERIC != $type) {
                continue;
            }
            if (isset($open[$level])) {
                $this->continuedStartAts[$key] = $open[$level][1];
            } else {
                $this->resumedStartAts[$key] = isset($ended[$level][$scheme]) ? $ended[$level][$scheme] + 1 : null;
                $this->continuedStartAts[$key] = $bullet->isBulletNumericContinue() ? $this->resumedStartAts[$key] : null;
            }
            $start = $bullet->getBulletNumericStartAt() ?? $this->continuedStartAts[$key] ?? 1;
            if (isset($open[$level]) && $open[$level][1] === $start) {
                ++$open[$level][2];
            } else {
                if (isset($open[$level])) {
                    $ended[$level][$scheme] = $open[$level][2];
                }
                $open[$level] = [$scheme, $start, $start];
            }
        }
    }

    /**
     * The start each numbered paragraph continues the numbering of, or null for one that begins
     * a numbering: the start a paragraph given none takes.
     *
     * @return array<int, null|int> keyed as the numbered paragraphs are
     */
    public function getContinuedStartAts(): array
    {
        return $this->continuedStartAts;
    }

    /**
     * The number each numbered paragraph that begins a numbering would take if it continued the
     * last numbering of its level and scheme, or null when there is none.
     *
     * @return array<int, null|int> keyed as the numbered paragraphs that begin a numbering are
     */
    public function getResumedStartAts(): array
    {
        return $this->resumedStartAts;
    }
}
