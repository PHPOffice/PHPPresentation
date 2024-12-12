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

namespace PhpOffice\PhpPresentation\Slide;

/**
 * Transition class.
 */
class Transition
{
    public const SPEED_FAST = 'fast';
    public const SPEED_MEDIUM = 'med';
    public const SPEED_SLOW = 'slow';

    public const TRANSITION_BLINDS_HORIZONTAL = 'blinds_horz';
    public const TRANSITION_BLINDS_VERTICAL = 'blinds_vert';
    public const TRANSITION_CHECKER_HORIZONTAL = 'checker_horz';
    public const TRANSITION_CHECKER_VERTICAL = 'checker_vert';
    public const TRANSITION_CIRCLE = 'circle';
    public const TRANSITION_COMB_HORIZONTAL = 'comb_horz';
    public const TRANSITION_COMB_VERTICAL = 'comb_vert';
    public const TRANSITION_COVER_DOWN = 'cover_d';
    public const TRANSITION_COVER_LEFT = 'cover_l';
    public const TRANSITION_COVER_LEFT_DOWN = 'cover_ld';
    public const TRANSITION_COVER_LEFT_UP = 'cover_lu';
    public const TRANSITION_COVER_RIGHT = 'cover_r';
    public const TRANSITION_COVER_RIGHT_DOWN = 'cover_rd';
    public const TRANSITION_COVER_RIGHT_UP = 'cover_ru';
    public const TRANSITION_COVER_UP = 'cover_u';
    public const TRANSITION_CUT = 'cut';
    public const TRANSITION_DIAMOND = 'diamond';
    public const TRANSITION_DISSOLVE = 'dissolve';
    public const TRANSITION_FADE = 'fade';
    public const TRANSITION_NEWSFLASH = 'newsflash';
    public const TRANSITION_PLUS = 'plus';
    public const TRANSITION_PULL_DOWN = 'pull_d';
    public const TRANSITION_PULL_LEFT = 'pull_l';
    public const TRANSITION_PULL_RIGHT = 'pull_r';
    public const TRANSITION_PULL_UP = 'pull_u';
    public const TRANSITION_PUSH_DOWN = 'push_d';
    public const TRANSITION_PUSH_LEFT = 'push_l';
    public const TRANSITION_PUSH_RIGHT = 'push_r';
    public const TRANSITION_PUSH_UP = 'push_u';
    public const TRANSITION_RANDOM = 'random';
    public const TRANSITION_RANDOMBAR_HORIZONTAL = 'randomBar_horz';
    public const TRANSITION_RANDOMBAR_VERTICAL = 'randomBar_vert';
    public const TRANSITION_SPLIT_IN_HORIZONTAL = 'split_in_horz';
    public const TRANSITION_SPLIT_OUT_HORIZONTAL = 'split_out_horz';
    public const TRANSITION_SPLIT_IN_VERTICAL = 'split_in_vert';
    public const TRANSITION_SPLIT_OUT_VERTICAL = 'split_out_vert';
    public const TRANSITION_STRIPS_LEFT_DOWN = 'strips_ld';
    public const TRANSITION_STRIPS_LEFT_UP = 'strips_lu';
    public const TRANSITION_STRIPS_RIGHT_DOWN = 'strips_rd';
    public const TRANSITION_STRIPS_RIGHT_UP = 'strips_ru';
    public const TRANSITION_WEDGE = 'wedge';
    public const TRANSITION_WIPE_DOWN = 'wipe_d';
    public const TRANSITION_WIPE_LEFT = 'wipe_l';
    public const TRANSITION_WIPE_RIGHT = 'wipe_r';
    public const TRANSITION_WIPE_UP = 'wipe_u';
    public const TRANSITION_ZOOM_IN = 'zoom_in';
    public const TRANSITION_ZOOM_OUT = 'zoom_out';

    /**
     * @var bool
     */
    protected $hasManualTrigger = false;

    /**
     * @var bool
     */
    protected $hasTimeTrigger = false;

    /**
     * @var null|int
     */
    protected $advanceTimeTrigger;

    /**
     * @var null|string
     */
    protected $speed;

    /**
     * @var null|string
     */
    protected $transitionType;

    public function setSpeed(?string $speed = self::SPEED_MEDIUM): self
    {
        if (in_array($speed, [self::SPEED_FAST, self::SPEED_MEDIUM, self::SPEED_SLOW])) {
            $this->speed = $speed;
        } else {
            $this->speed = null;
        }

        return $this;
    }

    public function getSpeed(): ?string
    {
        return $this->speed;
    }

    public function setManualTrigger(bool $value = false): self
    {
        $this->hasManualTrigger = $value;

        return $this;
    }

    public function hasManualTrigger(): bool
    {
        return $this->hasManualTrigger;
    }

    public function setTimeTrigger(bool $value = false, int $advanceTime = 1000): self
    {
        $this->hasTimeTrigger = $value;
        $this->advanceTimeTrigger = true === $value ? $advanceTime : null;

        return $this;
    }

    public function hasTimeTrigger(): bool
    {
        return $this->hasTimeTrigger;
    }

    public function getAdvanceTimeTrigger(): ?int
    {
        return $this->advanceTimeTrigger;
    }

    public function setTransitionType(?string $type = null): self
    {
        $this->transitionType = $type;

        return $this;
    }

    public function getTransitionType(): ?string
    {
        return $this->transitionType;
    }
}
