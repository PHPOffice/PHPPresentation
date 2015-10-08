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

namespace PhpOffice\PhpPresentation\Slide;

use PhpOffice\PhpPresentation\Slide;
use PhpOffice\PhpPresentation\Shape\RichText;

/**
 * Transition class
 */
class Transition
{
    const SPEED_FAST = 'fast';
    const SPEED_MEDIUM = 'med';
    const SPEED_SLOW = 'slow';

    const TRANSITION_BLINDS_HORIZONTAL = 'blinds_horz';
    const TRANSITION_BLINDS_VERTICAL = 'blinds_vert';
    const TRANSITION_CHECKER_HORIZONTAL = 'checker_horz';
    const TRANSITION_CHECKER_VERTICAL = 'checker_vert';
    const TRANSITION_CIRCLE_HORIZONTAL = 'circle_horz';
    const TRANSITION_CIRCLE_VERTICAL = 'circle_vert';
    const TRANSITION_COMB_HORIZONTAL = 'comb_horz';
    const TRANSITION_COMB_VERTICAL = 'comb_vert';
    const TRANSITION_COVER_DOWN = 'cover_d';
    const TRANSITION_COVER_LEFT = 'cover_l';
    const TRANSITION_COVER_LEFT_DOWN = 'cover_ld';
    const TRANSITION_COVER_LEFT_UP = 'cover_lu';
    const TRANSITION_COVER_RIGHT = 'cover_r';
    const TRANSITION_COVER_RIGHT_DOWN = 'cover_rd';
    const TRANSITION_COVER_RIGHT_UP = 'cover_ru';
    const TRANSITION_COVER_UP = 'cover_u';
    const TRANSITION_CUT = 'cut';
    const TRANSITION_DIAMOND = 'diamond';
    const TRANSITION_DISSOLVE = 'dissolve';
    const TRANSITION_FADE = 'fade';
    const TRANSITION_NEWSFLASH = 'newsflash';
    const TRANSITION_PLUS = 'plus';
    const TRANSITION_PULL_DOWN = 'pull_d';
    const TRANSITION_PULL_LEFT = 'pull_l';
    const TRANSITION_PULL_RIGHT = 'pull_r';
    const TRANSITION_PULL_UP = 'pull_u';
    const TRANSITION_PUSH_DOWN = 'push_d';
    const TRANSITION_PUSH_LEFT = 'push_l';
    const TRANSITION_PUSH_RIGHT = 'push_r';
    const TRANSITION_PUSH_UP = 'push_u';
    const TRANSITION_RANDOM = 'random';
    const TRANSITION_RANDOMBAR_HORIZONTAL = 'randomBar_horz';
    const TRANSITION_RANDOMBAR_VERTICAL = 'randomBar_vert';
    const TRANSITION_SPLIT_IN_HORIZONTAL = 'split_in_horz';
    const TRANSITION_SPLIT_OUT_HORIZONTAL = 'split_out_horz';
    const TRANSITION_SPLIT_IN_VERTICAL = 'split_in_vert';
    const TRANSITION_SPLIT_OUT_VERTICAL = 'split_out_vert';
    const TRANSITION_STRIPS_LEFT_DOWN = 'strips_ld';
    const TRANSITION_STRIPS_LEFT_UP = 'strips_lu';
    const TRANSITION_STRIPS_RIGHT_DOWN = 'strips_rd';
    const TRANSITION_STRIPS_RIGHT_UP = 'strips_ru';
    const TRANSITION_WEDGE = 'wedge';
    const TRANSITION_WIPE_DOWN = 'wipe_d';
    const TRANSITION_WIPE_LEFT = 'wipe_l';
    const TRANSITION_WIPE_RIGHT = 'wipe_r';
    const TRANSITION_WIPE_UP = 'wipe_u';
    const TRANSITION_ZOOM_IN = 'zoom_in';
    const TRANSITION_ZOOM_OUT = 'zoom_out';

    /**
     * @var bool
     */
    protected $hasManualTrigger = false;
    /**
     * @var bool
     */
    protected $hasTimeTrigger = false;
    /**
     * @var int
     */
    protected $advanceTimeTrigger = null;
    /**
     * @var null|self::SPEED_SLOW|self::SPEED_MEDIUM|self::SPEED_FAST
     */
    protected $speed = null;
    /**
     * @var null|self::TRANSITION_*
     */
    protected $transitionType = null;
    /**
     * @var array
     */
    protected $transitionOptions = array();

    public function setSpeed($speed = self::SPEED_MEDIUM)
    {
        if (in_array($speed, array(self::SPEED_FAST, self::SPEED_MEDIUM, self::SPEED_SLOW))) {
            $this->speed = $speed;
        } else {
            $this->speed = null;
        }

        return $this;
    }

    public function getSpeed()
    {
        return $this->speed;
    }

    public function setManualTrigger($value = false)
    {
        if (is_bool($value)) {
            $this->hasManualTrigger = $value;
        }
        return $this;
    }

    public function hasManualTrigger()
    {
        return $this->hasManualTrigger;
    }

    public function setTimeTrigger($value = false, $advanceTime = 1000)
    {
        if (is_bool($value)) {
            $this->hasTimeTrigger = $value;
        }
        $this->advanceTimeTrigger = null;
        if ($this->hasTimeTrigger === true) {
            $this->advanceTimeTrigger = (int) $advanceTime;
        }
        return $this;
    }

    public function hasTimeTrigger()
    {
        return $this->hasTimeTrigger;
    }

    public function getAdvanceTimeTrigger()
    {
        return $this->advanceTimeTrigger;
    }

    public function setTransitionType($type = null)
    {
        $this->transitionType = $type;
        return $this;
    }

    public function getTransitionType()
    {
        return $this->transitionType;
    }
}
