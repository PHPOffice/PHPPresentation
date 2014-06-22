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

namespace PhpOffice\PhpPowerpoint\Shared;

/**
 * ZipStreamWrapper
 */
class ZipStreamWrapper
{
    /**
     * Internal ZipArchive
     *
     * @var \ZipArchive
     */
    private $archive;

    /**
     * Filename in ZipArchive
     *
     * @var string
     */
    private $fileNameInArchive = '';

    /**
     * Position in file
     *
     * @var int
     */
    private $position = 0;

    /**
     * Data
     *
     * @var mixed
     */
    private $data = '';

    /**
     * Register wrapper
     */
    public static function register()
    {
        @stream_wrapper_unregister("zip");
        @stream_wrapper_register("zip", __CLASS__);
    }

    /**
     * Open stream
     *
     * Original param array $options and string $opened_path aren't used
     *
     * @param string $path
     * @param string $mode
     * @throws \Exception
     * @return bool
     */
    public function streamOpen($path, $mode)
    {
        // Check for mode
        if ($mode{0} != 'r') {
            throw new \Exception('Mode ' . $mode . ' is not supported. Only read mode is supported.');
        }

        // Parse URL
        $url = @parse_url($path);

        // Fix URL
        if (!is_array($url)) {
            $url['host'] = substr($path, strlen('zip://'));
            $url['path'] = '';
        }
        if (strpos($url['host'], '#') !== false) {
            if (!isset($url['fragment'])) {
                $url['fragment'] = substr($url['host'], strpos($url['host'], '#') + 1) . $url['path'];
                $url['host']     = substr($url['host'], 0, strpos($url['host'], '#'));
                unset($url['path']);
            }
        } else {
            $url['host'] = $url['host'] . $url['path'];
            unset($url['path']);
        }

        // Open archive
        $this->archive = new \ZipArchive();
        $this->archive->open($url['host']);

        $this->fileNameInArchive = $url['fragment'];
        $this->position          = 0;
        $this->data              = $this->archive->getFromName($this->fileNameInArchive);

        return true;
    }

    /**
     * Stat stream
     */
    public function streamStat()
    {
        return $this->archive->statName($this->fileNameInArchive);
    }

    /**
     * Read stream
     *
     * @param int $count
     * @return string
     */
    public function streamRead($count)
    {
        $ret = substr($this->data, $this->position, $count);
        $this->position += strlen($ret);

        return $ret;
    }

    /**
     * Tell stream
     */
    public function streamTell()
    {
        return $this->position;
    }

    /**
     * EOF stream
     */
    public function streamEof()
    {
        return $this->position >= strlen($this->data);
    }

    /**
     * Seek stream
     *
     * @param int $offset
     * @param int $whence
     * @return bool
     */
    public function streamSeek($offset, $whence)
    {
        switch ($whence) {
            case SEEK_SET:
                if ($offset < strlen($this->data) && $offset >= 0) {
                    $this->position = $offset;
                    return true;
                } else {
                    return false;
                }
                // Break intentionally omitted
            case SEEK_CUR:
                if ($offset >= 0) {
                    $this->position += $offset;
                    return true;
                } else {
                    return false;
                }
                // Break intentionally omitted
            case SEEK_END:
                if (strlen($this->data) + $offset >= 0) {
                    $this->position = strlen($this->data) + $offset;
                    return true;
                } else {
                    return false;
                }
                // Break intentionally omitted
            default:
                return false;
        }
    }
}
