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
 * PHPPowerPoint_Shared_ZipStreamWrapper
 */
class ZipStreamWrapper
{
    /**
     * Internal ZipAcrhive
     *
     * @var ZipAcrhive
     */
    private $_archive;

    /**
     * Filename in ZipAcrhive
     *
     * @var string
     */
    private $_fileNameInArchive = '';

    /**
     * Position in file
     *
     * @var int
     */
    private $_position = 0;

    /**
     * Data
     *
     * @var mixed
     */
    private $_data = '';

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
     */
    public function stream_open($path, $mode)
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
        $this->_archive = new ZipArchive();
        $this->_archive->open($url['host']);

        $this->_fileNameInArchive = $url['fragment'];
        $this->_position          = 0;
        $this->_data              = $this->_archive->getFromName($this->_fileNameInArchive);

        return true;
    }

    /**
     * Stat stream
     */
    public function stream_stat()
    {
        return $this->_archive->statName($this->_fileNameInArchive);
    }

    /**
     * Read stream
     *
     * @param int $count
     */
    public function stream_read($count)
    {
        $ret = substr($this->_data, $this->_position, $count);
        $this->_position += strlen($ret);

        return $ret;
    }

    /**
     * Tell stream
     */
    public function stream_tell()
    {
        return $this->_position;
    }

    /**
     * EOF stream
     */
    public function stream_eof()
    {
        return $this->_position >= strlen($this->_data);
    }

    /**
     * Seek stream
     *
     * @param int $offset
     * @param int $whence
     */
    public function stream_seek($offset, $whence)
    {
        switch ($whence) {
            case SEEK_SET:
                if ($offset < strlen($this->_data) && $offset >= 0) {
                    $this->_position = $offset;

                    return true;
                } else {
                    return false;
                }
                break;

            case SEEK_CUR:
                if ($offset >= 0) {
                    $this->_position += $offset;

                    return true;
                } else {
                    return false;
                }
                break;

            case SEEK_END:
                if (strlen($this->_data) + $offset >= 0) {
                    $this->_position = strlen($this->_data) + $offset;

                    return true;
                } else {
                    return false;
                }
                break;

            default:
                return false;
        }
    }
}
