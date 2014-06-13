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

if (!defined('DATE_W3C')) {
    define('DATE_W3C', 'Y-m-d\TH:i:sP');
}

/**
 * PHPPowerPoint_Shared_XMLWriter
 */
class XMLWriter
{
    /** Temporary storage method */
    const STORAGE_MEMORY = 1;
    const STORAGE_DISK = 2;

    /**
     * Internal XMLWriter
     *
     * @var XMLWriter
     */
    private $xmlWriter;

    /**
     * Temporary filename
     *
     * @var string
     */
    private $tempFileName = '';

    /**
     * Create a new PHPPowerPoint_Shared_XMLWriter instance
     *
     * @param int    $pTemporaryStorage       Temporary storage location
     * @param string $pTemporaryStorageFolder Temporary storage folder
     */
    public function __construct($pTemporaryStorage = self::STORAGE_MEMORY, $pTemporaryStorageDir = './')
    {
        // Create internal XMLWriter
        $this->xmlWriter = new \XMLWriter();

        // Open temporary storage
        if ($pTemporaryStorage == self::STORAGE_MEMORY) {
            $this->xmlWriter->openMemory();
        } else {
            // Create temporary filename
            $this->tempFileName = @tempnam($pTemporaryStorageDir, 'xml');

            // Open storage
            if ($this->xmlWriter->openUri($this->tempFileName) === false) {
                // Fallback to memory...
                $this->xmlWriter->openMemory();
            }
        }

        // Set default values
        $this->xmlWriter->setIndent(true);
    }

    /**
     * Destructor
     */
    public function __destruct()
    {
        // Desctruct XMLWriter
        unset($this->xmlWriter);

        // Unlink temporary files
        if ($this->tempFileName != '') {
            @unlink($this->tempFileName);
        }
    }

    /**
     * Catch function calls (and pass them to internal XMLWriter)
     *
     * @param mixed $function
     * @param mixed $args
     */
    public function __call($function, $args)
    {
        try {
            @call_user_func_array(array(
                $this->xmlWriter,
                $function
            ), $args);
        } catch (\Exception $ex) {
            // Do nothing!
        }
    }

    /**
     * Get written data
     *
     * @return $data
     */
    public function getData()
    {
        if ($this->tempFileName == '') {
            return $this->xmlWriter->outputMemory(true);
        } else {
            $this->xmlWriter->flush();

            return file_get_contents($this->tempFileName);
        }
    }

    /**
     * Fallback method for writeRaw, introduced in PHP 5.2
     *
     * @param  string $text
     * @return string
     */
    public function writeRaw($text)
    {
        if (isset($this->xmlWriter) && is_object($this->xmlWriter) && (method_exists($this->xmlWriter, 'writeRaw'))) {
            return $this->xmlWriter->writeRaw($text);
        }

        return $this->text($text);
    }
}
