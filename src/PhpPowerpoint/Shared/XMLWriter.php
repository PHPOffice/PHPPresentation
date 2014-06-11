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
    private $_xmlWriter;

    /**
     * Temporary filename
     *
     * @var string
     */
    private $_tempFileName = '';

    /**
     * Create a new PHPPowerPoint_Shared_XMLWriter instance
     *
     * @param int    $pTemporaryStorage       Temporary storage location
     * @param string $pTemporaryStorageFolder Temporary storage folder
     */
    public function __construct($pTemporaryStorage = self::STORAGE_MEMORY, $pTemporaryStorageFolder = './')
    {
        // Create internal XMLWriter
        $this->_xmlWriter = new \XMLWriter();

        // Open temporary storage
        if ($pTemporaryStorage == self::STORAGE_MEMORY) {
            $this->_xmlWriter->openMemory();
        } else {
            // Create temporary filename
            $this->_tempFileName = @tempnam($pTemporaryStorageFolder, 'xml');

            // Open storage
            if ($this->_xmlWriter->openUri($this->_tempFileName) === false) {
                // Fallback to memory...
                $this->_xmlWriter->openMemory();
            }
        }

        // Set default values
        $this->_xmlWriter->setIndent(true);
    }

    /**
     * Destructor
     */
    public function __destruct()
    {
        // Desctruct XMLWriter
        unset($this->_xmlWriter);

        // Unlink temporary files
        if ($this->_tempFileName != '') {
            @unlink($this->_tempFileName);
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
                $this->_xmlWriter,
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
        if ($this->_tempFileName == '') {
            return $this->_xmlWriter->outputMemory(true);
        } else {
            $this->_xmlWriter->flush();

            return file_get_contents($this->_tempFileName);
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
        if (isset($this->_xmlWriter) && is_object($this->_xmlWriter) && (method_exists($this->_xmlWriter, 'writeRaw'))) {
            return $this->_xmlWriter->writeRaw($text);
        }

        return $this->text($text);
    }
}
