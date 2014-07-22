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
 * XMLWriter
 *
 * @method bool endElement()
 * @method mixed flush(bool $empty = null)
 * @method bool openMemory()
 * @method string outputMemory(bool $flush = null)
 * @method bool setIndent(bool $indent)
 * @method bool startDocument(string $version = 1.0, string $encoding = null, string $standalone = null)
 * @method bool startElement(string $name)
 * @method bool text(string $content)
 * @method bool writeAttribute(string $name, mixed $value)
 * @method bool writeCData(string $content)
 * @method bool writeComment(string $content)
 * @method bool writeElement(string $name, string $content = null)
 * @method bool writeRaw(string $content)
 */
class XMLWriter
{
    /** Temporary storage method */
    const STORAGE_MEMORY = 1;
    const STORAGE_DISK = 2;

    /**
     * Internal XMLWriter
     *
     * @var \XMLWriter
     */
    private $xmlWriter;

    /**
     * Temporary filename
     *
     * @var string
     */
    private $tempFileName = '';

    /**
     * Create a new \PhpOffice\PhpPowerpoint\Shared\XMLWriter instance
     *
     * @param int $pTemporaryStorage Temporary storage location
     * @param string $pTemporaryStorageDir Temporary storage folder
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
            $this->xmlWriter->openUri($this->tempFileName);
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
     * @return string
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
}
