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
 * contributors, visit https://github.com/PHPOffice/PHPPowerPoint/contributors.
 *
 * @copyright   2010-2014 PHPPowerPoint contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 * @link        https://github.com/PHPOffice/PHPPowerPoint
 */

date_default_timezone_set('UTC');

// defining base dir for tests
if (!defined('PHPPOWERPOINT_TESTS_BASE_DIR')) {
    define('PHPPOWERPOINT_TESTS_BASE_DIR', realpath(__DIR__));
}

$vendor = realpath(__DIR__ . '/../vendor');

if (file_exists($vendor . "/autoload.php")) {
    require $vendor . "/autoload.php";
} else {
    $vendor = realpath(__DIR__ . '/../../../');
    if (file_exists($vendor . "/autoload.php")) {
        require $vendor . "/autoload.php";
    } else {
        throw new Exception("Unable to load dependencies");
    }
}

spl_autoload_register(function ($class) {
    $class = ltrim($class, '\\');
    $prefix = 'PhpOffice\\PhpPowerpoint\\Tests';
    if (strpos($class, $prefix) === 0) {
        $class = str_replace('\\', DIRECTORY_SEPARATOR, $class);
        $class = join(DIRECTORY_SEPARATOR, array('PhpPowerpoint', 'Tests', '_includes')) .
        substr($class, strlen($prefix));
        $file = __DIR__ . DIRECTORY_SEPARATOR . $class . '.php';
        if (file_exists($file)) {
            require_once $file;
        }
    }
});

require_once __DIR__ . "/../src/PhpPowerpoint/Autoloader.php";
\PhpOffice\PhpPowerpoint\Autoloader::register();
