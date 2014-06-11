<?php
/**
 * This file is part of PHPPowerPoint - A pure PHP library for reading and writing
* presentation documents.
*
* PHPPowerPoint is free software distributed under the terms of the GNU Lesser
* General Public License version 3 as published by the Free Software Foundation.
*
* For the full copyright and license information, please read the LICENSE
* file that was distributed with this source code. For the full list of
* contributors, visit https://github.com/PHPOffice/PHPPowerPoint/contributors. test bootstrap
*
* @link        https://github.com/PHPOffice/PHPPowerPoint
* @copyright   2010-2014 PHPPowerPoint contributors
* @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
*/

date_default_timezone_set('UTC');

// Define path to application directory
defined('APPLICATION_PATH')
    || define('APPLICATION_PATH', realpath(dirname(__FILE__) . '/../src'));

// Define path to application tests directory
defined('APPLICATION_TESTS_PATH')
    || define('APPLICATION_TESTS_PATH', realpath(dirname(__FILE__)));

// Define application environment
defined('APPLICATION_ENV') || define('APPLICATION_ENV', 'ci');

// Register autoloader
if (!defined('PHPPOWERPOINT_ROOT')) {
    define('PHPPOWERPOINT_ROOT', APPLICATION_PATH . '/');
}

spl_autoload_register(function ($class) {
    $class = ltrim($class, '\\');
    $prefix = 'PhpOffice\\PhpPowerpoint\\Tests';
    if (strpos($class, $prefix) === 0) {
        $class = str_replace('\\', DIRECTORY_SEPARATOR, $class);
        $class = join(DIRECTORY_SEPARATOR, array('PHPPowerPoint', 'Tests', '_includes')) .
        substr($class, strlen($prefix));
        $file = __DIR__ . DIRECTORY_SEPARATOR . $class . '.php';
        if (file_exists($file)) {
            require_once $file;
        }
    }
});

    require_once __DIR__ . "/../src/PHPPowerPoint/Autoloader.php";
    \PhpOffice\PhpPowerpoint\Autoloader::register();