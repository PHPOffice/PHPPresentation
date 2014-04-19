<?php
/**
 * $Id: bootstrap.php 2892 2011-08-14 15:11:50Z markbaker@PHPPowerPoint.net $
 *
 * @copyright   Copyright (C) 2011-2012 PHPPowerPoint. All rights reserved.
 * @package     PHPPowerPoint
 * @subpackage  PHPPowerPoint Unit Tests
 * @author      Mark Baker
 */

date_default_timezone_set('UTC');

// Define path to application directory
defined('APPLICATION_PATH')
    || define('APPLICATION_PATH', realpath(dirname(__FILE__) . '/../Classes'));

// Define path to application tests directory
defined('APPLICATION_TESTS_PATH')
    || define('APPLICATION_TESTS_PATH', realpath(dirname(__FILE__) ));

// Define application environment
defined('APPLICATION_ENV') || define('APPLICATION_ENV', 'ci');

// Register autoloader
if (!defined('PHPPOWERPOINT_ROOT'))
{
    define('PHPPOWERPOINT_ROOT', APPLICATION_PATH . '/');
}
require_once PHPPOWERPOINT_ROOT . 'PHPPowerPoint/Autoloader.php';
PHPPowerPoint_Autoloader::Register();

/**
 * @todo Sort out xdebug in vagrant so that this works in all sandboxes
 * For now, it is safer to test for it rather then remove it.
 */
echo "PHPPowerPoint tests beginning\n";

if (extension_loaded('xdebug')) {
    echo "Xdebug extension loaded and running\n";
    xdebug_enable();
} else {
    echo 'Xdebug not found, you should run the following at the command line: echo "zend_extension=/usr/lib64/php/modules/xdebug.so" > /etc/php.d/xdebug.ini' . "\n";
}
