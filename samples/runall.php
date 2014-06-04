<?php
/**
 * PHPPowerPoint
 *
 * Copyright (c) 2009 - 2010 PHPPowerPoint
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 * 
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 * 
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    ##VERSION##, ##DATE##
 */

/** Error reporting */
error_reporting(E_ALL);

// List of tests
$aTests = array(
	  '01simple.php'
	, '02presentation.php'
	, '03serialized.php'
	, '04inmemoryimage.php'
	, '05templated.php'
	, '06table.php'
	, '07chart.php'
	, '08chart.php'
);

// First, clear all results
foreach ($aTests as $sTest) {
	@unlink( str_replace('.php', '.ppt', 	$sTest) );
	@unlink( str_replace('.php', '.pptx', 	$sTest) );
	@unlink( str_replace('.php', '.phppt',	$sTest) );
	@unlink( str_replace('.php', '.htm',	$sTest) );
	@unlink( str_replace('.php', '.pdf',	$sTest) );
}

// Run all tests
foreach ($aTests as $sTest) {
	echo '============== TEST ==============' . "\r\n";
	echo 'Test name: ' . $sTest . "\r\n";
	echo "\r\n";
	echo shell_exec('php ' . $sTest);
	echo "\r\n";
	echo "\r\n";
}
