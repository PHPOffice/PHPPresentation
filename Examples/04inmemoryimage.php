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
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */

/** Error reporting */
error_reporting(E_ALL);

/** Include path **/
set_include_path(get_include_path() . PATH_SEPARATOR . '../Classes/');

/** PHPPowerPoint */
include 'PHPPowerPoint.php';


// Create new PHPPowerPoint object
echo date('H:i:s') . " Create new PHPPowerPoint object\n";
$objPHPPowerPoint = new PHPPowerPoint();

// Set properties
echo date('H:i:s') . " Set properties\n";
$objPHPPowerPoint->getProperties()->setCreator("Maarten Balliauw")
                                  ->setLastModifiedBy("Maarten Balliauw")
                                  ->setTitle("Office 2007 PPTX Test Document")
                                  ->setSubject("Office 2007 PPTX Test Document")
                                  ->setDescription("Test document for Office 2007 PPTX, generated using PHP classes.")
                                  ->setKeywords("office 2007 openxml php")
                                  ->setCategory("Test result file");

// Create slide
echo date('H:i:s') . " Create slide\n";
$currentSlide = $objPHPPowerPoint->getActiveSlide();

// Generate an image
echo date('H:i:s') . " Generate an image\n";
$gdImage = @imagecreatetruecolor(140, 20) or die('Cannot Initialize new GD image stream');
$textColor = imagecolorallocate($gdImage, 255, 255, 255);
imagestring($gdImage, 1, 5, 5,  'Created with PHPPowerPoint', $textColor);

// Add a drawing to the worksheet
echo date('H:i:s') . " Add a drawing to the worksheet\n";
$shape = new PHPPowerPoint_Shape_MemoryDrawing();
$shape->setName('Sample image')
      ->setDescription('Sample image')
      ->setImageResource($gdImage)
      ->setRenderingFunction(PHPPowerPoint_Shape_MemoryDrawing::RENDERING_JPEG)
      ->setMimeType(PHPPowerPoint_Shape_MemoryDrawing::MIMETYPE_DEFAULT)
      ->setHeight(36)
      ->setOffsetX(10)
      ->setOffsetY(10);
$currentSlide->addShape($shape);

// Save files
// @todo Error in ODT getPath
$basename = basename(__FILE__, '.php');
$formats = array('PowerPoint2007' => 'pptx');
foreach ($formats as $format => $extension) {
    echo date('H:i:s') . " Write to {$format} format\r\n";
    $objWriter = PHPPowerPoint_IOFactory::createWriter($objPHPPowerPoint, $format);
    $objWriter->save("results/{$basename}.{$extension}");
}

// Echo memory peak usage
echo date('H:i:s') . " Peak memory usage: " . (memory_get_peak_usage(true) / 1024 / 1024) . " MB\r\n";

// Echo done
echo date('H:i:s') . " Done writing file.\r\n";
