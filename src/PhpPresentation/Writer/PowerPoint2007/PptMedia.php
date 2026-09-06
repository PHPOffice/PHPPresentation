<?php

/**
 * This file is part of PHPPresentation - A pure PHP library for reading and writing
 * presentations documents.
 *
 * PHPPresentation is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPPresentation/contributors.
 *
 * @see        https://github.com/PHPOffice/PHPPresentation
 *
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

declare(strict_types=1);

namespace PhpOffice\PhpPresentation\Writer\PowerPoint2007;

use DK\OpenXml\OpenXmlPackage;
use PhpOffice\PhpPresentation\Shape\Drawing\AbstractDrawingAdapter;
use PhpOffice\PhpPresentation\Shape\Drawing\File;

class PptMedia extends AbstractDecoratorWriter
{
    public function render(): OpenXmlPackage
    {
        for ($i = 0; $i < $this->getDrawingHashTable()->count(); ++$i) {
            $shape = $this->getDrawingHashTable()->getByIndex($i);
            if (!$shape instanceof AbstractDrawingAdapter) {
                continue;
            }
            $name = '/ppt/media/' . $shape->getIndexedFilename();
            $type = $this->mediaContentType($shape->getExtension(), $shape->getMimeType());
            // A shape that already is a file on disk is handed over as one: the package reads it
            // when it saves, so the bytes never pass through a PHP string. Everything else -- a
            // base64 payload, a GD resource, an image inside another archive -- has no file to
            // point at and is handed over as contents.
            if ($shape instanceof File && is_file($shape->getPath())) {
                $this->oPackage->addPartFromPath($name, $type, $shape->getPath());
            } else {
                $this->oPackage->addPart($name, $type, $shape->getContents());
            }
        }

        return $this->oPackage;
    }
}
