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

use PhpOffice\Common\Adapter\Zip\ZipInterface;
use PhpOffice\Common\XMLWriter;
use PhpOffice\PhpPresentation\Shape\Comment;
use PhpOffice\PhpPresentation\Shape\Comment\Author;

class CommentAuthors extends AbstractDecoratorWriter
{
    public function render(): ZipInterface
    {
        /**
         * @var Author[]
         */
        $arrayAuthors = [];
        foreach ($this->getPresentation()->getAllSlides() as $oSlide) {
            foreach ($oSlide->getShapeCollection() as $oShape) {
                if (!($oShape instanceof Comment)) {
                    continue;
                }
                $oAuthor = $oShape->getAuthor();
                if (!($oAuthor instanceof Author)) {
                    continue;
                }
                if (array_key_exists($oAuthor->getHashCode(), $arrayAuthors)) {
                    continue;
                }
                $arrayAuthors[$oAuthor->getHashCode()] = $oAuthor;
            }
        }
        if (!empty($arrayAuthors)) {
            $this->getZip()->addFromString('ppt/commentAuthors.xml', $this->writeCommentsAuthors($arrayAuthors));
        }

        return $this->getZip();
    }

    /**
     * @param Author[] $arrayAuthors
     *
     * @return string
     */
    protected function writeCommentsAuthors($arrayAuthors)
    {
        $objWriter = new XMLWriter(XMLWriter::STORAGE_MEMORY);
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // p:cmAuthorLst
        $objWriter->startElement('p:cmAuthorLst');
        $objWriter->writeAttribute('xmlns:p', 'http://schemas.openxmlformats.org/presentationml/2006/main');

        $idxAuthor = 0;
        foreach ($arrayAuthors as $oAuthor) {
            $oAuthor->setIndex($idxAuthor++);

            // p:cmAuthor
            $objWriter->startElement('p:cmAuthor');
            $objWriter->writeAttribute('id', $oAuthor->getIndex());
            $objWriter->writeAttribute('name', $oAuthor->getName());
            $objWriter->writeAttribute('initials', $oAuthor->getInitials());
            $objWriter->writeAttribute('lastIdx', '2');
            $objWriter->writeAttribute('clrIdx', 0);
            $objWriter->endElement();
        }

        // ## p:cmAuthorLst
        $objWriter->endElement();

        return $objWriter->getData();
    }
}
