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

namespace PhpOffice\PhpPresentation\Writer\PDF;

use Dompdf\Dompdf as DomPDFLib;
use Dompdf\Options;
use PhpOffice\PhpPresentation\Writer\HTML;

class DomPDF extends HTML implements PDFWriterInterface
{
    /**
     * Save PhpPresentation to file.
     */
    public function save(string $filename): void
    {
        $this->isPDF = true;

        $html = $this->getHtmlContent();
        $html = str_replace(PHP_EOL, '', $html);

        $domPdf = new DomPDFLib(new Options());
        $domPdf->loadHtml($html);
        $domPdf->setPaper('a4', 'landscape');
        $domPdf->render();
        file_put_contents($filename, $domPdf->output());
    }
}
