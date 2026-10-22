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

namespace PhpOffice\PhpPresentation\Reader;

use ArrayIterator;
use ArrayObject;
use PhpOffice\Common\File;
use PhpOffice\PhpPresentation\DocumentLayout;
use PhpOffice\PhpPresentation\DocumentProperties;
use PhpOffice\PhpPresentation\Exception\FileNotFoundException;
use PhpOffice\PhpPresentation\Exception\InvalidFileFormatException;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\PresentationProperties;
use PhpOffice\PhpPresentation\Reader\Serialized\DataValidator;
use PhpOffice\PhpPresentation\Shape;
use PhpOffice\PhpPresentation\Shape\Drawing\AbstractDrawingAdapter;
use PhpOffice\PhpPresentation\Shape\Drawing\File as DrawingFile;
use PhpOffice\PhpPresentation\Slide;
use PhpOffice\PhpPresentation\Style;
use SplObjectStorage;
use Throwable;
use ZipArchive;

/**
 * Serialized format reader.
 */
class Serialized implements ReaderInterface
{
    private const MAX_GRAPH_DEPTH = 512;
    private const MAX_GRAPH_VALUES = 1000000;
    private const MAX_SERIALIZED_VALUES = 100000;

    /**
     * Only presentation model classes may be restored from file data.
     * GD rendering callbacks and ArrayObject iterator classes are validated separately.
     * Keep this list explicit: a namespace prefix also permits application gadgets.
     */
    private const ALLOWED_CLASSES = [
        ArrayObject::class,
        DocumentLayout::class,
        DocumentProperties::class,
        PhpPresentation::class,
        PresentationProperties::class,
        Shape\AutoShape::class,
        Shape\Chart\Axis::class,
        Shape\Chart\Gridlines::class,
        Shape\Chart\Legend::class,
        Shape\Chart\Marker::class,
        Shape\Chart\PlotArea::class,
        Shape\Chart\Series::class,
        Shape\Chart\Title::class,
        Shape\Chart\Type\Area::class,
        Shape\Chart\Type\Bar::class,
        Shape\Chart\Type\Bar3D::class,
        Shape\Chart\Type\Doughnut::class,
        Shape\Chart\Type\Line::class,
        Shape\Chart\Type\Pie::class,
        Shape\Chart\Type\Pie3D::class,
        Shape\Chart\Type\Radar::class,
        Shape\Chart\Type\Scatter::class,
        Shape\Chart\View3D::class,
        Shape\Chart::class,
        Shape\Comment\Author::class,
        Shape\Comment::class,
        Shape\Drawing\Base64::class,
        DrawingFile::class,
        Shape\Drawing\Gd::class,
        Shape\Drawing\ZipFile::class,
        Shape\Group::class,
        Shape\Hyperlink::class,
        Shape\Line::class,
        Shape\Media::class,
        Shape\Placeholder::class,
        Shape\RichText\BreakElement::class,
        Shape\RichText\Field::class,
        Shape\RichText\Numbering::class,
        Shape\RichText\Paragraph::class,
        Shape\RichText\Run::class,
        Shape\RichText\TextElement::class,
        Shape\RichText::class,
        Shape\Table\Cell::class,
        Shape\Table\Row::class,
        Shape\Table::class,
        Slide\Animation::class,
        Slide\Background\Color::class,
        Slide\Background\Image::class,
        Slide\Background\SchemeColor::class,
        Slide\Layout::class,
        Slide\Note::class,
        Slide\SlideLayout::class,
        Slide\SlideMaster::class,
        Slide\Transition::class,
        Slide::class,
        Style\Alignment::class,
        Style\Border::class,
        Style\Borders::class,
        Style\Bullet::class,
        Style\Color::class,
        Style\ColorMap::class,
        Style\Fill::class,
        Style\Font::class,
        Style\Outline::class,
        Style\SchemeColor::class,
        Style\Shadow::class,
        Style\TextStyle::class,
    ];

    /**
     * Can the current \PhpOffice\PhpPresentation\Reader\ReaderInterface read the file?
     */
    public function canRead(string $pFilename): bool
    {
        return $this->fileSupportsUnserializePhpPresentation($pFilename);
    }

    /**
     * Does a file support UnserializePhpPresentation ?
     */
    public function fileSupportsUnserializePhpPresentation(string $pFilename): bool
    {
        // Check if file exists
        if (!file_exists($pFilename)) {
            throw new FileNotFoundException($pFilename);
        }

        // File exists, does it contain PhpPresentation.xml?
        return File::fileExists("zip://$pFilename#PhpPresentation.xml");
    }

    /**
     * Loads PhpPresentation Serialized file.
     *
     * @SuppressWarnings("PHPMD.UnusedFormalParameter")
     */
    public function load(string $pFilename, int $flags = 0): PhpPresentation
    {
        // Check if file exists
        if (!file_exists($pFilename)) {
            throw new FileNotFoundException($pFilename);
        }

        // Unserialize... First make sure the file supports it!
        if (!$this->fileSupportsUnserializePhpPresentation($pFilename)) {
            throw new InvalidFileFormatException($pFilename, self::class);
        }

        try {
            return $this->loadSerialized($pFilename);
        } catch (InvalidFileFormatException $exception) {
            throw $exception;
        } catch (Throwable $exception) {
            throw new InvalidFileFormatException($pFilename, self::class, 'Invalid serialized presentation data');
        }
    }

    /**
     * Load PhpPresentation Serialized file.
     */
    private function loadSerialized(string $pFilename): PhpPresentation
    {
        $oArchive = new ZipArchive();
        if (true !== $oArchive->open($pFilename)) {
            throw new InvalidFileFormatException($pFilename, self::class);
        }

        try {
            $xmlContent = $oArchive->getFromName('PhpPresentation.xml');
            if (empty($xmlContent)) {
                throw new InvalidFileFormatException($pFilename, self::class, 'The file PhpPresentation.xml is malformed');
            }

            $xmlData = @simplexml_load_string($xmlContent, 'SimpleXMLElement', LIBXML_NONET);
            $serialized = $xmlData === false ? false : base64_decode((string) $xmlData->data, true);
            if ($serialized === false) {
                throw new InvalidFileFormatException($pFilename, self::class, 'Invalid serialized presentation data');
            }
            // Check internal restoration handlers before PHP can invoke them.
            (new DataValidator($serialized, self::ALLOWED_CLASSES, self::MAX_SERIALIZED_VALUES, self::MAX_GRAPH_DEPTH))->validate();
            $file = @unserialize($serialized, [
                'allowed_classes' => self::ALLOWED_CLASSES,
                'max_depth' => self::MAX_GRAPH_DEPTH,
            ]);
            $remaining = self::MAX_GRAPH_VALUES;
            if (!$file instanceof PhpPresentation || !$this->containsOnlyAllowedObjects($file, new SplObjectStorage(), $remaining)) {
                throw new InvalidFileFormatException($pFilename, self::class, 'Invalid serialized presentation data');
            }

            // Update media links
            for ($i = 0; $i < $file->getSlideCount(); ++$i) {
                foreach ($file->getSlide($i)->getShapeCollection() as $shape) {
                    if ($shape instanceof AbstractDrawingAdapter) {
                        $imgPath = 'zip://' . $pFilename . '#media/' . $shape->getImageIndex() . '/' . pathinfo($shape->getPath(), PATHINFO_BASENAME);
                        if ($shape instanceof DrawingFile) {
                            $shape->setPath($imgPath, false);
                        } else {
                            $shape->setPath($imgPath);
                        }
                    }
                }
            }

            return $file;
        } finally {
            $oArchive->close();
        }
    }

    /**
     * Reject incomplete objects anywhere in the graph, including cyclic properties.
     *
     * @param mixed $value
     * @param SplObjectStorage<object, null> $seen
     */
    private function containsOnlyAllowedObjects($value, SplObjectStorage $seen, int &$remaining, int $depth = 0): bool
    {
        if (--$remaining < 0 || $depth > self::MAX_GRAPH_DEPTH) {
            return false;
        }
        if (is_object($value)) {
            if (!in_array(get_class($value), self::ALLOWED_CLASSES, true)) {
                return false;
            }
            if ($seen->contains($value)) {
                return true;
            }
            $seen->attach($value);
            if ($value instanceof Shape\Drawing\Gd
                && !in_array($value->getRenderingFunction(), ['imagepng', 'imagegif', 'imagejpeg', 'imagebmp', 'imagewebp'], true)) {
                return false;
            }
            if ($value instanceof ArrayObject) {
                if ($value->getIteratorClass() !== ArrayIterator::class) {
                    return false;
                }
                $value = [$value->getArrayCopy(), get_object_vars($value)];
            } else {
                $value = (array) $value;
            }
        }
        if (is_array($value)) {
            foreach ($value as $child) {
                if (!$this->containsOnlyAllowedObjects($child, $seen, $remaining, $depth + 1)) {
                    return false;
                }
            }
        }

        return true;
    }
}
