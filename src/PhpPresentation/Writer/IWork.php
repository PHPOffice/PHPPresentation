<?php

declare(strict_types=1);

namespace PhpOffice\PhpPresentation\Writer;

use PhpOffice\Common\Adapter\Zip\ZipInterface;
use PhpOffice\PhpPresentation\Exception\DirectoryNotFoundException;
use PhpOffice\PhpPresentation\Exception\InvalidParameterException;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\Drawing\AbstractDrawingAdapter;
use PhpOffice\PhpPresentation\Common\Compression\Snappy;
use PhpOffice\PhpPresentation\Common\Protobuf\Message;
use ZipArchive;

class IWork extends AbstractWriter implements WriterInterface
{
    private string $bundlePath;
    private array $mediaFiles = [];
    private Snappy $snappy;
    private Message $protobuf;

    public function __construct(?PhpPresentation $pPhpPresentation = null)
    {
        $this->setPhpPresentation($pPhpPresentation ?? new PhpPresentation());
        $this->snappy = new Snappy();
        $this->protobuf = new Message();
    }

    public function save(string $pFilename): void
    {
        if (empty($pFilename)) {
            throw new InvalidParameterException('pFilename', '');
        }

        $this->bundlePath = $pFilename;

        try {
            // Create bundle structure
            $this->createBundleStructure();

            // Create Index.zip
            $this->createIndexZip();

            // Save media files
            $this->saveMediaFiles();

            // Create metadata
            $this->createMetadata();
        } catch (\Exception $e) {
            // Clean up on failure
            $this->cleanup();
            throw new \RuntimeException('Failed to save presentation: ' . $e->getMessage());
        }
    }

    /**
     * Create iWork bundle directory structure
     */
    private function createBundleStructure(): void
    {
        if (file_exists($this->bundlePath)) {
            throw new \RuntimeException('Destination path already exists');
        }

        if (
            !mkdir($this->bundlePath) ||
            !mkdir($this->bundlePath . '/Data') ||
            !mkdir($this->bundlePath . '/Metadata')
        ) {
            throw new DirectoryNotFoundException('Failed to create bundle structure');
        }
    }

    /**
     * Create and populate Index.zip
     */
    private function createIndexZip(): void
    {
        $zip = new ZipArchive();
        if ($zip->open($this->bundlePath . '/Index.zip', ZipArchive::CREATE) !== true) {
            throw new \RuntimeException('Failed to create Index.zip');
        }

        try {
            // Add document info
            $zip->addFromString('Document.iwa', $this->createDocumentIWA());

            // Add slides
            for ($i = 0; $i < $this->getPhpPresentation()->getSlideCount(); $i++) {
                $slide = $this->getPhpPresentation()->getSlide($i);
                $content = $this->createSlideIWA($slide);
                $zip->addFromString(sprintf('Slide-%d.iwa', $i + 1), $content);
            }

            $zip->close();
        } catch (\Exception $e) {
            $zip->close();
            throw $e;
        }
    }

    /**
     * Create Document IWA content
     */
    private function createDocumentIWA(): string
    {
        $properties = $this->getPhpPresentation()->getDocumentProperties();

        $data = [
            1 => [  // Document info
                1 => $properties->getTitle(),       // title
                2 => $properties->getCreator(),     // author
                3 => $properties->getCreated(),     // created
                4 => $properties->getModified(),    // modified
            ],
            2 => [  // Presentation properties
                1 => $this->getPhpPresentation()->getLayout()->getCX(),    // slideWidth
                2 => $this->getPhpPresentation()->getLayout()->getCY(),    // slideHeight
            ]
        ];

        return $this->createIWAContent($data);
    }

    /**
     * Create Slide IWA content
     */
    private function createSlideIWA($slide): string
    {
        $shapes = [];
        $index = 1;
        foreach ($slide->getShapeCollection() as $shape) {
            $shapes[$index++] = $this->convertShapeToIWA($shape);
        }

        $data = [
            1 => [  // Slide info
                1 => $slide->getName() ?? '',  // name
                2 => $slide->getSlideLayout() ? get_class($slide->getSlideLayout()) : '',  // layout
            ],
            2 => [  // Shapes
                1 => $shapes  // shapes array with positive indices
            ]
        ];

        // var_dump([
        //     'slideData' => $data,
        //     'shapeCount' => count($slide->getShapeCollection()),
        //     'shapes' => array_map(function ($shape) {
        //         return [
        //             'class' => get_class($shape),
        //             'width' => $shape->getWidth(),
        //             'height' => $shape->getHeight(),
        //             'offsetX' => $shape->getOffsetX(),
        //             'offsetY' => $shape->getOffsetY(),
        //         ];
        //     }, $slide->getShapeCollection())
        // ]);

        return $this->createIWAContent($data);
    }

    /**
     * Convert shape to IWA format
     */
    private function convertShapeToIWA($shape): array
    {
        $data = [
            1 => get_class($shape),  // type
            2 => [  // properties
                1 => $shape->getWidth(),    // width
                2 => $shape->getHeight(),   // height
                3 => $shape->getOffsetX(),  // offsetX
                4 => $shape->getOffsetY(),  // offsetY
                5 => $shape->getRotation(), // rotation
            ]
        ];

        if ($shape instanceof AbstractDrawingAdapter) {
            $mediaIndex = $this->addMediaFile($shape);
            $data[2][6] = $mediaIndex;  // mediaIndex
        }

        return $data;
    }

    /**
     * Create IWA content with Snappy compression
     */
    private function createIWAContent(array $data): string
    {
        try {
            // Convert to Protobuf
            $protobuf = $this->protobuf->encode($data);

            // Compress with Snappy
            return $this->snappy->compress($protobuf);
        } catch (\Exception $e) {
            throw new \RuntimeException('Failed to create IWA content: ' . $e->getMessage());
        }
    }

    private function addMediaFile($shape): int
    {
        $path = $shape->getPath();
        if (!isset($this->mediaFiles[$path])) {
            $this->mediaFiles[$path] = [
                'index' => count($this->mediaFiles) + 1,
                'name' => $shape->getName(),
                'description' => $shape->getDescription()
            ];
        }
        return $this->mediaFiles[$path]['index'];
    }

    /**
     * Save media files to Data directory
     */
    private function saveMediaFiles(): void
    {
        foreach ($this->mediaFiles as $path => $info) {
            $extension = pathinfo($path, PATHINFO_EXTENSION);
            $newPath = sprintf(
                '%s/Data/%d.%s',
                $this->bundlePath,
                $info['index'],
                $extension
            );

            if (!copy($path, $newPath)) {
                throw new \RuntimeException("Failed to copy media file: {$path}");
            }
        }
    }

    /**
     * Create metadata files
     */
    private function createMetadata(): void
    {
        $properties = $this->getPhpPresentation()->getDocumentProperties();

        $plist = $this->createPropertyList([
            'author' => $properties->getCreator(),
            'title' => $properties->getTitle(),
            'description' => $properties->getDescription(),
            'keywords' => $properties->getKeywords(),
            'category' => $properties->getCategory(),
            'created' => date('c', $properties->getCreated()),
            'modified' => date('c', $properties->getModified())
        ]);

        if (file_put_contents($this->bundlePath . '/Metadata/Properties.plist', $plist) === false) {
            throw new \RuntimeException('Failed to write Properties.plist');
        }
    }

    private function createPropertyList(array $properties): string
    {
        $output = '<?xml version="1.0" encoding="UTF-8"?>' . PHP_EOL;
        $output .= '<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">' . PHP_EOL;
        $output .= '<plist version="1.0">' . PHP_EOL;
        $output .= '<dict>' . PHP_EOL;

        foreach ($properties as $key => $value) {
            if ($value !== null && $value !== '') {
                $output .= sprintf(
                    "    <key>%s</key>\n    <string>%s</string>\n",
                    htmlspecialchars($key, ENT_XML1, 'UTF-8'),
                    htmlspecialchars($value, ENT_XML1, 'UTF-8')
                );
            }
        }

        $output .= '</dict>' . PHP_EOL;
        $output .= '</plist>' . PHP_EOL;

        return $output;
    }

    private function cleanup(): void
    {
        if (!empty($this->bundlePath) && file_exists($this->bundlePath)) {
            $this->removeDirectory($this->bundlePath);
        }
    }

    private function removeDirectory(string $path): void
    {
        if (!is_dir($path)) {
            return;
        }

        $files = new \FilesystemIterator($path, \FilesystemIterator::SKIP_DOTS);
        foreach ($files as $file) {
            if ($file->isDir()) {
                $this->removeDirectory($file->getPathname());
            } else {
                unlink($file->getPathname());
            }
        }
        rmdir($path);
    }
}
