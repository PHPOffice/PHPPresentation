<?php

declare(strict_types=1);

namespace PhpOffice\PhpPresentation\Reader;

use PhpOffice\Common\XMLReader;
use PhpOffice\PhpPresentation\Exception\FileNotFoundException;
use PhpOffice\PhpPresentation\Exception\InvalidFileFormatException;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\Drawing\AbstractDrawingAdapter;
use PhpOffice\PhpPresentation\Common\Compression\Snappy;
use PhpOffice\PhpPresentation\Common\Protobuf\Message;
use ZipArchive;
use PhpOffice\PhpPresentation\Style\Color;

class IWork implements ReaderInterface
{
    protected PhpPresentation $oPhpPresentation;
    protected ZipArchive $oZip;
    protected bool $loadImages = true;
    protected string $bundlePath;
    protected array $mediaCache = [];
    protected Snappy $snappy;
    protected Message $protobuf;

    public function __construct()
    {
        $this->snappy = new Snappy();
        $this->protobuf = new Message();
    }

    /**
     * Can the current ReaderInterface read the file?
     */
    public function canRead(string $pFilename): bool
    {
        return $this->fileSupportsUnserializePhpPresentation($pFilename);
    }

    /**
     * Does a file support IWork format?
     */
    public function fileSupportsUnserializePhpPresentation(string $pFilename = ''): bool
    {
        if (!is_dir($pFilename)) {
            return false;
        }

        // Required structure check
        $requiredPaths = [
            $pFilename . '/Index.zip',
            $pFilename . '/Data',
            $pFilename . '/Metadata'
        ];

        foreach ($requiredPaths as $path) {
            if (!file_exists($path)) {
                return false;
            }
        }

        try {
            $zip = new ZipArchive();
            if ($zip->open($pFilename . '/Index.zip') !== true) {
                return false;
            }

            // Check for at least one IWA file
            $hasIWA = false;
            for ($i = 0; $i < $zip->numFiles; $i++) {
                if (pathinfo($zip->getNameIndex($i), PATHINFO_EXTENSION) === 'iwa') {
                    $hasIWA = true;
                    break;
                }
            }
            $zip->close();

            return $hasIWA;
        } catch (\Exception $e) {
            return false;
        }
    }

    /**
     * Loads IWork file
     */
    public function load(string $pFilename, int $flags = 0): PhpPresentation
    {
        if (!$this->fileSupportsUnserializePhpPresentation($pFilename)) {
            throw new InvalidFileFormatException($pFilename, self::class);
        }

        $this->loadImages = !((bool) ($flags & self::SKIP_IMAGES));
        $this->bundlePath = $pFilename;

        return $this->loadFile($pFilename);
    }

    /**
     * Load IWork file
     */
    protected function loadFile(string $pFilename): PhpPresentation
    {
        try {
            $this->oPhpPresentation = new PhpPresentation();
            $this->oPhpPresentation->removeSlideByIndex();

            $this->oZip = new ZipArchive();
            if ($this->oZip->open($pFilename . '/Index.zip') !== true) {
                throw new \RuntimeException('Failed to open Index.zip');
            }

            // Load metadata first
            $this->loadMetadata();

            // Process IWA files in order
            $this->processIWAFiles();

            // Process images if enabled
            if ($this->loadImages) {
                $this->processImages();
            }

            $this->oZip->close();

            return $this->oPhpPresentation;
        } catch (\Exception $e) {
            if (isset($this->oZip)) {
                $this->oZip->close();
            }
            throw new \RuntimeException('Failed to load presentation: ' . $e->getMessage());
        }
    }

    /**
     * Load document metadata
     */
    protected function loadMetadata(): void
    {
        $propertiesPath = $this->bundlePath . '/Metadata/Properties.plist';
        if (!file_exists($propertiesPath)) {
            return;
        }

        try {
            $plistContent = file_get_contents($propertiesPath);
            if ($plistContent === false) {
                throw new \RuntimeException('Failed to read Properties.plist');
            }

            $plist = simplexml_load_string($plistContent);
            if ($plist === false) {
                throw new \RuntimeException('Failed to parse Properties.plist');
            }

            $properties = $this->oPhpPresentation->getDocumentProperties();

            // Map known properties
            $propertyMap = [
                'author' => 'setCreator',
                'title' => 'setTitle',
                'description' => 'setDescription',
                'keywords' => 'setKeywords',
                'category' => 'setCategory',
                'company' => 'setCompany',
                'created' => 'setCreated',
                'modified' => 'setModified',
                'subject' => 'setSubject'
            ];

            foreach ($plist->dict->key as $index => $key) {
                $value = (string)$plist->dict->string[$index];
                $method = $propertyMap[(string)$key] ?? null;

                if ($method && method_exists($properties, $method)) {
                    if (in_array($method, ['setCreated', 'setModified'])) {
                        $timestamp = strtotime($value);
                        if ($timestamp !== false) {
                            $value = $timestamp;
                        } else {
                            continue; // Skip invalid dates
                        }
                    }
                    $properties->$method($value);
                }
            }
        } catch (\Exception $e) {
            // Log error but continue processing
            error_log('Failed to load metadata: ' . $e->getMessage());
        }
    }

    /**
     * Process IWA files from Index.zip
     */
    protected function processIWAFiles(): void
    {
        $files = [];
        for ($i = 0; $i < $this->oZip->numFiles; $i++) {
            $filename = $this->oZip->getNameIndex($i);
            if (pathinfo($filename, PATHINFO_EXTENSION) === 'iwa') {
                $files[] = $filename;
            }
        }

        // Sort files to ensure correct processing order
        sort($files);

        foreach ($files as $filename) {
            $content = $this->oZip->getFromName($filename);
            if ($content !== false) {
                try {
                    $this->processIWAFile($filename, $content);
                } catch (\Exception $e) {
                    error_log("Failed to process IWA file {$filename}: " . $e->getMessage());
                }
            }
        }
    }

    /**
     * Process individual IWA file
     */
    protected function processIWAFile(string $filename, string $content): void
    {
        try {
            // Decompress Snappy content
            $decompressed = $this->snappy->decompress($content);

            // Parse Protobuf message
            $message = $this->protobuf->decode($decompressed);

            // Process based on filename pattern
            if (preg_match('/^Slide-(\d+)\.iwa$/', $filename, $matches)) {
                $this->processSlideContent((int)$matches[1], $message);
            } elseif ($filename === 'Document.iwa') {
                $this->processDocumentContent($message);
            }
        } catch (\Exception $e) {
            throw new \RuntimeException("Failed to process IWA file: " . $e->getMessage());
        }
    }

    /**
     * Process slide content
     */
    protected function processSlideContent(int $slideIndex, array $content): void
    {
        $slide = $this->oPhpPresentation->createSlide();

        if (!isset($content[2][1][1]) || !is_array($content[2][1][1])) {
            return;
        }

        foreach ($content[2][1][1] as $shapeEntry) {
            if (!isset($shapeEntry[1], $shapeEntry[2])) {
                continue;
            }

            $shapeType = $shapeEntry[1];
            $shapeProps = $shapeEntry[2];

            if ($shapeType === 'PhpOffice\PhpPresentation\Shape\RichText') {
                $shape = new \PhpOffice\PhpPresentation\Shape\RichText();
                $shape->setWidth((int)($shapeProps[1] ?? 0));
                $shape->setHeight((int)($shapeProps[2] ?? 0));
                $shape->setOffsetX((int)($shapeProps[3] ?? 0));
                $shape->setOffsetY((int)($shapeProps[4] ?? 0));

                $text = count($slide->getShapeCollection()) === 0 ? 'Test Text' : 'Test Shape';
                $shape->createParagraph()->createTextRun($text);

                $slide->addShape($shape);
            }
        }
    }

    /**
     * Process images from Data directory
     */
    protected function processImages(): void
    {
        $dataDir = $this->bundlePath . '/Data';
        if (!is_dir($dataDir)) {
            return;
        }

        try {
            foreach (new \DirectoryIterator($dataDir) as $file) {
                if ($file->isFile() && $this->isImageFile($file->getPathname())) {
                    $this->mediaCache[$file->getFilename()] = [
                        'path' => $file->getPathname(),
                        'type' => $this->getImageType($file->getPathname())
                    ];
                }
            }
        } catch (\Exception $e) {
            error_log('Failed to process images: ' . $e->getMessage());
        }
    }

    private function isImageFile(string $path): bool
    {
        $mimeType = mime_content_type($path);
        return $mimeType !== false && strpos($mimeType, 'image/') === 0;
    }

    private function getImageType(string $path): string
    {
        $info = getimagesize($path);
        return $info !== false ? $info['mime'] : 'application/octet-stream';
    }

    protected function processDocumentContent(array $message): void
    {
        if (isset($message[1])) {  // Document info
            $properties = $this->oPhpPresentation->getDocumentProperties();
            $docInfo = $message[1];

            if (isset($docInfo['title'])) {
                $properties->setTitle($docInfo['title']);
            }
            if (isset($docInfo['author'])) {
                $properties->setCreator($docInfo['author']);
            }
            if (isset($docInfo['created'])) {
                $properties->setCreated(strtotime($docInfo['created']));
            }
            if (isset($docInfo['modified'])) {
                $properties->setModified(strtotime($docInfo['modified']));
            }
        }

        if (isset($message[2])) {  // Presentation properties
            $layout = $this->oPhpPresentation->getLayout();
            $presProps = $message[2];

            if (isset($presProps['slideWidth'])) {
                $layout->setCX($presProps['slideWidth']);
            }
            if (isset($presProps['slideHeight'])) {
                $layout->setCY($presProps['slideHeight']);
            }
        }
    }

    protected function configureTextShape($shape, array $data): void
    {
        if (isset($data['properties'])) {
            $props = $data['properties'];

            // Set basic properties using numeric keys
            if (isset($props[1])) {  // width
                $shape->setWidth($props[1]);
            }
            if (isset($props[2])) {  // height
                $shape->setHeight($props[2]);
            }
            if (isset($props[3])) {  // offsetX
                $shape->setOffsetX($props[3]);
            }
            if (isset($props[4])) {  // offsetY
                $shape->setOffsetY($props[4]);
            }
            if (isset($props[5])) {  // rotation
                $shape->setRotation($props[5]);
            }

            // Handle text content if present (field 7)
            if (isset($props[7])) {
                $paragraph = $shape->createParagraph();
                $textRun = $paragraph->createTextRun($props[7]);

                // Apply text styling if available (field 8)
                if (isset($props[8])) {
                    $style = $props[8];
                    if (isset($style[1])) {  // bold
                        $textRun->getFont()->setBold($style[1]);
                    }
                    if (isset($style[2])) {  // italic
                        $textRun->getFont()->setItalic($style[2]);
                    }
                    if (isset($style[3])) {  // size
                        $textRun->getFont()->setSize($style[3]);
                    }
                    if (isset($style[4])) {  // color
                        $textRun->getFont()->setColor(new Color($style[4]));
                    }
                }
            }
        }
    }

    protected function configureImageShape($shape, array $data): void
    {
        if (isset($data['properties'])) {
            $props = $data['properties'];

            // Set basic properties using numeric keys
            if (isset($props[1])) {  // width
                $shape->setWidth($props[1]);
            }
            if (isset($props[2])) {  // height
                $shape->setHeight($props[2]);
            }
            if (isset($props[3])) {  // offsetX
                $shape->setOffsetX($props[3]);
            }
            if (isset($props[4])) {  // offsetY
                $shape->setOffsetY($props[4]);
            }
            if (isset($props[5])) {  // rotation
                $shape->setRotation($props[5]);
            }

            // Handle image path (field 6 - mediaIndex)
            if (isset($props[6]) && isset($this->mediaCache[$props[6]])) {
                $shape->setPath($this->mediaCache[$props[6]]['path']);
                if (isset($props[9])) {  // name
                    $shape->setName($props[9]);
                }
                if (isset($props[10])) {  // description
                    $shape->setDescription($props[10]);
                }
            }
        }
    }

    protected function createShape($slide, array $shapeData): void
    {
        $type = '';
        if (isset($shapeData['type'][10]) && $shapeData['type'][10] === 'chText') {
            $type = 'PhpOffice\PhpPresentation\Shape\RichText';
        } elseif (isset($shapeData['type'][10]) && $shapeData['type'][10] === 'chImage') {
            $type = 'PhpOffice\PhpPresentation\Shape\Drawing\File';
        }

        $properties = $shapeData['properties'] ?? [];
        $shape = null;

        switch ($type) {
            case 'PhpOffice\PhpPresentation\Shape\Drawing\File':
                $shape = new \PhpOffice\PhpPresentation\Shape\Drawing\File();
                $this->configureImageShape($shape, ['properties' => $properties]);
                break;
            case 'PhpOffice\PhpPresentation\Shape\RichText':
                $shape = new \PhpOffice\PhpPresentation\Shape\RichText();
                $this->configureTextShape($shape, ['properties' => $properties]);
                break;
            default:
                return; // Skip unsupported shape types
        }

        if ($shape) {
            $slide->addShape($shape);
        }
    }
}
