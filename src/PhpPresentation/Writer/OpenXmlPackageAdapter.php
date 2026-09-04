<?php

declare(strict_types=1);

namespace PhpOffice\PhpPresentation\Writer;

use DK\OpenXml\OpenXmlPackage;
use DOMDocument;
use DOMElement;
use PhpOffice\Common\Adapter\Zip\ZipInterface;

/**
 * Write a package rather than a zip.
 *
 * The Writers hand their parts over one string at a time and say nothing about what a part is, so
 * this collects them all and hands them to the package at the end, when `[Content_Types].xml` says
 * what each one is and the `.rels` parts say what points at what.
 */
class OpenXmlPackageAdapter implements ZipInterface
{
    /**
     * @var string
     */
    private $filename = '';

    /**
     * @var array<string, string>
     */
    private $parts = [];

    /**
     * @param string $filename
     *
     * @return $this
     */
    public function open($filename)
    {
        $this->filename = $filename;
        $this->parts = [];

        return $this;
    }

    /**
     * @return $this
     */
    public function addFromString(string $localname, string $contents, bool $withCompression = true)
    {
        $this->parts[ltrim($localname, '/')] = $contents;

        return $this;
    }

    /**
     * @return $this
     */
    public function close()
    {
        $package = OpenXmlPackage::create();

        [$defaults, $overrides] = $this->readContentTypes();

        // an extension the Writers gave a type of its own keeps it: a part whose type the
        // extension already yields is written without an override of its own
        foreach ($defaults as $extension => $contentType) {
            $package->setDefaultContentType($extension, $contentType);
        }

        foreach ($this->parts as $name => $contents) {
            if ('[Content_Types].xml' === $name || $this->isRelationshipPart($name)) {
                continue;
            }
            $package->addPart('/' . $name, $this->contentType($name, $defaults, $overrides), $contents);
        }

        foreach ($this->parts as $name => $contents) {
            if (!$this->isRelationshipPart($name)) {
                continue;
            }
            $this->addRelationships($package, $name, $contents);
        }

        $package->saveAs($this->filename);
        $this->parts = [];

        return $this;
    }

    /**
     * Read what the Writers said each part is.
     *
     * @return array{array<string, string>, array<string, string>}
     */
    private function readContentTypes(): array
    {
        $defaults = [];
        $overrides = [];

        if (!isset($this->parts['[Content_Types].xml'])) {
            return [$defaults, $overrides];
        }

        $document = new DOMDocument();
        $document->loadXML($this->parts['[Content_Types].xml']);

        foreach ($document->getElementsByTagName('Default') as $element) {
            if ($element instanceof DOMElement) {
                $defaults[strtolower($element->getAttribute('Extension'))] = $element->getAttribute('ContentType');
            }
        }
        foreach ($document->getElementsByTagName('Override') as $element) {
            if ($element instanceof DOMElement) {
                $overrides['/' . ltrim($element->getAttribute('PartName'), '/')] = $element->getAttribute('ContentType');
            }
        }

        return [$defaults, $overrides];
    }

    /**
     * @param array<string, string> $defaults
     * @param array<string, string> $overrides
     */
    private function contentType(string $name, array $defaults, array $overrides): string
    {
        if (isset($overrides['/' . $name])) {
            return $overrides['/' . $name];
        }

        $extension = strtolower((string) pathinfo($name, PATHINFO_EXTENSION));

        return $defaults[$extension] ?? 'application/octet-stream';
    }

    private function isRelationshipPart(string $name): bool
    {
        return '.rels' === substr($name, -5);
    }

    /**
     * Replay a `.rels` part as relationships of the part it belongs to.
     */
    private function addRelationships(OpenXmlPackage $package, string $name, string $contents): void
    {
        $source = $this->sourceOf($name);

        $document = new DOMDocument();
        $document->loadXML($contents);

        foreach ($document->getElementsByTagName('Relationship') as $element) {
            if (!$element instanceof DOMElement) {
                continue;
            }
            $package->addRelationship(
                $element->getAttribute('Type'),
                $element->getAttribute('Target'),
                'External' === $element->getAttribute('TargetMode'),
                $element->getAttribute('Id') ?: null,
                $source
            );
        }
    }

    /**
     * `ppt/_rels/presentation.xml.rels` belongs to `/ppt/presentation.xml`, and `_rels/.rels` to
     * the package itself.
     */
    private function sourceOf(string $name): ?string
    {
        $directory = trim((string) substr($name, 0, (int) strrpos($name, '_rels/')), '/');
        $basename = substr(basename($name), 0, -5);

        if ('' === $basename) {
            return null;
        }

        return '/' . ('' === $directory ? '' : $directory . '/') . $basename;
    }
}
