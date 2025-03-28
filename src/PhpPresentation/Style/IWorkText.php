<?php

declare(strict_types=1);

namespace PhpOffice\PhpPresentation\Style;

class IWorkText
{
    private array $textStyles = [];
    private array $paragraphStyles = [];

    public function addTextStyle(array $properties): int
    {
        $this->textStyles[] = [
            'font' => $properties['font'] ?? 'Helvetica',
            'size' => $properties['size'] ?? 12,
            'color' => $properties['color'] ?? '000000',
            'bold' => $properties['bold'] ?? false,
            'italic' => $properties['italic'] ?? false,
            'underline' => $properties['underline'] ?? false,
            'strikethrough' => $properties['strikethrough'] ?? false,
        ];
        return count($this->textStyles) - 1;
    }

    public function addParagraphStyle(array $properties): int
    {
        $this->paragraphStyles[] = [
            'alignment' => $properties['alignment'] ?? 'left',
            'lineSpacing' => $properties['lineSpacing'] ?? 1.0,
            'spaceBefore' => $properties['spaceBefore'] ?? 0,
            'spaceAfter' => $properties['spaceAfter'] ?? 0,
            'indentLeft' => $properties['indentLeft'] ?? 0,
            'indentRight' => $properties['indentRight'] ?? 0,
            'indentFirstLine' => $properties['indentFirstLine'] ?? 0,
        ];
        return count($this->paragraphStyles) - 1;
    }

    public function getTextStyle(int $index): array
    {
        return $this->textStyles[$index] ?? [];
    }

    public function getParagraphStyle(int $index): array
    {
        return $this->paragraphStyles[$index] ?? [];
    }
}
