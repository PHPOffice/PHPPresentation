<?php

declare(strict_types=1);

namespace PhpOffice\PhpPresentation\Slide\Layout;

class IWorkLayout
{
    private array $elements = [];
    private array $grid = [];
    private int $rows = 0;
    private int $columns = 0;

    public function addElement(string $type, array $properties): void
    {
        $this->elements[] = [
            'type' => $type,
            'properties' => $properties
        ];
    }

    public function setGrid(int $rows, int $columns): void
    {
        $this->rows = $rows;
        $this->columns = $columns;
        $this->grid = array_fill(0, $rows, array_fill(0, $columns, null));
    }

    public function placeElement(int $row, int $col, int $rowSpan = 1, int $colSpan = 1): void
    {
        // Validate placement
        if ($row + $rowSpan > $this->rows || $col + $colSpan > $this->columns) {
            throw new \InvalidArgumentException('Element placement out of grid bounds');
        }

        // Mark grid positions as occupied
        for ($r = $row; $r < $row + $rowSpan; $r++) {
            for ($c = $col; $c < $col + $colSpan; $c++) {
                $this->grid[$r][$c] = count($this->elements) - 1;
            }
        }
    }

    public function calculateElementBounds(float $slideWidth, float $slideHeight): array
    {
        $cellWidth = $slideWidth / $this->columns;
        $cellHeight = $slideHeight / $this->rows;
        $bounds = [];

        // Calculate bounds for each element
        foreach ($this->elements as $index => $element) {
            $elementBounds = $this->findElementBounds($index);
            if ($elementBounds) {
                $bounds[$index] = [
                    'x' => $elementBounds['col'] * $cellWidth,
                    'y' => $elementBounds['row'] * $cellHeight,
                    'width' => $elementBounds['colSpan'] * $cellWidth,
                    'height' => $elementBounds['rowSpan'] * $cellHeight,
                ];
            }
        }

        return $bounds;
    }

    private function findElementBounds(int $elementIndex): ?array
    {
        $found = false;
        $bounds = ['row' => 0, 'col' => 0, 'rowSpan' => 0, 'colSpan' => 0];

        // Find top-left corner
        for ($r = 0; $r < $this->rows; $r++) {
            for ($c = 0; $c < $this->columns; $c++) {
                if ($this->grid[$r][$c] === $elementIndex) {
                    $bounds['row'] = $r;
                    $bounds['col'] = $c;
                    $found = true;
                    break 2;
                }
            }
        }

        if (!$found) {
            return null;
        }

        // Calculate span
        $r = $bounds['row'];
        $c = $bounds['col'];
        while ($r < $this->rows && $this->grid[$r][$c] === $elementIndex) {
            $bounds['rowSpan']++;
            $r++;
        }
        while ($c < $this->columns && $this->grid[$bounds['row']][$c] === $elementIndex) {
            $bounds['colSpan']++;
            $c++;
        }

        return $bounds;
    }
}
