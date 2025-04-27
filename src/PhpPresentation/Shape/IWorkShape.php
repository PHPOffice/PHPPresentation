<?php

declare(strict_types=1);

namespace PhpOffice\PhpPresentation\Shape;

class IWorkShape
{
    private array $shapes = [];

    public function addShape(string $type, array $properties): int
    {
        $this->shapes[] = [
            'type' => $type,
            'properties' => array_merge([
                'x' => 0,
                'y' => 0,
                'width' => 0,
                'height' => 0,
                'rotation' => 0,
                'fill' => null,
                'stroke' => null,
                'shadow' => null,
            ], $properties)
        ];
        return count($this->shapes) - 1;
    }

    public function addImage(string $path, array $properties): int
    {
        return $this->addShape('image', array_merge([
            'path' => $path,
            'preserveAspectRatio' => true,
        ], $properties));
    }

    public function addTextBox(string $text, array $properties): int
    {
        return $this->addShape('textbox', array_merge([
            'text' => $text,
            'textStyle' => null,
            'paragraphStyle' => null,
        ], $properties));
    }

    public function getShape(int $index): ?array
    {
        return $this->shapes[$index] ?? null;
    }

    public function updateShape(int $index, array $properties): void
    {
        if (isset($this->shapes[$index])) {
            $this->shapes[$index]['properties'] = array_merge(
                $this->shapes[$index]['properties'],
                $properties
            );
        }
    }
}
