<?php

namespace PhpOffice\PhpPresentation\Shape\Drawing;

class Base64 extends AbstractDrawingAdapter
{
    /**
     * @var string
     */
    protected $data;

    /**
     * Unique name
     *
     * @var string
     */
    protected $uniqueName;

    /**
     * @var array<string, string>
     */
    protected $arrayMimeExtension = array(
        'image/jpeg' => 'jpg',
        'image/png' => 'png',
        'image/gif' => 'gif',
    );

    /**
     * @var string
     */
    protected $path;

    /**
     * Base64 constructor.
     */
    public function __construct()
    {
        parent::__construct();
        $this->uniqueName = md5(rand(0, 9999) . time() . rand(0, 9999));
        $this->data = '';
    }

    /**
     * @return string
     */
    public function getData(): string
    {
        return $this->data;
    }

    /**
     * @param string $data
     * @return self
     */
    public function setData(string $data): self
    {
        $this->data = $data;
        return $this;
    }

    /**
     * @return string
     */
    public function getContents(): string
    {
        list(, $imageContents) = explode(';', $this->getData());
        list(, $imageContents) = explode(',', $imageContents);
        return base64_decode($imageContents);
    }

    /**
     * @return string
     * @throws \Exception
     */
    public function getExtension(): string
    {
        list($data, ) = explode(';', $this->getData());
        list(, $mime) = explode(':', $data);

        if (!array_key_exists($mime, $this->arrayMimeExtension)) {
            throw new \Exception('Type Mime not found : "'.$mime.'"');
        }
        return $this->arrayMimeExtension[$mime];
    }

    /**
     * @return string
     * @throws \Exception
     */
    public function getIndexedFilename(): string
    {
        return $this->uniqueName . $this->getImageIndex() . '.' . $this->getExtension();
    }

    /**
     * @return string
     */
    public function getMimeType(): string
    {
        $sImage = $this->getContents();
        if (!function_exists('getimagesizefromstring')) {
            $uri = 'data://application/octet-stream;base64,' . base64_encode($sImage);
            $image = getimagesize($uri);
        } else {
            $image = getimagesizefromstring($sImage);
        }
        return image_type_to_mime_type($image[2]);
    }

    /**
     * Get Path
     *
     * @return string
     */
    public function getPath(): string
    {
        return $this->path;
    }

    public function setPath(string $path): self
    {
        $this->path = $path;

        return $this;
    }
}
