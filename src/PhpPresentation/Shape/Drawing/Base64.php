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
     * @var array
     */
    protected $arrayMimeExtension = array(
        'image/jpeg' => 'jpg',
        'image/png' => 'png',
        'image/gif' => 'gif',
    );

    /**
     * Base64 constructor.
     */
    public function __construct()
    {
        parent::__construct();
        $this->uniqueName = md5(rand(0, 9999) . time() . rand(0, 9999));
    }

    /**
     * @return mixed
     */
    public function getData()
    {
        return $this->data;
    }

    /**
     * @param mixed $data
     * @return Base64
     */
    public function setData($data)
    {
        $this->data = $data;
        return $this;
    }

    /**
     * @return string
     */
    public function getContents()
    {
        list(, $imageContents) = explode(';', $this->getData());
        list(, $imageContents) = explode(',', $imageContents);
        return base64_decode($imageContents);
    }

    /**
     * @return string
     * @throws \Exception
     */
    public function getExtension()
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
    public function getIndexedFilename()
    {
        return $this->uniqueName . $this->getImageIndex() . '.' . $this->getExtension();
    }

    /**
     * @return string
     */
    public function getMimeType()
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
}
