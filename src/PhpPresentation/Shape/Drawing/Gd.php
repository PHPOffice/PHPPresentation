<?php

namespace PhpOffice\PhpPresentation\Shape\Drawing;

class Gd extends AbstractDrawingAdapter
{
    /* Rendering functions */
    const RENDERING_DEFAULT = 'imagepng';
    const RENDERING_PNG = 'imagepng';
    const RENDERING_GIF = 'imagegif';
    const RENDERING_JPEG = 'imagejpeg';

    /* MIME types */
    const MIMETYPE_DEFAULT = 'image/png';
    const MIMETYPE_PNG = 'image/png';
    const MIMETYPE_GIF = 'image/gif';
    const MIMETYPE_JPEG = 'image/jpeg';

    /**
     * Image resource
     *
     * @var resource
     */
    protected $imageResource;

    /**
     * Rendering function
     *
     * @var string
     */
    protected $renderingFunction = self::RENDERING_DEFAULT;

    /**
     * Mime type
     *
     * @var string
     */
    protected $mimeType = self::MIMETYPE_DEFAULT;

    /**
     * Unique name
     *
     * @var string
     */
    protected $uniqueName;

    /**
     * Gd constructor
     */
    public function __construct()
    {
        parent::__construct();
        $this->uniqueName = md5(rand(0, 9999) . time() . rand(0, 9999));
    }

    /**
     * Get image resource
     *
     * @return resource
     */
    public function getImageResource()
    {
        return $this->imageResource;
    }

    /**
     * Set image resource
     *
     * @param $value resource
     * @return $this
     */
    public function setImageResource($value = null)
    {
        $this->imageResource = $value;

        if (!is_null($this->imageResource)) {
            // Get width/height
            $this->width  = imagesx($this->imageResource);
            $this->height = imagesy($this->imageResource);
        }

        return $this;
    }

    /**
     * Get rendering function
     *
     * @return string
     */
    public function getRenderingFunction()
    {
        return $this->renderingFunction;
    }

    /**
     * Set rendering function
     *
     * @param  string                            $value
     * @return $this
     */
    public function setRenderingFunction($value = self::RENDERING_DEFAULT)
    {
        $this->renderingFunction = $value;
        return $this;
    }

    /**
     * Get mime type
     *
     * @return string
     */
    public function getMimeType()
    {
        return $this->mimeType;
    }

    /**
     * Set mime type
     *
     * @param  string $value
     * @return $this
     */
    public function setMimeType($value = self::MIMETYPE_DEFAULT)
    {
        $this->mimeType = $value;
        return $this;
    }

    /**
     * @return string
     */
    public function getContents()
    {
        ob_start();
        call_user_func($this->getRenderingFunction(), $this->getImageResource());
        $imageContents = ob_get_contents();
        ob_end_clean();
        return $imageContents;
    }

    /**
     * @return string
     */
    public function getExtension()
    {
        $extension = strtolower($this->getMimeType());
        $extension = explode('/', $extension);
        $extension = $extension[1];
        return $extension;
    }

    /**
     * @return string
     */
    public function getIndexedFilename()
    {
        return $this->uniqueName . $this->getImageIndex() . '.' . $this->getExtension();
    }
}
