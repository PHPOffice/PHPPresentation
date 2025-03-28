<?php
$im = imagecreatetruecolor(1, 1);
imagepng($im, __DIR__ . '/test_image.png');
imagedestroy($im);
