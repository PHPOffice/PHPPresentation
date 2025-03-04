<?php
/**
 * Header file.
 */
use PhpOffice\PhpPresentation\Autoloader;
use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\Drawing;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Slide;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Writer\PDF\DomPDF;

ini_set('display_errors', 1);
ini_set('display_startup_errors', 1);
error_reporting(E_ALL);

define('CLI', (PHP_SAPI == 'cli') ? true : false);
define('EOL', CLI ? PHP_EOL : '<br />');
define('SCRIPT_FILENAME', basename($_SERVER['SCRIPT_FILENAME'], '.php'));
define('IS_INDEX', SCRIPT_FILENAME == 'index');

require_once __DIR__ . '/../src/PhpPresentation/Autoloader.php';
Autoloader::register();

if (is_file(dirname(__DIR__) . '/vendor/autoload.php')) {
    require_once dirname(__DIR__) . '/vendor/autoload.php';
} else {
    if (is_file(__DIR__ . '/../../Common/src/Common/Autoloader.php')) {
        include_once __DIR__ . '/../../Common/src/Common/Autoloader.php';
        PhpOffice\Common\Autoloader::register();
    } else {
        throw new Exception('Can not find the vendor or the common folder!');
    }
}
// do some checks to make sure the outputs are set correctly.
if (false === is_dir(__DIR__ . DIRECTORY_SEPARATOR . 'results')) {
    throw new Exception('The results folder is not present!');
}
if (false === is_writable(__DIR__ . DIRECTORY_SEPARATOR . 'results' . DIRECTORY_SEPARATOR)) {
    throw new Exception('The results folder is not writable!');
}
if (false === is_writable(__DIR__ . DIRECTORY_SEPARATOR)) {
    throw new Exception('The samples folder is not writable!');
}

// Set titles and names
$pageHeading = str_replace('_', ' ', SCRIPT_FILENAME);
$pageTitle = IS_INDEX ? 'Welcome to ' : "{$pageHeading} - ";
$pageTitle .= 'PHPPresentation';
$pageHeading = IS_INDEX ? '' : "<h1>{$pageHeading}</h1>";

$oShapeDrawing = new Drawing\File();
$oShapeDrawing->setName('PHPPresentation logo')
    ->setDescription('PHPPresentation logo')
    ->setPath(__DIR__ . '/resources/phppowerpoint_logo.gif')
    ->setHeight(36)
    ->setOffsetX(10)
    ->setOffsetY(10);
$oShapeDrawing->getShadow()->setVisible(true)
    ->setDirection(45)
    ->setDistance(10);
$oShapeDrawing->getHyperlink()->setUrl('https://github.com/PHPOffice/PHPPresentation/')->setTooltip('PHPPresentation');

// Create a shape (text)
$oShapeRichText = new RichText();
$oShapeRichText->setHeight(300)
    ->setWidth(600)
    ->setOffsetX(170)
    ->setOffsetY(180);
$oShapeRichText->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
$textRun = $oShapeRichText->createTextRun('Thank you for using PHPPresentation!');
$textRun->getFont()->setBold(true)
    ->setSize(60)
    ->setColor(new Color('FFE06B20'));

// Return to the caller script when runs by CLI
if (CLI) {
    return;
}

// Populate samples
$files = [];
if ($handle = opendir('.')) {
    while (false !== ($file = readdir($handle))) {
        if (preg_match('/^Sample_\d+_/', $file)) {
            $name = str_replace('_', ' ', preg_replace('/(Sample_|\.php)/', '', $file));
            $group = substr($name, 0, 1);
            $id = substr($name, 0, 2);
            if (!isset($files[$group])) {
                $files[$group] = [];
            }
            $files[$group][$name] = "<li><a href='{$file}'>{$name}</a></li>";
        }
    }
    closedir($handle);

    foreach ($files as $group => $a) {
        natsort($files[$group]);
    }
    ksort($files);
}

/**
 * Write documents.
 */
function write(PhpPresentation $phpPresentation, string $filename, array $writers = [
    'PowerPoint2007' => 'pptx',
    'ODPresentation' => 'odp',
    'HTML' => 'html',
    'PDF' => 'pdf',
]): string
{
    $result = '';

    // Write documents
    foreach ($writers as $writer => $extension) {
        $result .= date('H:i:s') . " Write to {$writer} format";
        if (null !== $extension) {
            $pathFile = __DIR__ . "/results/{$filename}.{$extension}";
            if (file_exists($pathFile)) {
                unlink($pathFile);
            }
            $ioWriter = IOFactory::createWriter($phpPresentation, $writer);
            $ioWriter->setPDFAdapter(new DomPDF());
            $ioWriter->save($pathFile);
        } else {
            $result .= ' ... NOT DONE!';
        }
        $result .= EOL;
    }

    $result .= getEndingNotes($writers);

    return $result;
}

/**
 * Get ending notes.
 *
 * @param array $writers
 *
 * @return string
 */
function getEndingNotes($writers)
{
    $result = '';

    // Do not show execution time for index
    if (!IS_INDEX) {
        $result .= date('H:i:s') . ' Done writing file(s)' . EOL;
        $result .= date('H:i:s') . ' Peak memory usage: ' . (memory_get_peak_usage(true) / 1024 / 1024) . ' MB' . EOL;
    }

    // Return
    if (CLI) {
        $result .= 'The results are stored in the "results" subdirectory.' . EOL;
    } else {
        if (!IS_INDEX) {
            $types = array_values($writers);
            $result .= '<p>&nbsp;</p>';
            $result .= '<p>Results: ';
            foreach ($types as $type) {
                if (null !== $type) {
                    $resultFile = 'results/' . SCRIPT_FILENAME . '.' . $type;
                    if (file_exists($resultFile)) {
                        $result .= "<a href='{$resultFile}' class='btn btn-primary'>{$type}</a> ";
                    }
                }
            }
            $result .= '</p>';
        }
    }

    return $result;
}

/**
 * Creates a templated slide.
 */
function createTemplatedSlide(PhpPresentation $objPHPPresentation): Slide
{
    // Create slide
    $slide = $objPHPPresentation->createSlide();

    // Add logo
    $shape = $slide->createDrawingShape();
    $shape->setName('PHPPresentation logo')
        ->setDescription('PHPPresentation logo')
        ->setPath(__DIR__ . '/resources/phppowerpoint_logo.gif')
        ->setHeight(36)
        ->setOffsetX(10)
        ->setOffsetY(10);
    $shape->getShadow()
        ->setVisible(true)
        ->setDirection(45)
        ->setDistance(10);

    // Return slide
    return $slide;
}
?>
<title><?php echo $pageTitle; ?></title>
<meta charset="utf-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<link rel="stylesheet" href="bootstrap/css/bootstrap.min.css" />
<link rel="stylesheet" href="bootstrap/css/font-awesome.min.css" />
<link rel="stylesheet" href="bootstrap/css/phppresentation.css" />
</head>
<body>
<div class="container">
<div class="navbar navbar-default" role="navigation">
    <div class="container-fluid">
        <div class="navbar-header">
            <button type="button" class="navbar-toggle" data-toggle="collapse" data-target=".navbar-collapse">
            <span class="sr-only">Toggle navigation</span>
            <span class="icon-bar"></span>
            <span class="icon-bar"></span>
            <span class="icon-bar"></span>
            </button>
            <a class="navbar-brand" href="./">PHPPresentation</a>
        </div>
        <div class="navbar-collapse collapse">
            <ul class="nav navbar-nav">
                <?php foreach ($files as $key => $groupfiles) { ?>
                <li class="dropdown active">
                    <a href="#" class="dropdown-toggle" data-toggle="dropdown"><i class="fa fa-code fa-lg"></i>&nbsp;Samples <?php echo $key; ?>x<strong class="caret"></strong></a>
                    <ul class="dropdown-menu"><?php echo implode('', $groupfiles); ?></ul>
                </li>
                <?php } ?>
            </ul>
            <ul class="nav navbar-nav navbar-right">
                <li><a href="https://github.com/PHPOffice/PHPPresentation"><i class="fa fa-github fa-lg" title="GitHub"></i>&nbsp;</a></li>
                <li><a href="https://phpoffice.github.io/PHPPresentation/"><i class="fa fa-book fa-lg" title="Docs"></i>&nbsp;</a></li>
                <li><a href="http://twitter.com/PHPOffice"><i class="fa fa-twitter fa-lg" title="Twitter"></i>&nbsp;</a></li>
            </ul>
        </div>
    </div>
</div>
<?php echo $pageHeading;
