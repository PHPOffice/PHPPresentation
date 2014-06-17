<?php
/**
 * Header file
*/
use PhpOffice\PhpPowerpoint\Autoloader;
use PhpOffice\PhpPowerpoint\Settings;
use PhpOffice\PhpPowerpoint\IOFactory;

error_reporting(E_ALL);
define('CLI', (PHP_SAPI == 'cli') ? true : false);
define('EOL', CLI ? PHP_EOL : '<br />');
define('SCRIPT_FILENAME', basename($_SERVER['SCRIPT_FILENAME'], '.php'));
define('IS_INDEX', SCRIPT_FILENAME == 'index');

require_once __DIR__ . '/../src/PhpPowerpoint/Autoloader.php';
Autoloader::register();

// Set writers
$writers = array('PowerPoint2007' => 'pptx', 'ODPresentation' => 'odp');

// Return to the caller script when runs by CLI
if (CLI) {
	return;
}

// Set titles and names
$pageHeading = str_replace('_', ' ', SCRIPT_FILENAME);
$pageTitle = IS_INDEX ? 'Welcome to ' : "{$pageHeading} - ";
$pageTitle .= 'PHPPowerPoint';
$pageHeading = IS_INDEX ? '' : "<h1>{$pageHeading}</h1>";

// Populate samples
$files = '';
if ($handle = opendir('.')) {
	while (false !== ($file = readdir($handle))) {
		if (preg_match('/^Sample_\d+_/', $file)) {
			$name = str_replace('_', ' ', preg_replace('/(Sample_|\.php)/', '', $file));
			$files .= "<li><a href='{$file}'>{$name}</a></li>";
		}
	}
	closedir($handle);
}

/**
 * Write documents
 *
 * @param \PhpOffice\PhpWord\PhpWord $phpWord
 * @param string $filename
 * @param array $writers
 */
function write($phpPowerPoint, $filename, $writers)
{
	$result = '';
	
	// Write documents
	foreach ($writers as $writer => $extension) {
		$result .= date('H:i:s') . " Write to {$writer} format";
		if (!is_null($extension)) {
			$xmlWriter = IOFactory::createWriter($phpPowerPoint, $writer);
			$xmlWriter->save(__DIR__ . "/{$filename}.{$extension}");
			rename(__DIR__ . "/{$filename}.{$extension}", __DIR__ . "/results/{$filename}.{$extension}");
		} else {
			$result .= ' ... NOT DONE!';
		}
		$result .= EOL;
	}

	$result .= getEndingNotes($writers);

	return $result;
}

/**
 * Get ending notes
 *
 * @param array $writers
 */
function getEndingNotes($writers)
{
	$result = '';

	// Do not show execution time for index
	if (!IS_INDEX) {
		$result .= date('H:i:s') . " Done writing file(s)" . EOL;
		$result .= date('H:i:s') . " Peak memory usage: " . (memory_get_peak_usage(true) / 1024 / 1024) . " MB" . EOL;
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
				if (!is_null($type)) {
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
 * Creates a templated slide
 *
 * @param PHPPowerPoint $objPHPPowerPoint
 * @return PHPPowerPoint_Slide
 */
function createTemplatedSlide(PhpOffice\PhpPowerpoint\PhpPowerpoint $objPHPPowerPoint)
{
	// Create slide
	$slide = $objPHPPowerPoint->createSlide();
	
	// Add logo
	$shape = $slide->createDrawingShape();
	$shape->setName('PHPPowerPoint logo')
		->setDescription('PHPPowerPoint logo')
		->setPath('./resources/phppowerpoint_logo.gif')
		->setHeight(36)
		->setOffsetX(10)
		->setOffsetY(10);
	$shape->getShadow()->setVisible(true)
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
<link rel="stylesheet" href="bootstrap/css/phppowerpoint.css" />
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
            <a class="navbar-brand" href="./">PHPPowerPoint</a>
        </div>
        <div class="navbar-collapse collapse">
            <ul class="nav navbar-nav">
                <li class="dropdown active">
                    <a href="#" class="dropdown-toggle" data-toggle="dropdown"><i class="fa fa-code fa-lg"></i>&nbsp;Samples<strong class="caret"></strong></a>
                    <ul class="dropdown-menu"><?php echo $files; ?></ul>
                </li>
            </ul>
            <ul class="nav navbar-nav navbar-right">
                <li><a href="https://github.com/PHPOffice/PHPPowerPoint"><i class="fa fa-github fa-lg" title="GitHub"></i>&nbsp;</a></li>
                <li><a href="http://phppowerpoint.readthedocs.org/en/develop/"><i class="fa fa-book fa-lg" title="Docs"></i>&nbsp;</a></li>
                <li><a href="http://twitter.com/PHPOffice"><i class="fa fa-twitter fa-lg" title="Twitter"></i>&nbsp;</a></li>
            </ul>
        </div>
    </div>
</div>
<?php echo $pageHeading; ?>