<?php
require 'vendor/autoload.php';

use \PhpOffice\PhpWord\PhpWord;
use \PhpOffice\PhpWord\Template;

$fileName = 'file.docx';

$tagsWanted = array(
    'NUNBER_OF_FTE',
    'ACTUAL_NUMBER_OF_HACCP_plans',
    'EFFECTIVE_ON_SITE_AUDIT_DURATION',
    'FSSC_Company_Number'
);

if (file_exists($fileName)) {

    $word = \PhpOffice\PhpWord\IOFactory::load($fileName);
    $tags = $word->getTags();
    foreach ($tags as $tag) {
        if (in_array($tag['property']['tag'], $tagsWanted)) {
            echo $tag['property']['alias'] . ': ' . $tag['content']['content'] . PHP_EOL;
        }
    }

    // Extracts document.xml
    $template = new Template($fileName);
    file_put_contents(basename($fileName) . '.xml', $template->documentXML);
} else {
    echo "File doens't exists $fileName";
}
exit;
