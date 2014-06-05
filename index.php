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
    echo '<pre>';
    foreach ($tags as $tag) {
        $isTag = isset($tag['property']['tag']) && in_array($tag['property']['tag'], $tagsWanted);
        if (isset($tag['property']) && isset($tag['property']['alias']) &&
            (
                $isTag ||
                isset($tag['property']['is_checkbox']) && $tag['property']['is_checkbox'] ||
                isset($tag['property']['is_date']) && $tag['property']['is_date']
            )
        ) {
            echo $tag['property']['alias'] . ': ' .
                $tag['content']['content'];

            if ($isTag) {

                /*if (is_float($tag['content']['content'])) {
                    echo ' Is float: true';
                } else {
                    echo ' Is float: false';
                }*/
            }

            echo PHP_EOL;
        }
    }

    // Extracts document.xml
    $template = new Template($fileName);
    file_put_contents(basename($fileName) . '.xml', $template->documentXML);
} else {
    echo "File doesn't exists $fileName";
}
exit;
