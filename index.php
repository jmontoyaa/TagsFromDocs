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

$table = array();
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

            $float = null;
            if ($isTag) {
                if (
                    is_numeric($tag['content']['content'])
                ) {
                    $float = (float) $tag['content']['content'];
                    if (is_float($float)) {
                        $float = ' Is float: true';
                    } else {
                        $float = ' Is float: false';
                    }

                } else {
                    $float = ' Is float: false';
                }
            }
            $row = array(
                $tag['property']['alias'],
                $tag['content']['content'],
                $float
            );
            $table[] = $row;
        }
    }

    echo '<table border="1" style="border-collapse:collapse;">';
    foreach ($table as $row) {
        echo '<tr><td>';
        echo $row[0];
        echo '<td>';
        echo $row[1];
        echo '</td><td>';
        echo $row[2];
        echo '</td></tr>';
    }

    // Extracts document.xml
    $template = new Template($fileName);
    file_put_contents(basename($fileName) . '.xml', $template->documentXML);
} else {
    echo "File doesn't exists $fileName";
}
exit;
