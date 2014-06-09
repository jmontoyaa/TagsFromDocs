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

$header = array(
    'Alias',
    'Tag',
    'Content',
    'Checkbox selected',
    'Is float'
);

$table = array();
$table[] = $header;

if (file_exists($fileName)) {

    $word = \PhpOffice\PhpWord\IOFactory::load($fileName);
    $tags = $word->getTags();
    if (!empty($tags)) {
        foreach ($tags as $tag) {

            $isTag = isset($tag['property']['tag']) && in_array(
                $tag['property']['tag'],
                $tagsWanted
            );

            if (isset($tag['property']) &&
                isset($tag['property']['alias']) &&
                isset($tag['property']['tag'])
/**             &&
                (
                    $isTag ||
                    isset($tag['property']['is_checkbox']) && $tag['property']['is_checkbox'] ||
                    isset($tag['property']['is_date']) && $tag['property']['is_date']
                )
**/
            ) {
                $isCheckBox = isset($tag['property']['is_checkbox']) && $tag['property']['is_checkbox'];

                $float = null;
                if ($isTag) {
                    if (is_numeric($tag['content']['content'])) {
                        $float = (float)$tag['content']['content'];
                        if (is_float($float)) {
                            $float = 'true';
                        } else {
                            $float = 'false';
                        }
                    } else {
                        $float = 'false';
                    }
                }

                $checkBoxResult = null;
                if ($isCheckBox) {
                    $checkBoxResult = $tag['property']['checkbox_selected'];
                }

                $row = array(
                    $tag['property']['alias'],
                    $tag['property']['tag'],
                    $tag['content']['content'],
                    $checkBoxResult,
                    $float
                );

                if (!($checkBoxResult == '0')) {  
                    $table[] = $row;
                }
            }
        }

        if (!empty($table)) {
            echo '<table border="1" style="border-collapse:collapse;">';
            foreach ($table as $row) {
                echo '<tr><td>';
                echo $row[0];
                echo '<td>';
                echo $row[1];
                echo '<td>';
                echo $row[2];
                echo '</td><td>';
                echo $row[3];
                echo '</td><td>';
                echo $row[4];
                echo '</td></tr>';
            }
        }
    }

    // Extracts the document.xml file
    $template = new Template($fileName);
    file_put_contents(basename($fileName) . '.xml', $template->documentXML);
} else {
    echo "File doesn't exists $fileName";
}
exit;
