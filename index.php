<?php include 'phpexcel/Classes/PHPExcel.php';

define('MIME_XLS', 'application/vnd.ms-excel');
define('MIME_XLSX', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

// Initialize

$default_config = array(
    'export_name' => 'excel.xls',
    'export_type' => 'Excel5', // can be Excel2007
    'export_mime' => MIME_XLS, // can be MIME_XLSX
    'row_delimiter' => "\r",
    'col_delimiter' => "\t",
);

$user_config = isset($_POST['config']) ? json_decode($_POST['config'], true) : array();
$config = array_merge($default_config, $user_config);


$objPHPExcel = new PHPExcel();
if(!empty($_FILES['template']) && $_FILES['template']['error'] === 0) {

    $file_uploaded = $_FILES['template'];

    // Init Template
    try {
        $objPHPExcel = PHPExcel_IOFactory::load($file_uploaded['tmp_name']);
    } catch(Exception $e) {}
	
    if(!isset($user_config['export_mime']) and $file_uploaded['type'] == MIME_XLSX) {
		$config['export_type'] = 'Excel2007';
		$config['export_mime'] = MIME_XLSX;
	}

	if(!isset($user_config['export_name'])) {
		$config['export_name'] = $file_uploaded['name'];
	}
}


// Fill
$worksheet = $objPHPExcel->setActiveSheetIndex(0);

function xlsapi_fill(&$worksheet, $xlscript) {

    global $config;

    foreach(explode($config['row_delimiter'], $_POST['xlscript']) as $row) {
        $args = explode($config['col_delimiter'], trim($row));
        if($args[0] == 'SELECT_WORKSHEET') {
            $index = intval($args[1]);
            $worksheet = $objPHPExcel->setActiveSheetIndex($index);
        } elseif ($args[0] == 'FILL') {
            $cell = $args[1];
            $content = $args[2];
            $worksheet->setCellValue($cell, $content);
        } elseif ($args[0] == 'FILL2') {
            $col = intval($args[1]);
            $row = intval($args[2]);
            $content = $args[3];
            $worksheet->setCellValueByColumnAndRow($col, $row, $content);
        } elseif ($args[0] == 'MERGE') {
        } elseif ($args[0] == 'STYLE') {
        } elseif ($args[0] == 'FONTSIZE') {
        } elseif ($args[0] == 'WRAP_TEXT') {
            $begin = intval($args[1]);
            $end = intval($args[2]);
            $wrap = $args[3] != '0';
            $worksheet->getStyle("$begin:$end")->getAlignment()->setWrapText($wrap);
        } elseif ($args[0] == 'SET_URL') {
            $cell = $args[1];
            $url = $args[2];
            $worksheet->getCell($cell)->getHyperlink()->setUrl($url);
        }
    }

}


if(!empty($_POST['xlscript'])) {
    xlsapi_fill($worksheet, $_POST['xlscript']);
}


// Export
// Redirect output to a client's web browser
$ua = $_SERVER["HTTP_USER_AGENT"];
$export_name_encoded = str_replace("+", "%20",urlencode($config['export_name']));
if (preg_match("/MSIE/", $ua)) {
    header('Content-Disposition: attachment; filename="' . $export_name_encoded . '"');
} else if (preg_match("/Firefox/", $ua)) {
    header('Content-Disposition: attachment; filename*="utf8\'\'' . $export_name_encoded . '"');
} else if (preg_match("/python/i", $ua)) {
    header('Content-Disposition: attachment; filename="' . $export_name_encoded . '"');
} else {
    header('Content-Disposition: attachment; filename="' . $config['export_name'] . '"');
}
//header("Content-Disposition: attachment;filename=\"".urlencode($config['export_name'])."\"");
header("Content-Type: {$config['export_mime']}");
header("Cache-Control: max-age=0");

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, $config['export_type']);
$objWriter->save('php://output');
exit;
