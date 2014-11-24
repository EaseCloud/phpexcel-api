<?php include 'phpexcel/Classes/PHPExcel.php';

//print_r($_FILES);
//exit();

define('MIME_XLS', 'application/vnd.ms-excel');
define('MIME_XLSX', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

// Initialize
$export_name = 'excel.xls';
$export_type = 'Excel5';
$export_mime = MIME_XLS;

$default_config = array(
    'export_name' => 'excel.xls',
    'export_type' => 'Excel5', // can be Excel2007
    'export_mime' => MIME_XLS, // can be MIME_XLSX
    'row_delimeter' => "\r",
    'col_delimeter' => "\t",
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
		$export_type = 'Excel2007';
		$export_mime = MIME_XLSX;
	}

	if(!isset($user_config['export_name'])) {
		$export_name = $file_uploaded['name'];
	}
}


// Fill
$worksheet = $objPHPExcel->setActiveSheetIndex(0);
if(!empty($_POST['xlscript'])) {
    foreach(explode($config['row_delimeter'], $_POST['xlscript']) as $row) {
        $args = explode($config['col_delimeter'], trim($row));
		if($args[0] == 'FILL') {
			$cell = $args[1];
			$content = $args[2];
			$worksheet->setCellValue($cell, $content);
		} elseif ($args[0] == 'FILL2') {
		} elseif ($args[0] == 'MERGE') {
		} elseif ($args[0] == 'STYLE') {
		} elseif ($args[0] == 'FONTSIZE') {
		}
    }
}


// Export
// Redirect output to a client's web browser
$ua = $_SERVER["HTTP_USER_AGENT"];
$export_name_encoded = str_replace("+", "%20",urlencode($export_name));
if (preg_match("/MSIE/", $ua)) {
    header('Content-Disposition: attachment; filename="' . $export_name_encoded . '"');
} else if (preg_match("/Firefox/", $ua)) {
    header('Content-Disposition: attachment; filename*="utf8\'\'' . $export_name_encoded . '"');
} else if (preg_match("/python/i", $ua)) {
    header('Content-Disposition: attachment; filename="' . $export_name_encoded . '"');
} else {
    header('Content-Disposition: attachment; filename="' . $export_name . '"');
}
//header("Content-Disposition: attachment;filename=\"".urlencode($export_name)."\"");
header("Content-Type: $export_mime");
header("Cache-Control: max-age=0");

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, $export_type);
$objWriter->save('php://output');
exit;