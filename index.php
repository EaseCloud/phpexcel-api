<?php include 'phpexcel/Classes/PHPExcel.php';

define('MIME_XLS', 'application/vnd.ms-excel');
define('MIME_XLSX', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

// Initialize
$export_name = 'excel.xls';
$export_type = 'Excel5';
$export_mime = MIME_XLS;
$objPHPExcel = new PHPExcel();

if(!empty($_FILES['template']) && $_FILES['template']['error'] === 0) {

    $file_uploaded = $_FILES['template'];

    // Init Template
    try {
        $objPHPExcel = PHPExcel_IOFactory::load($file_uploaded['tmp_name']);
    } catch(Exception $e) {}

    if($file_uploaded['type'] == MIME_XLSX) {
        $export_type = 'Excel2007';
        $export_mime = MIME_XLSX;
    }

    $export_name = $file_uploaded['name'];
}


// Fill
$worksheet = $objPHPExcel->setActiveSheetIndex(0);
if(!empty($_POST['data'])) {
    foreach(explode(chr(13), $_POST['data']) as $row) {
        $args = explode(chr(9), trim($row));
        $worksheet->setCellValue($args[0], $args[1]);
    }
}

// Export
// Redirect output to a client's web browser
header("Content-Type: $export_mime");
header("Content-Disposition: attachment;filename=\"$export_name\"");
header("Cache-Control: max-age=0");

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, $export_type);
$objWriter->save('php://output');
exit;