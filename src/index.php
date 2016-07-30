<?php

/**
 * Routine bootstrap code.
 * 1. Constant define and global php settings
 * 2. Import PHPExcel Library
 * 3. Load the config object from $_POST['config']
 * 4. Initialize the $objPHPExcel object
 *
 * Const variables:
 * 1. MIME_XLS
 * 2. MIME_XLSX
 *
 * Global vars generated:
 * 1. $objPHPExcel: The operating php excel object
 * 2. $config: The overall merged config objects
 *  - export_name: the download file name, default: excel.xls
 *  - export_type: either Excel5 or Excel2007, default: Excel5
 *  - export_MIME: standard http Content-Type, default: MIME_XLS
 *  - row_delimiter: xlscript row delimiter, default: \n
 *  - col_delimiter: xlscript column delimiter, default: '|'
 *  - timezone: php timezone
 *  - debug: whether to show error
 *
 */

error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);


// 1. Constant define
define('MIME_JSON', 'application/json');
define('MIME_XLS', 'application/vnd.ms-excel');
define('MIME_XLSX', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
define('XLSCRIPT_ROOT', __DIR__ . '/');


// 2. Import PHPExcel Library
require XLSCRIPT_ROOT . "../PHPExcel/Classes/PHPExcel.php";
require 'inc/functions.php';


// 3. Load config
global $config;
$default_config = include('inc/config.php');
$user_config = json_decode(@$_POST['config'] ?: '[]', true);
$config = array_merge($default_config, $user_config);

date_default_timezone_set(@$config['timezone']);
//if (@$config['debug']) {
//    error_reporting(E_ALL);
//}


// 4. Initialize excel object
$objPHPExcel = new PHPExcel();
if (!empty($_FILES['template']) && $_FILES['template']['error'] === 0) {

    // Init Template if uploaded
    $file_uploaded = @$_FILES['template'];
    try {
        $objPHPExcel = PHPExcel_IOFactory::load($file_uploaded['tmp_name']);
    } catch (Exception $e) {
        var_dump($e);
        die;
    }

    if (!isset($user_config['export_mime']) and $file_uploaded['type'] == MIME_XLSX) {
        $config['export_type'] = 'Excel2007';
        $config['export_mime'] = MIME_XLSX;
    }

    if (!isset($user_config['export_name'])) {
        $config['export_name'] = $file_uploaded['name'];
    }

}


// 5. Read the xlscript from input
$xlscript = @$_POST['xlscript'] ?: '';

// Normalize the default row delimiter
$xlscript = str_replace("\r\n", "\n", $xlscript);
$xlscript = str_replace("\r", "\n", $xlscript);


// 6. Do the specified action
$action = @$_POST['action'] ?: 'write';

if ($action === 'write') {
    // Apply xlscript on the uploaded file then render output as download.
    xlsapi_fill($objPHPExcel, $xlscript);
    xlsapi_render($objPHPExcel);
} elseif (@$_POST['action'] === 'read') {
    // Read all the uploaded files and render the data as json.
    $data = xlsapi_read($objPHPExcel);
    header("Content-Type: " . MIME_JSON);
    exit(json_encode($data));
}



