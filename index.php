<?php 

include 'PHPExcel/Classes/PHPExcel.php';

define('MIME_XLS', 'application/vnd.ms-excel');
define('MIME_XLSX', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

date_default_timezone_set('Asia/Shanghai');

// Initialize

$default_config = array(
    'export_name' => 'excel.xls',
    'export_type' => 'Excel5', // can be Excel2007
    'export_mime' => MIME_XLS, // can be MIME_XLSX
    'row_delimiter' => "\r",
    'col_delimiter' => "|",
);

$user_config = isset($_POST['config']) ? json_decode($_POST['config'], true) : array();
$config = array_merge($default_config, $user_config);


$objPHPExcel = new PHPExcel();
if (!empty($_FILES['template']) && $_FILES['template']['error'] === 0) {

    $file_uploaded = $_FILES['template'];

    // Init Template
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


// Fill

/**
 * @param $objPHPExcel PHPExcel
 * @param $xlscript string
 */
function xlsapi_fill(&$objPHPExcel, $xlscript)
{

    global $config;

    $worksheet = $objPHPExcel->setActiveSheetIndex(0);

    foreach (explode($config['row_delimiter'], $xlscript) as $row) {

        $args = explode($config['col_delimiter'], trim($row));

        if ($args[0] == 'SELECT_WORKSHEET') {

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

            // 合并后的内容会以$begin单元格的内容填充
            $begin = $args[1];
            $end = $args[2];
            $worksheet->mergeCells("$begin:$end");

        } elseif ($args[0] == 'ALIGN') {

            /* Horizontal alignment styles */
//            const HORIZONTAL_GENERAL				= 'general';
//            const HORIZONTAL_LEFT					= 'left';
//            const HORIZONTAL_RIGHT					= 'right';
//            const HORIZONTAL_CENTER					= 'center';
//            const HORIZONTAL_CENTER_CONTINUOUS		= 'centerContinuous';
//            const HORIZONTAL_JUSTIFY				= 'justify';
//            const HORIZONTAL_FILL				    = 'fill';
//            const HORIZONTAL_DISTRIBUTED		    = 'distributed';        // Excel2007 only
            $begin = $args[1];
            $end = $args[2];
            $pValue = $args[3];
            $worksheet->getStyle("$begin:$end")->getAlignment()->setHorizontal($pValue);

        } elseif ($args[0] == 'VALIGN') {
            
            /* Vertical alignment styles */
//            const VERTICAL_BOTTOM					= 'bottom';
//            const VERTICAL_TOP						= 'top';
//            const VERTICAL_CENTER					= 'center';
//            const VERTICAL_JUSTIFY					= 'justify';
//            const VERTICAL_DISTRIBUTED		        = 'distributed';        // Excel2007 only
            $begin = $args[1];
            $end = $args[2];
            $pValue = $args[3];
            $worksheet->getStyle("$begin:$end")->getAlignment()->setVertical($pValue);
            
        } elseif ($args[0] == 'SET_BORDER') {

            $begin = $args[1];
            $end = $args[2];

            $borders = $worksheet->getStyle("$begin:$end")->getBorders();

            $border_position = @$args[3] ?: 'all';  // top left bottom right diagonal

            switch ($border_position) {
                case 'all':
                    $border = $borders->getAllBorders();
                    break;
                case 'outline':
                    $border = $borders->getOutline();
                    break;
                case 'inside':
                    $border = $borders->getInside();
                    break;
                case 'horizontal':
                    $border = $borders->getHorizontal();
                    break;
                case 'vertical':
                    $border = $borders->getVertical();
                    break;
                case 'top':
                    $border = $borders->getTop();
                    break;
                case 'right':
                    $border = $borders->getRight();
                    break;
                case 'bottom':
                    $border = $borders->getBottom();
                    break;
                case 'left':
                    $border = $borders->getLeft();
                    break;
                case 'diagonal':
                    $border = $borders->getDiagonal();
                    break;
                default:
                    $border = false;
            }
            if (!$border) continue;

            $border_style = @$args[4];
            if ($border_style) {
//                const BORDER_NONE				= 'none';
//                const BORDER_DASHDOT			= 'dashDot';
//                const BORDER_DASHDOTDOT			= 'dashDotDot';
//                const BORDER_DASHED				= 'dashed';
//                const BORDER_DOTTED				= 'dotted';
//                const BORDER_DOUBLE				= 'double';
//                const BORDER_HAIR				= 'hair';
//                const BORDER_MEDIUM				= 'medium';
//                const BORDER_MEDIUMDASHDOT		= 'mediumDashDot';
//                const BORDER_MEDIUMDASHDOTDOT	= 'mediumDashDotDot';
//                const BORDER_MEDIUMDASHED		= 'mediumDashed';
//                const BORDER_SLANTDASHDOT		= 'slantDashDot';
//                const BORDER_THICK				= 'thick';
//                const BORDER_THIN				= 'thin';
                $border->setBorderStyle($border_style);
            }

            $border_color = @$args[5];  // of type 'AARRGGBB'
            if ($border_color) {
                $border->getColor()->setARGB($border_color);
            }

        } elseif ($args[0] == 'STYLE') {
            // 设置字体样式

            $begin = $args[1];
            $end = $args[2];
            
            $style_type = strtoupper($args[3]);
            
            $style = $worksheet->getStyle("$begin:$end")->getFont();

            $toggle = substr($style_type, 0, 1) != '~';
            $action = preg_replace('/^~/', '', $style_type);

            switch ($action) {
                case 'BOLD': $style->setBold($toggle); break;
                case 'ITALIC': $style->setItalic($toggle); break;
                case 'UNDERLINE': $style->setUnderline($toggle); break;
                default: continue;
            }

        } elseif ($args[0] == 'SET_WIDTH') {

            $col = $args[1];
            $width = doubleval($args[2]);

            $worksheet->getColumnDimension($col)->setWidth($width);

        } elseif ($args[0] == 'SET_HEIGHT') {

            $row = $args[1];
            $height = doubleval($args[2]);

            $worksheet->getRowDimension($row)->setRowHeight($height);

        } elseif ($args[0] == 'FONT_SIZE') {

            $begin = $args[1];
            $end = $args[2];
            $font_size = doubleval($args[3]);

            $worksheet->getStyle("$begin:$end")->getFont()->setSize($font_size);


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


if (@$_POST['xlscript']) {
    xlsapi_fill($objPHPExcel, $_POST['xlscript']);
}

//exit('----');

// Export
// Redirect output to a client's web browser
$ua = $_SERVER["HTTP_USER_AGENT"];
$export_name_encoded = str_replace("+", "%20", urlencode($config['export_name']));
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
