<?php
require_once dirname(__FILE__) . './Classes/PHPExcel.php';
date_default_timezone_set('Europe/London');
define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');
$objPHPExcel = new PHPExcel();
// Set document properties
echo date('H:i:s') , " Set document properties" , EOL;
//$objPHPExcel->getProperties()->setCreator("Maarten Balliauw")
//    ->setLastModifiedBy("Maarten Balliauw")
//    ->setTitle("PHPExcel Test Document")
//    ->setSubject("PHPExcel Test Document")
//    ->setDescription("Test document for PHPExcel, generated using PHP classes.")
//    ->setKeywords("office PHPExcel php")
//    ->setCategory("Test result file");
//set default font
//$objPHPExcel->getDefaultStyle()->getFont()->setName('Arial')
//    ->setSize(12);
// Add some data
echo date('H:i:s') , " Add some data" , EOL;

$objPHPExcel->getActiveSheet()->getStyle()->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER); //单个单元格居中
$objPHPExcel->getDefaultStyle()->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getDefaultStyle()->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

$objPHPExcel->setActiveSheetIndex(0)
    ->setCellValue('A1', '姓名')
    ->setCellValue('B1', '年龄')
    ->setCellValue('C1', '性别')
    ->setCellValue('D1', '工资')
    ->setCellValue('E1', '日期');

// Miscellaneous glyphs, UTF-8
$time_val = date('Y-m-d H:i:s');
$objPHPExcel->setActiveSheetIndex(0)
    ->setCellValue('A2', '小明')
    ->setCellValue('b2', '23')
    ->setCellValue('c2', '男')
    ->setCellValue('d2', '2000')
    ->setCellValue('e2', $time_val);
$dataArray = array(
    array('xih','54','女',2070,$time_val),
    array('xiff','54','女',2060,$time_val),
    array('xfdgi','54','女',2090,$time_val),
    array('xifg','54','女',2500,$time_val),
    array('xif','54','女',2200,$time_val),
    array('xia','54','女',2230,$time_val),
);
$objPHPExcel->getActiveSheet()->fromArray($dataArray, NULL, 'A2');
$objPHPExcel->getActiveSheet()->getStyle('A1:E1')->getFont()->setBold(true); //设置title字体加粗
$objPHPExcel->getActiveSheet()->getStyle('A1:E1')->getFont()->setSize(16);  //设置title字体大小

//$objPHPExcel->getActiveSheet()->setAutoFilter($objPHPExcel->getActiveSheet()->calculateWorksheetDimension());   //增加自动筛选
//$objPHPExcel->getActiveSheet()->getStyle('C9')->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_YYYYMMDD2); //Y-m-d
//
//$objPHPExcel->getActiveSheet()->setCellValue('A10', 'Date/Time')
//    ->setCellValue('B10', 'Time')
//    ->setCellValue('C10', PHPExcel_Shared_Date::PHPToExcel( $dateTimeNow ));
//$objPHPExcel->getActiveSheet()->getStyle('C10')->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_TIME4); //H:i:s
//
//$objPHPExcel->getActiveSheet()->setCellValue('A11', 'Date/Time')
//    ->setCellValue('B11', 'Date and Time')
//    ->setCellValue('C11', PHPExcel_Shared_Date::PHPToExcel( $dateTimeNow ));
//$objPHPExcel->getActiveSheet()->getStyle('C11')->getNumberFormat()->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_DATETIME);//  d/m/y h:i
//getRowDimension()默认第一行
$objPHPExcel->getActiveSheet()->getRowDimension()->setRowHeight(-1);

//$objPHPExcel->getActiveSheet()->getStyle('A1')->getAlignment()->setWrapText(true);


$value = "-ValueA -Value B-Value Cccccccccccccccccccc";
$objPHPExcel->getActiveSheet()->setCellValue('A10', $value);
$objPHPExcel->getActiveSheet()->getRowDimension(10)->setRowHeight(-1);
$objPHPExcel->getActiveSheet()->getStyle('A10')->getAlignment()->setWrapText(true); //单元格内换行
$objPHPExcel->getActiveSheet()->getStyle('A10')->setQuotePrefix(true);

$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setAutoSize(true);    //宽度自适应




// Rename worksheet
$objPHPExcel->getActiveSheet()->setTitle('abc');


// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$objPHPExcel->setActiveSheetIndex(0);


// Save Excel 2007 file
echo date('H:i:s') , " Write to Excel2007 format" , EOL;
$callStartTime = microtime(true);

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save(str_replace('.php', '.xlsx', __FILE__));
$callEndTime = microtime(true);
$callTime = $callEndTime - $callStartTime;   //执行时间

// Save Excel 95 file
echo date('H:i:s') , " Write to Excel5 format" , EOL;
$callStartTime = microtime(true);

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save(str_replace('.php', '.xls', __FILE__));
