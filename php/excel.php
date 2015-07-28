<?php
//	パスを通す
set_include_path('/_lib/Classes/');
//	インクルード
include_once ( dirname ( __FILE__ ) . '/_lib/Classes/PHPExcel.php' );
include_once ( dirname ( __FILE__ ) . '/_lib/Classes/PHPExcel/IOFactory.php');


//	テンプレート読み込み
$reader		= PHPExcel_IOFactory::createReader ( 'Excel5' );
$template	= $reader->load ( dirname ( __FILE__ ) . '/template.xls' );

print ( dirname ( __FILE__ ) . '\template.xls' );

$templateSheet	= clone $template->getActiveSheet ();

$excel		= new PHPExcel ();

$needlessSheet = $excel->getActiveSheet ();

$sheet = $excel->addExternalSheet ( $templateSheet );
$sheet->setTitle ( 'テスト' );


//	出力

//  必要ないシートを削除
$sheetIndex = $excel->GetIndex ( $needlessSheet );
$excel->removeSheetByIndex ( $sheetIndex );

$writer	= PHPExcel_IOFactory::createWriter ( $excel, 'Excel5' );
$outputFileName = '/2015_07.xls';
$writer->save ( dirname ( __FILE__ ) . $outputFileName );

print ( dirname ( __FILE__ ) );
print ( $outputFileName );
?>