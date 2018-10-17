<?php
error_reporting(E_ALL);
ini_set('display_errors', FALSE);
ini_set('display_startup_errors', FALSE);
if (PHP_SAPI == 'cli') {
	die('This example should only be run from a Web Browser');
}

$data1 = getData1();
$data2 = getData2();

foreach ($data1 as $k => $v) {
	$sn = $v[1];
	foreach ($data2 as $kk => $vv) {
		if ($sn == $vv[3]) {
			$data1[$k][15] = $vv['5'];
			$data1[$k][16] = $vv['8'];
			break;
		}
	}
}
writeData($data1);


function getData1() {
	$data = array();
	//读取excel
	set_time_limit(0); //设置页面等待时间
	error_reporting(E_ALL);
	date_default_timezone_set('Asia/ShangHai'); 


	$type='xlsx';
	$uploadfile = './songfang/a.xlsx';
	if ($uploadfile) {
	    require_once './PHPExcel-1.8.1/Classes/PHPExcel/IOFactory.php';
	    if($type=='xlsx'||$type=='xls' ){
	        $reader = \PHPExcel_IOFactory::createReader('Excel2007'); // 读取 excel 文档
	    }else if( $type=='csv' ){
	        $reader = \PHPExcel_IOFactory::createReader('CSV'); // 读取 excel 文档
	    }else{
	        die('Not supported file types!');
	    }

	    $PHPExcel = $reader->load($uploadfile); // 文档名称
	    $objWorksheet = $PHPExcel->getActiveSheet();
	    $highestRow = $objWorksheet->getHighestRow(); // 取得总行数‎⁨mac⁩ ▸ ⁨用户⁩ ▸ ⁨zhangpeng⁩ ▸ ⁨work⁩ ▸ ⁨laywork⁩ ▸ ⁨PHPExcel-1.8.1⁩ ▸ ⁨Classes⁩ ▸ ⁨PHPExcel⁩

	    $arr = array(1 => 'A', 2 => 'B', 3 => 'C', 4 => 'D', 5 => 'E', 6 => 'F', 7 => 'G', 8 => 'H', 9 => 'I', 10 => 'J', 11 => 'K', 12 => 'L', 13 => 'M', 14 => 'N', 15 => 'O', 16 => 'P', 17 => 'Q', 18 => 'R', 19 => 'S', 20 => 'T', 21 => 'U', 22 => 'V', 23 => 'W', 24 => 'X', 25 => 'Y', 26 => 'Z');
	    // 一次读取一列
	    for ($row = 2; $row <= $highestRow; $row++) {
	        for ($column = 0; $arr[$column+1] != 'P'; $column++) {
	            $val = $objWorksheet->getCellByColumnAndRow($column, $row)->getValue();
	            $data[$row][$column] = $val;
	        }
	    }
	    //print_r($data);
	}
	return $data;
}

function getData2() {
	$data = array();
	//读取excel
	set_time_limit(0); //设置页面等待时间
	error_reporting(E_ALL);
	date_default_timezone_set('Asia/ShangHai');

	$type='xlsx';
	$uploadfile = './songfang/b.xlsx';
	if ($uploadfile) {
	    require_once './PHPExcel-1.8.1/Classes/PHPExcel/IOFactory.php';
	    if($type=='xlsx'||$type=='xls' ){
	        $reader = \PHPExcel_IOFactory::createReader('Excel2007'); // 读取 excel 文档
	    }else if( $type=='csv' ){
	        $reader = \PHPExcel_IOFactory::createReader('CSV'); // 读取 excel 文档
	    }else{
	        die('Not supported file types!');
	    }

	    $PHPExcel = $reader->load($uploadfile); // 文档名称
	    $objWorksheet = $PHPExcel->getActiveSheet();
	    $highestRow = $objWorksheet->getHighestRow(); // 取得总行数

	    $arr = array(1 => 'A', 2 => 'B', 3 => 'C', 4 => 'D', 5 => 'E', 6 => 'F', 7 => 'G', 8 => 'H', 9 => 'I', 10 => 'J', 11 => 'K', 12 => 'L', 13 => 'M', 14 => 'N', 15 => 'O', 16 => 'P', 17 => 'Q', 18 => 'R', 19 => 'S', 20 => 'T', 21 => 'U', 22 => 'V', 23 => 'W', 24 => 'X', 25 => 'Y', 26 => 'Z');
	    // 一次读取一列
	    for ($row = 2; $row <= $highestRow; $row++) {
	        for ($column = 0; $arr[$column+1] != 'K'; $column++) {
	            $val = $objWorksheet->getCellByColumnAndRow($column, $row)->getValue();
	            $data[$row][$column] = $val;
	        }
	    }
	    //print_r($data);
	}
	return $data;
}

function writeData($data) {

	require_once './PHPExcel-1.8.1/Classes/PHPExcel.php';


	// Create new PHPExcel object
	$objPHPExcel = new PHPExcel();

	// Set document properties
	$objPHPExcel->getProperties()->setCreator("genglei")
								 ->setLastModifiedBy("genglei")
								 ->setTitle("Office 2007 XLSX Test Document")
								 ->setSubject("Office 2007 XLSX Test Document")
								 ->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
								 ->setKeywords("office 2007 openxml php")
								 ->setCategory("report file");



	// Add some data
	$objPHPExcel->setActiveSheetIndex(0)
	            ->setCellValue('A1', 'ID')
	            ->setCellValue('B1', 'SN号')
	            ->setCellValue('C1', '商户号')
	            ->setCellValue('D1', '商户名称')
	            ->setCellValue('E1', '机具类型')
	            ->setCellValue('F1', '进件人ID')
	            ->setCellValue('G1', '进件人姓名')
	            ->setCellValue('H1', '提货人ID')
	            ->setCellValue('I1', '提货人姓名')
	            ->setCellValue('J1', '是否新增商户')
	            ->setCellValue('K1', '是否有效商户')
	            ->setCellValue('L1', '商户审核通过日期')
	            ->setCellValue('M1', '终端绑定日期')
	            ->setCellValue('N1', '缉拿押金生效日期')
	            ->setCellValue('O1', '是否激活机具')
	            ->setCellValue('P1', '返货')
	            ->setCellValue('Q1', '开通代理')
	            ->setCellValue('R1', '返现日期');
						
	$i=2;

	$merge = $out = array();
	foreach($data as $item){
		if (!empty($item[16])) {
			$is = '是';
		} else {
			$is = '否';
		}

		$objPHPExcel->setActiveSheetIndex(0)
		            ->setCellValue('A'.$i, $item[0])
		            ->setCellValue('B'.$i, "\t".$item[1])
		            ->setCellValue('C'.$i, "\t".$item[2])
		            ->setCellValue('D'.$i, $item[3])
		            ->setCellValue('E'.$i, $item[4])
		            ->setCellValue('F'.$i, $item[5])
		            ->setCellValue('G'.$i, $item[6])
		            ->setCellValue('H'.$i, $item[7])
		            ->setCellValue('I'.$i, $item[8])
		            ->setCellValue('J'.$i, $item[9])
		            ->setCellValue('K'.$i, $item[10])
		            ->setCellValue('L'.$i, $item[11])
		            ->setCellValue('M'.$i, $item[12])
		            ->setCellValue('N'.$i, $item[13])
		            ->setCellValue('O'.$i, $item[14])
		            ->setCellValue('P'.$i, $is)
		            ->setCellValue('Q'.$i, $item[15])
		            ->setCellValue('R'.$i, $item[16]);
		$i++;	
	}					
		
	//$objPHPExcel->getActiveSheet()->getStyle('A1:I1')->getFont()->setBold(true);
	$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(10); 
	$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(10); 
	$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(30); 
	$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(10); 
	$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(20); 
	$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(12); 
	$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(20); 
	$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth(12);
	$objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth(12);
	$objPHPExcel->getActiveSheet()->getColumnDimension('J')->setWidth(12);
	$objPHPExcel->getActiveSheet()->setTitle('统计');


	// Set active sheet index to the first sheet, so Excel opens this as the first sheet
	$objPHPExcel->setActiveSheetIndex(0);

	header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
	header('Content-Disposition: attachment;filename="report_'.time().'.xlsx"');
	header('Cache-Control: max-age=0');

	$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
	$objWriter->save('php://output');
	exit;
}



?>