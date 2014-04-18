<?php
require_once "PHPExcel.php";
require_once "PHPExcel/IOFactory.php";
require_once "PHPExcel/Reader/Excel5.php";


class LogSaveToExcel {
	
	public function readLog($filePath){
		$f = fopen($filePath, 'r');
		$data = array();
		$i = 0;
		while(!feof($f)){
			$line = fgets($f);  //读取log中一行的内容
			$row = explode(',', $line); //把一行内容分割
			$data[$i]=$row;
			$i++;
		}
		return $data;
	}
	
	public function dataprocess($data){
		//创建Excel对象
		$objExcel = new PHPExcel();
		//表头S1,S2,S3,S4,S5,S6,S7,S8,S9,S10,S11,S12,S13,S14,S15,S16,S17,S18,S19,S21,S22,S23,S24,S25,S26,S27,S28,S29,S32,S33,S34,S35,S36,S37,S38,S39,S40,S41,S42,S45,S46,S47,S48,S49,S50,S51,S53
		$columnName = array();
		//不取happendate,表头写死Date
		for($i = 1 ; $i < count($data[0]); $i++){
			$columnName[$i-1] = $data[0][$i];	
		}
		//不取第一行表头数据
		//三维数组存数据
		$value = array(array());
		$i = 0;
		$j = 0;
		for($r = 1 ; $r < count($data); $r++){
			//判断是不是同一个月，且最后一行数据的处理
			if($r == count($data)-1){
				$value[$i][$j] = $data[$r];
			}else{
				$first = explode('-', $data[$r][0]);
				$second = explode('-', $data[$r+1][0]);
				if($first[1] == $second[1]){
					$value[$i][$j] = $data[$r];
					$j++;
				}else{
					$value[$i][$j] = $data[$r];
					$i++;
					$j = 0;
				
				}
			}	
		}
		for($row = 0 ; $row < count($value);$row++){
			$this->saveToExcel($objExcel, $value[$row], $columnName, $row);  //多少个二维数组，做多少个工作表
		}
		//判断有几个月的数据
		if(count($value)==1){
			$dateValue = explode('-', $value[0][0][0]);
			$fileName = $dateValue[1].'月份'.'.xls';
		}else{
		    $maxValue = count($value);
		    $firstName = explode('-', $value[0][0][0]);
		    $lastName =  explode('-', $value[$maxValue-1][0][0]);
		    $fileName = $firstName[1].'-'.$lastName[1].'月份'.'.xls';
		}
		//输出到浏览器
		Header('content-Type:application/vnd.ms-excel;charset=gb2312');
		header("Content-Type: application/force-download");
		header("Content-Type: application/octet-stream");
		header("Content-Type: application/download");
		header('Content-Disposition:inline;filename="'.$fileName.'"');
		header("Content-Transfer-Encoding: binary");
		header("Expires: Mon, 26 Jul 1997 05:00:00 GMT");
		header("Last-Modified: " . gmdate("D, d M Y H:i:s") . " GMT");
		header("Cache-Control: must-revalidate, post-check=0, pre-check=0");
		header("Pragma: no-cache");
		$objWriter = new PHPExcel_Writer_Excel5($objExcel);//2003的流对象
		$objWriter->save('php://output');//参数-表示直接输出到浏览器，供客户端下载
		
	}
	
	 public function saveToExcel($objExcel,$data,$columnName,$i){
	 	//创建Excel对象
// 	 	if($objExcel == null){
// 	 		$objExcel = new PHPExcel();
// 	 	}
		$objActSheet = $objExcel->createSheet($i);
// 		$objExcel->setActiveSheetIndex(0);   //设置当前sheet索引，sheet默认0
// 		$objActSheet = $objExcel->getActiveSheet();  //获取当前sheet对象
		$strDate = $data[1][0];
		$arrDate = explode('-', $strDate);
		$sheetName = $arrDate[1].'.'.$arrDate[0];
		$title = $arrDate[0].'-'.$arrDate[1].' VN BILL';
		$fileName = $title.'.xls';
		$objActSheet->setTitle($sheetName);  //设置sheet名字
		
		$highestColumn = count($data[0])-1;//取得最大列数
		//$highestColumnIndex = PHPExcel_Cell::columnIndexFromString($highestColumn);//返回最大列数的索引
		$maxcolName = PHPExcel_Cell::stringFromColumnIndex($highestColumn);//返回最大的列名字
		//合并单元格并填值
		$objActSheet->mergeCells("A1:".$maxcolName."1");
		$objActSheet->setCellValue("A1",$title);
		
		$objActSheet->mergeCells("A2:A3");
		$objActSheet->setCellValue("A2",'Data');
		
		//画S1,S2,S3,S4,S5,S6,S7,S8,S9,S10,S11,S12,S13,S14,S15,S16,S17,S18,S19,S21,S22,S23,S24,S25,S26,S27,S28,S29,S32,S33,S34,S35,S36,S37,S38,S39,S40,S41,S42,S45,S46,S47,S48,S49,S50,S51,S53
		for($col = 0; $col <count($columnName)-1;$col++){
			$colName= PHPExcel_Cell::stringFromColumnIndex($col+1);//返回列名字
			$objActSheet->setCellValue($colName.'2',$columnName[$col]);
		}
		
		$objActSheet->setCellValue($maxcolName.'2','TOTAL');
		
		//画Real Gold
		for($col = 1; $col <count($data[0]);$col++){
			$colName= PHPExcel_Cell::stringFromColumnIndex($col);//返回列名字
			$objActSheet->setCellValue($colName.'3','Real Gold');
		}
		
		//填值
		$rowNumber = 4;
		for($row = 0 ; $row < count($data); $row++){
			for($column = 0; $column < count($data[$row]) ; $column++){
 				$colName= PHPExcel_Cell::stringFromColumnIndex($column);//返回列名字
 				$objActSheet->setCellValue($colName.$rowNumber,$data[$row][$column]);
				
			}
			$str ='=';
			//求每一行的和
			for($i = 1; $i < count($data[$row])-1; $i++){
				$colName= PHPExcel_Cell::stringFromColumnIndex($i);//返回列名字
				if($i == count($data[$row])-2){
					$str = $str.$colName.$rowNumber;
				}else{
					$str = $str.$colName.$rowNumber.'+';
				}	
			}
			$colName= PHPExcel_Cell::stringFromColumnIndex($i);//返回列名字
			$objActSheet->setCellValue($colName.$rowNumber,$str);
			$rowNumber++;
		}
		
		$maxRow = count($data)+4;
		$objActSheet->setCellValue('A'.$maxRow,'Total');
		//求每一列的和
		for($column = 1; $column < count($data[0]); $column++){
			$rowNumber=4;
			$colSumStr='=';
			$colName= PHPExcel_Cell::stringFromColumnIndex($column);//返回列名字
			for($row= 0; $row < count($data);$row++){
				if($row == count($data)-1){
					$colSumStr = $colSumStr.$colName.$rowNumber;
				}else{
					$colSumStr = $colSumStr.$colName.$rowNumber.'+';
				}
				$rowNumber++;
			}
			//echo $colSumStr.'</br>';
			$objActSheet->setCellValue($colName.$maxRow,$colSumStr);
		}
		
		
		//设置样式
		$tilteStyle = array(
				'font' => array(
						'bold' => true,
				),
				'alignment' => array(
						'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
						'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
				),
				'fill' => array(
						'type' => PHPExcel_Style_Fill::FILL_SOLID,
						'startcolor' => array(
								'argb' => 'FFFF83FA',
						),
						'endcolor' => array(
								'argb' => 'FFFF83FA',
						),
				),
		);
		
		$thStyle = array(
				'font' => array(
						'bold' => true,
				),
				'alignment' => array(
						'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
						'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
				),
		);
		
		$bodyStyle = array(
				'alignment' => array(
						'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
						'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
				),
				'borders' => array(
						'allborders' => array(
								'style' => PHPExcel_Style_Border::BORDER_THIN,
						),
				)
		);
		
		$tailStyle = array(
				'font' => array(
						'bold' => true,
				),
				'alignment' => array(
						'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
						'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
				),
		);
		
		$hightColumnStyle = array(
				'font' => array(
						'bold' => true,
				),
				'alignment' => array(
						'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
						'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
				),
				'borders' => array(
						'allborders' => array(
								'style' => PHPExcel_Style_Border::BORDER_THIN,
						),
				)
		);
		
		//执行样式
		$objActSheet->getStyle('A1:'.$maxcolName.'1')->applyFromArray($tilteStyle);
		$objActSheet->getStyle('B2:'.$maxcolName.'2')->applyFromArray($thStyle);
		$objActSheet->getStyle('A2:'.$maxcolName.$maxRow)->applyFromArray($bodyStyle);
		$objActSheet->getStyle($maxcolName.'2'.':'.$maxcolName.$maxRow)->applyFromArray($hightColumnStyle);
		$objActSheet->getStyle('A'.$maxRow.':'.$maxcolName.$maxRow)->applyFromArray($tailStyle);
	 }
}