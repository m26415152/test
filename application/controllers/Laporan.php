<?php
defined('BASEPATH') OR exit('No direct script access allowed');

// Load library phpspreadsheet
require('./excel/vendor/autoload.php');

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
// End load library phpspreadsheet

class Laporan extends CI_Controller {

	// Load model
	public function __construct(){
		parent::__construct();
		$this->load->model('Provinsi_model');
	}

	// Main page
	public function index(){
		$provinsi = $this->Provinsi_model->listing();
		$data = array( 'title' => 'Test Export Laporan Excel ','provinsi' => $provinsi);
		//$this->load->view('laporan_view', $data, FALSE);
		$this->detail();
		$this->getarray();
	}

	// Export ke excel
	public function export(){
		$provinsi = $this->Provinsi_model->listing();
		// Create new Spreadsheet object
		$spreadsheet = new Spreadsheet();

		// Set document properties
		$spreadsheet->getProperties()	->setCreator('Test')
										->setLastModifiedBy('Test')
										->setTitle('Test')
										->setSubject('Test')
										->setDescription('Test')
										->setKeywords('Test')
										->setCategory('Test');

		// Add some data
		$spreadsheet->setActiveSheetIndex(0)
												->setCellValue('A1', 'KODE PROVINSI')
												->setCellValue('B1', 'NAMA PROVINSI');

		// Miscellaneous glyphs, UTF-8
		$i=2; foreach($provinsi as $provinsi) {
			$spreadsheet->setActiveSheetIndex(0)
			->setCellValue('A'.$i, $provinsi->id_provinsi)
			->setCellValue('B'.$i, $provinsi->nama_provinsi);
			$i++;
		}

		// Rename worksheet
		$spreadsheet->getActiveSheet()->setTitle('Report Excel '.date('d-m-Y'));

		// Set active sheet index to the first sheet, so Excel opens this as the first sheet
		$spreadsheet->setActiveSheetIndex(0);

		// Redirect output to a client’s web browser (Xlsx)
		header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
		header('Content-Disposition: attachment;filename="Laporan Excel.xlsx"');
		header('Cache-Control: max-age=0');
		// If you're serving to IE 9, then the following may be needed
		header('Cache-Control: max-age=1');

		// If you're serving to IE over SSL, then the following may be needed
		header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
		header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
		header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
		header('Pragma: public'); // HTTP/1.0

		$writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
		$writer->save('php://output');
		exit;
	}


public function detailpoo(){ 

		$objReader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader('Xlsx');
		//changed true if want to read dataonly not the border, cell width , cell height and stuff..
		$objReader->setReadDataOnly(false);

		//FileName and Sheet Name
		$objPHPExcel = $objReader->load('F_PUR_INVBLI.xlsx');
		
		//dapatkan sheet data
		//https://stackoverflow.com/questions/52007144/phpspreadsheet-foreach-loop-through-multiple-sheets
		$spreadsheet = $objPHPExcel;
		$spreadsheet_clone = $spreadsheet;

		//print_r($spreadsheet);
		//$sheetCount = $spreadsheet->getSheetCount();
		//$sheetName = $spreadsheet->getSheetNames();
		//for ($i = 0; $i < $sheetCount; $i++) {
			//print_r($sheetName[$i]."<br/>");
		   	//$sheet = $spreadsheet->getSheet($i);

		  	//$sheetData = $sheet->toArray(null, true, true, true);
		    //https://phpspreadsheet.readthedocs.io/en/develop/topics/reading-and-writing-to-file
		    //Embedding generated HTML in a web page
		  	//$writer = new \PhpOffice\PhpSpreadsheet\Writer\Html($spreadsheet_clone);
			//
			//tampilkan worksheet data sebagai html
			//echo $writer->generateSheetData();
			//$sheetIndex = $spreadsheet_clone->getIndex(
			//    $spreadsheet_clone->getSheetByName($sheetName[$i])
			//);
			//$spreadsheet_clone->removeSheetByIndex($sheetIndex);

			//agar worksheet tidak bug
			//$myWorkSheet = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($spreadsheet, 'My Data');
			//$spreadsheet_clone->addSheet($myWorkSheet);
		//}
		$writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
		$writer->save("05featuredemo.xlsx");
	}

public function detail(){ 

		$objReader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader('Xlsx');
		//$objReader->setReadDataOnly(true);

		//FileName and Sheet Name
		$objPHPExcel = $objReader->load('F_PUR_INVBLI.xlsx');

		//dapatkan sheet data
		//https://stackoverflow.com/questions/52007144/phpspreadsheet-foreach-loop-through-multiple-sheets
		$spreadsheet = $objPHPExcel;
		$spreadsheet_clone = $spreadsheet;

		//print_r($spreadsheet);
		$sheetCount = $spreadsheet->getSheetCount();
		$sheetName = $spreadsheet->getSheetNames();

		foreach($spreadsheet->getNamedRanges() as $name => $namedRange) {
		   print_r($name) . "</br>"; 
		   print_r($namedRange->getRange());
		   
		   //print_r($spreadsheet->getActiveSheet()->rangeToArray($namedRange->getRange()));
		   //print_r($spreadsheet->rangeToArray("A1"));
		}

		for ($i = 0; $i < $sheetCount; $i++) {
		    $sheet = $spreadsheet->getSheet($i);

		/*
		//print_r($sheet->getColumnDimension());
		foreach ($sheet->getRowIterator() as $row) {
		  $cellIterator = $row->getCellIterator();
		  $cellIterator->setIterateOnlyExistingCells(FALSE);
		  foreach ($cellIterator as $key => $cell) {
		  	//print_r($objPHPExcel->getActiveSheet()->getColumnDimension('A'));
		  	$vWidth = $objPHPExcel->getActiveSheet()->getColumnDimension($key)->getWidth();
		  	$objPHPExcel->getActiveSheet()->getPageSetup()->setFitToWidth($vWidth);
		  }
		  break;
		}*/


		    // $sheetData = $sheet->toArray(null, true, true, true);

			//print_r($sheet);		    
		    //https://phpspreadsheet.readthedocs.io/en/develop/topics/reading-and-writing-to-file
		    //Embedding generated HTML in a web page
		    $writer = new \PhpOffice\PhpSpreadsheet\Writer\Html($spreadsheet_clone);
			
			//tampilkan worksheet data
			/*print_r($sheetName[$i]."<br/>");
			echo $writer->generateSheetData();
			$sheetIndex = $spreadsheet_clone->getIndex(
			    $spreadsheet_clone->getSheetByName($sheetName[$i])
			);
			$spreadsheet_clone->removeSheetByIndex($sheetIndex);

			//agar worksheet tidak bug
			$myWorkSheet = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($spreadsheet, 'My Data');
			$spreadsheet_clone->addSheet($myWorkSheet);

			*/

		}

		// $writer2 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($objPHPExcel, "Xls");
	 //    header('Content-Disposition: attachment; filename="file.xls"');
	 //    $writer2->save("php://output");	 
	}

public function getarray(){
		$arr_excel = array();
		$arr_excel_temp = array();

		$objReader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader('Xlsx');
		$objReader->setReadDataOnly(false);

		//FileName and Sheet Name
		$objPHPExcel = $objReader->load('F_PUR_INVBLI.xlsx');

		$worksheet = $objPHPExcel->getSheetByName('#.Var1#.19');

		$highestRow = $worksheet->getHighestRow(); // e.g. 12
		$highestColumn = $worksheet->getHighestColumn(); // e.g M' 

		$highestColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestColumn); // e.g. 7


		//GET TEMPLATE FROM EXCEL
			for ($col = 1; $col <= $highestColumnIndex; ++$col){
			    for ($row = 1; $row <= $highestRow; ++$row){
			    	$arr_excel[$col][$row] =  $worksheet->getCellByColumnAndRow($col, $row)->getValue();
				}
			}

		//ECHO TEMPLATE TO WEBSITE
		echo "<h3>Template</h3>";
		foreach ($arr_excel as $obj_key => $coll){
			echo "Coloumn $obj_key <br>";
				
			foreach ($coll as $key=>$value){
				echo "$key : $value <br>";
				}

			echo "<br>";
			}


		//DUPLICATE DATA
		$data_detail = 3;
		$col_temp = 1;
		$row_temp = 1;
		for ($header = 1; $header <= 3; $header++){
			$col_temp = 1;
			$row_temp = ($header * $highestRow) + 1;
			//echo "Header : ".$header."<br><br>";
			for ($col = 1; $col <= $highestColumnIndex; ++$col){
			    for ($row = 1; $row <= $highestRow; ++$row){
			    	if($header == 2 && $data_detail > 1 && $row == 16){
			    		for($rownew = 1;$rownew <= $data_detail; $rownew++){
			    			$arr_excel_temp[$col_temp][$row_temp] =  $arr_excel[$col][$row];
			    			$row_temp++;
			    		}
			    	}
			    	else{
			    		$arr_excel_temp[$col_temp][$row_temp] =  $arr_excel[$col][$row];
			    		$row_temp++;
			    	}

			    	//echo "C:" .$col_temp;
			    	//echo " ";
			    	//echo "R:" .$row_temp;
			    	//echo " ";
				}
				//echo "<br>";
				$row_temp += -(1 * $highestRow);
				$col_temp++;
			}
			//echo "<br><br>";
		}

		//OUTPUT DUPLICATE TO WEB
		echo "<h3>OUTPUT</h3>";
		foreach ($arr_excel_temp as $obj_key_temp => $coll_temp){
			echo "Coloumn $obj_key_temp <br>";
				
			foreach ($coll_temp as $key_temp => $value_temp){
				echo "$key_temp : $value_temp <br>";
				}

			echo "<br>";
		}
			
		//SAVE DUPLICATE TO EXCEL
		$tempspreadsheet = new Spreadsheet();
		$i=1; foreach($arr_excel_temp as $Temporary) {
			$tempspreadsheet->setActiveSheetIndex(0)
							->setCellValue('A'.$i, $Temporary)
							->setCellValue('B'.$i, $Temporary)
							->setCellValue('C'.$i, $Temporary)
							->setCellValue('D'.$i, $Temporary)
							->setCellValue('E'.$i, $Temporary)
							->setCellValue('F'.$i, $Temporary)
							->setCellValue('G'.$i, $Temporary)
							->setCellValue('H'.$i, $Temporary)
							->setCellValue('I'.$i, $Temporary)
							->setCellValue('J'.$i, $Temporary)
							->setCellValue('K'.$i, $Temporary)
							->setCellValue('L'.$i, $Temporary);
			$i++;
		}


		

		$objPHPExcel->getActiveSheet()->duplicateStyle($objPHPExcel->getActiveSheet()->getStyle('A1:A21'),'A22');

		$writer = IOFactory::createWriter($objPHPExcel, 'Xlsx');
		$writer->save("Test.xlsx");

		$writer2 = IOFactory::createWriter($tempspreadsheet, 'Xlsx');
		$writer2->save("Test.xlsx");
	}
}

/* End of file Laporan.php */
/* Location: ./application/controllers/Laporan.php */

////////