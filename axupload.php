<?php
// include library
include 'vendor/autoload.php';

// get uploaded file
$nama_file = $_FILES["file"]["name"];
$type_file = $_FILES["file"]["type"];
$temp_file = $_FILES["file"]["tmp_name"];
$size_file = $_FILES["file"]["size"];

// read uploaded file 
$objPHPExcel = PHPExcel_IOFactory::load($temp_file);

// get sheet 0
$sheet = $objPHPExcel->getSheet(0);
// get highest row
$highestRow = $sheet->getHighestRow();
// get highest column
$highestColumn = $sheet->getHighestColumn();

// define array for save data
$array_data = array();

// looping row excel 
// karena row 1 title jadi dimulai dengan row 2
for($row = 2; $row <= $highestRow; $row++)
{
    // get data excel and change to array
    $data = $sheet->rangeToArray('A' . $row . ':' . $highestColumn . $row, NULL, TRUE, FALSE);
    // get data excel 0 = A | 1 = B dst 
    $area = $data[0][0];
    $status = $data[0][1];
    // cek area sudah masuk ke variable array data jika belum maka masukan data
    if(!in_array($area, array_column($array_data, "area")))
    {
        // memasukan data
        $array_data[] = array(
            "area"      => $area,
            "close"     => 0,
            "backend"   => 0,
            "total"     => 0
        );
    }

    // cari index array area diatas untuk keperluan input datanya
    $index_array = array_search($area, array_column($array_data, "area"));
    if($status == "CLOSE")
    {
        $array_data[$index_array]["close"] += 1;
        $array_data[$index_array]["total"] += 1;
    }
    elseif($status == "BACKEND")
    {
        $array_data[$index_array]["backend"] += 1;
        $array_data[$index_array]["total"] += 1;
    }
    else
    {
        $array_data[$index_array]["close"] += 1;
        $array_data[$index_array]["total"] += 1;
    }


}

// buat sheet baru
$newSheet = $objPHPExcel->createSheet();

// isi data header/titlenya
$newSheet->setCellValue('A1', "AREA");
$newSheet->setCellValue('B1', "BACKEND");
$newSheet->setCellValue('C1', "CLOSE");
$newSheet->setCellValue('D1', "GRAND TOTAL");

// initiasi variable mulai rownya di sheet yang baru 
$row = 2;
// mendefinisikan variable total
$ttl_backend = 0;
$ttl_close = 0;

// looping dari variable penyimpanan data yang sudah dibuat dari proses diatas
foreach($array_data as $row_data)
{
    // mengisikan data
    $newSheet->setCellValue('A'.$row, $row_data["area"]);
    $newSheet->setCellValue('B'.$row, $row_data["backend"]);
    $newSheet->setCellValue('C'.$row, $row_data["close"]);
    $newSheet->setCellValue('D'.$row, $row_data["total"]);

    $ttl_backend += $row_data["backend"];
    $ttl_close += $row_data["close"];
    $row++;
}
// mengisikan total data
$newSheet->setCellValue('B'.$row, $ttl_backend);
$newSheet->setCellValue('C'.$row, $ttl_close);

// We'll be outputting an excel file
header('Content-type: application/vnd.ms-excel');

// It will be called file.xls
header('Content-Disposition: attachment; filename="laporan.xls"');

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
// Write file to the browser
$objWriter->save('php://output');
?>