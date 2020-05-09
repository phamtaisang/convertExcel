<?php
require '../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

//cấu hình
$FIELDS_MAP_CONFIG = [
    'name' => 'ten_san_pham',
    'quantity' => 'so_luong',
    'ma_sp' => 'sku',
    'mo_ta' => 'description',
];

$spreadsheet = new Spreadsheet();
$spreadsheet_output = new Spreadsheet();
$inputFileType = 'Xlsx';
$inputFileName = '../file/file_A.xlsx';
$inputFileOutput = '../file/file_B.xlsx';
$sheetname = 'Data Sheet #3';
//đọc file excel A
$reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader($inputFileType);
$spreadsheet = $reader->load($inputFileName);
$worksheet = $spreadsheet->getActiveSheet();
$worksheetData = $reader->listWorksheetInfo($inputFileName);
$header_row = $worksheet->toArray()[0];
$row = [];
$rowCount = 0;

//validate
$keyConfig = array_keys($FIELDS_MAP_CONFIG);
//check lỗi tên trường sai or thiếu tên trường
$missFields = array_diff($keyConfig, $header_row);
//check lỗi tồn tại tên trường
$unique = array_unique(array_diff_assoc($header_row, array_unique($header_row)));
if (($key = array_search("", $unique)) !== false) {
    unset($unique[$key]);
}

if ($missFields != null) {
    $error = implode(', ', $missFields);
    echo "not found : " . $error;
    exit;
} elseif ($unique != null) {
    $error = implode(', ', $unique);
    echo "2 columns already exist : " . $error;
    exit;
}

foreach ($worksheet->toArray() as $row) {
    $data_fill = array_combine($header_row, $row);
    foreach ($data_fill as $col => $df) {
        if (!in_array($col, array_keys($FIELDS_MAP_CONFIG))) {
            unset($data_fill[$col]);
        }
    }

    // order column
    $data_fill = array_merge(array_flip(array_keys($FIELDS_MAP_CONFIG)), $data_fill);

    // write file B
    $sheet = $spreadsheet_output->getActiveSheet();
    $spreadsheet->setActiveSheetIndex(0);
    $colAlphabet = 'A';
    foreach ($FIELDS_MAP_CONFIG as $toHeader) {
        $spreadsheet_output->getActiveSheet()->setCellValue("{$colAlphabet}1", $toHeader);
        if (substr($colAlphabet, strlen($colAlphabet) - 1) === 'Z') {
            $colAlphabet = substr($colAlphabet, 0, strlen($colAlphabet) - 1) . 'AA';
        } else {
            $colAlphabet++;
        }
    }
    $sheet = $spreadsheet_output->getActiveSheet();
    $colAlphabet = 'A';
    $rowCount++;
    foreach ($data_fill as $key => $data) {
        // Add some data
        $spreadsheet_output->getActiveSheet()->setCellValue($colAlphabet . $rowCount, $data);
        if (substr($colAlphabet, strlen($colAlphabet) - 1) === 'Z') {
            $colAlphabet = substr($colAlphabet, 0, strlen($colAlphabet) - 1) . 'AA';
        } else {
            $colAlphabet++;
        }
    }
}

//save
$writer = new Xlsx($spreadsheet_output);
$writer->save($inputFileOutput);

