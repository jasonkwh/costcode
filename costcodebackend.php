<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$target_dir = "xlsxdownloads/";
if (!is_dir($target_dir)) {
    mkdir($target_dir, 0755, true);
}
$target_dir = "xlsxuploads/";
if (!is_dir($target_dir)) {
    mkdir($target_dir, 0755, true);
}
$filename = $target_dir . basename($_FILES["fileToUpload"]["name"]);
$imageFileType = pathinfo($filename,PATHINFO_EXTENSION);
$uploadOk = 1;
$errormsg = "";
// Allow certain file formats
if(strtolower($imageFileType) != "xls" && strtolower($imageFileType) != "xlsx") {
    $errormsg = "error";
    $uploadOk = 0;
}
// Check if $uploadOk is set to 0 by an error
if ($uploadOk == 0) {
    $errormsg = "error";
} else {
    if (move_uploaded_file($_FILES["fileToUpload"]["tmp_name"], $filename)) {
        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        $spreadsheet = $reader->load($filename);
        $worksheet = $spreadsheet->getSheet(0);
        $temparray = array();
        $startworking = 0;
        $employee = "";
        for($i=1;$i<=$worksheet->getHighestRow();$i++) {
            if($startworking == 1) {
                if((trim($worksheet->getCell('K' . $i)->getValue())!="COST CODE") && (trim($worksheet->getCell('K' . $i)->getValue())!="") && (trim($worksheet->getCell('J' . $i)->getValue())!="")) {
                    $temparray[trim($worksheet->getCell('K' . $i)->getValue())][trim($worksheet->getCell('J' . $i)->getValue())] += (float)trim($worksheet->getCell('L' . $i)->getValue()) + (float)trim($worksheet->getCell('M' . $i)->getValue()) + (float)trim($worksheet->getCell('N' . $i)->getValue()) + (float)trim($worksheet->getCell('O' . $i)->getValue()) + (float)trim($worksheet->getCell('P' . $i)->getValue()) + (float)trim($worksheet->getCell('Q' . $i)->getValue()) + (float)trim($worksheet->getCell('R' . $i)->getValue());
                }
            }
            if(trim($worksheet->getCell('K' . $i)->getValue())=="COST CODE") {
                $startworking = 1;
            }
            if(trim($worksheet->getCell('A' . $i)->getValue())=="EMPLOYEE NAME:") {
                $employee = str_replace(" ","_",str_replace("'","_",trim($worksheet->getCell('B' . $i)->getValue())));
            }
        }
        $spreadsheet2 = new Spreadsheet();
        $sheet = $spreadsheet2->getActiveSheet();
        $i = 2;
        $j = 2;
        $temperarray = array();
        $tempindex = 0;
        $lastkey = sizeof($temparray)+1;
        $alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        foreach ($temparray as $key => $values) {
            $sheet->setCellValue('A' . $i, $key);
            foreach($values as $valuekey => $value) {
                if(!in_array($valuekey,$temperarray)) {
                    $sheet->setCellValueByColumnAndRow($j,1, $valuekey);
                    $sheet->setCellValueByColumnAndRow($j,$lastkey+1, "=SUM(" . (string)$alphabet[$j-1] . "2:" . (string)$alphabet[$j-1] . (string)$lastkey . ")");
                    $sheet->getColumnDimensionByColumn($j)->setAutoSize(true);
                    array_push($temperarray,$valuekey);
                    $j++;
                }
                $tempindex = array_search($valuekey, $temperarray) + 2;
                $sheet->setCellValueByColumnAndRow($tempindex,$i, $value);
            }
            $i++;
        }
        $sheet->getColumnDimensionByColumn(1)->setAutoSize(true);
        $writer = new Xlsx($spreadsheet2);
        $writer->save('xlsxdownloads/' . $employee . '.xlsx');
        $spreadsheet->disconnectWorksheets();
        $spreadsheet2->disconnectWorksheets();
        unset($spreadsheet);
        unset($spreadsheet2);
        $errormsg = (isset($_SERVER['HTTPS']) ? "https" : "http") . "://" . $_SERVER['HTTP_HOST'] . "/test/xlsxdownloads/" . $employee . ".xlsx";
    } else {
        $errormsg = "error";
    }
}
echo $errormsg;
?>
