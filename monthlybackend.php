<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$activityfile = "";
$monthlyfile = "";

$target_dir = "xlsxdownloads/";
if (!is_dir($target_dir)) {
    mkdir($target_dir, 0755, true);
}
$target_dir = "xlsxuploads/";
if (!is_dir($target_dir)) {
    mkdir($target_dir, 0755, true);
}

//Upload Activity Summary
$target_file = $target_dir . basename($_FILES["activityupload"]["name"]);
$imageFileType = pathinfo($target_file,PATHINFO_EXTENSION);
$uploadOk = 1;
$filename = $target_dir . time() . "1." . $imageFileType;
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
    if (move_uploaded_file($_FILES["activityupload"]["tmp_name"], $filename)) {
        $activityfile = $filename;
    } else {
        $errormsg = "error";
    }
}

//Upload GWA Monthly
$target_file = $target_dir . basename($_FILES["monthlyupload"]["name"]);
$imageFileType = pathinfo($target_file,PATHINFO_EXTENSION);
$uploadOk = 1;
$filename = $target_dir . time() . "2." . $imageFileType;
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
    if (move_uploaded_file($_FILES["monthlyupload"]["tmp_name"], $filename)) {
        $monthlyfile = $filename;
    } else {
        $errormsg = "error";
    }
}
if(($activityfile!="") && ($monthlyfile!="")) {
    $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
    $activityworkbook = $reader->load($activityfile);
    $activitysheet = $activityworkbook->getSheet(0);
    $monthlyworkbook = $reader->load($monthlyfile);
    $monthlysheet = $monthlyworkbook->getSheet(0);
    $activitysheetdate = explode(" To ",trim($activitysheet->getCell('D8')->getValue()))[0];
    $activitysheetdatearray = explode("/",$activitysheetdate);
    $date = "Wage-" . DateTime::createFromFormat('!m',$activitysheetdatearray[1])->format('F') . " " . $activitysheetdatearray[2];
    $exportfilename = "GWA Salary Monthly - " . $activitysheetdatearray[2] . " " . $activitysheetdatearray[1];
    $checkingcol = 0;
    for($col=1;$col<=200;$col++) {
        if(trim($monthlysheet->getCellByColumnAndRow($col,8)->getValue())=="Checking") {
            $checkingcol = $col;
            break;
        }
    }
    $monthlysheet->getColumnDimensionByColumn($checkingcol-2)->setVisible(false);
    $monthlysheet->insertNewColumnBeforeByIndex($checkingcol,1);
    $checkingcol = $checkingcol + 1;
    $monthlysheet->setCellValueByColumnAndRow($checkingcol-1,8, $date);
    $monthlysheet->getColumnDimensionByColumn($checkingcol-1)->setAutoSize(true);
    $writer = new Xlsx($monthlyworkbook);
    $writer->save('xlsxdownloads/' . $exportfilename . '.xlsx');
    $errormsg = (isset($_SERVER['HTTPS']) ? "https" : "http") . "://" . $_SERVER['HTTP_HOST'] . "/test/xlsxdownloads/" . $exportfilename . ".xlsx";
} else {
    $errormsg = "error";
}
echo $errormsg;
?>
