<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$activityfile = "";
$monthlyfile = "";
$leavefile = "";

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

//Upload Annual Leave
$target_file = $target_dir . basename($_FILES["leaveupload"]["name"]);
$imageFileType = pathinfo($target_file,PATHINFO_EXTENSION);
$uploadOk = 1;
$filename = $target_dir . time() . "3." . $imageFileType;
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
    if (move_uploaded_file($_FILES["leaveupload"]["tmp_name"], $filename)) {
        $leavefile = $filename;
    } else {
        $errormsg = "error";
    }
}

if(($activityfile!="") && ($monthlyfile!="") && ($leavefile!="")) {
    $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
    $activityworkbook = $reader->load($activityfile);
    $activitysheet = $activityworkbook->getSheet(0);
    $monthlyworkbook = $reader->load($monthlyfile);
    $monthlysheet = $monthlyworkbook->getSheet(0);
    $activitysheetdate = explode(" To ",trim($activitysheet->getCell('D8')->getValue()))[0];
    $activitysheetdatearray = explode("/",$activitysheetdate);
    $date = "Wage-" . DateTime::createFromFormat('!m',$activitysheetdatearray[1])->format('F') . " " . $activitysheetdatearray[2];
    $leavedate = DateTime::createFromFormat('!m',$activitysheetdatearray[1])->format('M') . "-" . substr($activitysheetdatearray[2],2);
    $exportfilename = "GWA Salary Monthly - " . $activitysheetdatearray[2] . " " . $activitysheetdatearray[1];
    $checkingcol = 0;
    for($col=1;$col<=\PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($monthlysheet->getHighestColumn());$col++) {
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
    $activityarray = array();
    for($row=11;$row<=$activitysheet->getHighestRow();$row++) {
        if(trim($activitysheet->getCellByColumnAndRow(2,$row)->getValue())!="") {
            $activityarray[trim($activitysheet->getCellByColumnAndRow(2,$row)->getValue())]['Wages'] = trim($activitysheet->getCellByColumnAndRow(3,$row)->getValue());
            $activityarray[trim($activitysheet->getCellByColumnAndRow(2,$row)->getValue())]['Deductions'] = trim($activitysheet->getCellByColumnAndRow(4,$row)->getValue());
            $activityarray[trim($activitysheet->getCellByColumnAndRow(2,$row)->getValue())]['Taxes'] = trim($activitysheet->getCellByColumnAndRow(5,$row)->getValue());
            $activityarray[trim($activitysheet->getCellByColumnAndRow(2,$row)->getValue())]['Net Pay'] = trim($activitysheet->getCellByColumnAndRow(6,$row)->getValue());
            $activityarray[trim($activitysheet->getCellByColumnAndRow(2,$row)->getValue())]['Expenses'] = trim($activitysheet->getCellByColumnAndRow(7,$row)->getValue());
        } else {
            break;
        }
    }
    $activityworkbook->disconnectWorksheets();
    unset($activityworkbook);
    $leaveworkbook = $reader->load($leavefile);
    $leavesheet = $leaveworkbook->getSheet(0);
    $leavearray = array();
    $balancecol = 0;
    for($col=1;$col<=\PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($leavesheet->getHighestColumn());$col++) {
        if(trim($leavesheet->getCellByColumnAndRow($col,1)->getFormattedValue())==$leavedate) {
            $balancecol = $col+2;
            break;
        }
    }
    for($row=3;$row<=$leavesheet->getHighestRow();$row++) {
        if((trim($leavesheet->getCellByColumnAndRow(1,$row)->getValue())!="") && (trim($leavesheet->getCellByColumnAndRow(1,$row)->getValue())!="Total")) {
            $leavearray[trim($leavesheet->getCellByColumnAndRow(1,$row)->getValue())] = (float)trim($leavesheet->getCellByColumnAndRow($balancecol,$row)->getCalculatedValue());
        }
    }
    $leaveworkbook->disconnectWorksheets();
    unset($leaveworkbook);
    $monthlyarray = array();
    for($row=9;$row<=$monthlysheet->getHighestRow();$row++) {
        if(trim($monthlysheet->getCellByColumnAndRow(3,$row)->getValue())!="") {
            $monthlysheet->setCellValueByColumnAndRow($checkingcol-1,$row,$activityarray[trim($monthlysheet->getCellByColumnAndRow(3,$row)->getValue())]['Wages']); //Wages
            $monthlysheet->setCellValueByColumnAndRow($checkingcol+2,$row,$activityarray[trim($monthlysheet->getCellByColumnAndRow(3,$row)->getValue())]['Deductions']); //Deductions
            $monthlysheet->setCellValueByColumnAndRow($checkingcol+3,$row,$activityarray[trim($monthlysheet->getCellByColumnAndRow(3,$row)->getValue())]['Taxes']); //Taxes
            $monthlysheet->setCellValueByColumnAndRow($checkingcol+4,$row,$activityarray[trim($monthlysheet->getCellByColumnAndRow(3,$row)->getValue())]['Net Pay']); //Net Pay
            $monthlysheet->setCellValueByColumnAndRow($checkingcol+5,$row,$activityarray[trim($monthlysheet->getCellByColumnAndRow(3,$row)->getValue())]['Expenses']); //Expenses
            $monthlysheet->setCellValueByColumnAndRow($checkingcol+6,$row,$leavearray[trim($monthlysheet->getCellByColumnAndRow(3,$row)->getValue())]); //Annual Leave
            $monthlysheet->setCellValueByColumnAndRow($checkingcol,$row,'=IF(' . \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($checkingcol-2) . (string)$row . '=' . \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($checkingcol-1) . (string)$row . ',"Match","Not Match")');
            $monthlyarray[trim($monthlysheet->getCellByColumnAndRow(3,$row)->getValue())] = 1;
            unset($activityarray[trim($monthlysheet->getCellByColumnAndRow(3,$row)->getValue())]);
        } else {
            break;
        }
    }
    if(!empty($activityarray)) {
        $monthlyarray = array_merge($monthlyarray,$activityarray);
        ksort($monthlyarray);
        $positionindex = 0;
        $positionarray = array();
        foreach($monthlyarray as $key => $value) {
            if(isset($activityarray[$key])) {
                $positionarray[$key] = $positionindex;
            }
            $positionindex++;
        }
        foreach($positionarray as $key => $value) {
            $monthlysheet->insertNewRowBefore($value+10,1);
            $monthlysheet->setCellValueByColumnAndRow(3,$value+10,$key);
            $monthlysheet->setCellValueByColumnAndRow($checkingcol-1,$value+10,$activityarray[$key]['Wages']);
            $monthlysheet->setCellValueByColumnAndRow($checkingcol+2,$value+10,$activityarray[$key]['Deductions']);
            $monthlysheet->setCellValueByColumnAndRow($checkingcol+3,$value+10,$activityarray[$key]['Taxes']);
            $monthlysheet->setCellValueByColumnAndRow($checkingcol+4,$value+10,$activityarray[$key]['Net Pay']);
            $monthlysheet->setCellValueByColumnAndRow($checkingcol+5,$value+10,$activityarray[$key]['Expenses']);
            $monthlysheet->setCellValueByColumnAndRow($checkingcol+6,$value+10,$leavearray[$key]);
            $monthlysheet->setCellValueByColumnAndRow($checkingcol,$value+10,'=IF(' . \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($checkingcol-2) . (string)($value+10) . '=' . \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($checkingcol-1) . (string)($value+10) . ',"Match","Not Match")');
        }
    }

    $writer = new Xlsx($monthlyworkbook);
    $writer->save('xlsxdownloads/' . $exportfilename . '.xlsx');
    $monthlyworkbook->disconnectWorksheets();
    unset($monthlyworkbook);
    $errormsg = (isset($_SERVER['HTTPS']) ? "https" : "http") . "://" . $_SERVER['HTTP_HOST'] . "/test/xlsxdownloads/" . $exportfilename . ".xlsx";
} else {
    $errormsg = "error";
}
echo $errormsg;
?>
