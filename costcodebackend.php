<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

function ContainsNumbers($String){
    return preg_match('/\\d/', $String) > 0;
}

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
        $costcodeopt = $_REQUEST['costcodeopt'];
        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        $spreadsheet = $reader->load($filename);
        $worksheet = $spreadsheet->getSheet(0);
        $temparray = array();
        $startworking = 0;
        $employee = "";
        $overtimetotal = 0.00;
        $j = 0;
        $valueincrement = 0;
        $checkcostcodefill = 0;
        if($costcodeopt==0) {
            for($i=1;$i<=$worksheet->getHighestRow();$i++) {
                if($startworking == 1) {
                    if((trim($worksheet->getCell('K' . $i)->getValue())!="COST CODE") && (trim($worksheet->getCell('K' . $i)->getValue())!="") && (trim($worksheet->getCell('J' . $i)->getValue())!="")) {
                        $temparray[trim($worksheet->getCell('K' . $i)->getCalculatedValue())][trim($worksheet->getCell('J' . $i)->getCalculatedValue())] += (float)trim($worksheet->getCell('L' . $i)->getValue()) + (float)trim($worksheet->getCell('M' . $i)->getValue()) + (float)trim($worksheet->getCell('N' . $i)->getValue()) + (float)trim($worksheet->getCell('O' . $i)->getValue()) + (float)trim($worksheet->getCell('P' . $i)->getValue()) + (float)trim($worksheet->getCell('Q' . $i)->getValue()) + (float)trim($worksheet->getCell('R' . $i)->getValue());
                    }
                    if((trim($worksheet->getCell('K' . $i)->getValue())!="COST CODE") && (trim($worksheet->getCell('K' . $i)->getValue())=="") && (trim($worksheet->getCell('J' . $i)->getValue())!="") && (trim($worksheet->getCell('S' . $i)->getCalculatedValue())>0.0)) {
                        $checkcostcodefill = 1;
                    }
                }
                if(trim($worksheet->getCell('K' . $i)->getValue())=="COST CODE") {
                    $startworking = 1;
                }
                if(trim($worksheet->getCell('A' . $i)->getValue())=="EMPLOYEE NAME:") {
                    $employee = str_replace(" ","_",str_replace("'","_",trim($worksheet->getCell('B' . $i)->getValue())));
                }
                if(strtolower(trim($worksheet->getCell('P' . $i)->getValue()))=="overtime this pay period") {
                    $overtimetotal = (float)(trim($worksheet->getCell('S' . $i)->getCalculatedValue()));
                }
            }
            $j = 2;
            $valueincrement = 2;
        } else {
            for($i=1;$i<=$worksheet->getHighestRow();$i++) {
                if($startworking == 1) {
                    if((trim($worksheet->getCell('K' . $i)->getValue())!="JOB CODE") && (trim($worksheet->getCell('K' . $i)->getValue())!="") && (trim($worksheet->getCell('J' . $i)->getValue())!="") && (trim($worksheet->getCell('L' . $i)->getValue())!="")) {
                        $temparray[explode(" - ",trim($worksheet->getCell('K' . $i)->getCalculatedValue()))[0] . " - " . explode(" - ",trim($worksheet->getCell('L' . $i)->getCalculatedValue()))[0]][trim($worksheet->getCell('J' . $i)->getValue())] += (float)trim($worksheet->getCell('M' . $i)->getValue()) + (float)trim($worksheet->getCell('N' . $i)->getValue()) + (float)trim($worksheet->getCell('O' . $i)->getValue()) + (float)trim($worksheet->getCell('P' . $i)->getValue()) + (float)trim($worksheet->getCell('Q' . $i)->getValue()) + (float)trim($worksheet->getCell('R' . $i)->getValue()) + (float)trim($worksheet->getCell('S' . $i)->getValue());
                    }
                }
                if(trim($worksheet->getCell('K' . $i)->getValue())=="JOB CODE") {
                    $startworking = 1;
                }
                if(trim($worksheet->getCell('A' . $i)->getValue())=="EMPLOYEE NAME:") {
                    $employee = str_replace(" ","_",str_replace("'","_",trim($worksheet->getCell('B' . $i)->getValue())));
                }
                if(strtolower(trim($worksheet->getCell('Q' . $i)->getValue()))=="overtime this pay period") {
                    $overtimetotal = (float)(trim($worksheet->getCell('T' . $i)->getCalculatedValue()));
                }
            }
            $j = 3;
            $valueincrement = 3;
        }
        $spreadsheet2 = new Spreadsheet();
        $sheet = $spreadsheet2->getActiveSheet();
        $i = 2;
        $temperarray = array();
        $tempindex = 0;
        $lastkey = sizeof($temparray)+1;
        $alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        $styleArray = [
            'font' => [
                'bold' => true,
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT,
            ],
            'borders' => [
                'top' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                ],
            ],
        ];
        $validatearray = array();
        foreach ($temparray as $key => $values) {
            if($costcodeopt==0) {
                $sheet->setCellValue('A' . $i, $key);
            } else {
                $tempindexarray = explode(" - ",$key);
                $sheet->setCellValue('A' . $i, $tempindexarray[0]);
                $sheet->setCellValue('B' . $i, $tempindexarray[1]);
            }
            foreach($values as $valuekey => $value) {
                if(!in_array($valuekey,$temperarray)) {
                    if(ContainsNumbers($valuekey)) {
                        array_push($validatearray,$j);
                    }
                    $sheet->setCellValueByColumnAndRow($j,1, $valuekey);
                    $sheet->setCellValueByColumnAndRow($j,$lastkey+1, "=SUM(" . (string)$alphabet[$j-1] . "2:" . (string)$alphabet[$j-1] . (string)$lastkey . ")");
                    $sheet->getStyleByColumnAndRow($j,$lastkey+1)->applyFromArray($styleArray);
                    $sheet->getColumnDimensionByColumn($j)->setAutoSize(true);
                    array_push($temperarray,$valuekey);
                    $j++;
                }
                $tempindex = array_search($valuekey, $temperarray) + $valueincrement;
                $sheet->setCellValueByColumnAndRow($tempindex,$i, $value);
            }
            $i++;
        }
        $totalhours = 0.00;
        foreach ($validatearray as $col) {
            $totalhours += (float)($sheet->getCellByColumnAndRow($col,$lastkey+1)->getCalculatedValue());
        }
        if(($totalhours!=$overtimetotal) || ($checkcostcodefill==1)) {
            if($totalhours!=$overtimetotal) {
                $sheet->setCellValueByColumnAndRow(1,$lastkey+3,"WARNING: OVERTIME TOTAL NOT MATCH!");
            }
            if($checkcostcodefill==1) {
                $sheet->setCellValueByColumnAndRow(1,$lastkey+3,"WARNING: PLEASE CHECK THE ALLOWANCE!");
            }
            $sheet->getStyleByColumnAndRow(1,$lastkey+3)->getFont()->setBold(true);
            $sheet->getStyleByColumnAndRow(1,$lastkey+3)->getFont()->setColor( new \PhpOffice\PhpSpreadsheet\Style\Color( \PhpOffice\PhpSpreadsheet\Style\Color::COLOR_RED ) );
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
