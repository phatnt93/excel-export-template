<?php 
namespace ExcelExportTemplate;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class EETemplate
{
    private $types = [
        '[[' => '2D',
        '[' => 'D',
        '{' => 'S'
    ];
    
    private $errors = [];
    private $paramsCells = [];

    function __construct()
    {
        
    }

    public function getTypes($in = null){
        if (isset($this->types[$in])) {
            return $this->types[$in];
        }
        return $this->types;
    }

    public function getErrors(){
        return $this->errors;
    }

    public function setError($err = ''){
        array_push($this->errors, $err);
    }

    public function setErrors($errs = []){
        $this->errors = array_merge($this->errors, $errs);
    }

    public function hasError(){
        if (count($this->getErrors()) > 0) {
            return true;
        }
        return false;
    }

    public function clean(){
        $this->errors = [];
    }
    

    // Run
    public function run($opt = []){
        try {
            $opt = array_merge([
                'params' => [],
                'template' => '',
                'target' => '',
                'output' => 'download'
            ], $opt);
            extract($opt);
            $spreadsheet = $this->loadSpreadsheetFromTemplate($template);
            $sheet = $this->getSheet($spreadsheet);
            // Loop each key of params to insert into file
            foreach ($params as $kp => $data) {
                // Find coordinate
                $cellCoor = $this->findKeyCellCoordinate($sheet, $kp);
                if (!empty($cellCoor)) {
                    // Add param cells
                    $this->paramsCells[$kp] = [
                        'cell' => $cellCoor
                    ];
                    // Detect type of data insert
                    $typeData = $this->detectType($kp);
                    // Insert data into template
                    $sheet = $this->insertData($sheet, $cellCoor, $kp, $data, $typeData);
                }
            }
            if ($output == 'file') {
                $this->outputFile($spreadsheet, $target);
            }else{
                $this->outputDownload($spreadsheet, $target);
            }
        } catch (\Exception $e) {
            $this->setError($e->getMessage());
        }
    }

    public function loadSpreadsheetFromTemplate($template){
        $inputFileType = IOFactory::identify($template);
        $reader = IOFactory::createReader($inputFileType);
        $spreadsheet = $reader->load($template);
        return $spreadsheet;
    }

    public function getSheet($spreadsheet, $byName = ''){
        if ($byName !== '') {
            return $spreadsheet->getSheetByName($byName);
        }
        return $spreadsheet->getActiveSheet();
    }

    public function findKeyCellCoordinate($sheet, $key = ''){
        $res = null;
        $highestRow = $sheet->getHighestRow();
        $highestCol = $sheet->getHighestColumn();
        for ($i=1; $i <= $highestRow; $i++) { 
            for ($j='A'; $j <= $highestCol; $j++) { 
                $cellCoordinate = $j . $i;
                $stringCheck = $sheet->getCell($cellCoordinate)->getValue();
                if (strpos($stringCheck, $key) > -1) {
                    $res = $cellCoordinate;
                    break;
                }
            }
            if ($res !== null) {
                break;
            }
        }
        return $res;
    }

    public function detectType($key){
        if (strpos($key, '[[') > -1) {
            return $this->getTypes('[[');
        }elseif (strpos($key, '[') > -1) {
            return $this->getTypes('[');
        }elseif (strpos($key, '{') > -1) {
            return $this->getTypes('{');
        }else{
            return null;
        }
    }

    public function insertData($sheet, $cellCoor, $find, $data, $type){
        switch ($type) {
            case '2D':{
                $sheet = $this->insertData2D($sheet, $cellCoor, $find, $data);
                break;
            }
            case 'D':{
                $sheet = $this->insertDataD($sheet, $cellCoor, $find, $data);
                break;
            }
            case 'S':{
                $sheet = $this->insertDataS($sheet, $cellCoor, $find, $data);
                break;
            }
            default:
                break;
        }
        return $sheet;
    }

    public function insertData2D($sheet, $cellCoor, $find, $data){
        $rowsAppend = count($data) - 1;
        $sheet->insertNewRowBefore($cellCoor[1] + 1, $rowsAppend);
        $sheet->fromArray($data, null, $cellCoor);
        return $sheet;
    }

    public function insertDataD($sheet, $cellCoor, $find, $data){
        $data = array_chunk($data, 1);
        $rowsAppend = count($data) - 1;
        $sheet->insertNewRowBefore($cellCoor[1] + 1, $rowsAppend);
        $sheet->fromArray($data, null, $cellCoor);
        return $sheet;
    }

    public function insertDataS($sheet, $cellCoor, $find, $data){
        $cellValue = $sheet->getCell($cellCoor)->getValue();
        $cellValue = str_replace($find, $data, $cellValue);
        $sheet->getCell($cellCoor)->setValue($cellValue);
        return $sheet;
    }
    
    public function outputFile($spreadsheet, $target){
        $writer = new Xlsx($spreadsheet);
        $writer->save($target);
    }

    public function outputDownload($spreadsheet, $target){
        // Export to download
        $fileName = $target;
        // Redirect output to a client’s web browser (Excel2007)
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="'. $fileName .'"');
        header('Cache-Control: max-age=0');
        // If you're serving to IE 9, then the following may be needed
        header('Cache-Control: max-age=1');

        // If you're serving to IE over SSL, then the following may be needed
        header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
        header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
        header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
        header ('Pragma: public'); // HTTP/1.0
        
        $writer = new Xlsx($spreadsheet);
        $writer->save('php://output');
        exit;
    }
}