<?php
/**
 * Created by PhpStorm.
 * User: fahmihilmansyah
 * Date: 25/10/18
 * Time: 15.00
 */

namespace app\mylib;
use \yii;

class parsingExcel
{
    
    /***
     * @param $pathToFile Directory of path excel
     * @param bool $allSheet option for get parsing allsheet
     * @param int $numberSheet sheet selected
     * @return array return with array data after parsing
     * @throws \PHPExcel_Exception
     * @throws \PHPExcel_Reader_Exception
     */
    static function parsingExcel($pathToFile,$allSheet=false,$numberSheet=0)
    {
        require_once(Yii::getAlias('@vendor/phpoffice/phpexcel/Classes/PHPExcel.php'));
        $ext = 'xlsx';
        if($ext == 'xls') {
            $objReader = \PHPExcel_IOFactory::createReader('Excel5'); // untuk tipe file .xls
        } else {
            $objReader = \PHPExcel_IOFactory::createReader('Excel2007'); // untuk tipe file .xlsx
        }
        $objPHPExcelUpdate = $objReader->load($pathToFile);
        $objWorksheetUpdate = $objPHPExcelUpdate->getActiveSheet();
        $dataSheet=array();
        foreach ($objPHPExcelUpdate->getWorksheetIterator() as $worksheet) {
            $sheetNumber = $objPHPExcelUpdate->getIndex($worksheet);
            if($allSheet == true){
                $dataHasil = self::prosesParsing($objWorksheetUpdate);
                $dataSheet['sheet_'.$sheetNumber]=$dataHasil;
            }else{
                if($numberSheet == $sheetNumber){
                    $dataSheet = self::prosesParsing($objWorksheetUpdate);
                }
            }
        }
        return $dataSheet;
    }

    private function prosesParsing($objWorksheet){
        $highestRow = $objWorksheet->getHighestRow(); // e.g. 10
        $highestColumn = $objWorksheet->getHighestColumn(); // e.g 'F'
        $dataHasil = array();
        for ($row = 1; $row <= $highestRow; ++$row) {
            $rowData = $objWorksheet->rangeToArray('A' . $row . ':' . $highestColumn . $row, NULL, TRUE, FALSE);
            $dataHasil[] = $rowData[0];
        }
        return $dataHasil;
    }
}