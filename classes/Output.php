<?php

namespace Simple\Excel;

/**
 * Output.php
 *
 * Excel文件输出
 * @author lishuo <letwhip@gmail.com>
 * @version 0.0.1
 * @license MIT
 */
class Output
{
    /**
     * extName
     *
     * 允许的扩展类型
     * @var enmu
     */
    private $extName = [];

    /**
     * title 
     * 
     * worksheetName
     * @var string
     * @access private
     */
    private $title = 'worksheet';

    public function __construct()
    {
        $this->extName = ['xls', 'xlsx'];
    }

    /**
     * setWorkSheet 
     * 
     * 设置工作簿名称
     * @access public
     * @return void
     */
    public function setTitle($name)
    {
        if (is_string($name)) {
            $this->title = $name;
        }
    }

    /**
     * doExport
     *
     * example
     * $title = '保表输出结果'
     * @ $data 表头数组
     * $data['title']['A'] = 'DATE';
     * $data['title']['B'] = 'PV';
     * $data['title']['C'] = 'UV;
     * $data['title']['D'] = 'RATE';
     *
     * @ $data 数据数组
     * $data['data'][0] = ['2016-04-05', '123,456' ,'123,456', '30%'];
     * $data['data'][1] = ['2016-04-06', '123,456' ,'123,456', '30%'];
     * $data['data'][2] = ['2016-04-06', '123,456' ,'123,456', '30%'];
     *
     * @ $data 新的工作簿
     * $data['newSheet'][1]['sheetName'] = '新的工作簿';
     * $data['newSheet'][1]['data'][0] = ['DATE', 'PV', 'UV', 'RATE'];
     * $data['newSheet'][1]['data'][1] = ['2016-04-06', '123,456', '123,456', '30%'];
     * $data['newSheet'][1]['filter'] = 'percent2';
     *
     * @$extName 扩展名
     * $extName = 'xls';
     *
     * @param string $title
     * @param array $data
     * @param string $extName
     * @access public
     * @return void
     */
    public function doExport($data, $sheetName = '工组薄', $extName = 'xlsx')
    {
        // 将输出文件格式锁定在正确格式中
        if (!in_array($extName, $this->extName)){
            $outputFileType = 'xls';
        } else {
            $outputFileType = $extName;
        }
        $data['title'] || $data['title']['A'] = '';
        $data['data'] || $data['data'][0] = (int) 0;

        $objExcel = new \PHPExcel();

        if ($outputFileType == 'xls') {
            $objWriter = new \PHPExcel_Writer_Excel5($objExcel);
        } else {
            $objWriter = new \PHPExcel_Writer_Excel2007($objExcel);
        }

        $objProps = $objExcel->getProperties();
        $objProps->setCreator("Excel");
        $objProps->setTitle("Office $outputFileType Document");
        $objProps->setSubject("Office $outputFileType Document");
        $objProps->setDescription("generated by simple-excel.");
        $objProps->setKeywords("office excel PHPExcel");
        $objExcel->setActiveSheetIndex(0);
        $objActSheet = $objExcel->getSheet();

        //设置worksheet标题
        $objActSheet->setTitle($sheetName);

        //当前仅允许24列
        $data['title'] = array_slice($data['title'], 0, 24);

        // 表头
        foreach ($data['title'] as $k => $v) {
            $objActSheet->setCellValue($k.'1', $v);
        }

        // 根据数据返回范围
        $rows = $this->getRange(count($data['title']));

        // 数据数组
        foreach ($data['data'] as $index => $value) {
            $index++;
            foreach ($value as $num => $res) {
                $objActSheet->setCellValue($rows[$num].$index, $res);
            }
        }

        // 存在更多sheet
        if ($data['newSheet']) {
            foreach ($data['newSheet'] as $key => $val) {
                $myWorkSheet = new \PHPExcel_Worksheet($objExcel, $val['sheetName']);
                $objExcel->addSheet($myWorkSheet, $key);
                $objExcel->setActiveSheetIndex($key);
                $objExcel->getActiveSheet()->fromArray($val['data'], $val['filter'], 'A1');
            }
        }

        $outputFileName = $this->title . '.' . $outputFileType;
        header("Content-Type: application/force-download");
        header("Content-Type: application/octet-stream");
        header("Content-Type: application/download");
        header('Content-Disposition:inline;filename="'. $outputFileName .'"');
        header("Content-Transfer-Encoding: binary");
        header("Last-Modified: " . gmdate("D, d M Y H:i:s") . " GMT");
        header("Cache-Control: must-revalidate, post-check=0, pre-check=0");
        header("Pragma: no-cache");
        $objWriter->save('php://output');
        $objExcel->disconnectWorksheets();
        unset($objWriter);
    }

    /**
     * getRange
     *
     * 获得范围
     * @param mixed $length
     * @access protected
     * @return array
     */
    protected function getRange($length)
    {
        $range = range('A', 'Z');
        $range = array_slice($range, 0, $length);
        return $range;
    }
}
