<?php

namespace Simple\Excel\Tests;

use Simple\Excel\Output;
/**
 * Writer.php
 *
 * Excel文件输出
 * @author lishuo <letwhip@gmail.com>
 * @version 0.0.1
 * @license MIT
 */

$data = [];
/**
 * 表头
 */
$data['title'] = [
    'A' => 'DATE',
    'B' => 'PV',
    'C' => 'UV',
    'D' => 'RATE'
];
/**
 * 数据
 */
$data['data'] = [
    0 => [
        '2016-04-05',
        '123,456',
        '123,456',
        '100%'
    ],
    1 => [
        '2016-04-06',
        '123,456',
        '123,456',
        '100%'
    ],
    2 => [
        '2016-04-07',
        '123,456',
        '123,456',
        '100%'
    ]
];

$sheetName = '工作簿';
$title = '报表输出结果';
$extName = 'xlsx';
$output = new Output();
$output->doExport($data, $sheetName, $title, $extName);
