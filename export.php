<?php 
require "./../composer/vendor/autoload.php";
require "./ExcelExportTemplate.php";

use ExcelExportTemplate\EETemplate;

$data = [
    [1, 'data1', 21],
    [2, 'data2', 22],
    [3, 'data3', 23],
    [4, 'data4', 24]
];
$dataTen = ['data1', 'data2', 'data3', 'data4'];
$target = './ExcelExportTemplate' . time() . '.xlsx';

$eet = new EETemplate();
$eet->run([
    'params' => [
        '[[table]]' => $data,
        '{ngay}' => date('d/m/Y'),
        '{tong_tuoi}' => 90,
        '[ds_ten]' => $dataTen
    ],
    'template' => './ExcelExportTemplate.xlsx',
    'target' => $target,
    'output' => 'download'
]);

echo "OK";