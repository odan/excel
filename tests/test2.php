<?php

use Odan\Excel\ExcelWriter;
use Odan\Excel\ZipFile;

require __DIR__ . '/../vendor/autoload.php';

echo (memory_get_usage(true) / 1024 / 1024) . "\n";
echo (memory_get_peak_usage(true) / 1024 / 1024) . "\n";

$file = new ZipFile();
$excel = new ExcelWriter($file);

$excel->setSheetName('My Sheet');

$head = ['Date', 'Name', 'Amount'];
$excel->writeHead($head);

$data = [
    ['2003-12-31', 'James', '220'],
    ['2003-8-23', 'Mike', '153.5'],
    ['2003-06-01', 'John', '34.12'],
];

for ($i = 0; $i < 100000; $i++) {
    $data[] = ['2003-12-31', 'James' . $i, $i];
}

// Write data
foreach ($data as $rowData) {
    $excel->writeRow($rowData);
}

$excel->generate();

$stream = $file->getStream();
rewind($stream);
$data = stream_get_contents($stream);
file_put_contents(__DIR__ . '/excel.xlsx', $data);

echo (memory_get_usage(true) / 1024 / 1024) . "\n";
echo (memory_get_peak_usage(true) / 1024 / 1024) . "\n";