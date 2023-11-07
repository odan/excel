<?php

namespace App\Excel\Test;

use Odan\Excel\ExcelWriter;
use Odan\Excel\ZipFile;
use PHPUnit\Framework\TestCase;

final class ExcelWriterTest extends TestCase
{
    public function test(): void
    {
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

        // Write data
        foreach ($data as $rowData) {
            $excel->writeRow($rowData);
        }

        $excel->generate();

        $stream = $file->getStream();
        $data = stream_get_contents($stream);
        file_put_contents(__DIR__ . '/file.xlsx', $data);

        $this->assertTrue(true);
    }
}
