<?php

namespace App\Excel\Test;

use DOMDocument;
use Odan\Excel\ExcelStream;
use Odan\Excel\ExcelWorkbook;
use PHPUnit\Framework\TestCase;
use RecursiveDirectoryIterator;
use RecursiveIteratorIterator;

final class ExcelWriterTest extends TestCase
{
    public function test(): void
    {
        $file = new ExcelStream();
        $workbook = new ExcelWorkbook($file);
        $sheet = $workbook->createSheet('My Sheet');

        $head = ['Date', 'Name', 'Amount'];
        $sheet->addHeader($head);

        $data = [
            ['2003-12-31', 'James', '220'],
            ['2003-8-23', 'Mike', '153.5'],
            ['2003-06-01', 'John', '34.12'],
        ];

        // Write data
        foreach ($data as $rowData) {
            $sheet->addRow($rowData);
        }

        $workbook->generate();

        $stream = $file->getStream();
        $data = stream_get_contents($stream);
        // file_put_contents(__DIR__ . '/file.xlsx', $data);

        $this->assertStringStartsWith('PK', $data);
        // $this->assertSame(3915, strlen($data));
    }

    public function formatXlsx(): void
    {
        $directoryToProcess = __DIR__ . '/file2';
        $this->processXmlFilesInDirectory($directoryToProcess);

        $this->assertEmpty([]);
    }

    // Function to format and save an XML file
    public function formatAndSaveXmlFile($xmlFilePath)
    {
        // Load the XML file
        $document = new DOMDocument();
        $document->preserveWhiteSpace = false;
        $document->formatOutput = true;

        $document->load($xmlFilePath);
        // Format the XML content
        $formattedXml = $document->saveXML();

        // Save the formatted XML back to the original file
        file_put_contents($xmlFilePath, $formattedXml);
    }

    // Function to process XML files recursively in a directory
    public function processXmlFilesInDirectory($directory)
    {
        $iterator = new RecursiveIteratorIterator(new RecursiveDirectoryIterator($directory));

        foreach ($iterator as $file) {
            if ($file->isFile()) {
                $xmlFilePath = $file->getPathname();
                $this->formatAndSaveXmlFile($xmlFilePath);
            }
        }
    }
}
