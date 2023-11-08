<?php

namespace App\Excel\Test;

use DOMDocument;
use Odan\Excel\ExcelFile;
use Odan\Excel\ExcelWorkbook;
use PHPUnit\Framework\TestCase;
use RecursiveDirectoryIterator;
use RecursiveIteratorIterator;

final class ExcelWriterTest extends TestCase
{
    public function test(): void
    {
        $workbook = new ExcelWorkbook();
        $sheet = $workbook->addSheet('My Sheet');

        // Header columns
        $columns = ['Date', 'Name', 'Amount'];
        $sheet->addColumns($columns);

        $data = [
            ['2003-12-31', 'James', '220'],
            ['2003-8-23', 'Mike', '153.5'],
            ['2003-06-01', 'John', '34.12'],
        ];

        // Write data
        foreach ($data as $rowData) {
            $sheet->addRow($rowData);
        }

        $sheet2 = $workbook->addSheet('My Sheet 2');

        // Header columns
        $columns = ['Date', 'Name', 'Amount', 'Fee'];
        $sheet2->addColumns($columns);

        $data = [
            ['2023-12-31', 'Max', '220', '0.1'],
            ['2023-8-23', 'John', '1234.5', '0.3'],
            ['2023-06-01', 'Daniel', '6789.12', '1.4'],
        ];

        // Write data
        foreach ($data as $rowData) {
            $sheet2->addRow($rowData);
        }

        $file = new ExcelFile();
        $workbook->save($file);

        $data = stream_get_contents($file->readStream());
        file_put_contents(__DIR__ . '/file.xlsx', $data);

        $this->assertStringStartsWith('PK', $data);
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
