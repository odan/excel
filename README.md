# odan/excel

Extreme fast in-memory Excel (XLSX) file writer.

## Requirements

* PHP 8.1+

## Features

- Optimized for minimal memory usage and high performance.
- Compatibility with Microsoft Excel 2007-365 (ISO/IEC 29500-1:2016).
- Compatibility with LibreOffice / OpenOffice Calc.
- In-memory operation by default.
- Optional hard disk access, when memory limitations are reached.
- Multiple sheets in a workbook.
- Header columns with bold font.
- Custom worksheet name.
- Data types for rows: string, int, float
- Data types for columns: string

## Limitations

The purpose of this package is to provide a very fast and
memory efficient Excel (XLSX) file generator. It is designed for
very fast data output, but not for fancy worksheet styles.
If you need more layout and color options, you may better use a
different package, such as PhpSpreadsheet.

* Number of workbooks: Limited by available memory and system resources.
* Sheets in a workbook: Limited by available memory (default is 1 sheet).
* Maximal number of columns: 16,384 (specification limit)
* Maximal number of rows: 1,048,576 (specification limit)
* Font styles: 2 (normal for rows and **bold** for columns)

## Installation

```bash
composer require odan/excel
```

## Usage

```php
use Odan\Excel\ExcelWorkbook;
use Odan\Excel\ZipDeflateStream;

$workbook = new ExcelWorkbook();
$sheet = $workbook->addSheet('My Sheet');

// Write header columns
$columns = ['Date', 'Name', 'Amount'];
$sheet->addColumns($columns);

// Write data
$rows = [
    ['2023-01-31', 'James', 220],
    ['2023-03-28', 'Mike', 153.5],
    ['2024-07-02', 'Sally', 34.12],
];

foreach ($rows as $row) {
    $sheet->addRow($row);
}

// Save as Excel file in memory
$file = new ZipDeflateStream();
$workbook->save($file);
```

**Generating only In-Memory Excel file**

This data is a pure in-memory stream `php://memory` (default)
that never overflows onto the hard disk, 
regardless of the amount of data written.

```php
use Odan\Excel\ZipDeflateStream;

// ...

$file = new ZipDeflateStream();
$workbook->save($file);
```

**Generating temporary files**

The `php://temp` stream is designed for temporary data storage in memory.

However, if the amount of data written exceeds a certain threshold 
(usually around 2KB or 8KB, depending on PHP versions and configurations), 
PHP may automatically switch to using temporary files on disk to store the data. 
This is done to conserve memory when dealing with large amounts of data.

This kind of stream is suitable for most scenarios where you need temporary 
in-memory storage, but it should automatically switch to using temporary files 
on disk to store the excess data when it overflows a certain threshold.

```php
use Odan\Excel\ZipDeflateStream;

// ...

$file = new ZipDeflateStream('php://temp');
$workbook->save($file);
```

The memory limit of `php://temp` can be controlled by appending `/maxmemory:NN`,
where NN is the maximum amount of data to keep in memory before using a temporary file, in bytes.

This optional parameter allows setting the memory limit before `php://temp` starts using a temporary file.

```php
use Odan\Excel\ZipDeflateStream;

// ...

// Set the limit to 5 MB.
$maxMb = 5 * 1024 * 1024;
$file = new ZipDeflateStream('php://temp/maxmemory:' . $maxMb);

$workbook->save($file);
```

**Save file in filesystem**

If the file does not exist, it will be created.
If it already exists, its content will be truncated (cleared)
when you write data to it.
Make sure the server has write permissions.

Directly as file stream...

```php
use Odan\Excel\ZipDeflateStream;

// ...

$file = new ZipDeflateStream('example.xlsx');
$workbook->save($file);
```

... or with stream_get_contents.

```php
use Odan\Excel\ZipDeflateStream;

// ...

$file = new ZipDeflateStream();
$workbook->save($file);

$data = stream_get_contents($file->getStream());
file_put_contents('filename.xlsx', $data);
```

**Generating Excel file on hard disk with write permissions**

```php
use Odan\Excel\ZipDeflateStream;

// ...

$filename = 'example.xlsx';

// Create an empty file using touch
touch($filename);

// Set write permissions to the file
chmod($filename, 0644);

$file = new ZipDeflateStream($filename);
$workbook->save($file);
```

**Reading the stream contents as string**

```php
use Odan\Excel\ZipDeflateStream;

// ...

$file = new ZipDeflateStream();
$workbook->save($file);

// Read contents of stream into a string
$data = stream_get_contents($file->getStream());
```

**Stream directly to the HTTP response**

To send an existing stream directly to the HTTP response,
you can use the `fpassthru` function. This function reads from
an open file pointer and sends the contents directly to the output buffer.

Here's an example of how to do this:

```php
<?php

// Set the content type to Excel
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment; filename="example.xlsx"');

// ...
$stream = $file->getStream();

// Send the stream contents directly to the HTTP response
fpassthru($stream);

// Close the stream
fclose($stream);
```

**Stream directly to the PSR-7 HTTP response**

To stream a file directly to an PSR-7 HTTP response using
the [Nyholm PSR-7](https://github.com/Nyholm/psr7) package, you may use it as follows:

```php
use Nyholm\Psr7\Response;
use Nyholm\Psr7\Stream;
use Odan\Excel\ZipDeflateStream;

// ...

$file = new ZipDeflateStream();
$workbook->save($file);

// Generate safe filename
$outputFilename = rawurlencode(basename('example.xlsx'));
$contentDisposition = sprintf("attachment; filename*=UTF-8''%s", $outputFilename);

// Add the response headers
$response = $response
    ->withHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    ->withHeader('Content-Disposition', $contentDisposition)
    ->withHeader('Pragma', 'private')
    ->withHeader('Cache-Control', 'private, must-revalidate')
    ->withHeader('Content-Transfer-Encoding', 'binary');

// Set the response body to the file stream
$response = $response->withBody(new Stream($file->getStream()));
```

Change the filename accordingly.

## Using the ZipStream-PHP package

When working with very large Excel files, typically over 4 GB in size, 
you can use the [ZipStream-PHP](https://github.com/maennchen/ZipStream-PHP) package to 
create Excel files in the ZIP64 format, which is designed for handling such large files.

**Installation**

```bash
composer require maennchen/zipstream-php
```

Next, use the `Odan\Excel\Zip64Stream` class for creating Excel 
files that offer improved compatibility and support larger file sizes.

```php
use Odan\Excel\ExcelWorkbook;
use Odan\Excel\Zip64Stream;

$workbook = new ExcelWorkbook();
$sheet = $workbook->addSheet('My Sheet');

// Write data
$rows = [
    ['2023-01-31', 'James', 220],
    ['2023-03-28', 'Mike', 153.5],
    ['2024-07-02', 'Sally', 34.12],
];

foreach ($rows as $row) {
    $sheet->addRow($row);
}

// Save as Excel file
$file = new Zip64Stream('filename.xlsx');
$workbook->save($file);
```

## License

The MIT License (MIT). Please see [License File](LICENSE) for more information.
