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
- Header columns with bold font.
- Custom worksheet name.

## Limitations

The purpose of this package is to provide a very fast and 
memory efficient XLSX file generator. It's designed for 
very fast data output, but not for fancy worksheet styles.
If you need more flexibility in terms of multiple 
sheets and colorful designs, you may use a 
different package, such as PhpSpreadsheet.

## Installation

```bash
composer require odan/excel
```

## Usage

```php
use Odan\Excel\ExcelWorkbook;
use Odan\Excel\ExcelFile;

$workbook = new ExcelWorkbook();
$sheet = $workbook->addSheet('My Sheet');

// Write header columns
$columns = ['Date', 'Name', 'Amount'];
$sheet->addColumns($columns);

// Write data
$rows = [
    ['2023-01-31', 'James', '220'],
    ['2023-03-28', 'Mike', '153.5'],
    ['2024-07-02', 'John', '34.12'],
];

foreach ($rows as $row) {
    $sheet->writeRow($row);
}

// Save as Excel file
$file = new ExcelFile();
$workbook->save($file);

// Save file in filesystem
$data = stream_get_contents($file->readStream());
file_put_contents(__DIR__ . '/filename.xlsx', $data);
```

## License

The MIT License (MIT). Please see [License File](LICENSE) for more information.
