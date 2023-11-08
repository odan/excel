# odan/excel

Extreme fast in-memory Excel (XLSX) file writer.

## Requirements

* PHP 8.1+

## Features

- Extreme performance and minimal memory usage.
- No hard disk access required.
- Header columns with bold font.
- Custom sheet name.

## Limitations

The purpose of this package is to provide a very fast and 
memory efficient XLSX file generator. It is designed for 
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
use Odan\Excel\ExcelWriter;
use Odan\Excel\ZipFile;

$file = new ZipFile();
$excel = new ExcelWriter($file);

// Change sheet name
$excel->setSheetName('My Sheet');

// Write headers
$head = ['Date', 'Name', 'Amount'];
$excel->writeHead($head);

// Write data
$rows = [
    ['2023-01-31', 'James', '220'],
    ['2023-03-28', 'Mike', '153.5'],
    ['2024-07-02', 'John', '34.12'],
];

foreach ($rows as $row) {
    $excel->writeRow($row);
}

// Generate Excel file
$excel->generate();

// Save as Excel file
$data = stream_get_contents($file->getStream());
file_put_contents(__DIR__ . '/excel.xlsx', $data);
```

## License

The MIT License (MIT). Please see [License File](LICENSE) for more information.
