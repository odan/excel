<?php

namespace Odan\Excel;

use DOMDocument;
use DOMElement;
use DOMNode;

final class ExcelWorksheet
{
    private DOMDocument $sheetXml;

    private DOMNode $sheetData;

    private int $rowIndex = 0;

    private string $sheetName = 'Sheet1';

    private SharedStrings $sharedStrings;

    private DOMElement $dimension;
    private int $boundRow = 1;
    private int $boundColumn = 1;

    public function __construct(SharedStrings $sharedStrings)
    {
        $this->sharedStrings = $sharedStrings;
        $this->initSheetXml();
    }

    public function setSheetName(string $sheetName): void
    {
        $this->sheetName = $sheetName;
    }

    public function addColumns(array $values): void
    {
        $row = $this->sheetXml->createElement('row');
        $this->sheetData->appendChild($row);
        $row->setAttribute('r', (string)++$this->rowIndex);

        foreach ($values as $colIndex => $value) {
            $column = $this->sheetXml->createElement('c');
            $column->setAttribute('r', $this->getCoordinate($this->rowIndex, $colIndex + 1));
            $row->appendChild($column);

            // Apply the cell style by referencing it through the s attribute
            // 1 = bold style
            $column->setAttribute('s', '1');
            $column->setAttribute('t', 's');
            $sharedStringIndex = $this->sharedStrings->add($value);
            $valueElement = $this->sheetXml->createElement('v', (string)$sharedStringIndex);
            $column->appendChild($valueElement);
        }
    }

    public function addRow(array $values): void
    {
        $row = $this->sheetXml->createElement('row');
        $this->sheetData->appendChild($row);
        $row->setAttribute('r', (string)++$this->rowIndex);

        foreach ($values as $colIndex => $value) {
            $column = $this->sheetXml->createElement('c');
            $column->setAttribute('r', $this->getCoordinate($this->rowIndex, $colIndex + 1));
            $row->appendChild($column);

            // s = 0 = Normal font (see styles.xml)
            $column->setAttribute('s', '0');
            $column->setAttribute('t', 's');
            $sharedStringIndex = $this->sharedStrings->add($value);
            $valueNode = $this->sheetXml->createElement('v', (string)$sharedStringIndex);
            $column->appendChild($valueNode);
        }
    }

    private function getCoordinate(int $row, int $column): string
    {
        $columnLetter = '';

        // Maximum limit
        $column = min($column, 16384);

        $this->boundRow = max($this->boundRow, $row);
        $this->boundColumn = max($this->boundColumn, $column);

        while ($column > 0) {
            $remainder = ($column - 1) % 26;
            $columnLetter = chr(65 + $remainder) . $columnLetter;
            $column = floor(($column - $remainder) / 26);
        }

        // Combine the column letter(s) and row number to form the string
        return $columnLetter . $row;
    }

    public function createSheetXml(): string
    {
        // The row and column bounds of all cells in this worksheet
        $bound = $this->getCoordinate($this->boundRow, $this->boundColumn);
        $this->dimension->setAttribute('ref', 'A1:' . $bound);

        return (string)$this->sheetXml->saveXML();
    }

    private function initSheetXml(): void
    {
        // https://learn.microsoft.com/en-us/office/open-xml/working-with-sheets

        $this->sheetXml = new DOMDocument('1.0', 'UTF-8');
        $this->sheetXml->formatOutput = true;
        $this->sheetXml->xmlStandalone = true;

        $worksheet = $this->sheetXml->createElement('worksheet');
        $this->sheetXml->appendChild($worksheet);
        $worksheet->setAttribute('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
        $worksheet->setAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
        $worksheet->setAttribute('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006');
        $worksheet->setAttribute('xmlns:x14ac', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac');
        $worksheet->setAttribute('xmlns:xr', 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision');
        $worksheet->setAttribute('xmlns:xr2', 'http://schemas.microsoft.com/office/spreadsheetml/2015/revision2');
        $worksheet->setAttribute('xmlns:xr3', 'http://schemas.microsoft.com/office/spreadsheetml/2016/revision3');
        $worksheet->setAttribute('mc:Ignorable', 'x14ac xr xr2 xr3');
        $worksheet->setAttribute('xr:uid', '{00000000-0001-0000-0000-000000000000}');

        $this->dimension = $this->sheetXml->createElement('dimension');
        $worksheet->appendChild($this->dimension);

        // The row and column bounds of all cells in this worksheet
        $this->dimension->setAttribute('ref', 'A1:A1');

        $sheetViews = $this->sheetXml->createElement('sheetViews');
        $worksheet->appendChild($sheetViews);

        $sheetView = $this->sheetXml->createElement('sheetView');
        $sheetViews->appendChild($sheetView);
        $sheetView->setAttribute('tabSelected', '1');
        $sheetView->setAttribute('workbookViewId', '0');

        $sheetFormatPr = $this->sheetXml->createElement('sheetFormatPr');
        $worksheet->appendChild($sheetFormatPr);
        $sheetFormatPr->setAttribute('defaultRowHeight', '15');

        $this->sheetData = $this->sheetXml->createElement('sheetData');
        $worksheet->appendChild($this->sheetData);

        $pageMargins = $this->sheetXml->createElement('pageMargins');
        $worksheet->appendChild($pageMargins);
        $pageMargins->setAttribute('left', '0.7');
        $pageMargins->setAttribute('right', '0.7');
        $pageMargins->setAttribute('top', '0.75');
        $pageMargins->setAttribute('bottom', '0.75');
        $pageMargins->setAttribute('header', '0.3');
        $pageMargins->setAttribute('footer', '0.3');
    }

    public function getName(): string
    {
        return $this->sheetName;
    }
}
