<?php

namespace Odan\Excel;

use DOMDocument;
use DOMElement;
use DOMNode;

final class ExcelWriter
{
    private FileWriterInterface $file;
    private DOMDocument $sheetXml;
    private DOMNode $sheetData;
    private int $rowIndex = 0;
    private array $sharedStrings = [];
    private string $sheetName = 'Sheet1';

    public function __construct(FileWriterInterface $file)
    {
        $this->file = $file;
        $this->initSheetXml();
    }

    public function setSheetName(string $sheetName): void
    {
        $this->sheetName = $sheetName;
    }

    public function writeHead(array $values)
    {
        $row = $this->sheetData->appendChild($this->sheetXml->createElement('row'));
        $row->setAttribute('r', ++$this->rowIndex);

        foreach ($values as $colIndex => $value) {
            $column = $row->appendChild($this->sheetXml->createElement('c'));
            $column->setAttribute('r', $this->mapRowColumnToString($this->rowIndex, $colIndex + 1));
            // Apply the cell style by referencing it through the s attribute
            // 1 = bold style
            $column->setAttribute('s', '1');
            $column->setAttribute('t', 's');
            $sharedStringIndex = $this->createSharedStringIndex($value);
            $column->appendChild($this->sheetXml->createElement('v', $sharedStringIndex));
        }
    }

    public function writeRow(array $values): void
    {
        $row = $this->sheetData->appendChild($this->sheetXml->createElement('row'));
        $row->setAttribute('r', ++$this->rowIndex);

        foreach ($values as $colIndex => $value) {
            $column = $row->appendChild($this->sheetXml->createElement('c'));
            $column->setAttribute('r', $this->mapRowColumnToString($this->rowIndex, $colIndex + 1));
            // s = 0 = Normal font (see styles.xml)
            $column->setAttribute('s', '0');
            $column->setAttribute('t', 's');
            $sharedStringIndex = $this->createSharedStringIndex($value);
            $column->appendChild($this->sheetXml->createElement('v', $sharedStringIndex));
        }
    }

    private function mapRowColumnToString(int $row, int $column)
    {
        $columnLetter = '';

        while ($column > 0) {
            $remainder = ($column - 1) % 26;
            $columnLetter = chr(65 + $remainder) . $columnLetter;
            $column = floor(($column - $remainder) / 26);
        }

        // Combine the column letter(s) and row number to form the string
        return $columnLetter . $row;
    }

    public function generate(): void
    {
        $this->file->addFile('[Content_Types].xml', $this->createContentTypesXml());
        $this->file->addFile('_rels/.rels', $this->createRelsXml());
        $this->file->addFile('docProps/app.xml', $this->createDocPropsAppXml());
        $this->file->addFile('docProps/core.xml', $this->createDocPropsCoreXml());
        $this->file->addFile('xl/_rels/workbook.xml.rels', $this->createWorkbookRelsXml());
        $this->file->addFile('xl/styles.xml', $this->createStylesXml());
        $this->file->addFile('xl/workbook.xml', $this->createWorkbookXml());
        $this->file->addFile('xl/sharedStrings.xml', $this->createSharedStringsXml());
        $this->file->addFile('xl/worksheets/sheet1.xml', $this->createSheetXml());
    }

    private function createSharedStringsXml(): string
    {
        $dom = new DOMDocument('1.0', 'utf-8');
        $dom->formatOutput = true;
        $dom->xmlStandalone = true;

        $sst = $dom->appendChild($dom->createElement('sst'));
        $sst->setAttribute('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
        $sst->setAttribute('count', count($this->sharedStrings));
        $sst->setAttribute('uniqueCount', count($this->sharedStrings));

        foreach ($this->sharedStrings as $sharedString => $key) {
            $si = $sst->appendChild($dom->createElement('si'));
            $t = $si->appendChild($dom->createElement('t', $sharedString));
            $t->setAttribute('xml:space', 'preserve');
        }

        return $dom->saveXML();
    }

    private function createSheetXml(): string
    {
        $data = $this->sheetXml->saveXML();

        return $data;
    }

    private function createContentTypesXml(): string
    {
        $dom = new DOMDocument('1.0', 'utf-8');
        $dom->formatOutput = true;

        // Create the root element <Types> with the xmlns attribute
        $types = $dom->appendChild($dom->createElement('Types'));
        $types->setAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/content-types');

        $defaultAttributes = [
            ['Extension' => 'xml', 'ContentType' => 'application/xml'],
            ['Extension' => 'rels', 'ContentType' => 'application/vnd.openxmlformats-package.relationships+xml'],
            ['Extension' => 'png', 'ContentType' => 'image/png'],
            ['Extension' => 'jpeg', 'ContentType' => 'image/jpeg'],
        ];

        $this->createElements($dom, $types, 'Default', $defaultAttributes);

        $overrideAttributes = [
            [
                'PartName' => '/_rels/.rels',
                'ContentType' => 'application/vnd.openxmlformats-package.relationships+xml',
            ],
            [
                'PartName' => '/xl/workbook.xml',
                'ContentType' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
            ],
            [
                'PartName' => '/xl/styles.xml',
                'ContentType' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
            ],
            [
                'PartName' => '/xl/worksheets/sheet1.xml',
                'ContentType' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
            ],
            [
                'PartName' => '/xl/sharedStrings.xml',
                'ContentType' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
            ],
            [
                'PartName' => '/xl/_rels/workbook.xml.rels',
                'ContentType' => 'application/vnd.openxmlformats-package.relationships+xml',
            ],
            [
                'PartName' => '/docProps/core.xml',
                'ContentType' => 'application/vnd.openxmlformats-package.core-properties+xml',
            ],
            [
                'PartName' => '/docProps/app.xml',
                'ContentType' => 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
            ],
        ];

        $this->createElements($dom, $types, 'Override', $overrideAttributes);

        return $dom->saveXML();
    }

    private function createWorkbookXml(): string
    {
        $dom = new DOMDocument('1.0', 'utf-8');
        $dom->formatOutput = true;
        $dom->xmlStandalone = true;

        $workbook = $dom->appendChild($dom->createElement('workbook'));
        $workbook->setAttribute('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
        $workbook->setAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');

        // Create child elements and set their attributes
        $fileVersion = $dom->createElement('fileVersion');
        $fileVersion->setAttribute('appName', 'Calc');

        $workbookPr = $dom->createElement('workbookPr');
        $workbookPr->setAttribute('backupFile', 'false');
        $workbookPr->setAttribute('showObjects', 'all');
        $workbookPr->setAttribute('date1904', 'false');

        $workbookProtection = $dom->createElement('workbookProtection');

        $bookViews = $dom->createElement('bookViews');
        $workbookView = $dom->createElement('workbookView');
        $workbookView->setAttribute('showHorizontalScroll', 'true');
        $workbookView->setAttribute('showVerticalScroll', 'true');
        $workbookView->setAttribute('showSheetTabs', 'true');
        $workbookView->setAttribute('xWindow', '0');
        $workbookView->setAttribute('yWindow', '0');
        $workbookView->setAttribute('windowWidth', '16384');
        $workbookView->setAttribute('windowHeight', '8192');
        $workbookView->setAttribute('tabRatio', '500');
        $workbookView->setAttribute('firstSheet', '0');
        $workbookView->setAttribute('activeTab', '0');

        $sheets = $dom->createElement('sheets');
        $sheet = $dom->createElement('sheet');
        $sheet->setAttribute('name', $this->sheetName);
        $sheet->setAttribute('sheetId', '1');
        $sheet->setAttribute('state', 'visible');
        $sheet->setAttribute('r:id', 'rId2');

        $calcPr = $dom->createElement('calcPr');
        $calcPr->setAttribute('iterateCount', '100');
        $calcPr->setAttribute('refMode', 'A1');
        $calcPr->setAttribute('iterate', 'false');
        $calcPr->setAttribute('iterateDelta', '0.001');

        $extLst = $dom->createElement('extLst');
        $ext = $dom->createElement('ext');
        $ext->setAttribute('xmlns:loext', 'http://schemas.libreoffice.org/');
        $ext->setAttribute('uri', '{7626C862-2A13-11E5-B345-FEFF819CDC9F}');
        $loext = $dom->createElement('loext:extCalcPr');
        $loext->setAttribute('stringRefSyntax', 'CalcA1');

        // Append child elements to the <workbook> element in the desired order
        $workbook->appendChild($fileVersion);
        $workbook->appendChild($workbookPr);
        $workbook->appendChild($workbookProtection);
        $workbook->appendChild($bookViews);
        $bookViews->appendChild($workbookView);
        $workbook->appendChild($sheets);
        $sheets->appendChild($sheet);
        $workbook->appendChild($calcPr);
        $workbook->appendChild($extLst);
        $extLst->appendChild($ext);
        $ext->appendChild($loext);

        return $dom->saveXML();
    }

    private function createStylesXml(): string
    {
        $dom = new DOMDocument('1.0', 'utf-8');
        $dom->formatOutput = true;
        $dom->xmlStandalone = true;

        $styleSheet = $dom->appendChild($dom->createElement('styleSheet'));
        $styleSheet->setAttribute('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');

        // Create the <fonts> element with a count attribute of "2"
        $fonts = $dom->createElement('fonts');
        $fonts->setAttribute('count', '2');

        // Create the first <font> element for the default font (not bold)
        $font1 = $dom->createElement('font');
        $sz1 = $dom->createElement('sz');
        $sz1->setAttribute('val', '11');
        $name1 = $dom->createElement('name');
        $name1->setAttribute('val', 'Calibri');
        $family1 = $dom->createElement('family');
        $family1->setAttribute('val', '2');
        $b1 = $dom->createElement('b');
        $b1->setAttribute('val', 'false');
        $font1->appendChild($sz1);
        $font1->appendChild($name1);
        $font1->appendChild($family1);
        $font1->appendChild($b1);

        // Create the second <font> element for the bold font
        $font2 = $dom->createElement('font');
        $sz2 = $dom->createElement('sz');
        $sz2->setAttribute('val', '11');
        $name2 = $dom->createElement('name');
        $name2->setAttribute('val', 'Calibri');
        $family2 = $dom->createElement('family');
        $family2->setAttribute('val', '2');
        $b2 = $dom->createElement('b');
        $b2->setAttribute('val', 'true');
        $font2->appendChild($sz2);
        $font2->appendChild($name2);
        $font2->appendChild($family2);
        $font2->appendChild($b2);

        // Append the <font> elements to the <fonts> element
        $fonts->appendChild($font1);
        $fonts->appendChild($font2);

        // Create the <cellXfs> element with a count attribute of "2"
        $cellXfs = $dom->createElement('cellXfs');
        $cellXfs->setAttribute('count', '2');

        // Create the first <xf> element for the default font (fontId="0")
        $xf1 = $dom->createElement('xf');
        $xf1->setAttribute('numFmtId', '0');
        $xf1->setAttribute('fontId', '0');
        $xf1->setAttribute('fillId', '0');
        $xf1->setAttribute('borderId', '0');
        $xf1->setAttribute('applyFont', 'true'); // Apply default font

        // Create the second <xf> element for the bold font (fontId="1")
        $xf2 = $dom->createElement('xf');
        $xf2->setAttribute('numFmtId', '0');
        $xf2->setAttribute('fontId', '1');
        $xf2->setAttribute('fillId', '0');
        $xf2->setAttribute('borderId', '0');
        $xf2->setAttribute('applyFont', 'true'); // Apply bold font

        // Append the <xf> elements to the <cellXfs> element
        $cellXfs->appendChild($xf1);
        $cellXfs->appendChild($xf2);

        // Append the <fonts> and <cellXfs> elements to the <styleSheet> element
        $styleSheet->appendChild($fonts);
        $styleSheet->appendChild($cellXfs);

        return $dom->saveXML();
    }

    private function createWorkbookRelsXml(): string
    {
        $dom = new DOMDocument('1.0', 'utf-8');
        $dom->formatOutput = true;

        // Create the root element <Relationships> with the xmlns attribute
        $relationships = $dom->appendChild($dom->createElement('Relationships'));
        $relationships->setAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

        $relationshipData = [
            [
                'Id' => 'rId1',
                'Type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
                'Target' => 'styles.xml',
            ],
            [
                'Id' => 'rId2',
                'Type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
                'Target' => 'worksheets/sheet1.xml',
            ],
            [
                'Id' => 'rId3',
                'Type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
                'Target' => 'sharedStrings.xml',
            ],
        ];

        $this->createElements($dom, $relationships, 'Relationship', $relationshipData);

        return $dom->saveXML();
    }

    private function createRelsXml(): string
    {
        $dom = new DOMDocument('1.0', 'utf-8');
        $dom->formatOutput = true;

        $relationships = $dom->appendChild($dom->createElement('Relationships'));
        $relationships->setAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

        $relationshipItems = [
            [
                'Id' => 'rId1',
                'Type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
                'Target' => 'xl/workbook.xml',
            ],
            [
                'Id' => 'rId2',
                'Type' => 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
                'Target' => 'docProps/core.xml',
            ],
            [
                'Id' => 'rId3',
                'Type' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
                'Target' => 'docProps/app.xml',
            ],
        ];

        $this->createElements($dom, $relationships, 'Relationship', $relationshipItems);

        return $dom->saveXML();
    }

    private function createDocPropsAppXml(): string
    {
        $dom = new DOMDocument('1.0', 'utf-8');
        $dom->formatOutput = true;
        $dom->xmlStandalone = true;

        // Create the root element <Properties> with the required namespaces
        $properties = $dom->appendChild($dom->createElement('Properties'));
        $properties->setAttribute('xmlns', 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties');
        $properties->setAttribute('xmlns:vt', 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes');

        // Create child elements and set their text content
        $properties->appendChild($dom->createElement('Template'));
        $properties->appendChild($dom->createElement('TotalTime', '1'));
        $properties->appendChild(
            $dom->createElement(
                'Application',
                'LibreOffice/7.4.3.2$Windows_X86_64 LibreOffice_project/1048a8393ae2eeec98dff31b5c133c5f1d08b890'
            )
        );
        $properties->appendChild($dom->createElement('AppVersion', '15.0000'));

        return $dom->saveXML();
    }

    private function createDocPropsCoreXml(): string
    {
        $dom = new DOMDocument('1.0', 'utf-8');
        $dom->formatOutput = true;
        $dom->xmlStandalone = true;

        $coreProperties = $dom->createElement('cp:coreProperties');
        $coreProperties->setAttribute(
            'xmlns:cp',
            'http://schemas.openxmlformats.org/package/2006/metadata/core-properties'
        );
        $coreProperties->setAttribute('xmlns:dc', 'http://purl.org/dc/elements/1.1/');
        $coreProperties->setAttribute('xmlns:dcterms', 'http://purl.org/dc/terms/');
        $coreProperties->setAttribute('xmlns:dcmitype', 'http://purl.org/dc/dcmitype/');
        $coreProperties->setAttribute('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance');

        // Create child elements and set their attributes and text content
        $dctermsCreated = $dom->createElement('dcterms:created');
        $dctermsCreated->setAttribute('xsi:type', 'dcterms:W3CDTF');
        $dctermsCreated->nodeValue = '2023-11-04T22:53:36Z';

        $dcCreator = $dom->createElement('dc:creator');
        $dcDescription = $dom->createElement('dc:description');
        $dcLanguage = $dom->createElement('dc:language');
        $dcLanguage->nodeValue = 'de-DE';

        $cpLastModifiedBy = $dom->createElement('cp:lastModifiedBy');

        $dctermsModified = $dom->createElement('dcterms:modified');
        $dctermsModified->setAttribute('xsi:type', 'dcterms:W3CDTF');
        $dctermsModified->nodeValue = '2023-11-04T22:54:48Z';

        $cpRevision = $dom->createElement('cp:revision');
        $cpRevision->nodeValue = '1';

        $dcSubject = $dom->createElement('dc:subject');
        $dcTitle = $dom->createElement('dc:title');

        $coreProperties->appendChild($dctermsCreated);
        $coreProperties->appendChild($dcCreator);
        $coreProperties->appendChild($dcDescription);
        $coreProperties->appendChild($dcLanguage);
        $coreProperties->appendChild($cpLastModifiedBy);
        $coreProperties->appendChild($dctermsModified);
        $coreProperties->appendChild($cpRevision);
        $coreProperties->appendChild($dcSubject);
        $coreProperties->appendChild($dcTitle);

        $dom->appendChild($coreProperties);

        return $dom->saveXML();
    }

    private function initSheetXml(): void
    {
        // https://learn.microsoft.com/en-us/office/open-xml/working-with-sheets

        $this->sheetXml = new DOMDocument('1.0', 'utf-8');
        $this->sheetXml->formatOutput = true;

        $worksheet = $this->sheetXml->appendChild($this->sheetXml->createElement('worksheet'));
        $worksheet->setAttribute('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
        $worksheet->setAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
        $worksheet->setAttribute('xmlns:xdr', 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing');
        $worksheet->setAttribute('xmlns:x14', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main');
        $worksheet->setAttribute('xmlns:xr2', 'http://schemas.microsoft.com/office/spreadsheetml/2015/revision2');
        $worksheet->setAttribute('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006');

        // $sheetPr = $worksheet->appendChild($this->sheetXml->createElement('sheetPr'));
        // $sheetPr->setAttribute('filterMode', 'false');

        // $pageSetUpPr = $sheetPr->appendChild($this->sheetXml->createElement('pageSetUpPr'));
        // $pageSetUpPr->setAttribute('fitToPage', 'false');

        // $dimension = $sheetPr->appendChild($this->sheetXml->createElement('dimension'));
        // $dimension->setAttribute('ref', 'A1:C3');

        $this->sheetData = $worksheet->appendChild($this->sheetXml->createElement('sheetData'));
    }

    private function createSharedStringIndex(string $string): int
    {
        $index = $this->sharedStrings[$string] ?? null;
        if ($index !== null) {
            return $index;
        }

        $newIndex = count($this->sharedStrings);
        $this->sharedStrings[$string] = $newIndex;

        return $newIndex;
    }

    public function createElements(DOMDocument $dom, DOMElement $parentElement, $tagName, $items)
    {
        foreach ($items as $item) {
            $element = $this->createElementWithAttributes($dom, $tagName, $item);
            $parentElement->appendChild($element);
        }
    }

    public function createElementWithAttributes(DOMDocument $dom, string $tagName, array $attributes = []): DOMElement
    {
        $element = $dom->createElement($tagName);
        foreach ($attributes as $key => $value) {
            $element->setAttribute($key, $value);
        }

        return $element;
    }
}
