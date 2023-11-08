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
    /** @var array<string, int> */
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

    public function writeHead(array $values): void
    {
        $row = $this->sheetXml->createElement('row');
        $this->sheetData->appendChild($row);
        $row->setAttribute('r', (string)++$this->rowIndex);

        // @todo dynamic spans
        $row->setAttribute('spans', '1:3');

        foreach ($values as $colIndex => $value) {
            $column = $this->sheetXml->createElement('c');
            $column->setAttribute('r', $this->mapRowColumnToString($this->rowIndex, $colIndex + 1));
            $row->appendChild($column);

            // Apply the cell style by referencing it through the s attribute
            // 1 = bold style
            $column->setAttribute('s', '1');
            $column->setAttribute('t', 's');
            $sharedStringIndex = $this->createSharedStringIndex($value);
            $valueElement = $this->sheetXml->createElement('v', (string)$sharedStringIndex);
            $column->appendChild($valueElement);
        }
    }

    public function writeRow(array $values): void
    {
        $row = $this->sheetXml->createElement('row');
        $this->sheetData->appendChild($row);
        $row->setAttribute('r', (string)++$this->rowIndex);
        // @todo dynamic spans
        $row->setAttribute('spans', '1:3');

        foreach ($values as $colIndex => $value) {
            $column = $this->sheetXml->createElement('c');
            $column->setAttribute('r', $this->mapRowColumnToString($this->rowIndex, $colIndex + 1));
            $row->appendChild($column);

            // s = 0 = Normal font (see styles.xml)
            $column->setAttribute('s', '0');
            $column->setAttribute('t', 's');
            $sharedStringIndex = $this->createSharedStringIndex($value);
            $valueNode = $this->sheetXml->createElement('v', (string)$sharedStringIndex);
            $column->appendChild($valueNode);
        }
    }

    private function mapRowColumnToString(int $row, int $column): string
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
        $dom = new DOMDocument('1.0', 'UTF-8');
        $dom->formatOutput = true;
        $dom->xmlStandalone = true;

        $sst = $dom->createElement('sst');
        $dom->appendChild($sst);
        $sst->setAttribute('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
        $sst->setAttribute('count', (string)count($this->sharedStrings));
        $sst->setAttribute('uniqueCount', (string)count($this->sharedStrings));

        foreach ($this->sharedStrings as $sharedString => $key) {
            $si = $dom->createElement('si');
            $sst->appendChild($si);
            $textNode = $dom->createElement('t', $sharedString);
            // $textNode->setAttribute('xml:space', 'preserve');
            $si->appendChild($textNode);
        }

        return (string)$dom->saveXML();
    }

    private function createSheetXml(): string
    {
        return (string)$this->sheetXml->saveXML();
    }

    private function createContentTypesXml(): string
    {
        $dom = new DOMDocument('1.0', 'UTF-8');
        $dom->formatOutput = true;
        $dom->xmlStandalone = true;

        // Create the root element <Types> with the xmlns attribute
        $types = $dom->createElement('Types');
        $dom->appendChild($types);
        $types->setAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/content-types');

        $defaultAttributes = [
            ['Extension' => 'rels', 'ContentType' => 'application/vnd.openxmlformats-package.relationships+xml'],
            ['Extension' => 'xml', 'ContentType' => 'application/xml'],
        ];

        $this->createElements($dom, $types, 'Default', $defaultAttributes);

        $overrideAttributes = [
            /* [
                 'PartName' => '/_rels/.rels',
                 'ContentType' => 'application/vnd.openxmlformats-package.relationships+xml',
             ],*/
            [
                'PartName' => '/xl/workbook.xml',
                'ContentType' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
            ],
            [
                'PartName' => '/docProps/core.xml',
                'ContentType' => 'application/vnd.openxmlformats-package.core-properties+xml',
            ],
            [
                'PartName' => '/docProps/app.xml',
                'ContentType' => 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
            ],
            [
                'PartName' => '/xl/worksheets/sheet1.xml',
                'ContentType' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
            ],
            [
                'PartName' => '/xl/styles.xml',
                'ContentType' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
            ],

            [
                'PartName' => '/xl/sharedStrings.xml',
                'ContentType' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
            ],
            /* [
                 'PartName' => '/xl/_rels/workbook.xml.rels',
                 'ContentType' => 'application/vnd.openxmlformats-package.relationships+xml',
             ],*/
        ];

        $this->createElements($dom, $types, 'Override', $overrideAttributes);

        return (string)$dom->saveXML();
    }

    private function createWorkbookXml(): string
    {
        $dom = new DOMDocument('1.0', 'UTF-8');
        $dom->formatOutput = true;
        $dom->xmlStandalone = true;

        $workbook = $dom->createElement('workbook');

        $dom->appendChild($workbook);
        $workbook->setAttribute('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
        $workbook->setAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');

        $workbook->setAttribute('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006');
        $workbook->setAttribute('xmlns:x15', 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main');
        $workbook->setAttribute('xmlns:xr', 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision');
        $workbook->setAttribute('xmlns:xr6', 'http://schemas.microsoft.com/office/spreadsheetml/2016/revision6');
        $workbook->setAttribute('xmlns:xr10', 'http://schemas.microsoft.com/office/spreadsheetml/2016/revision10');
        $workbook->setAttribute('xmlns:xr2', 'http://schemas.microsoft.com/office/spreadsheetml/2015/revision2');
        $workbook->setAttribute('mc:Ignorable', 'x15 xr xr6 xr10 xr2');

        // Create child elements and set their attributes
        $fileVersion = $dom->createElement('fileVersion');
        $workbook->appendChild($fileVersion);
        $fileVersion->setAttribute('appName', 'xl');
        $fileVersion->setAttribute('lastEdited', '7');
        $fileVersion->setAttribute('lowestEdited', '4');
        $fileVersion->setAttribute('rupBuild', '27031');

        $workbookPr = $dom->createElement('workbookPr');
        $workbook->appendChild($workbookPr);
        $workbookPr->setAttribute('defaultThemeVersion', '166925');

        $revisionPtr = $dom->createElement('xr:revisionPtr');
        $workbook->appendChild($revisionPtr);
        $revisionPtr->setAttribute('revIDLastSave', '0');
        $revisionPtr->setAttribute('documentId', '8_{D45FB324-B00D-43AB-BE0A-CC2F30BE489D}');
        $revisionPtr->setAttribute('xr6:coauthVersionLast', '47');
        $revisionPtr->setAttribute('xr6:coauthVersionMax', '47');
        $revisionPtr->setAttribute('xr10:uidLastSave', '{00000000-0000-0000-0000-000000000000}');

        $bookViews = $dom->createElement('bookViews');
        $workbook->appendChild($bookViews);

        $workbookView = $dom->createElement('workbookView');
        $bookViews->appendChild($workbookView);
        $workbookView->setAttribute('xWindow', '240');
        $workbookView->setAttribute('yWindow', '105');
        $workbookView->setAttribute('windowWidth', '14805');
        $workbookView->setAttribute('windowHeight', '8010');
        $workbookView->setAttribute('xr2:uid', '{00000000-000D-0000-FFFF-FFFF00000000}');

        $sheets = $dom->createElement('sheets');
        $workbook->appendChild($sheets);

        $sheet = $dom->createElement('sheet');
        $sheets->appendChild($sheet);
        $sheet->setAttribute('name', $this->sheetName);
        $sheet->setAttribute('sheetId', '1');
        $sheet->setAttribute('r:id', 'rId2');

        $calcPr = $dom->createElement('calcPr');
        $workbook->appendChild($calcPr);
        $calcPr->setAttribute('calcId', '191028');

        $extLst = $dom->createElement('extLst');
        $workbook->appendChild($extLst);

        $ext = $dom->createElement('ext');
        $extLst->appendChild($ext);

        $ext->setAttribute('xmlns:x15', 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main');
        $ext->setAttribute('uri', '{140A7094-0E35-4892-8432-C4D2E57EDEB5}');
        $loext = $dom->createElement('x15:workbookPr');
        $ext->appendChild($loext);
        $loext->setAttribute('chartTrackingRefBase', '1');

        return (string)$dom->saveXML();
    }

    private function createStylesXml(): string
    {
        $dom = new DOMDocument('1.0', 'UTF-8');
        $dom->formatOutput = true;
        $dom->xmlStandalone = true;

        $styleSheet = $dom->createElement('styleSheet');
        $dom->appendChild($styleSheet);
        $styleSheet->setAttribute('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
        $styleSheet->setAttribute('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006');
        $styleSheet->setAttribute('xmlns:x14ac', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac');
        $styleSheet->setAttribute('xmlns:x16r2', 'http://schemas.microsoft.com/office/spreadsheetml/2015/02/main');
        $styleSheet->setAttribute('xmlns:xr', 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision');
        $styleSheet->setAttribute('mc:Ignorable', 'x14ac x16r2 xr');

        // Create the <fonts> element with a count attribute of "2"
        $fonts = $dom->createElement('fonts');
        $styleSheet->appendChild($fonts);
        $fonts->setAttribute('count', '2');

        // Create the first <font> element for the default font (not bold)
        $font1 = $dom->createElement('font');
        $fonts->appendChild($font1);
        $sz1 = $dom->createElement('sz');
        $font1->appendChild($sz1);
        $sz1->setAttribute('val', '11');
        $name1 = $dom->createElement('name');
        $font1->appendChild($name1);
        $name1->setAttribute('val', 'Calibri');
        $family1 = $dom->createElement('family');
        $font1->appendChild($family1);
        $family1->setAttribute('val', '2');
        $font1->appendChild($family1);

        // Create the second <font> element for the bold font
        $font2 = $dom->createElement('font');
        $fonts->appendChild($font2);
        $b2 = $dom->createElement('b');
        $font2->appendChild($b2);

        $sz2 = $dom->createElement('sz');
        $font2->appendChild($sz2);
        $sz2->setAttribute('val', '11');
        $name2 = $dom->createElement('name');
        $font2->appendChild($name2);
        $name2->setAttribute('val', 'Calibri');
        $family2 = $dom->createElement('family');
        $font2->appendChild($family2);
        $family2->setAttribute('val', '2');

        // ----

        // Create the root element <fills>
        $fills = $dom->createElement('fills');
        $styleSheet->appendChild($fills);
        $fills->setAttribute('count', '2');

        // Create <fill> elements
        $fill1 = $dom->createElement('fill');
        $fills->appendChild($fill1);
        $patternFill1 = $dom->createElement('patternFill');
        $patternFill1->setAttribute('patternType', 'none');
        $fill1->appendChild($patternFill1);

        $fill2 = $dom->createElement('fill');
        $fills->appendChild($fill2);
        $patternFill2 = $dom->createElement('patternFill');
        $patternFill2->setAttribute('patternType', 'gray125');
        $fill2->appendChild($patternFill2);

        // Create the root element <borders>
        $borders = $dom->createElement('borders');
        $styleSheet->appendChild($borders);

        $borders->setAttribute('count', '1');

        // Create <border> element
        $border = $dom->createElement('border');
        $borders->appendChild($border);
        $borderElements = ['left', 'right', 'top', 'bottom', 'diagonal'];

        foreach ($borderElements as $element) {
            $borderElement = $dom->createElement($element);
            $border->appendChild($borderElement);
        }

        // Create the root element <cellStyleXfs>
        $cellStyleXfs = $dom->createElement('cellStyleXfs');
        $styleSheet->appendChild($cellStyleXfs);
        $cellStyleXfs->setAttribute('count', '1');

        // Create <xf> element
        $xf = $dom->createElement('xf');
        $cellStyleXfs->appendChild($xf);
        $xf->setAttribute('numFmtId', '0');
        $xf->setAttribute('fontId', '0');
        $xf->setAttribute('fillId', '0');
        $xf->setAttribute('borderId', '0');

        // Create the <cellXfs> element with a count attribute of "2"
        $cellXfs = $dom->createElement('cellXfs');
        $styleSheet->appendChild($cellXfs);
        $cellXfs->setAttribute('count', '2');

        // Create the first <xf> element for the default font (fontId="0")
        $xf1 = $dom->createElement('xf');
        $cellXfs->appendChild($xf1);
        $xf1->setAttribute('numFmtId', '0');
        $xf1->setAttribute('fontId', '0');
        $xf1->setAttribute('fillId', '0');
        $xf1->setAttribute('borderId', '0');
        $xf1->setAttribute('xfId', '0');

        // Create the second <xf> element for the bold font (fontId="1")
        $xf2 = $dom->createElement('xf');
        $cellXfs->appendChild($xf2);
        $xf2->setAttribute('numFmtId', '0');
        $xf2->setAttribute('fontId', '1');
        $xf2->setAttribute('fillId', '0');
        $xf2->setAttribute('borderId', '0');
        $xf2->setAttribute('xfId', '0');
        $xf2->setAttribute('applyFont', '1'); // Apply bold font

        // ---
        // Create the root element <cellStyles>
        $cellStyles = $dom->createElement('cellStyles');
        $styleSheet->appendChild($cellStyles);
        $cellStyles->setAttribute('count', '1');

        // Create <cellStyle> element
        $cellStyle = $dom->createElement('cellStyle');
        $cellStyles->appendChild($cellStyle);
        $cellStyle->setAttribute('name', 'Normal');
        $cellStyle->setAttribute('xfId', '0');
        $cellStyle->setAttribute('builtinId', '0');

        // Create <dxfs> element
        $dxfs = $dom->createElement('dxfs');
        $styleSheet->appendChild($dxfs);
        $dxfs->setAttribute('count', '0');

        // Create <tableStyles> element
        $tableStyles = $dom->createElement('tableStyles');
        $styleSheet->appendChild($tableStyles);
        $tableStyles->setAttribute('count', '0');
        $tableStyles->setAttribute('defaultTableStyle', 'TableStyleMedium2');
        $tableStyles->setAttribute('defaultPivotStyle', 'PivotStyleLight16');

        $extLst = $dom->createElement('extLst');
        $styleSheet->appendChild($extLst);

        $ext1 = $dom->createElement('ext');
        $extLst->appendChild($ext1);
        $ext1->setAttribute('xmlns:x14', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main');
        $ext1->setAttribute('uri', '{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}');

        $x14SlicerStyles = $dom->createElement('x14:slicerStyles');
        $ext1->appendChild($x14SlicerStyles);
        $x14SlicerStyles->setAttribute('defaultSlicerStyle', 'SlicerStyleLight1');

        $ext2 = $dom->createElement('ext');
        $ext2->setAttribute('xmlns:x15', 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main');
        $ext2->setAttribute('uri', '{9260A510-F301-46a8-8635-F512D64BE5F5}');

        $x15TimelineStyles = $dom->createElement('x15:timelineStyles');
        $ext2->appendChild($x15TimelineStyles);
        $x15TimelineStyles->setAttribute('defaultTimelineStyle', 'TimeSlicerStyleLight1');

        $extLst->appendChild($ext2);

        return (string)$dom->saveXML();
    }

    private function createWorkbookRelsXml(): string
    {
        $dom = new DOMDocument('1.0', 'UTF-8');
        $dom->formatOutput = true;
        $dom->xmlStandalone = true;

        // Create the root element <Relationships> with the xmlns attribute
        $relationships = $dom->createElement('Relationships');
        $dom->appendChild($relationships);
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

        return (string)$dom->saveXML();
    }

    private function createRelsXml(): string
    {
        $dom = new DOMDocument('1.0', 'UTF-8');
        $dom->formatOutput = true;
        $dom->xmlStandalone = true;

        $relationships = $dom->createElement('Relationships');
        $dom->appendChild($relationships);
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

        return (string)$dom->saveXML();
    }

    private function createDocPropsAppXml(): string
    {
        $dom = new DOMDocument('1.0', 'UTF-8');
        $dom->formatOutput = true;
        $dom->xmlStandalone = true;

        // Create the root element <Properties> with the required namespaces
        $properties = $dom->createElement('Properties');
        $dom->appendChild($properties);
        $properties->setAttribute('xmlns', 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties');
        $properties->setAttribute('xmlns:vt', 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes');

        // Create child elements and set their text content
        $properties->appendChild($dom->createElement('Application', 'Microsoft Excel Online'));
        $properties->appendChild($dom->createElement('Manager'));
        $properties->appendChild($dom->createElement('Company'));
        $properties->appendChild($dom->createElement('HyperlinkBase'));
        $properties->appendChild($dom->createElement('AppVersion', '16.0300'));

        return (string)$dom->saveXML();
    }

    private function createDocPropsCoreXml(): string
    {
        $dom = new DOMDocument('1.0', 'UTF-8');
        $dom->formatOutput = true;
        $dom->xmlStandalone = true;

        $coreProperties = $dom->createElement('cp:coreProperties');
        $dom->appendChild($coreProperties);
        $coreProperties->setAttribute(
            'xmlns:cp',
            'http://schemas.openxmlformats.org/package/2006/metadata/core-properties'
        );
        $coreProperties->setAttribute('xmlns:dc', 'http://purl.org/dc/elements/1.1/');
        $coreProperties->setAttribute('xmlns:dcterms', 'http://purl.org/dc/terms/');
        $coreProperties->setAttribute('xmlns:dcmitype', 'http://purl.org/dc/dcmitype/');
        $coreProperties->setAttribute('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance');

        $coreProperties->appendChild($dom->createElement('dc:title'));
        $coreProperties->appendChild($dom->createElement('dc:subject'));
        $coreProperties->appendChild($dom->createElement('dc:creator'));
        $coreProperties->appendChild($dom->createElement('cp:keywords'));
        $coreProperties->appendChild($dom->createElement('dc:description'));
        $coreProperties->appendChild($dom->createElement('cp:lastModifiedBy'));
        $coreProperties->appendChild($dom->createElement('cp:revision'));

        // Create child elements and set their attributes and text content
        $dctermsCreated = $dom->createElement('dcterms:created', date('Y-m-d\TH:i:s\Z'));
        $coreProperties->appendChild($dctermsCreated);
        $dctermsCreated->setAttribute('xsi:type', 'dcterms:W3CDTF');

        $dctermsModified = $dom->createElement('dcterms:modified', date('Y-m-d\TH:i:s\Z'));
        $coreProperties->appendChild($dctermsModified);
        $dctermsModified->setAttribute('xsi:type', 'dcterms:W3CDTF');

        $coreProperties->appendChild($dom->createElement('cp:category'));
        $coreProperties->appendChild($dom->createElement('cp:contentStatus'));

        return (string)$dom->saveXML();
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

        $dimension = $this->sheetXml->createElement('dimension');
        $worksheet->appendChild($dimension);
        // @todo make dynamic
        $dimension->setAttribute('ref', 'A1:C4');

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

    public function createElements(DOMDocument $dom, DOMElement $parentElement, string $tagName, array $items): void
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
