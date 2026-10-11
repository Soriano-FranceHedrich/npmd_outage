<?php
declare(strict_types=1);
session_start();
require_once __DIR__ . '/db.php';

$pdo = db();
$action = $_GET['action'] ?? 'dashboard';

function column_letter_to_index(string $column): int {
    $index = 0;
    $letters = str_split(strtoupper($column));
    foreach ($letters as $letter) {
        $index = $index * 26 + (ord($letter) - 64);
    }
    return $index - 1;
}

function outage_region_color(string $region): string {
    $region = strtoupper(trim($region));
    $colors = [
        'REGION 10' => '164D78',
        'REGION 4B' => '4D8A68',
        'REGION 7' => 'D26A3A',
        'REGION 4A' => '8B5E83',
        'REGION 13' => 'C49A3A',
        'NCR' => '4F7CAC',
        'REGION 5' => '2A9D8F',
        'REGION 6' => 'AA4B55',
        'REGION 1' => '5B8DB8',
        'REGION 2' => '8A9A3A',
        'REGION 3' => 'E09F3E',
        'REGION 8' => '2F7F82',
        'REGION 9' => '9C6644',
        'REGION 11' => '6D597A',
        'REGION 12' => '588157',
        'CAR' => '6C757D',
        'CARAGA' => 'C06C84',
        'BARMM' => '7F5539',
    ];
    if (isset($colors[$region])) {
        return $colors[$region];
    }

    $palette = array_values($colors);
    return $palette[abs(crc32($region)) % count($palette)];
}

function normalize_import_header(string $name): string {
    $name = preg_replace('/^\xEF\xBB\xBF/', '', trim($name)) ?? trim($name);
    $name = strtolower($name);
    $name = preg_replace('/[^a-z0-9]+/', '_', $name) ?? $name;
    $name = trim($name, '_');
    $aliases = [
        'job_order_number' => 'job_order_no', 'job_order' => 'job_order_no', 'job_order_no_' => 'job_order_no',
        'jo_no' => 'job_order_no', 'jo_number' => 'job_order_no', 'jo' => 'job_order_no',
        'project_office_name' => 'project_name', 'project_office' => 'project_name', 'office_name' => 'project_name',
        'site_name' => 'project_name', 'assigned_team_name' => 'assigned_team', 'team_name' => 'assigned_team',
        'remarks_notes' => 'remarks', 'comments' => 'remarks', 'comment' => 'remarks',
    ];
    return $aliases[$name] ?? $name;
}

function format_import_matrix(array $rows): array {
    $rows = array_values(array_filter($rows, static function ($row): bool {
        foreach ($row as $value) if (trim((string)$value) !== '') return true;
        return false;
    }));
    if (!$rows) return [];

    $knownHeaders = ['job_order_no', 'project_name', 'province', 'region', 'location', 'assigned_team', 'team', 'remarks', 'down_time', 'up_time', 'status', 'priority', 'root_cause', 'project_id'];
    $headerIndex = 0;
    $bestScore = 0;
    foreach (array_slice($rows, 0, 15, true) as $rowIndex => $candidate) {
        $headers = array_map(static fn($value) => normalize_import_header((string)$value), $candidate);
        $score = count(array_intersect($knownHeaders, $headers));
        if ($score > $bestScore) {
            $bestScore = $score;
            $headerIndex = $rowIndex;
        }
    }

    $header = array_map(static fn($value) => normalize_import_header((string)$value), $rows[$headerIndex]);
    $formatted = [];
    for ($i = $headerIndex + 1; $i < count($rows); $i++) {
        $record = [];
        foreach ($header as $index => $key) {
            if ($key !== '') $record[$key] = trim((string)($rows[$i][$index] ?? ''));
        }
        if (array_filter($record, static fn($value) => $value !== '')) $formatted[] = $record;
    }
    return $formatted;
}

function normalize_job_order_no(string $value): string {
    $value = strtoupper(trim(preg_replace('/\s+/u', ' ', $value) ?? $value));
    // Excel can expose an integer identifier as a decimal with a zero-only fraction.
    if (preg_match('/^([0-9]+)\.0+$/', $value, $matches)) {
        $value = $matches[1];
    }
    return $value;
}

function normalize_project_name(string $value): string {
    $value = trim(preg_replace('/\s+/u', ' ', $value) ?? $value);
    return function_exists('mb_strtolower') ? mb_strtolower($value, 'UTF-8') : strtolower($value);
}

function read_xlsx_zip_entry(string $archivePath, string $entryName): ?string {
    $archive = file_get_contents($archivePath);
    if ($archive === false) return null;

    $endOffset = strrpos($archive, "PK\x05\x06");
    if ($endOffset === false || strlen($archive) < $endOffset + 22) return null;
    $end = unpack('vdisk/vstart_disk/ventries_disk/ventries/Vdirectory_size/Vdirectory_offset/vcomment_length', substr($archive, $endOffset + 4, 18));
    if (!$end || $end['directory_offset'] + $end['directory_size'] > strlen($archive)) return null;

    $offset = (int)$end['directory_offset'];
    $limit = $offset + (int)$end['directory_size'];
    while ($offset + 46 <= $limit && substr($archive, $offset, 4) === "PK\x01\x02") {
        $header = unpack('vmade/vneeded/vflags/vmethod/vtime/vdate/Vcrc/Vcompressed/Vuncompressed/vname_length/vextra_length/vcomment_length/vdisk/vinternal/Vexternal/Vlocal_offset', substr($archive, $offset + 4, 42));
        if (!$header) return null;
        $nameStart = $offset + 46;
        $name = substr($archive, $nameStart, $header['name_length']);
        if ($name === $entryName) {
            if (($header['flags'] & 1) !== 0) return null;
            $localOffset = (int)$header['local_offset'];
            if (substr($archive, $localOffset, 4) !== "PK\x03\x04") return null;
            $local = unpack('vneeded/vflags/vmethod/vtime/vdate/Vcrc/Vcompressed/Vuncompressed/vname_length/vextra_length', substr($archive, $localOffset + 4, 26));
            if (!$local) return null;
            $dataOffset = $localOffset + 30 + $local['name_length'] + $local['extra_length'];
            $compressedData = substr($archive, $dataOffset, $header['compressed']);
            if ($header['method'] === 0) return $compressedData;
            if ($header['method'] === 8) {
                $data = @gzinflate($compressedData);
                return $data === false ? null : $data;
            }
            return null;
        }
        $offset = $nameStart + $header['name_length'] + $header['extra_length'] + $header['comment_length'];
    }
    return null;
}

function parse_uploaded_csv(string $tmpName): array {
    $handle = fopen($tmpName, 'rb');
    if ($handle === false) {
        return [];
    }

    $rows = [];
    $separators = [',', ';', "\t"];
    $firstLine = fgets($handle, 8192);
    if ($firstLine === false) {
        fclose($handle);
        return [];
    }
    $sample = trim($firstLine);
    $separator = ',';
    foreach ($separators as $candidate) {
        if (substr_count($sample, $candidate) >= 1) {
            $separator = $candidate;
            break;
        }
    }
    rewind($handle);

    while (($data = fgetcsv($handle, 10000, $separator)) !== false) {
        if ($data === [null] || $data === []) {
            continue;
        }
        $clean = array_map(static function ($value) {
            return is_string($value) ? trim($value) : trim((string)$value);
        }, $data);
        if (!array_filter($clean, static fn($value) => $value !== '')) {
            continue;
        }
        $rows[] = $clean;
    }

    fclose($handle);
    return format_import_matrix($rows);
}

function parse_uploaded_xlsx(string $tmpName): array {
    $zip = null;
    if (class_exists('ZipArchive')) {
        $zip = new ZipArchive();
        if ($zip->open($tmpName) !== true) return [];
        $getEntry = static fn(string $name) => $zip->getFromName($name);
        $hasEntry = static fn(string $name): bool => $zip->locateName($name) !== false;
    } else {
        $getEntry = static fn(string $name) => read_xlsx_zip_entry($tmpName, $name);
        $hasEntry = static fn(string $name): bool => read_xlsx_zip_entry($tmpName, $name) !== null;
    }

    $sharedStrings = [];
    $sharedStringsXml = $getEntry('xl/sharedStrings.xml');
    if (is_string($sharedStringsXml)) {
        $sharedDom = new DOMDocument();
        libxml_use_internal_errors(true);
        $sharedDom->loadXML($sharedStringsXml);
        libxml_clear_errors();
        foreach ($sharedDom->getElementsByTagName('si') as $si) {
            $parts = [];
            foreach ($si->getElementsByTagName('t') as $textNode) {
                $parts[] = $textNode->textContent;
            }
            $sharedStrings[] = implode('', $parts);
        }
    }

    $sheetPath = 'xl/worksheets/sheet1.xml';
    $workbookXml = $getEntry('xl/workbook.xml');
    $relsXml = $getEntry('xl/_rels/workbook.xml.rels');
    if (is_string($workbookXml) && is_string($relsXml)) {
        $workbookDom = new DOMDocument();
        $relsDom = new DOMDocument();
        libxml_use_internal_errors(true);
        $workbookLoaded = $workbookDom->loadXML($workbookXml);
        $relsLoaded = $relsDom->loadXML($relsXml);
        libxml_clear_errors();
        if ($workbookLoaded && $relsLoaded) {
            $sheet = $workbookDom->getElementsByTagName('sheet')->item(0);
            if ($sheet && $sheet->hasAttribute('r:id')) {
                $relationshipId = $sheet->getAttribute('r:id');
                foreach ($relsDom->getElementsByTagName('Relationship') as $relationship) {
                    if ($relationship->getAttribute('Id') !== $relationshipId) continue;
                    $target = str_replace('\\', '/', $relationship->getAttribute('Target'));
                    $target = str_starts_with($target, '/') ? ltrim($target, '/') : 'xl/' . $target;
                    $parts = [];
                    foreach (explode('/', $target) as $part) {
                        if ($part === '' || $part === '.') continue;
                        if ($part === '..') array_pop($parts);
                        else $parts[] = $part;
                    }
                    $resolvedPath = implode('/', $parts);
                    if ($hasEntry($resolvedPath)) $sheetPath = $resolvedPath;
                    break;
                }
            }
        }
    }

    $sheetXml = $getEntry($sheetPath);
    if ($zip instanceof ZipArchive) $zip->close();
    if (!is_string($sheetXml)) {
        return [];
    }

    $dom = new DOMDocument();
    libxml_use_internal_errors(true);
    $dom->loadXML($sheetXml);
    libxml_clear_errors();

    $rows = [];
    foreach ($dom->getElementsByTagName('row') as $rowNode) {
        $values = [];
        foreach ($rowNode->getElementsByTagName('c') as $cellNode) {
            $cellRef = $cellNode->getAttribute('r');
            preg_match('/([A-Z]+)(\d+)/i', $cellRef, $matches);
            $columnIndex = $matches ? column_letter_to_index($matches[1]) : 0;
            while (count($values) <= $columnIndex) {
                $values[] = '';
            }
            $cellType = $cellNode->getAttribute('t');
            $value = '';
            if ($cellType === 's') {
                $vNode = $cellNode->getElementsByTagName('v')->item(0);
                if ($vNode) {
                    $index = (int)$vNode->nodeValue;
                    $value = $sharedStrings[$index] ?? '';
                }
            } elseif ($cellType === 'inlineStr') {
                $isNode = $cellNode->getElementsByTagName('is')->item(0);
                if ($isNode) {
                    $value = $isNode->textContent;
                }
            } else {
                $vNode = $cellNode->getElementsByTagName('v')->item(0);
                if ($vNode) {
                    $value = $vNode->nodeValue;
                }
            }
            $values[$columnIndex] = (string)$value;
        }
        $rows[] = $values;
    }

    return format_import_matrix($rows);
}

function parse_uploaded_xls(string $tmpName): array {
    $contents = file_get_contents($tmpName);
    if ($contents === false || trim($contents) === '') {
        return [];
    }

    $dom = new DOMDocument();
    libxml_use_internal_errors(true);
    $loaded = $dom->loadHTML($contents);
    libxml_clear_errors();
    if (!$loaded) {
        return [];
    }

    $table = $dom->getElementsByTagName('table')->item(0);
    if (!$table) {
        return [];
    }

    $rows = [];
    foreach ($table->getElementsByTagName('tr') as $rowNode) {
        $values = [];
        foreach ($rowNode->childNodes as $cellNode) {
            if (!in_array(strtolower($cellNode->nodeName), ['th', 'td'], true)) {
                continue;
            }
            $values[] = trim(preg_replace('/\s+/', ' ', $cellNode->textContent) ?? '');
        }
        if ($values) {
            if (array_filter($values, static fn($value) => $value !== '')) {
                $rows[] = $values;
            }
        }
    }

    return format_import_matrix($rows);
}

function parse_uploaded_file(string $tmpName, string $originalName): array {
    $ext = strtolower(pathinfo($originalName, PATHINFO_EXTENSION));
    if ($ext === 'csv') {
        return parse_uploaded_csv($tmpName);
    }
    if ($ext === 'xlsx') {
        return parse_uploaded_xlsx($tmpName);
    }
    if ($ext === 'xls') {
        return parse_uploaded_xls($tmpName);
    }
    return [];
}

function normalize_import_datetime(?string $value): ?string {
    $value = trim((string)($value ?? ''));
    if ($value === '') {
        return null;
    }
    if (is_numeric($value) && (float)$value >= 1 && (float)$value < 300000) {
        $timestamp = strtotime('1899-12-30 00:00:00 UTC') + (int)round((float)$value * 86400);
        return gmdate('Y-m-d H:i:s', $timestamp);
    }
    $timestamp = strtotime(str_replace('T', ' ', $value));
    return $timestamp === false ? null : date('Y-m-d H:i:s', $timestamp);
}

function set_xlsx_cell_value(DOMDocument $sheet, string $reference, string|int $value, ?string $style = null): void {
    $namespace = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    $xpath = new DOMXPath($sheet);
    $xpath->registerNamespace('m', $namespace);
    if (!preg_match('/^([A-Z]+)(\d+)$/', $reference, $matches)) {
        throw new RuntimeException('Invalid spreadsheet cell reference.');
    }
    $rowNumber = (int)$matches[2];
    $row = $xpath->query('//m:sheetData/m:row[@r="' . $rowNumber . '"]')->item(0);
    if (!$row) {
        $row = $sheet->createElementNS($namespace, 'row');
        $row->setAttribute('r', (string)$rowNumber);
        $sheetData = $xpath->query('//m:sheetData')->item(0);
        if (!$sheetData) {
            throw new RuntimeException('The spreadsheet template has no sheet data.');
        }
        $insertBefore = null;
        foreach ($xpath->query('./m:row', $sheetData) as $existingRow) {
            if ((int)$existingRow->getAttribute('r') > $rowNumber) {
                $insertBefore = $existingRow;
                break;
            }
        }
        $sheetData->insertBefore($row, $insertBefore);
    }

    $cell = $xpath->query('./m:c[@r="' . $reference . '"]', $row)->item(0);
    if (!$cell) {
        $cell = $sheet->createElementNS($namespace, 'c');
        $cell->setAttribute('r', $reference);
        $columnIndex = column_letter_to_index($matches[1]);
        $insertBefore = null;
        foreach ($xpath->query('./m:c', $row) as $existingCell) {
            preg_match('/^([A-Z]+)/', $existingCell->getAttribute('r'), $existingColumn);
            if ($existingColumn && column_letter_to_index($existingColumn[1]) > $columnIndex) {
                $insertBefore = $existingCell;
                break;
            }
        }
        $row->insertBefore($cell, $insertBefore);
    }
    if ($style !== null) {
        $cell->setAttribute('s', $style);
    }
    while ($cell->firstChild) {
        $cell->removeChild($cell->firstChild);
    }
    $cell->removeAttribute('t');

    if (is_int($value)) {
        $cell->appendChild($sheet->createElementNS($namespace, 'v', (string)$value));
        return;
    }

    $cell->setAttribute('t', 'inlineStr');
    $inline = $sheet->createElementNS($namespace, 'is');
    $text = $sheet->createElementNS($namespace, 't');
    if (preg_match('/^\s|\s$/u', $value)) {
        $text->setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
    }
    $text->appendChild($sheet->createTextNode($value));
    $inline->appendChild($text);
    $cell->appendChild($inline);
}

function clear_xlsx_cell_value(DOMElement $cell): void {
    while ($cell->firstChild) {
        $cell->removeChild($cell->firstChild);
    }
    $cell->removeAttribute('t');
}

function update_xlsx_chart_series(string $chartXml, array $categories, array $values, string $categoryFormula, string $valueFormula, array $pointColors = []): string {
    $chart = new DOMDocument();
    if (!$chart->loadXML($chartXml)) {
        throw new RuntimeException('A chart in the spreadsheet template could not be read.');
    }
    $xpath = new DOMXPath($chart);
    $xpath->registerNamespace('c', 'http://schemas.openxmlformats.org/drawingml/2006/chart');
    $xpath->registerNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
    $series = $xpath->query('//c:ser')->item(0);
    if (!$series) {
        return $chartXml;
    }

    $chartNamespace = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
    $drawingNamespace = 'http://schemas.openxmlformats.org/drawingml/2006/main';
    foreach (iterator_to_array($xpath->query('./c:dPt', $series)) as $point) {
        $index = $xpath->query('./c:idx', $point)->item(0);
        if (!$index || (int)$index->getAttribute('val') >= count($pointColors)) {
            $series->removeChild($point);
        }
    }
    foreach ($pointColors as $index => $color) {
        $point = null;
        foreach ($xpath->query('./c:dPt', $series) as $existingPoint) {
            $existingIndex = $xpath->query('./c:idx', $existingPoint)->item(0);
            if ($existingIndex && (int)$existingIndex->getAttribute('val') === $index) {
                $point = $existingPoint;
                break;
            }
        }
        if (!$point) {
            $point = $chart->createElementNS($chartNamespace, 'c:dPt');
            $pointIndex = $chart->createElementNS($chartNamespace, 'c:idx');
            $pointIndex->setAttribute('val', (string)$index);
            $point->appendChild($pointIndex);
            $shape = $chart->createElementNS($chartNamespace, 'c:spPr');
            $fill = $chart->createElementNS($drawingNamespace, 'a:solidFill');
            $rgb = $chart->createElementNS($drawingNamespace, 'a:srgbClr');
            $fill->appendChild($rgb);
            $shape->appendChild($fill);
            $point->appendChild($shape);
            $insertBefore = $xpath->query('./c:dLbls | ./c:cat', $series)->item(0);
            $series->insertBefore($point, $insertBefore);
        }
        $rgb = $xpath->query('.//a:srgbClr', $point)->item(0);
        if (!$rgb) {
            throw new RuntimeException('A chart color in the spreadsheet template could not be updated.');
        }
        $rgb->setAttribute('val', $color);
    }

    foreach ([
        ['expression' => './c:cat//c:f', 'cache' => './c:cat//c:strCache', 'formula' => $categoryFormula, 'values' => $categories],
        ['expression' => './c:val//c:f', 'cache' => './c:val//c:numCache', 'formula' => $valueFormula, 'values' => $values],
    ] as $seriesPart) {
        $formula = $xpath->query($seriesPart['expression'], $series)->item(0);
        if (!$formula) {
            continue;
        }
        $formula->nodeValue = $seriesPart['formula'];
        $cache = $xpath->query($seriesPart['cache'], $series)->item(0);
        if (!$cache) {
            continue;
        }
        $pointCount = $xpath->query('./c:ptCount', $cache)->item(0);
        if (!$pointCount) {
            $pointCount = $chart->createElementNS('http://schemas.openxmlformats.org/drawingml/2006/chart', 'c:ptCount');
            $cache->insertBefore($pointCount, $cache->firstChild);
        }
        $pointCount->setAttribute('val', (string)count($seriesPart['values']));
        foreach (iterator_to_array($cache->childNodes) as $child) {
            if ($child instanceof DOMElement && $child->localName === 'pt') {
                $cache->removeChild($child);
            }
        }
        foreach ($seriesPart['values'] as $index => $value) {
            $point = $chart->createElementNS('http://schemas.openxmlformats.org/drawingml/2006/chart', 'c:pt');
            $point->setAttribute('idx', (string)$index);
            $pointValue = $chart->createElementNS('http://schemas.openxmlformats.org/drawingml/2006/chart', 'c:v');
            $pointValue->appendChild($chart->createTextNode((string)$value));
            $point->appendChild($pointValue);
            $cache->appendChild($point);
        }
    }
    return (string)$chart->saveXML();
}

function build_xlsx_from_template(string $templatePath, array $replacements): string {
    $archive = file_get_contents($templatePath);
    if ($archive === false) {
        throw new RuntimeException('The Excel report template could not be read.');
    }
    $endOffset = strrpos($archive, "PK\x05\x06");
    if ($endOffset === false || strlen($archive) < $endOffset + 22) {
        throw new RuntimeException('The Excel report template is not a valid workbook.');
    }
    $end = unpack('vdisk/vstart_disk/ventries_disk/ventries/Vdirectory_size/Vdirectory_offset/vcomment_length', substr($archive, $endOffset + 4, 18));
    if (!$end || $end['entries'] === 0xffff || $end['directory_offset'] + $end['directory_size'] > strlen($archive)) {
        throw new RuntimeException('The Excel report template uses an unsupported archive format.');
    }

    $entries = [];
    $directoryOffset = (int)$end['directory_offset'];
    $directoryLimit = $directoryOffset + (int)$end['directory_size'];
    while ($directoryOffset + 46 <= $directoryLimit && substr($archive, $directoryOffset, 4) === "PK\x01\x02") {
        $header = unpack('vmade/vneeded/vflags/vmethod/vtime/vdate/Vcrc/Vcompressed/Vuncompressed/vname_length/vextra_length/vcomment_length/vdisk/vinternal/Vexternal/Vlocal_offset', substr($archive, $directoryOffset + 4, 42));
        if (!$header || $header['disk'] !== 0 || ($header['flags'] & 1) !== 0) {
            throw new RuntimeException('The Excel report template contains an unsupported archive entry.');
        }
        $nameStart = $directoryOffset + 46;
        $name = substr($archive, $nameStart, $header['name_length']);
        $localOffset = (int)$header['local_offset'];
        if (substr($archive, $localOffset, 4) !== "PK\x03\x04") {
            throw new RuntimeException('The Excel report template contains an invalid archive entry.');
        }

        if (array_key_exists($name, $replacements)) {
            $data = $replacements[$name];
        } else {
            $local = unpack('vneeded/vflags/vmethod/vtime/vdate/Vcrc/Vcompressed/Vuncompressed/vname_length/vextra_length', substr($archive, $localOffset + 4, 26));
            if (!$local || ($local['flags'] & 1) !== 0) {
                throw new RuntimeException('The Excel report template contains an unsupported archive entry.');
            }
            $dataOffset = $localOffset + 30 + $local['name_length'] + $local['extra_length'];
            $compressedData = substr($archive, $dataOffset, $header['compressed']);
            if ($header['method'] === 0) {
                $data = $compressedData;
            } elseif ($header['method'] === 8) {
                $data = gzinflate($compressedData);
            } else {
                throw new RuntimeException('The Excel report template uses an unsupported compression method.');
            }
            if ($data === false) {
                throw new RuntimeException('The Excel report template contains unreadable data.');
            }
        }

        $compressedData = gzdeflate($data, 6);
        if ($compressedData === false) {
            throw new RuntimeException('The Excel report could not be compressed.');
        }
        $flags = $header['flags'] & 0x800;
        $crc = crc32($data);
        $localHeader = pack('VvvvvvVVVvv', 0x04034b50, 20, $flags, 8, $header['time'], $header['date'], $crc, strlen($compressedData), strlen($data), strlen($name), 0);
        $entries[] = [
            'name' => $name,
            'local' => $localHeader . $name . $compressedData,
            'made' => $header['made'],
            'flags' => $flags,
            'time' => $header['time'],
            'date' => $header['date'],
            'crc' => $crc,
            'compressed' => strlen($compressedData),
            'uncompressed' => strlen($data),
            'external' => $header['external'],
        ];
        $directoryOffset = $nameStart + $header['name_length'] + $header['extra_length'] + $header['comment_length'];
    }
    if (!$entries || $directoryOffset > $directoryLimit) {
        throw new RuntimeException('The Excel report template archive is incomplete.');
    }

    $output = '';
    $centralDirectory = '';
    foreach ($entries as $entry) {
        $localOffset = strlen($output);
        $output .= $entry['local'];
        $centralDirectory .= pack(
            'VvvvvvvVVVvvvvvVV',
            0x02014b50,
            $entry['made'],
            20,
            $entry['flags'],
            8,
            $entry['time'],
            $entry['date'],
            $entry['crc'],
            $entry['compressed'],
            $entry['uncompressed'],
            strlen($entry['name']),
            0,
            0,
            0,
            0,
            $entry['external'],
            $localOffset
        ) . $entry['name'];
    }
    $centralOffset = strlen($output);
    $output .= $centralDirectory;
    $output .= pack('VvvvvVVv', 0x06054b50, 0, 0, count($entries), count($entries), strlen($centralDirectory), $centralOffset, 0);
    return $output;
}

function update_xlsx_dimension(DOMDocument $sheet, string $range): void {
    $xpath = new DOMXPath($sheet);
    $xpath->registerNamespace('m', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
    $dimension = $xpath->query('//m:dimension')->item(0);
    if ($dimension) {
        $dimension->setAttribute('ref', $range);
    }
}

function add_xlsx_fill_styles(DOMDocument $styles, array $colors): array {
    $namespace = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    $xpath = new DOMXPath($styles);
    $xpath->registerNamespace('m', $namespace);
    $fills = $xpath->query('//m:fills')->item(0);
    $cellFormats = $xpath->query('//m:cellXfs')->item(0);
    if (!$fills || !$cellFormats) {
        throw new RuntimeException('The spreadsheet template does not have cell styles.');
    }

    $fillCount = (int)$fills->getAttribute('count');
    $styleCount = (int)$cellFormats->getAttribute('count');
    $styleIds = [];
    foreach (array_unique($colors) as $color) {
        $fill = $styles->createElementNS($namespace, 'fill');
        $pattern = $styles->createElementNS($namespace, 'patternFill');
        $pattern->setAttribute('patternType', 'solid');
        $foreground = $styles->createElementNS($namespace, 'fgColor');
        $foreground->setAttribute('rgb', 'FF' . strtoupper($color));
        $background = $styles->createElementNS($namespace, 'bgColor');
        $background->setAttribute('indexed', '64');
        $pattern->appendChild($foreground);
        $pattern->appendChild($background);
        $fill->appendChild($pattern);
        $fills->appendChild($fill);

        $format = $styles->createElementNS($namespace, 'xf');
        $format->setAttribute('numFmtId', '0');
        $format->setAttribute('fontId', '0');
        $format->setAttribute('fillId', (string)$fillCount);
        $format->setAttribute('borderId', '0');
        $format->setAttribute('xfId', '0');
        $format->setAttribute('applyFill', '1');
        $cellFormats->appendChild($format);
        $styleIds[$color] = (string)$styleCount;
        $fillCount++;
        $styleCount++;
    }
    $fills->setAttribute('count', (string)$fillCount);
    $cellFormats->setAttribute('count', (string)$styleCount);
    return $styleIds;
}

if ($action === 'export_projects') {
    $projects = $pdo->query("SELECT * FROM projects ORDER BY project_name")->fetchAll();

    header('Content-Type: application/vnd.ms-excel; charset=UTF-8');
    header('Content-Disposition: attachment; filename="psa-project-offices-' . date('Y-m-d') . '.xls"');
    echo '<html><head><meta charset="UTF-8"><style>.job-order-no { mso-number-format:"\\@"; }</style></head><body>';
    echo '<table border="1"><thead><tr>';
    foreach (['Job Order No.', 'Project Name', 'Province', 'Region', 'Location', 'Assigned Team', 'Remarks'] as $heading) {
        echo '<th>' . h($heading) . '</th>';
    }
    echo '</tr></thead><tbody>';
    foreach ($projects as $project) {
        echo '<tr>';
        echo '<td class="job-order-no">' . h((string)$project['job_order_no']) . '</td>';
        foreach ([$project['project_name'], $project['province'], $project['region'], $project['location'], $project['assigned_team'], $project['remarks']] as $value) {
            echo '<td>' . h((string)$value) . '</td>';
        }
        echo '</tr>';
    }
    echo '</tbody></table></body></html>';
    exit;
}

if ($action === 'export_outages') {
    $outages = $pdo->query("
        SELECT o.*, p.job_order_no, p.project_name, p.province, p.region, p.location, p.assigned_team
        FROM outages o
        JOIN projects p ON p.id=o.project_id
        ORDER BY o.down_time DESC
    ")->fetchAll();

    header('Content-Type: application/vnd.ms-excel; charset=UTF-8');
    header('Content-Disposition: attachment; filename="psa-outage-log-' . date('Y-m-d') . '.xls"');
    echo '<html><head><meta charset="UTF-8"><style>.job-order-no { mso-number-format:"\\@"; }</style></head><body>';
    echo '<table border="1"><thead><tr>';
    foreach (['Project / Office', 'Region', 'Team', 'Status', 'Priority', 'Root Cause', 'Down Time', 'Up Time', 'Duration', 'Remarks'] as $heading) {
        echo '<th>' . h($heading) . '</th>';
    }
    echo '</tr></thead><tbody>';
    foreach ($outages as $outage) {
        $duration = 'OPEN';
        if ($outage['up_time'] && $outage['down_time']) {
            $seconds = max(0, strtotime($outage['up_time']) - strtotime($outage['down_time']));
            $duration = floor($seconds / 3600) . 'h ' . floor(($seconds % 3600) / 60) . 'm';
        }
        echo '<tr>';
        foreach ([$outage['project_name'], $outage['region'], $outage['assigned_team'], $outage['status'], $outage['priority'], $outage['root_cause'], format_dt($outage['down_time']), format_dt($outage['up_time']), $duration, $outage['remarks']] as $value) {
            echo '<td>' . h((string)$value) . '</td>';
        }
        echo '</tr>';
    }
    echo '</tbody></table></body></html>';
    exit;
}

if ($action === 'export_analysis_report') {
    $outages = $pdo->query(
        "SELECT o.*, p.job_order_no, p.project_name, p.region, p.province, p.location
         FROM outages o
         JOIN projects p ON p.id=o.project_id
         ORDER BY o.down_time DESC"
    )->fetchAll();

    $totalEvents = count($outages);
    $resolved = 0;
    $open = 0;
    $regions = [];
    $sites = [];
    $siteRegions = [];
    $monthlyCounts = [];
    $weeklyCounts = [];
    foreach ($outages as $outage) {
        if ($outage['up_time']) {
            $resolved++;
        } else {
            $open++;
        }
        $region = trim((string)$outage['region']);
        $projectName = (string)$outage['project_name'];
        $regions[$region] = ($regions[$region] ?? 0) + 1;
        $sites[$projectName] = ($sites[$projectName] ?? 0) + 1;
        $siteRegions[$projectName] = $region;

        $downTime = new DateTimeImmutable((string)$outage['down_time']);
        $month = $downTime->format('Y-m');
        $monthlyCounts[$month] = ($monthlyCounts[$month] ?? 0) + 1;
        $weekStart = $downTime->modify('-' . ((int)$downTime->format('N') - 1) . ' days')->format('Y-m-d');
        $weeklyCounts[$weekStart] = ($weeklyCounts[$weekStart] ?? 0) + 1;
    }
    arsort($regions);
    arsort($sites);

    $regionColors = [];
    foreach (array_keys($regions) as $region) {
        $regionColors[$region] = outage_region_color($region);
    }
    $regionRows = [];
    foreach ($regions as $region => $count) {
        $regionRows[] = [$region !== '' ? $region : '(Unspecified)', $count];
    }
    $siteRows = [];
    foreach ($sites as $projectName => $count) {
        $siteRows[] = [
            $siteRegions[$projectName] !== '' ? $siteRegions[$projectName] : '(Unspecified)',
            $projectName,
            $count,
        ];
    }
    usort($siteRows, static function (array $left, array $right): int {
        return ($right[2] <=> $left[2]) ?: strcmp($left[1], $right[1]);
    });

    $monthlyLabels = [];
    $monthlyValues = [];
    $currentMonth = new DateTimeImmutable(date('Y-m-01'));
    for ($offset = 11; $offset >= 0; $offset--) {
        $month = $currentMonth->modify('-' . $offset . ' months');
        $key = $month->format('Y-m');
        $monthlyLabels[] = $month->format('M y');
        $monthlyValues[] = $monthlyCounts[$key] ?? 0;
    }

    $weeklyLabels = [];
    $weeklyValues = [];
    $thisWeek = new DateTimeImmutable('today');
    $thisWeek = $thisWeek->modify('-' . ((int)$thisWeek->format('N') - 1) . ' days');
    for ($offset = 7; $offset >= 0; $offset--) {
        $week = $thisWeek->modify('-' . $offset . ' weeks');
        $key = $week->format('Y-m-d');
        $weeklyLabels[] = $week->format('M d, Y');
        $weeklyValues[] = $weeklyCounts[$key] ?? 0;
    }

    $templatePath = __DIR__ . '/Excel_report_template.xlsx';
    if (!is_file($templatePath) || !is_readable($templatePath)) {
        error_log('Analysis Excel export failed: Excel_report_template.xlsx is missing or unreadable.');
        http_response_code(500);
        header('Content-Type: text/plain; charset=UTF-8');
        echo 'Unable to generate the Excel report because its template is unavailable.';
        exit;
    }

    try {
        $templateEntries = [];
        foreach (['xl/worksheets/sheet3.xml', 'xl/worksheets/sheet4.xml', 'xl/worksheets/sheet5.xml'] as $entryName) {
            $xml = read_xlsx_zip_entry($templatePath, $entryName);
            if (!is_string($xml)) {
                throw new RuntimeException('The Excel report template is missing a report worksheet.');
            }
            $templateEntries[$entryName] = $xml;
        }
        $stylesXml = read_xlsx_zip_entry($templatePath, 'xl/styles.xml');
        $styles = new DOMDocument();
        if (!is_string($stylesXml) || !$styles->loadXML($stylesXml)) {
            throw new RuntimeException('The cell styles in the Excel report template could not be read.');
        }
        $legendStyles = add_xlsx_fill_styles($styles, array_values($regionColors));
        $templateEntries['xl/styles.xml'] = (string)$styles->saveXML();

        $analysis = new DOMDocument();
        if (!$analysis->loadXML($templateEntries['xl/worksheets/sheet3.xml'])) {
            throw new RuntimeException('The Analysis worksheet in the template could not be read.');
        }
        set_xlsx_cell_value($analysis, 'A8', 'Generated: ' . date('F d, Y'));
        foreach ([
            'B11' => $totalEvents,
            'B12' => $resolved,
            'B13' => $open,
            'B14' => count($regions),
            'B15' => count($sites),
        ] as $reference => $value) {
            set_xlsx_cell_value($analysis, $reference, $value);
        }

        $analysisXPath = new DOMXPath($analysis);
        $analysisXPath->registerNamespace('m', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
        $regionTableEnd = max(23, 17 + count($regionRows));
        for ($row = 18; $row <= $regionTableEnd; $row++) {
            foreach (['A', 'B'] as $column) {
                $cell = $analysisXPath->query('//m:sheetData/m:row[@r="' . $row . '"]/m:c[@r="' . $column . $row . '"]')->item(0);
                if ($cell) {
                    clear_xlsx_cell_value($cell);
                }
            }
        }
        foreach ($regionRows as $index => [$region, $count]) {
            $row = 18 + $index;
            set_xlsx_cell_value($analysis, 'A' . $row, $region, '29');
            set_xlsx_cell_value($analysis, 'B' . $row, $count, '30');
        }

        $regionLegendEnd = max(15, 9 + count($regionRows));
        for ($row = 10; $row <= $regionLegendEnd; $row++) {
            $colorCell = $analysisXPath->query('//m:sheetData/m:row[@r="' . $row . '"]/m:c[@r="I' . $row . '"]')->item(0);
            if ($colorCell) {
                $colorCell->setAttribute('s', '37');
            }
            foreach (['J', 'K'] as $column) {
                $cell = $analysisXPath->query('//m:sheetData/m:row[@r="' . $row . '"]/m:c[@r="' . $column . $row . '"]')->item(0);
                if ($cell) {
                    clear_xlsx_cell_value($cell);
                }
            }
        }
        foreach ($regionRows as $index => [$region, $count]) {
            $row = 10 + $index;
            $color = outage_region_color($region === '(Unspecified)' ? '' : $region);
            set_xlsx_cell_value($analysis, 'I' . $row, '', $legendStyles[$color]);
            set_xlsx_cell_value($analysis, 'J' . $row, $region, '37');
            set_xlsx_cell_value($analysis, 'K' . $row, $count, '38');
        }

        $siteTableEnd = max(48, 29 + count($siteRows));
        for ($row = 30; $row <= $siteTableEnd; $row++) {
            foreach (['K', 'L', 'M', 'N'] as $column) {
                $cell = $analysisXPath->query('//m:sheetData/m:row[@r="' . $row . '"]/m:c[@r="' . $column . $row . '"]')->item(0);
                if ($cell) {
                    clear_xlsx_cell_value($cell);
                }
            }
        }
        foreach ($siteRows as $index => [$region, $projectName, $count]) {
            $row = 30 + $index;
            $color = outage_region_color($region === '(Unspecified)' ? '' : $region);
            set_xlsx_cell_value($analysis, 'K' . $row, '', $legendStyles[$color]);
            set_xlsx_cell_value($analysis, 'L' . $row, $region, '37');
            set_xlsx_cell_value($analysis, 'M' . $row, $projectName, '37');
            set_xlsx_cell_value($analysis, 'N' . $row, $count, '38');
        }
        update_xlsx_dimension($analysis, 'A1:N' . max($siteTableEnd, $regionTableEnd));
        $templateEntries['xl/worksheets/sheet3.xml'] = (string)$analysis->saveXML();

        $chartData = new DOMDocument();
        if (!$chartData->loadXML($templateEntries['xl/worksheets/sheet4.xml'])) {
            throw new RuntimeException('The Chart Data worksheet in the template could not be read.');
        }
        $chartDataXPath = new DOMXPath($chartData);
        $chartDataXPath->registerNamespace('m', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
        $clearThrough = max(20, count($regionRows) + 1);
        for ($row = 2; $row <= $clearThrough; $row++) {
            foreach (['A', 'B', 'C'] as $column) {
                $cell = $chartDataXPath->query('//m:sheetData/m:row[@r="' . $row . '"]/m:c[@r="' . $column . $row . '"]')->item(0);
                if ($cell) clear_xlsx_cell_value($cell);
            }
        }
        for ($row = 2; $row <= max(20, count($siteRows) + 1); $row++) {
            foreach (['D', 'E', 'F'] as $column) {
                $cell = $chartDataXPath->query('//m:sheetData/m:row[@r="' . $row . '"]/m:c[@r="' . $column . $row . '"]')->item(0);
                if ($cell) clear_xlsx_cell_value($cell);
            }
        }
        for ($row = 2; $row <= 13; $row++) {
            foreach (['G', 'H'] as $column) {
                $cell = $chartDataXPath->query('//m:sheetData/m:row[@r="' . $row . '"]/m:c[@r="' . $column . $row . '"]')->item(0);
                if ($cell) clear_xlsx_cell_value($cell);
            }
        }
        for ($row = 2; $row <= 9; $row++) {
            foreach (['J', 'K'] as $column) {
                $cell = $chartDataXPath->query('//m:sheetData/m:row[@r="' . $row . '"]/m:c[@r="' . $column . $row . '"]')->item(0);
                if ($cell) clear_xlsx_cell_value($cell);
            }
        }
        foreach ($regionRows as $index => [$region, $count]) {
            $row = 2 + $index;
            set_xlsx_cell_value($chartData, 'A' . $row, $region);
            set_xlsx_cell_value($chartData, 'B' . $row, $count);
            set_xlsx_cell_value($chartData, 'C' . $row, outage_region_color($region === '(Unspecified)' ? '' : $region));
        }
        foreach ($siteRows as $index => [$region, $projectName, $count]) {
            $row = 2 + $index;
            set_xlsx_cell_value($chartData, 'D' . $row, $projectName);
            set_xlsx_cell_value($chartData, 'E' . $row, $count);
            set_xlsx_cell_value($chartData, 'F' . $row, outage_region_color($region === '(Unspecified)' ? '' : $region));
        }
        foreach ($monthlyLabels as $index => $label) {
            $row = 2 + $index;
            set_xlsx_cell_value($chartData, 'G' . $row, $label);
            set_xlsx_cell_value($chartData, 'H' . $row, $monthlyValues[$index]);
        }
        foreach ($weeklyLabels as $index => $label) {
            $row = 2 + $index;
            set_xlsx_cell_value($chartData, 'J' . $row, $label);
            set_xlsx_cell_value($chartData, 'K' . $row, $weeklyValues[$index]);
        }
        update_xlsx_dimension($chartData, 'A1:K' . max(20, count($regionRows) + 1, count($siteRows) + 1));
        $templateEntries['xl/worksheets/sheet4.xml'] = (string)$chartData->saveXML();

        $outageDetails = new DOMDocument();
        if (!$outageDetails->loadXML($templateEntries['xl/worksheets/sheet5.xml'])) {
            throw new RuntimeException('The Outage Details worksheet in the template could not be read.');
        }
        $detailsXPath = new DOMXPath($outageDetails);
        $detailsXPath->registerNamespace('m', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
        foreach ($detailsXPath->query('//m:sheetData/m:row') as $rowNode) {
            if ((int)$rowNode->getAttribute('r') < 8) {
                continue;
            }
            foreach ($detailsXPath->query('./m:c', $rowNode) as $cell) {
                clear_xlsx_cell_value($cell);
            }
            if ((int)$rowNode->getAttribute('r') > count($outages) + 7) {
                $rowNode->parentNode->removeChild($rowNode);
            }
        }
        foreach ($outages as $index => $outage) {
            $row = 8 + $index;
            $values = [
                'A' => (string)$outage['job_order_no'],
                'B' => (string)$outage['project_name'],
                'C' => trim((string)$outage['region']),
                'D' => (string)$outage['province'],
                'E' => (string)$outage['location'],
                'F' => format_dt($outage['down_time']),
                'G' => format_dt($outage['up_time']),
                'H' => $outage['up_time'] ? 'Resolved' : 'Open',
                'I' => (string)$outage['priority'],
                'J' => (string)$outage['remarks'],
            ];
            foreach ($values as $column => $value) {
                set_xlsx_cell_value($outageDetails, $column . $row, $value, $column === 'E' ? '46' : '29');
            }
        }
        $lastDetailsRow = max(7, 7 + count($outages));
        update_xlsx_dimension($outageDetails, 'A1:J' . $lastDetailsRow);
        $templateEntries['xl/worksheets/sheet5.xml'] = (string)$outageDetails->saveXML();

        $regionPointColors = array_map(static function (array $row): string {
            $region = $row[0] === '(Unspecified)' ? '' : $row[0];
            return outage_region_color($region);
        }, $regionRows);
        $sitePointColors = array_map(static function (array $row): string {
            $region = $row[0] === '(Unspecified)' ? '' : $row[0];
            return outage_region_color($region);
        }, $siteRows);
        foreach ([
            'xl/charts/chart1.xml' => [
                array_column($regionRows, 0),
                array_column($regionRows, 1),
                "'Chart Data'!\$A\$2:\$A\$" . max(2, count($regionRows) + 1),
                "'Chart Data'!\$B\$2:\$B\$" . max(2, count($regionRows) + 1),
                $regionPointColors,
            ],
            'xl/charts/chart2.xml' => [
                array_column($siteRows, 1),
                array_column($siteRows, 2),
                "'Chart Data'!\$D\$2:\$D\$" . max(2, count($siteRows) + 1),
                "'Chart Data'!\$E\$2:\$E\$" . max(2, count($siteRows) + 1),
                $sitePointColors,
            ],
            'xl/charts/chart3.xml' => [
                $monthlyLabels,
                $monthlyValues,
                "'Chart Data'!\$G\$2:\$G\$13",
                "'Chart Data'!\$H\$2:\$H\$13",
                [],
            ],
            'xl/charts/chart4.xml' => [
                $weeklyLabels,
                $weeklyValues,
                "'Chart Data'!\$J\$2:\$J\$9",
                "'Chart Data'!\$K\$2:\$K\$9",
                [],
            ],
        ] as $entryName => [$categories, $values, $categoryFormula, $valueFormula, $pointColors]) {
            $xml = read_xlsx_zip_entry($templatePath, $entryName);
            if (!is_string($xml)) {
                throw new RuntimeException('A chart in the Excel report template is missing.');
            }
            $templateEntries[$entryName] = update_xlsx_chart_series($xml, $categories ?: [''], $values ?: [0], $categoryFormula, $valueFormula, $pointColors);
        }

        $workbookXml = read_xlsx_zip_entry($templatePath, 'xl/workbook.xml');
        if (!is_string($workbookXml)) {
            throw new RuntimeException('The workbook metadata in the template could not be read.');
        }
        $workbook = new DOMDocument();
        if (!$workbook->loadXML($workbookXml)) {
            throw new RuntimeException('The workbook metadata in the template is invalid.');
        }
        $workbookXPath = new DOMXPath($workbook);
        $workbookXPath->registerNamespace('m', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
        $calculation = $workbookXPath->query('//m:calcPr')->item(0);
        if ($calculation) {
            $calculation->setAttribute('calcMode', 'auto');
            $calculation->setAttribute('fullCalcOnLoad', '1');
            $calculation->setAttribute('forceFullCalc', '1');
            $templateEntries['xl/workbook.xml'] = (string)$workbook->saveXML();
        }

        $report = build_xlsx_from_template($templatePath, $templateEntries);
    } catch (RuntimeException $e) {
        error_log('Analysis Excel export failed: ' . $e->getMessage());
        http_response_code(500);
        header('Content-Type: text/plain; charset=UTF-8');
        echo 'Unable to generate the Excel report. Please check the server log.';
        exit;
    }

    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment; filename="psa-outage-analysis-report-' . date('Y-m-d') . '.xlsx"');
    header('Content-Length: ' . strlen($report));
    echo $report;
    exit;
}

/* ---------- POST actions ---------- */
if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $postAction = $_POST['action'] ?? '';

    // Keep each feature's POST workflow in its own file for easier debugging.
    require __DIR__ . '/project_actions.php';
    require __DIR__ . '/outage_actions.php';
}

/* ---------- Dashboard data ---------- */
$today = date('Y-m-d');
$monthStart = date('Y-m-01');
$monthEnd = date('Y-m-t');

$totalProjects = (int)$pdo->query("SELECT COUNT(*) FROM projects")->fetchColumn();
$activeProjects = $totalProjects;
$openOutages = (int)$pdo->query("SELECT COUNT(*) FROM outages WHERE LOWER(TRIM(status)) IN ('open', 'ongoing')")->fetchColumn();
$monthOutages = (int)$pdo->query("SELECT COUNT(*) FROM outages WHERE DATE(down_time) BETWEEN '$monthStart' AND '$monthEnd'")->fetchColumn();

$recentOutages = $pdo->query("
    SELECT o.*, p.job_order_no, p.project_name, p.province, p.region, p.location, p.assigned_team
    FROM outages o
    JOIN projects p ON p.id=o.project_id
    ORDER BY o.down_time DESC
    LIMIT 5
")->fetchAll();

$outagePageSize = 10;
$outagePage = max(1, (int)($_GET['outage_page'] ?? 1));
$outageCount = (int)$pdo->query("SELECT COUNT(*) FROM outages")->fetchColumn();
$outagePageCount = max(1, (int)ceil($outageCount / $outagePageSize));
$outagePage = min($outagePage, $outagePageCount);
$outageOffset = ($outagePage - 1) * $outagePageSize;
$outageListQuery = $pdo->prepare("
    SELECT o.*, p.job_order_no, p.project_name, p.province, p.region, p.location, p.assigned_team
    FROM outages o
    JOIN projects p ON p.id=o.project_id
    ORDER BY o.down_time DESC
    LIMIT :limit OFFSET :offset
");
$outageListQuery->bindValue(':limit', $outagePageSize, PDO::PARAM_INT);
$outageListQuery->bindValue(':offset', $outageOffset, PDO::PARAM_INT);
$outageListQuery->execute();
$outageList = $outageListQuery->fetchAll();
$outageFrom = $outageCount > 0 ? $outageOffset + 1 : 0;
$outageTo = min($outageOffset + count($outageList), $outageCount);
$currentDateTime = date('Y-m-d\TH:i');

$projects = $pdo->query("SELECT * FROM projects ORDER BY project_name")->fetchAll();
$projectPageSize = 10;
$projectPage = max(1, (int)($_GET['page'] ?? 1));
$projectCount = count($projects);
$projectPageCount = max(1, (int)ceil($projectCount / $projectPageSize));
$projectPage = min($projectPage, $projectPageCount);
$projectOffset = ($projectPage - 1) * $projectPageSize;
$projectList = $pdo->prepare("SELECT * FROM projects ORDER BY project_name LIMIT :limit OFFSET :offset");
$projectList->bindValue(':limit', $projectPageSize, PDO::PARAM_INT);
$projectList->bindValue(':offset', $projectOffset, PDO::PARAM_INT);
$projectList->execute();
$projectList = $projectList->fetchAll();

$officeAnalysis = $pdo->query("SELECT p.job_order_no, p.project_name, p.region, COUNT(o.id) AS outage_count
    FROM projects p LEFT JOIN outages o ON o.project_id=p.id
    GROUP BY p.id ORDER BY outage_count DESC, p.project_name")->fetchAll();
$officePageSize = 10;
$officeCount = count($officeAnalysis);
$officePageCount = max(1, (int)ceil($officeCount / $officePageSize));
$officePage = min(max(1, (int)($_GET['office_page'] ?? 1)), $officePageCount);
$officeOffset = ($officePage - 1) * $officePageSize;
$officeAnalysisPage = array_slice($officeAnalysis, $officeOffset, $officePageSize);
$regionAnalysis = $pdo->query("SELECT p.region, COUNT(o.id) AS outage_count
    FROM projects p JOIN outages o ON o.project_id=p.id
    GROUP BY p.region ORDER BY outage_count DESC, p.region")->fetchAll();
$monthlyAnalysis = $pdo->query("SELECT DATE_FORMAT(down_time, '%Y-%m') AS outage_month,
    COUNT(*) AS outage_count, COUNT(DISTINCT project_id) AS site_count
    FROM outages GROUP BY DATE_FORMAT(down_time, '%Y-%m') ORDER BY outage_month DESC")->fetchAll();
$monthlyTrendMap = [];
foreach ($pdo->query("SELECT DATE_FORMAT(down_time, '%Y-%m') AS period, COUNT(*) AS outage_count FROM outages GROUP BY DATE_FORMAT(down_time, '%Y-%m')")->fetchAll() as $row) {
    $monthlyTrendMap[$row['period']] = (int)$row['outage_count'];
}
$monthlyTrendData = [];
$monthCursor = new DateTimeImmutable('first day of this month');
for ($i = 11; $i >= 0; $i--) {
    $period = $monthCursor->modify('-' . $i . ' month')->format('Y-m');
    $monthlyTrendData[] = [
        'period' => $period,
        'outage_count' => $monthlyTrendMap[$period] ?? 0
    ];
}
$weeklyTrendMap = [];
foreach ($pdo->query("SELECT DATE_FORMAT(DATE_SUB(down_time, INTERVAL ((DAYOFWEEK(down_time)+5) % 7) DAY), '%Y-%m-%d') AS week_start, COUNT(*) AS outage_count
    FROM outages
    GROUP BY DATE_FORMAT(DATE_SUB(down_time, INTERVAL ((DAYOFWEEK(down_time)+5) % 7) DAY), '%Y-%m-%d')
    ORDER BY week_start ASC")->fetchAll() as $row) {
    $weeklyTrendMap[$row['week_start']] = (int)$row['outage_count'];
}
$weeklyTrendData = [];
$weekCursor = new DateTimeImmutable('monday this week');
for ($i = 7; $i >= 0; $i--) {
    $weekStart = $weekCursor->modify('-' . $i . ' week')->format('Y-m-d');
    $weeklyTrendData[] = [
        'week_start' => $weekStart,
        'outage_count' => $weeklyTrendMap[$weekStart] ?? 0
    ];
}
$monthlySites = $pdo->query("SELECT DATE_FORMAT(o.down_time, '%Y-%m') AS outage_month,
    p.project_name FROM outages o JOIN projects p ON p.id=o.project_id
    GROUP BY outage_month, p.id ORDER BY outage_month DESC, p.project_name")->fetchAll();
$sitesByMonth = [];
foreach ($monthlySites as $site) {
    $sitesByMonth[$site['outage_month']][] = $site['project_name'];
}
$monthlyTrendLabels = [];
$monthlyTrendValues = [];
foreach ($monthlyTrendData as $row) {
    $monthlyTrendLabels[] = date('M Y', strtotime($row['period'] . '-01'));
    $monthlyTrendValues[] = (int)$row['outage_count'];
}
$weeklyTrendLabels = [];
$weeklyTrendValues = [];
foreach (array_reverse($weeklyTrendData) as $row) {
    $weeklyTrendLabels[] = date('M d', strtotime($row['week_start']));
    $weeklyTrendValues[] = (int)$row['outage_count'];
}

$flash = get_flash();
?>
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title><?= h(APP_NAME) ?></title>
<link rel="stylesheet" href="style.css?v=<?= (int)filemtime(__DIR__ . '/style.css') ?>">
</head>
<body>
<div class="app">
    <aside class="sidebar">
        <div class="brand">
            <img class="brand-logo" src="logo.png" alt="Royal Cable NPMD">
            <span class="brand-collapsed-title">NPMD</span>
            <div>
                <strong>Royal Cable</strong>
                <small>Daily Internet Monitoring</small>
            </div>
        </div>
        <nav>
            <a class="<?= $action==='dashboard'?'active':'' ?>" href="?action=dashboard"><svg class="nav-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="m3 10 9-7 9 7v10a1 1 0 0 1-1 1h-5v-6H9v6H4a1 1 0 0 1-1-1z"/></svg>Dashboard</a>
            <a class="<?= $action==='projects'?'active':'' ?>" href="?action=projects"><svg class="nav-icon" viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/></svg>Project Offices</a>
            <a class="<?= $action==='outages'?'active':'' ?>" href="?action=outages"><svg class="nav-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="m13 2-9 12h7l-1 8 9-12h-7z"/></svg>Outage Log</a>
            <a class="<?= $action==='analysis'?'active':'' ?>" href="?action=analysis"><svg class="nav-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M4 19V5M4 19h16"/><path d="m7 15 3-4 3 2 5-7"/></svg>Data Analysis</a>
        </nav>
    </aside>

    <main class="main">
        <header class="topbar">
            <div>
                <h1>
                <?php
                    echo $action === 'projects' ? 'Project Offices' :
                        ($action === 'outages' ? 'Outage Log' :
                        ($action === 'analysis' ? 'Data Analysis' : 'NPMD Outage Monitoring'));
                ?>
                </h1>
                <span><?= date('l, F d, Y') ?></span>
            </div>
        </header>

        <?php if ($flash): ?>
            <div class="flash-toast <?= h($flash['type']) ?>" role="status" aria-live="polite">
                <svg class="flash-toast-icon" viewBox="0 0 24 24" aria-hidden="true">
                    <?php if ($flash['type'] === 'success'): ?>
                        <path d="M20 6 9 17l-5-5" />
                    <?php else: ?>
                        <path d="M12 8v5M12 17h.01" /><circle cx="12" cy="12" r="9" />
                    <?php endif; ?>
                </svg>
                <span><?= h($flash['message']) ?></span>
            </div>
        <?php endif; ?>

        <?php if ($action === 'projects'): ?>
            <?php
            $projectRegionLabels = ['NCR' => 'National Capital Region (NCR)', 'CAR' => 'Cordillera Administrative Region (CAR)', 'Region 1' => 'Region I – Ilocos Region', 'Region 2' => 'Region II – Cagayan Valley', 'Region 3' => 'Region III – Central Luzon', 'Region 4A' => 'Region IV-A – CALABARZON', 'Region 4B' => 'MIMAROPA Region', 'Region 5' => 'Region V – Bicol Region', 'Region 6' => 'Region VI – Western Visayas', 'NIR' => 'Negros Island Region (NIR)', 'Region 7' => 'Region VII – Central Visayas', 'Region 8' => 'Region VIII – Eastern Visayas', 'Region 9' => 'Region IX – Zamboanga Peninsula', 'Region 10' => 'Region X – Northern Mindanao', 'Region 11' => 'Region XI – Davao Region', 'Region 12' => 'Region XII – SOCCSKSARGEN', 'Region 13' => 'Region XIII – Caraga', 'BARMM' => 'Bangsamoro Autonomous Region in Muslim Mindanao (BARMM)'];
            $projectProvincesByRegion = ['NCR' => [], 'CAR' => ['Abra', 'Apayao', 'Benguet', 'Ifugao', 'Kalinga', 'Mountain Province'], 'Region 1' => ['Ilocos Norte', 'Ilocos Sur', 'La Union', 'Pangasinan'], 'Region 2' => ['Batanes', 'Cagayan', 'Isabela', 'Nueva Vizcaya', 'Quirino'], 'Region 3' => ['Aurora', 'Bataan', 'Bulacan', 'Nueva Ecija', 'Pampanga', 'Tarlac', 'Zambales'], 'Region 4A' => ['Batangas', 'Cavite', 'Laguna', 'Quezon', 'Rizal'], 'Region 4B' => ['Marinduque', 'Occidental Mindoro', 'Oriental Mindoro', 'Palawan', 'Romblon'], 'Region 5' => ['Albay', 'Camarines Norte', 'Camarines Sur', 'Catanduanes', 'Masbate', 'Sorsogon'], 'Region 6' => ['Aklan', 'Antique', 'Capiz', 'Iloilo', 'Guimaras'], 'NIR' => ['Negros Occidental', 'Negros Oriental', 'Siquijor'], 'Region 7' => ['Bohol', 'Cebu'], 'Region 8' => ['Biliran', 'Eastern Samar', 'Leyte', 'Northern Samar', 'Samar', 'Southern Leyte'], 'Region 9' => ['Sulu', 'Zamboanga del Norte', 'Zamboanga del Sur', 'Zamboanga Sibugay'], 'Region 10' => ['Bukidnon', 'Camiguin', 'Lanao del Norte', 'Misamis Occidental', 'Misamis Oriental'], 'Region 11' => ['Davao de Oro', 'Davao del Norte', 'Davao del Sur', 'Davao Occidental', 'Davao Oriental'], 'Region 12' => ['Cotabato', 'South Cotabato', 'Sultan Kudarat', 'Sarangani'], 'Region 13' => ['Agusan del Norte', 'Agusan del Sur', 'Dinagat Islands', 'Surigao del Norte', 'Surigao del Sur'], 'BARMM' => ['Basilan', 'Lanao del Sur', 'Maguindanao del Norte', 'Maguindanao del Sur', 'Tawi-Tawi']];
            $projectRegions = array_keys($projectRegionLabels);
            $edit = null;
            if (isset($_GET['edit'])) {
                $stmt = $pdo->prepare("SELECT * FROM projects WHERE id=?");
                $stmt->execute([(int)$_GET['edit']]);
                $edit = $stmt->fetch() ?: null;
            }
            ?>
            <?php if ($edit): ?>
            <div class="view-modal edit-modal" role="dialog" aria-modal="true" aria-labelledby="edit-project-title">
                <div class="view-modal-backdrop" aria-hidden="true"></div>
                <section class="view-modal-dialog edit-card">
                    <div class="section-head">
                        <div><h2 id="edit-project-title">Edit Project / Office</h2></div>
                        <a class="btn secondary" href="?action=projects">Cancel</a>
                    </div>
                    <form method="post" class="grid-form">
                        <input type="hidden" name="action" value="save_project">
                        <input type="hidden" name="id" value="<?= (int)$edit['id'] ?>">
                        <input type="hidden" name="return_page" value="<?= $projectPage ?>">
                        <label>Project Name
                            <input required name="project_name" value="<?= h($edit['project_name']) ?>" placeholder="e.g. PSA - CVEA">
                        </label>
                        <label>Region
                            <select class="project-region" name="region" data-provinces="<?= h(json_encode($projectProvincesByRegion, JSON_UNESCAPED_UNICODE)) ?>"><option value="">Select region</option><?php if ($edit['region'] !== '' && !isset($projectRegionLabels[$edit['region']])): ?><option selected value="<?= h($edit['region']) ?>"><?= h($edit['region']) ?></option><?php endif; ?><?php foreach ($projectRegionLabels as $regionValue => $regionLabel): ?><option value="<?= h($regionValue) ?>" <?= $edit['region'] === $regionValue ? 'selected' : '' ?>><?= h($regionLabel) ?></option><?php endforeach; ?></select>
                        </label>
                        <label>Province
                            <select class="project-province" name="province" data-value="<?= h($edit['province']) ?>"><option value="">Select province</option><?php if ($edit['province'] !== '' && !in_array($edit['province'], $projectProvincesByRegion[$edit['region']] ?? [], true)): ?><option selected value="<?= h($edit['province']) ?>"><?= h($edit['province']) ?></option><?php endif; ?><?php foreach ($projectProvincesByRegion[$edit['region']] ?? [] as $provinceOption): ?><option value="<?= h($provinceOption) ?>" <?= $edit['province'] === $provinceOption ? 'selected' : '' ?>><?= h($provinceOption) ?></option><?php endforeach; ?></select>
                        </label>
                        <label class="full">Location
                            <input name="location" value="<?= h($edit['location']) ?>" placeholder="e.g. Quezon City">
                        </label>
                        <label>Job Order No.
                            <input required name="job_order_no" value="<?= h($edit['job_order_no']) ?>" placeholder="e.g.03123123">
                        </label>
                        <label>Assigned Team
                            <input name="assigned_team" value="<?= h($edit['assigned_team']) ?>" placeholder="e.g. NPMD / Field Team">
                        </label>
                        <label class="full">Remarks
                            <textarea name="remarks" rows="3" placeholder="Optional notes"><?= h($edit['remarks']) ?></textarea>
                        </label>
                        <div class="full"><button class="btn" type="submit"><svg class="btn-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M5 4h12l3 3v13H4V4z"/><path d="M8 4v5h8V4M8 16h8"/></svg>Save Changes</button></div>
                    </form>
                </section>
            </div>
            <?php else: ?>
            <?php
            $pendingImports = $_SESSION['pending_project_imports'] ?? [];
            $importValue = static function (array $row, array $keys): string {
                foreach ($keys as $key) {
                    if (isset($row[$key]) && trim((string)$row[$key]) !== '') {
                        return trim((string)$row[$key]);
                    }
                }
                return '';
            };
            ?>
            <?php if ($pendingImports): ?>
            <div class="view-modal edit-modal" role="dialog" aria-modal="true" aria-labelledby="import-project-title">
                <div class="view-modal-backdrop" aria-hidden="true"></div>
                <section class="view-modal-dialog edit-card import-review-dialog">
                    <div class="section-head">
                        <div><h2 id="import-project-title">Review Imported Project / Office</h2><p class="muted">Review all <?= count($pendingImports) ?> imported projects before saving.</p></div>
                        <form method="post"><input type="hidden" name="action" value="cancel_import_project"><button class="btn secondary" type="submit">Cancel</button></form>
                    </div>
                    <form method="post" class="grid-form">
                        <input type="hidden" name="action" value="confirm_import_project">
                        <?php foreach ($pendingImports as $importIndex => $importRow): ?>
                        <div class="full import-project-row" style="position:relative">
                            <div class="import-project-row-head" style="padding-right:88px;margin-bottom:14px"><h3>Project <?= $importIndex + 1 ?></h3><button class="btn small danger remove-import-project" type="button" style="position:absolute;top:16px;right:16px">Delete</button></div>
                            <div class="grid-form">
                                <label>Project Name
                                    <input required name="projects[<?= $importIndex ?>][project_name]" value="<?= h($importValue($importRow, ['project_name', 'project', 'project_office', 'office_name'])) ?>" placeholder="e.g. PSA - CVEA">
                                </label>
                                <label>Region
                                    <?php $importRegion = $importValue($importRow, ['region']); ?>
                                    <select class="project-region" name="projects[<?= $importIndex ?>][region]" data-provinces="<?= h(json_encode($projectProvincesByRegion, JSON_UNESCAPED_UNICODE)) ?>"><option value="">Select region</option><?php if ($importRegion !== '' && !isset($projectRegionLabels[$importRegion])): ?><option selected value="<?= h($importRegion) ?>"><?= h($importRegion) ?></option><?php endif; ?><?php foreach ($projectRegionLabels as $regionValue => $regionLabel): ?><option value="<?= h($regionValue) ?>" <?= $importRegion === $regionValue ? 'selected' : '' ?>><?= h($regionLabel) ?></option><?php endforeach; ?></select>
                                </label>
                                <label>Province
                                    <?php $importProvince = $importValue($importRow, ['province']); ?>
                                    <select class="project-province" name="projects[<?= $importIndex ?>][province]" data-value="<?= h($importProvince) ?>"><option value="">Select province</option><?php if ($importProvince !== '' && !in_array($importProvince, $projectProvincesByRegion[$importRegion] ?? [], true)): ?><option selected value="<?= h($importProvince) ?>"><?= h($importProvince) ?></option><?php endif; ?><?php foreach ($projectProvincesByRegion[$importRegion] ?? [] as $provinceOption): ?><option value="<?= h($provinceOption) ?>" <?= $importProvince === $provinceOption ? 'selected' : '' ?>><?= h($provinceOption) ?></option><?php endforeach; ?></select>
                                </label>
                                <label class="full">Location
                                    <input name="projects[<?= $importIndex ?>][location]" value="<?= h($importValue($importRow, ['location'])) ?>" placeholder="e.g. Quezon City">
                                </label>
                                <label>Job Order No.
                                    <input required name="projects[<?= $importIndex ?>][job_order_no]" value="<?= h($importValue($importRow, ['job_order_no', 'job_order', 'job_order_no_'])) ?>" placeholder="e.g.03123123">
                                </label>
                                <label>Assigned Team
                                    <input name="projects[<?= $importIndex ?>][assigned_team]" value="<?= h($importValue($importRow, ['assigned_team', 'team'])) ?>" placeholder="e.g. NPMD / Field Team">
                                </label>
                                <label class="full">Remarks
                                    <textarea name="projects[<?= $importIndex ?>][remarks]" rows="3" placeholder="Optional notes"><?= h($importValue($importRow, ['remarks'])) ?></textarea>
                                </label>
                            </div>
                        </div>
                        <?php endforeach; ?>
                        <div class="full actions"><button class="btn" type="submit">Save All Projects</button></div>
                    </form>
                </section>
            </div>
            <?php endif; ?>
            <section class="card collapse-card">
                <details class="collapse-panel" open>
                    <summary>
                        <span class="collapse-title">Add Project / Office</span>
                        <span class="collapse-toggle-icon" aria-hidden="true">▾</span>
                    </summary>
                    <div class="collapse-body">
                        <form method="post" class="grid-form">
                            <input type="hidden" name="action" value="save_project">
                            <input type="hidden" name="id" value="0">
                            <label>Project Name
                                <input required name="project_name" value="<?= h($edit['project_name'] ?? '') ?>" placeholder="e.g. PSA - CVEA">
                            </label>
                            <label>Region
                                <select class="project-region" name="region" data-provinces="<?= h(json_encode($projectProvincesByRegion, JSON_UNESCAPED_UNICODE)) ?>"><option value="">Select region</option><?php foreach ($projectRegionLabels as $regionValue => $regionLabel): ?><option value="<?= h($regionValue) ?>"><?= h($regionLabel) ?></option><?php endforeach; ?></select>
                            </label>
                            <label>Province
                                <select class="project-province" name="province"><option value="">Select region first</option></select>
                            </label>
                            <label class="full">Location
                                <input name="location" value="" placeholder="e.g. Quezon City">
                            </label>
                            <label>Job Order No.
                                <input required name="job_order_no" value="<?= h($edit['job_order_no'] ?? '') ?>" placeholder="e.g.03123123">
                            </label>
                            <label>Assigned Team
                                <input name="assigned_team" value="<?= h($edit['assigned_team'] ?? '') ?>" placeholder="e.g. NPMD / Field Team">
                            </label>
                            <label class="full">Remarks
                                <textarea name="remarks" rows="3" placeholder="Optional notes"><?= h($edit['remarks'] ?? '') ?></textarea>
                            </label>
                            <div class="full"><button class="btn" type="submit">Add Project</button></div>
                        </form>
                    </div>
                </details>
            </section>
            <?php endif; ?>

            <section class="card">
                <div class="section-head">
                    <div><h2>Project / Office List</h2></div>
                    <div class="actions uploader-actions">
                        <a class="btn secondary" href="?action=export_projects"><svg class="btn-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3v12M7 10l5 5 5-5M5 21h14"/></svg>Export</a>
                        <form method="post" enctype="multipart/form-data" class="upload-form">
                            <input type="hidden" name="action" value="upload_projects">
                            <label class="file-upload-inline" for="project_csv_upload">
                                <span class="btn secondary"><svg class="btn-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M12 21V9M7 14l5-5 5 5M5 3h14"/></svg>Upload Excel/CSV</span>
                                <input id="project_csv_upload" type="file" name="project_csv" accept=".csv,.xls,.xlsx" />
                            </label>
                            <button class="btn secondary upload-submit" type="submit">Import</button>
                        </form>
                    </div>
                    <label class="table-search-wrap"><span class="sr-only">Search projects</span><input class="table-search" type="search" data-table-search placeholder="Search projects..." aria-label="Search projects"></label>
                </div>
                <div class="table-wrap">
                <table data-filter-table>
                    <thead><tr>
                        <th class="row-number">No.</th>
                        <th>Job Order No.</th>
                        <th>Project</th>
                        <th>Province</th>
                        <th>Region</th>
                        <th>Location</th>
                        <th>Team</th>
                        <th>Actions</th>
                    </tr></thead>
                    <tbody>
                    <?php foreach ($projectList as $projectIndex => $p): ?>
                        <tr id="project-row-<?= (int)$p['id'] ?>">
                            <td class="row-number"><?= $projectOffset + $projectIndex + 1 ?></td>
                            <td><?= h($p['job_order_no']) ?></td>
                            <td><b><?= h($p['project_name']) ?></b><div class="muted"><?= h($p['remarks']) ?></div></td>
                            <td><?= h($p['province']) ?></td>
                            <td><?= h($p['region']) ?></td>
                            <td><?= h($p['location']) ?></td>
                            <td><?= h($p['assigned_team']) ?></td>
                            <td class="actions outage-actions">
                                <a class="btn small secondary" href="?action=projects&edit=<?= (int)$p['id'] ?>&page=<?= $projectPage ?>">Edit</a>
                                <form method="post" onsubmit="return confirm('Delete this project and its outage records?');">
                                    <input type="hidden" name="action" value="delete_project">
                                    <input type="hidden" name="id" value="<?= (int)$p['id'] ?>">
                                    <button class="btn small danger" type="submit">Delete</button>
                                </form>
                            </td>
                        </tr>
                    <?php endforeach; ?>
                    <?php if (!$projects): ?><tr><td colspan="9" class="empty">No projects yet. Add your first PSA office above.</td></tr><?php endif; ?>
                    </tbody>
                </table>
                </div>
                <?php if ($projectPageCount > 1): ?>
                    <div class="pagination" aria-label="Project list pagination">
                        <?php if ($projectPage > 1): ?>
                            <a class="btn small secondary" href="?action=projects&page=<?= $projectPage - 1 ?>">Previous</a>
                        <?php endif; ?>
                        <span>Page <?= $projectPage ?> of <?= $projectPageCount ?></span>
                        <?php if ($projectPage < $projectPageCount): ?>
                            <a class="btn small secondary" href="?action=projects&page=<?= $projectPage + 1 ?>">Next</a>
                        <?php endif; ?>
                    </div>
                <?php endif; ?>
            </section>

        <?php elseif ($action === 'outages'): ?>
            <?php
            $pendingOutages = $_SESSION['pending_outage_imports'] ?? [];
            $outageImportValue = static function (array $row, array $keys, string $default = ''): string {
                foreach ($keys as $key) {
                    if (isset($row[$key]) && trim((string)$row[$key]) !== '') {
                        return trim((string)$row[$key]);
                    }
                }
                return $default;
            };
            $outageImportDate = static function (array $row, array $keys): string {
                foreach ($keys as $key) {
                    if (isset($row[$key]) && trim((string)$row[$key]) !== '') {
                        $date = normalize_import_datetime((string)$row[$key]);
                        return $date ? date('Y-m-d\\TH:i', strtotime($date)) : '';
                    }
                }
                return '';
            };
            $outageImportProject = static function (array $row) use ($projects, $outageImportValue): int {
                $selectedId = (int)$outageImportValue($row, ['project_id']);
                if ($selectedId > 0) return $selectedId;
                $jobOrder = normalize_job_order_no($outageImportValue($row, ['job_order_no', 'job_order', 'job_order_no_']));
                $projectName = strtolower(trim($outageImportValue($row, ['project_name', 'project_office', 'project', 'office', 'office_name'])));
                foreach ($projects as $project) {
                    if ($jobOrder !== '' && normalize_job_order_no((string)$project['job_order_no']) === $jobOrder) return (int)$project['id'];
                }
                foreach ($projects as $project) {
                    if ($projectName !== '' && strtolower(trim((string)$project['project_name'])) === $projectName) return (int)$project['id'];
                }
                return 0;
            };
            ?>
            <?php if ($pendingOutages): ?>
            <div class="view-modal edit-modal" role="dialog" aria-modal="true" aria-labelledby="import-outage-title">
                <div class="view-modal-backdrop" aria-hidden="true"></div>
                <section class="view-modal-dialog edit-card import-review-dialog">
                    <div class="section-head">
                        <div><h2 id="import-outage-title">Review Imported Outages</h2><p class="muted">Review all <?= count($pendingOutages) ?> outage records before saving.</p></div>
                        <form method="post"><input type="hidden" name="action" value="cancel_import_outages"><button class="btn secondary" type="submit">Cancel</button></form>
                    </div>
                    <form method="post" class="grid-form">
                        <input type="hidden" name="action" value="confirm_import_outages">
                        <?php foreach ($pendingOutages as $importIndex => $importRow): ?>
                        <?php $selectedProjectId = $outageImportProject($importRow); ?>
                        <div class="full import-project-row">
                            <h3>Outage <?= $importIndex + 1 ?></h3>
                            <div class="grid-form">
                                <label>Project / Office
                                    <select required name="outages[<?= $importIndex ?>][project_id]">
                                        <option value="">Select project...</option>
                                        <?php foreach ($projects as $project): ?>
                                        <option value="<?= (int)$project['id'] ?>" <?= $selectedProjectId === (int)$project['id'] ? 'selected' : '' ?>><?= h($project['job_order_no'] . ' — ' . $project['project_name']) ?></option>
                                        <?php endforeach; ?>
                                    </select>
                                </label>
                                <label>Status
                                    <select name="outages[<?= $importIndex ?>][status]">
                                        <?php foreach (['Open', 'Ongoing', 'Restored', 'Closed'] as $status): ?>
                                        <option <?= $outageImportValue($importRow, ['status'], 'Open') === $status ? 'selected' : '' ?>><?= h($status) ?></option>
                                        <?php endforeach; ?>
                                    </select>
                                </label>
                                <label>Priority
                                    <select name="outages[<?= $importIndex ?>][priority]">
                                        <?php foreach (['Low', 'Normal', 'High', 'Critical'] as $priority): ?>
                                        <option <?= $outageImportValue($importRow, ['priority'], 'Normal') === $priority ? 'selected' : '' ?>><?= h($priority) ?></option>
                                        <?php endforeach; ?>
                                    </select>
                                </label>
                                <label>Root Cause
                                    <select name="outages[<?= $importIndex ?>][root_cause]">
                                        <?php foreach (['Power Failure', 'Equipment Failure', 'Maintenance Activity', 'Weather', 'Configuration Error', 'Third-Party Damage', 'Unknown', 'Other'] as $cause): ?>
                                        <option <?= $outageImportValue($importRow, ['root_cause'], 'Unknown') === $cause ? 'selected' : '' ?>><?= h($cause) ?></option>
                                        <?php endforeach; ?>
                                    </select>
                                </label>
                                <label>Down Time
                                    <input required type="datetime-local" name="outages[<?= $importIndex ?>][down_time]" value="<?= h($outageImportDate($importRow, ['down_time', 'down'])) ?>">
                                </label>
                                <label>Up Time
                                    <input type="datetime-local" name="outages[<?= $importIndex ?>][up_time]" value="<?= h($outageImportDate($importRow, ['up_time', 'up'])) ?>">
                                </label>
                                <label class="full">Remarks
                                    <textarea name="outages[<?= $importIndex ?>][remarks]" rows="2" placeholder="Optional notes"><?= h($outageImportValue($importRow, ['remarks'])) ?></textarea>
                                </label>
                            </div>
                        </div>
                        <?php endforeach; ?>
                        <div class="full actions"><button class="btn" type="submit">Save All Outages</button></div>
                    </form>
                </section>
            </div>
            <?php endif; ?>
            <section class="card collapse-card">
                <details class="collapse-panel" open>
                    <summary>
                        <span class="collapse-title">Record Internet Outage</span>
                        <span class="collapse-toggle-icon" aria-hidden="true">▾</span>
                    </summary>
                    <div class="collapse-body">
                        <form method="post" class="grid-form">
                            <input type="hidden" name="action" value="save_outage">
                            <label>Project / Office
                                <select required name="project_id">
                                    <option value="">Select project...</option>
                                    <?php foreach ($projects as $p): ?>
                                        <option value="<?= (int)$p['id'] ?>"><?= h($p['project_name']) ?></option>
                                    <?php endforeach; ?>
                                </select>
                            </label>
                            <label>Status
                                <select name="status">
                                    <?php foreach (['Open','Ongoing','Restored','Closed'] as $v): ?>
                                        <option <?= ($v === 'Open') ? 'selected' : '' ?>><?= $v ?></option>
                                    <?php endforeach; ?>
                                </select>
                            </label>
                            <label>Priority
                                <select name="priority">
                                    <?php foreach (['Low','Normal','High','Critical'] as $v): ?>
                                        <option <?= ($v === 'Normal') ? 'selected' : '' ?>><?= $v ?></option>
                                    <?php endforeach; ?>
                                </select>
                            </label>
                            <label class="datetime-label"><span class="datetime-label-text"><svg class="datetime-label-icon" viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="4" width="18" height="17" rx="2"/><path d="M7 2v4M17 2v4M3 9h18"/></svg>Down Time</span>
                                <input type="datetime-local" name="down_time" value="<?= h(date('Y-m-d\TH:i', strtotime($currentDateTime))) ?>" required>
                            </label>
                            <label class="datetime-label"><span class="datetime-label-text"><svg class="datetime-label-icon" viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="4" width="18" height="17" rx="2"/><path d="M7 2v4M17 2v4M3 9h18"/></svg>Up Time <span class="muted">(optional)</span></span>
                                <input type="datetime-local" name="up_time" value="">
                            </label>
                            <label>Root Cause
                                <select name="root_cause">
                                    <?php foreach (['Power Failure','Equipment Failure','Maintenance Activity','Weather','Configuration Error','Third-Party Damage','Unknown','Other'] as $v): ?>
                                        <option><?= $v ?></option>
                                    <?php endforeach; ?>
                                </select>
                            </label>
                            <label class="full">Outage Remarks
                                <input name="remarks" placeholder="e.g. PLDT link down, power issue, fiber cut, provider maintenance">
                            </label>
                            <div class="full"><button class="btn" type="submit">Save Outage</button></div>
                        </form>
                    </div>
                </details>
            </section>

            <?php
            $editOutage = null;
            if (isset($_GET['edit_outage'])) {
                $stmt = $pdo->prepare("SELECT * FROM outages WHERE id=?");
                $stmt->execute([(int)$_GET['edit_outage']]);
                $editOutage = $stmt->fetch() ?: null;
            }
            $viewOutage = null;
            if (isset($_GET['view_outage'])) {
                $stmt = $pdo->prepare("SELECT o.*, p.job_order_no, p.project_name, p.province, p.region, p.location, p.assigned_team
                    FROM outages o JOIN projects p ON p.id=o.project_id WHERE o.id=?");
                $stmt->execute([(int)$_GET['view_outage']]);
                $viewOutage = $stmt->fetch() ?: null;
            }
            ?>
            <?php if ($viewOutage): ?>
            <?php
                $viewOutageDuration = 'OPEN';
                if ($viewOutage['down_time'] && $viewOutage['up_time']) {
                    $seconds = max(0, strtotime($viewOutage['up_time']) - strtotime($viewOutage['down_time']));
                    $viewOutageDuration = floor($seconds / 3600) . 'h ' . floor(($seconds % 3600) / 60) . 'm';
                }
            ?>
            <div class="view-modal" role="dialog" aria-modal="true" aria-labelledby="view-outage-title">
                <div class="view-modal-backdrop" aria-hidden="true"></div>
                <section class="view-modal-dialog">
                    <div class="section-head">
                        <div><h2 id="view-outage-title">Outage Details</h2><p>Viewing <?= h($viewOutage['project_name']) ?>.</p></div>
                        <a class="btn secondary" href="?action=outages">Close</a>
                    </div>
                    <div class="detail-grid">
                        <div><span>Job Order No.</span><strong><?= h($viewOutage['job_order_no']) ?></strong></div>
                        <div><span>Project / Location</span><strong><?= h($viewOutage['project_name']) ?> / <?= h($viewOutage['location']) ?></strong></div>
                        <div><span>Region / Team</span><strong><?= h($viewOutage['region']) ?> / <?= h($viewOutage['assigned_team']) ?></strong></div>
                        <div><span>Status</span><strong><?= h($viewOutage['status'] ?? 'Open') ?></strong></div>
                        <div><span>Priority</span><strong><?= h($viewOutage['priority'] ?? 'Normal') ?></strong></div>
                        <div><span>Root Cause</span><strong><?= h($viewOutage['root_cause'] ?? 'Unknown') ?></strong></div>
                        <div><span>Down Time</span><strong><?= h(format_dt($viewOutage['down_time'])) ?></strong></div>
                        <div><span>Up Time</span><strong><?= h(format_dt($viewOutage['up_time'])) ?></strong></div>
                        <div><span>Duration</span><strong><?= h($viewOutageDuration) ?></strong></div>
                        <div class="detail-wide"><span>Remarks</span><strong><?= h($viewOutage['remarks'] ?: 'No remarks') ?></strong></div>
                    </div>
                    <div class="modal-actions">
                        <a class="btn secondary" href="?action=outages&edit_outage=<?= (int)$viewOutage['id'] ?>"><svg class="btn-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="m4 16-.8 4.8L8 20l11.5-11.5a2.1 2.1 0 0 0-3-3Z"/><path d="m14.5 6.5 3 3"/></svg>Edit</a>
                        <?php if (!$viewOutage['up_time']): ?>
                            <form method="post">
                                <input type="hidden" name="action" value="close_outage">
                                <input type="hidden" name="id" value="<?= (int)$viewOutage['id'] ?>">
                                <input type="hidden" name="up_time" value="<?= date('Y-m-d H:i') ?>">
                                <button class="btn" type="submit"><svg class="btn-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3v14M7 12l5 5 5-5M5 21h14"/></svg>Mark Up Now</button>
                            </form>
                        <?php endif; ?>
                        <form method="post" onsubmit="return confirm('Delete this outage record?');">
                            <input type="hidden" name="action" value="delete_outage">
                            <input type="hidden" name="id" value="<?= (int)$viewOutage['id'] ?>">
                            <button class="btn danger" type="submit"><svg class="btn-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M4 7h16M10 11v6M14 11v6M6 7l1 14h10l1-14M9 7V4h6v3"/></svg>Delete</button>
                        </form>
                    </div>
                </section>
            </div>
            <?php endif; ?>
            <?php if ($editOutage): ?>
            <div class="view-modal edit-modal" role="dialog" aria-modal="true" aria-labelledby="edit-outage-title">
                <div class="view-modal-backdrop" aria-hidden="true"></div>
                <section class="view-modal-dialog edit-card">
                <div class="section-head">
                    <div><h2 id="edit-outage-title">Edit Outage Record</h2><p>Correct the project, Down Time, Up Time, or remarks.</p></div>
                    <a class="btn secondary" href="?action=outages">Cancel</a>
                </div>
                <form method="post" class="grid-form">
                    <input type="hidden" name="action" value="update_outage">
                    <input type="hidden" name="id" value="<?= (int)$editOutage['id'] ?>">
                    <label>Project / Office
                        <select required name="project_id">
                            <option value="">Select project...</option>
                            <?php foreach ($projects as $p): ?>
                                <option value="<?= (int)$p['id'] ?>" <?= ((int)$editOutage['project_id']===(int)$p['id'])?'selected':'' ?>>
                                    <?= h($p['project_name']) ?>
                                </option>
                            <?php endforeach; ?>
                        </select>
                    </label>
                    <label>Status
                        <select name="status">
                            <?php foreach (['Open','Ongoing','Restored','Closed'] as $v): ?>
                                <option <?= (($editOutage['status'] ?? 'Open') === $v) ? 'selected' : '' ?>><?= $v ?></option>
                            <?php endforeach; ?>
                        </select>
                    </label>
                    <label>Priority
                        <select name="priority">
                            <?php foreach (['Low','Normal','High','Critical'] as $v): ?>
                                <option <?= (($editOutage['priority'] ?? 'Normal') === $v) ? 'selected' : '' ?>><?= $v ?></option>
                            <?php endforeach; ?>
                        </select>
                    </label>
                    <label class="datetime-label"><span class="datetime-label-text"><svg class="datetime-label-icon" viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="4" width="18" height="17" rx="2"/><path d="M7 2v4M17 2v4M3 9h18"/></svg>Down Time <span class="muted">(leave empty if outage has not started)</span></span>
                        <input type="datetime-local" name="down_time" value="<?= $editOutage['down_time'] ? h(date('Y-m-d\TH:i', strtotime($editOutage['down_time']))) : '' ?>">
                    </label>
                    <label class="datetime-label"><span class="datetime-label-text"><svg class="datetime-label-icon" viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="4" width="18" height="17" rx="2"/><path d="M7 2v4M17 2v4M3 9h18"/></svg>Up Time <span class="muted">(leave empty if still down)</span></span>
                        <input type="datetime-local" name="up_time" value="<?= $editOutage['up_time'] ? h(date('Y-m-d\TH:i', strtotime($editOutage['up_time']))) : '' ?>">
                    </label>
                    <label>Root Cause
                        <select name="root_cause">
                            <?php foreach (['Power Failure','Equipment Failure','Maintenance Activity','Weather','Configuration Error','Third-Party Damage','Unknown','Other'] as $v): ?>
                                <option <?= (($editOutage['root_cause'] ?? 'Unknown') === $v) ? 'selected' : '' ?>><?= $v ?></option>
                            <?php endforeach; ?>
                        </select>
                    </label>
                    <label class="full">Outage Remarks
                        <input name="remarks" value="<?= h($editOutage['remarks']) ?>">
                    </label>
                    <div class="full"><button class="btn" type="submit"><svg class="btn-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M5 4h12l3 3v13H4V4z"/><path d="M8 4v5h8V4M8 16h8"/></svg>Save Changes</button></div>
                </form>
                </section>
            </div>
            <?php endif; ?>

            <section class="card">
                <div class="section-head">
                    <div><h2>Outage History</h2></div>
                    <div class="actions uploader-actions">
                        <a class="btn secondary" href="?action=export_outages"><svg class="btn-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3v12M7 10l5 5 5-5M5 21h14"/></svg>Export</a>
                        <form method="post" enctype="multipart/form-data" class="upload-form">
                            <input type="hidden" name="action" value="upload_outages">
                            <label class="file-upload-inline" for="outage_csv_upload">
                                <span class="btn secondary"><svg class="btn-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M12 21V9M7 14l5-5 5 5M5 3h14"/></svg>Upload Excel/CSV</span>
                                <input id="outage_csv_upload" type="file" name="outage_csv" accept=".csv,.xls,.xlsx" required />
                            </label>
                            <button class="btn secondary upload-submit" type="submit">Import</button>
                        </form>
                    </div>
                    <label class="table-search-wrap"><span class="sr-only">Search outage history</span><input class="table-search" type="search" data-table-search placeholder="Search outage history..." aria-label="Search outage history"></label>
                </div>
                <div class="table-wrap outage-history-wrap">
                <table class="outage-history-table" data-filter-table>
                    <thead><tr>
                        <th class="row-number">No.</th>
                        <th>Job Order No.</th>
                        <th>Project Code</th>
                        <th>Project / Location</th>
                        <th>Down</th>
                        <th></th>
                    </tr></thead>
                    <tbody>
                    <?php foreach ($outageList as $outageIndex => $o): ?>
                        <?php
                        $duration = 'OPEN';
                        if ($o['up_time'] && $o['down_time']) {
                            $seconds = max(0, strtotime($o['up_time']) - strtotime($o['down_time']));
                            $duration = floor($seconds/3600).'h '.floor(($seconds%3600)/60).'m';
                        }
                        ?>
                        <tr>
                            <td class="row-number"><?= $outageOffset + $outageIndex + 1 ?></td>
                            <td><?= h($o['job_order_no']) ?></td>
                            <td><b><?= h($o['outage_code']) ?></b></td>
                            <td><b><?= h($o['project_name']) ?></b></td>
                            <td><?= h(format_dt($o['down_time'])) ?></td>
                            <td class="actions outage-actions">
                                <a class="btn small secondary" href="?action=outages&view_outage=<?= (int)$o['id'] ?>"><svg class="btn-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M2.5 12s3.5-6 9.5-6 9.5 6 9.5 6-3.5 6-9.5 6-9.5-6-9.5-6Z"/><circle cx="12" cy="12" r="2.5"/></svg>View</a>
                            </td>
                        </tr>
                    <?php endforeach; ?>
                    <?php if (!$outageCount): ?><tr><td colspan="6" class="empty">No outage records yet.</td></tr><?php endif; ?>
                    </tbody>
                </table>
                </div>
                <?php if ($outagePageCount > 1): ?>
                    <div class="pagination" aria-label="Outage history pagination">
                        <?php if ($outagePage > 1): ?>
                            <a class="btn small secondary" href="?action=outages&outage_page=<?= $outagePage - 1 ?>"><span aria-hidden="true">‹</span> Previous</a>
                        <?php endif; ?>
                        <span>Page <?= $outagePage ?> of <?= $outagePageCount ?></span>
                        <?php if ($outagePage < $outagePageCount): ?>
                            <a class="btn small secondary" href="?action=outages&outage_page=<?= $outagePage + 1 ?>">Next <span aria-hidden="true">›</span></a>
                        <?php endif; ?>
                    </div>
                <?php endif; ?>
            </section>

        <?php elseif ($action === 'analysis'): ?>
            <section class="stats">
                <div class="stat"><svg class="stat-icon" viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="4" width="18" height="16" rx="2"/><path d="M7 8h10M7 12h10M7 16h6"/></svg><span>Total outage events</span><strong><?= array_sum(array_map('intval', array_column($monthlyAnalysis, 'outage_count'))) ?></strong></div>
                <div class="stat"><svg class="stat-icon" viewBox="0 0 24 24" aria-hidden="true"><circle cx="9" cy="8" r="3"/><path d="M3 19c.4-3 2.5-5 6-5s5.6 2 6 5M16 11a3 3 0 1 0 0-6M16 14c2.8 0 4.5 1.7 5 4"/></svg><span>Offices with outages</span><strong><?= count(array_filter($officeAnalysis, static fn(array $row): bool => (int)$row['outage_count'] > 0)) ?></strong></div>
                <div class="stat"><svg class="stat-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3v18M3 12h18"/><circle cx="12" cy="12" r="8"/></svg><span>Regions with outages</span><strong><?= count($regionAnalysis) ?></strong></div>
                <div class="stat"><svg class="stat-icon" viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="5" width="18" height="16" rx="2"/><path d="M7 3v4M17 3v4M3 10h18M7 14h3M14 14h3M7 17h3"/></svg><span>Months with outages</span><strong><?= count($monthlyAnalysis) ?></strong></div>
            </section>

            <section class="card quick-actions-card">
                <div class="section-head">
                    <div><h2>Analysis Summary</h2></div>
                    <div class="actions">
                        <a class="btn secondary" href="?action=export_analysis_report"><svg class="btn-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3v12M7 10l5 5 5-5M5 21h14"/></svg>Export Excel Report</a>
                    </div>
                </div>
            </section>

            <section class="chart-grid">
                <section class="card chart-card">
                    <div class="section-head"><div><h2>Outages per Region</h2><p>Total outage events grouped by region.</p></div></div>
                    <div class="chart-container region-chart-container"><canvas id="region-pie-chart" aria-label="Donut chart showing outages per region"></canvas></div>
                    <div id="region-pie-legend" class="chart-legend region-chart-legend"></div>
                </section>
                <section class="card chart-card">
                    <div class="section-head"><div><h2>Outages per Site</h2><p>Total outage events grouped by office.</p></div></div>
                    <div class="chart-container"><canvas id="site-bar-chart" aria-label="Bar chart showing outages per site"></canvas></div>
                </section>
            </section>

            <section class="chart-grid">
                <section class="card chart-card">
                    <div class="section-head"><div><h2>Monthly Outage Trend</h2><p>Outages by month for the last 12 months.</p></div></div>
                    <div class="chart-container"><canvas id="monthly-trend-chart" aria-label="Line chart showing outages by month"></canvas></div>
                </section>
                <section class="card chart-card">
                    <div class="section-head"><div><h2>Weekly Outage Trend</h2><p>Outages by week for the latest period.</p></div></div>
                    <div class="chart-container"><canvas id="weekly-trend-chart" aria-label="Line chart showing outages by week"></canvas></div>
                </section>
            </section>

            <section class="card" id="outages-by-office">
                <div class="section-head"><div><h2>Outages by Office</h2></div><label class="table-search-wrap"><span class="sr-only">Search outages by office</span><input class="table-search" type="search" data-table-search placeholder="Search offices..." aria-label="Search outages by office"></label></div>
                <div class="table-wrap">
                <table data-filter-table>
                    <thead><tr>
                        <th class="row-number">No.</th>
                        <th>Job Order No.</th>
                        <th>Project / Office</th>
                        <th>Region</th>
                        <th>Outage Count</th>
                    </tr></thead>
                    <tbody>
                    <?php foreach ($officeAnalysisPage as $analysisIndex => $row): ?>
                        <tr><td class="row-number"><?= $officeOffset + $analysisIndex + 1 ?></td><td><?= h($row['job_order_no']) ?></td><td><?= h($row['project_name']) ?></td><td><?= h($row['region']) ?></td><td><?= (int)$row['outage_count'] ?></td></tr>
                    <?php endforeach; ?>
                    <?php if (!$officeAnalysis): ?><tr><td colspan="5" class="empty">No offices or outage records yet.</td></tr><?php endif; ?>
                    </tbody>
                </table>
                </div>
                <?php if ($officePageCount > 1): ?>
                    <div class="pagination" aria-label="Outages by Office pagination">
                        <?php if ($officePage > 1): ?><a class="btn small secondary" href="?action=<?= h($action) ?>&office_page=<?= $officePage - 1 ?>#outages-by-office">Previous</a><?php endif; ?>
                        <span>Page <?= $officePage ?> of <?= $officePageCount ?> - Showing <?= $officeOffset + 1 ?> to <?= min($officeOffset + $officePageSize, $officeCount) ?> of <?= $officeCount ?></span>
                        <?php if ($officePage < $officePageCount): ?><a class="btn small secondary" href="?action=<?= h($action) ?>&office_page=<?= $officePage + 1 ?>#outages-by-office">Next</a><?php endif; ?>
                    </div>
                <?php endif; ?>
            </section>

            <section class="analysis-grid">
                <section class="card">
                    <div class="section-head"><div><h2>Outages by Region</h2></div><label class="table-search-wrap"><span class="sr-only">Search outages by region</span><input class="table-search" type="search" data-table-search placeholder="Search regions..." aria-label="Search outages by region"></label></div>
                    <div class="table-wrap">
                    <table data-filter-table>
                        <thead><tr><th class="row-number">No.</th><th>Region</th><th>Outage Count</th></tr></thead>
                        <tbody>
                        <?php foreach ($regionAnalysis as $analysisIndex => $row): ?><tr><td class="row-number"><?= $analysisIndex + 1 ?></td><td><?= h($row['region'] ?: 'Unassigned') ?></td><td><?= (int)$row['outage_count'] ?></td></tr><?php endforeach; ?>
                        <?php if (!$regionAnalysis): ?><tr><td colspan="3" class="empty">No outage records yet.</td></tr><?php endif; ?>
                        </tbody>
                    </table>
                    </div>
                </section>
                <section class="card">
                    <div class="section-head"><div><h2>Monthly Summary</h2></div><label class="table-search-wrap"><span class="sr-only">Search monthly summary</span><input class="table-search" type="search" data-table-search placeholder="Search months..." aria-label="Search monthly summary"></label></div>
                    <div class="table-wrap">
                    <table data-filter-table>
                        <thead><tr><th class="row-number">No.</th><th>Month</th><th>Outages</th><th>Sites</th></tr></thead>
                        <tbody>
                        <?php foreach ($monthlyAnalysis as $analysisIndex => $row): ?><tr><td class="row-number"><?= $analysisIndex + 1 ?></td><td><?= h(date('F Y', strtotime($row['outage_month'] . '-01'))) ?></td><td><?= (int)$row['outage_count'] ?></td><td><?= (int)$row['site_count'] ?></td></tr><?php endforeach; ?>
                        <?php if (!$monthlyAnalysis): ?><tr><td colspan="4" class="empty">No outage records yet.</td></tr><?php endif; ?>
                        </tbody>
                    </table>
                    </div>
                </section>
            </section>

            <section class="card">
                <div class="section-head"><div><h2>Sites with Outages per Month</h2></div><label class="table-search-wrap"><span class="sr-only">Search sites by month</span><input class="table-search" type="search" data-table-search placeholder="Search sites..." aria-label="Search sites by month"></label></div>
                <div class="table-wrap">
                <table data-filter-table>
                    <thead><tr><th class="row-number">No.</th><th>Month</th><th>Sites with outage</th></tr></thead>
                    <tbody>
                    <?php foreach ($monthlyAnalysis as $analysisIndex => $row): ?><tr><td class="row-number"><?= $analysisIndex + 1 ?></td><td><?= h(date('F Y', strtotime($row['outage_month'] . '-01'))) ?></td><td><?= h(implode(', ', $sitesByMonth[$row['outage_month']] ?? [])) ?></td></tr><?php endforeach; ?>
                    <?php if (!$monthlyAnalysis): ?><tr><td colspan="3" class="empty">No outage records yet.</td></tr><?php endif; ?>
                    </tbody>
                </table>
                </div>
            </section>

        <?php else: ?>
            <section class="stats">
                <div class="stat"><svg class="stat-icon" viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/></svg><span>Total Projects</span><strong><?= $totalProjects ?></strong></div>
                <div class="stat"><svg class="stat-icon" viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="8"/><path d="m8 12 3 3 5-6"/></svg><span>Active Projects</span><strong><?= $activeProjects ?></strong></div>
                <div class="stat"><svg class="stat-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="m13 2-9 12h7l-1 8 9-12h-7z"/></svg><span>Open Outages</span><strong class="<?= $openOutages ? 'open' : '' ?>"><?= $openOutages ?></strong></div>
                <div class="stat"><svg class="stat-icon" viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="5" width="18" height="16" rx="2"/><path d="M7 3v4M17 3v4M3 10h18M7 14h3M14 14h3M7 17h3"/></svg><span>Outages This Month</span><strong><?= $monthOutages ?></strong></div>
            </section>

            <section class="card quick-actions-card">
                <div class="section-head">
                    <div><h2>Quick Actions⚡</h2></div>
                    <div class="actions">
                        <a class="btn" href="?action=outages"><svg class="btn-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M12 5v14M5 12h14"/></svg>Record Outage</a>
                        <a class="btn secondary" href="?action=export_outages"><svg class="btn-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3v12M7 10l5 5 5-5M5 21h14"/></svg>Export Outage Log</a>
                        <a class="btn secondary" href="?action=projects"><svg class="btn-icon" viewBox="0 0 24 24" aria-hidden="true"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/></svg>Manage Projects</a>
                    </div>
                </div>
            </section>

            <section class="card">
                <div class="section-head"><div><h2>Recent Outages</h2></div><label class="table-search-wrap"><span class="sr-only">Search recent outages</span><input class="table-search" type="search" data-table-search placeholder="Search recent outages..." aria-label="Search recent outages"></label></div>
                <div class="table-wrap">
                <table data-filter-table>
                    <thead><tr>
                        <th class="row-number">No.</th>
                        <th>Job Order No.</th>
                        <th>Project</th>
                        <th>Down</th>
                        <th>Up</th>
                        <th>Duration</th>
                        <th>Status</th>
                    </tr></thead>
                    <tbody>
                    <?php foreach ($recentOutages as $recentIndex => $o): ?>
                        <?php
                        $duration='OPEN';
                        if ($o['up_time'] && $o['down_time']) {
                            $seconds=max(0,strtotime($o['up_time'])-strtotime($o['down_time']));
                            $duration=floor($seconds/3600).'h '.floor(($seconds%3600)/60).'m';
                        }
                        ?>
                        <tr>
                            <td class="row-number"><?= $recentIndex + 1 ?></td>
                            <td><?= h($o['job_order_no']) ?></td>
                            <td><b><?= h($o['project_name']) ?></b></td>
                            <td><?= h(format_dt($o['down_time'])) ?></td>
                            <td><?= h(format_dt($o['up_time'])) ?></td>
                            <td><?= h($duration) ?></td>
                            <td><?= $o['up_time'] ? '<span class="badge">Resolved</span>' : '<span class="badge danger-badge">OPEN</span>' ?></td>
                        </tr>
                    <?php endforeach; ?>
                    <?php if (!$recentOutages): ?><tr><td colspan="7" class="empty">No outage records yet.</td></tr><?php endif; ?>
                    </tbody>
                </table>
                </div>
            </section>
        <?php endif; ?>
    </main>
</div>
<script>
(function () {
    function renumberImportProjects() {
        document.querySelectorAll('.import-review-dialog .import-project-row').forEach(function (row, index) {
            var heading = row.querySelector('h3');
            if (heading) heading.textContent = 'Project ' + (index + 1);
            row.querySelectorAll('[name]').forEach(function (field) {
                field.name = field.name.replace(/^projects\[\d+\]/, 'projects[' + index + ']');
            });
        });
    }

    document.querySelectorAll('.remove-import-project').forEach(function (button) {
        button.addEventListener('click', function () {
            var row = button.closest('.import-project-row');
            if (row) row.remove();
            renumberImportProjects();
        });
    });

    document.querySelectorAll('.project-region').forEach(function (regionSelect) {
        var form = regionSelect.closest('.grid-form');
        var provinceSelect = form ? form.querySelector('.project-province') : null;
        if (!provinceSelect) return;
        var provincesByRegion = JSON.parse(regionSelect.getAttribute('data-provinces') || '{}');
        function refreshProvinces(preserveValue) {
            var region = regionSelect.value;
            if (!Object.prototype.hasOwnProperty.call(provincesByRegion, region)) return;
            var previous = preserveValue ? (provinceSelect.getAttribute('data-value') || provinceSelect.value) : '';
            var provinces = provincesByRegion[region];
            provinceSelect.innerHTML = '<option value="">' + (region === 'NCR' ? 'No provinces in NCR' : 'Select province') + '</option>';
            provinces.forEach(function (province) {
                var option = document.createElement('option');
                option.value = province;
                option.textContent = province;
                provinceSelect.appendChild(option);
            });
            if (previous && provinces.indexOf(previous) === -1) {
                var legacyOption = document.createElement('option');
                legacyOption.value = previous;
                legacyOption.textContent = previous;
                provinceSelect.appendChild(legacyOption);
            }
            provinceSelect.value = previous;
            provinceSelect.removeAttribute('data-value');
        }
        refreshProvinces(true);
        regionSelect.addEventListener('change', function () { refreshProvinces(false); });
    });

    function pad(value) {
        return String(value).padStart(2, '0');
    }

    function parseDateValue(value) {
        if (!value) return null;
        var match = String(value).match(/^(\d{4})-(\d{2})-(\d{2})(?:T|\s)(\d{2}):(\d{2})/);
        var slashMatch = String(value).match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})(?:\s+(\d{1,2}):(\d{2})\s*(AM|PM)?)?$/i);
        var date;
        if (match) {
            date = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]), Number(match[4]), Number(match[5]), 0, 0);
        } else if (slashMatch) {
            var slashHour = Number(slashMatch[4] || 0);
            var slashMinute = Number(slashMatch[5] || 0);
            var meridiem = (slashMatch[6] || '').toUpperCase();
            if (meridiem === 'PM' && slashHour < 12) slashHour += 12;
            if (meridiem === 'AM' && slashHour === 12) slashHour = 0;
            date = new Date(Number(slashMatch[3]), Number(slashMatch[1]) - 1, Number(slashMatch[2]), slashHour, slashMinute, 0, 0);
        } else {
            date = new Date(value);
        }
        return Number.isNaN(date.getTime()) ? null : date;
    }

    function formatForInput(date) {
        return date.getFullYear() + '-' + pad(date.getMonth() + 1) + '-' + pad(date.getDate()) + 'T' + pad(date.getHours()) + ':' + pad(date.getMinutes());
    }

    function formatDisplayValue(value) {
        var date = parseDateValue(value);
        if (!date) return 'Select date & time';
        return pad(date.getMonth() + 1) + '/' + pad(date.getDate()) + '/' + date.getFullYear() + ' ' + get12Hour(date.getHours()) + ':' + pad(date.getMinutes()) + ' ' + getMeridiem(date.getHours());
    }

    function get12Hour(hour) {
        var value = hour % 12;
        return value === 0 ? 12 : value;
    }

    function getMeridiem(hour) {
        return hour >= 12 ? 'PM' : 'AM';
    }

    function renderPicker(widget) {
        var input = widget.querySelector('input[type="hidden"]');
        var displayButton = widget.querySelector('.datetime-display');
        var monthYear = widget.querySelector('.picker-month-year');
        var dayGrid = widget.querySelector('.picker-days');
        var hourSelect = widget.querySelector('.picker-hour');
        var minuteSelect = widget.querySelector('.picker-minute');
        var meridiemSelect = widget.querySelector('.picker-meridiem');
        var state = widget.__dateTimeState || {
            selectedDate: parseDateValue(input.value) || new Date(),
            activeMonth: new Date((parseDateValue(input.value) || new Date()).getFullYear(), (parseDateValue(input.value) || new Date()).getMonth(), 1)
        };
        widget.__dateTimeState = state;

        if (!state.selectedDate && input.value) {
            state.selectedDate = parseDateValue(input.value) || new Date();
        }
        if (!state.activeMonth) {
            state.activeMonth = new Date(state.selectedDate.getFullYear(), state.selectedDate.getMonth(), 1);
        }

        displayButton.value = input.value ? formatDisplayValue(input.value) : '';
        monthYear.textContent = state.activeMonth.toLocaleString([], { month: 'long', year: 'numeric' });

        var firstOfMonth = new Date(state.activeMonth.getFullYear(), state.activeMonth.getMonth(), 1);
        var startDate = new Date(firstOfMonth);
        startDate.setDate(startDate.getDate() - firstOfMonth.getDay());

        dayGrid.innerHTML = '';
        for (var i = 0; i < 42; i++) {
            var dayDate = new Date(startDate);
            dayDate.setDate(startDate.getDate() + i);

            var cell = document.createElement('button');
            cell.type = 'button';
            cell.className = 'picker-day';
            if (dayDate.getMonth() !== state.activeMonth.getMonth()) {
                cell.classList.add('other-month');
            }
            if (dayDate.getFullYear() === state.selectedDate.getFullYear() && dayDate.getMonth() === state.selectedDate.getMonth() && dayDate.getDate() === state.selectedDate.getDate()) {
                cell.classList.add('selected');
            }
            var today = new Date();
            if (dayDate.getFullYear() === today.getFullYear() && dayDate.getMonth() === today.getMonth() && dayDate.getDate() === today.getDate()) {
                cell.classList.add('today');
            }
            cell.textContent = String(dayDate.getDate());
            cell.addEventListener('click', function (event) {
                var clickedDate = new Date(Number(event.currentTarget.dataset.dateValue));
                var baseDate = parseDateValue(input.value) || new Date();
                clickedDate.setHours(baseDate.getHours(), baseDate.getMinutes(), 0, 0);
                state.selectedDate = clickedDate;
                input.value = formatForInput(clickedDate);
                state.activeMonth = new Date(clickedDate.getFullYear(), clickedDate.getMonth(), 1);
                renderPicker(widget);
            });
            cell.dataset.dateValue = String(dayDate.getTime());
            dayGrid.appendChild(cell);
        }

        hourSelect.innerHTML = '';
        for (var hour = 1; hour <= 12; hour++) {
            var hourOption = document.createElement('option');
            hourOption.value = String(hour);
            hourOption.textContent = pad(hour);
            if (hour === get12Hour(state.selectedDate.getHours())) {
                hourOption.selected = true;
            }
            hourSelect.appendChild(hourOption);
        }

        minuteSelect.innerHTML = '';
        for (var minute = 0; minute < 60; minute += 5) {
            var minuteOption = document.createElement('option');
            minuteOption.value = String(minute);
            minuteOption.textContent = pad(minute);
            if (minute === state.selectedDate.getMinutes()) {
                minuteOption.selected = true;
            }
            minuteSelect.appendChild(minuteOption);
        }

        meridiemSelect.innerHTML = '<option value="AM">AM</option><option value="PM">PM</option>';
        meridiemSelect.value = getMeridiem(state.selectedDate.getHours());

        hourSelect.onchange = function () {
            var date = state.selectedDate || new Date();
            var hour = Number(hourSelect.value);
            var minute = Number(minuteSelect.value);
            var meridiem = meridiemSelect.value;
            if (meridiem === 'PM' && hour !== 12) {
                hour += 12;
            }
            if (meridiem === 'AM' && hour === 12) {
                hour = 0;
            }
            date.setHours(hour, minute, 0, 0);
            state.selectedDate = date;
            input.value = formatForInput(date);
            renderPicker(widget);
        };

        minuteSelect.onchange = hourSelect.onchange;
        meridiemSelect.onchange = hourSelect.onchange;

        widget.querySelector('.picker-nav.prev').onclick = function () {
            state.activeMonth = new Date(state.activeMonth.getFullYear(), state.activeMonth.getMonth() - 1, 1);
            renderPicker(widget);
        };
        widget.querySelector('.picker-nav.next').onclick = function () {
            state.activeMonth = new Date(state.activeMonth.getFullYear(), state.activeMonth.getMonth() + 1, 1);
            renderPicker(widget);
        };
        widget.querySelector('.picker-clear').onclick = function () {
            input.value = '';
            state.selectedDate = new Date();
            state.activeMonth = new Date(state.selectedDate.getFullYear(), state.selectedDate.getMonth(), 1);
            displayButton.value = '';
            widget.classList.remove('open');
            renderPicker(widget);
        };
        widget.querySelector('.picker-today').onclick = function () {
            var today = new Date();
            state.selectedDate = today;
            state.activeMonth = new Date(today.getFullYear(), today.getMonth(), 1);
            input.value = formatForInput(today);
            renderPicker(widget);
        };
    }

    function initializeDateTimePicker(widget) {
        var input = widget.querySelector('input[type="hidden"]');
        var displayButton = widget.querySelector('.datetime-display');
        if (!input || !displayButton) return;

        var initialDate = parseDateValue(input.value) || new Date();
        widget.__dateTimeState = {
            selectedDate: initialDate,
            activeMonth: new Date(initialDate.getFullYear(), initialDate.getMonth(), 1)
        };
        renderPicker(widget);

        displayButton.addEventListener('click', function (event) {
            event.stopPropagation();
            document.querySelectorAll('.custom-datetime-picker').forEach(function (other) {
                if (other !== widget) other.classList.remove('open');
            });
            widget.classList.toggle('open');
        });

        function commitTypedDate() {
            var typedDate = parseDateValue(displayButton.value.trim());
            if (!displayButton.value.trim()) {
                input.value = '';
                renderPicker(widget);
                return;
            }
            if (!typedDate) {
                displayButton.value = input.value ? formatDisplayValue(input.value) : '';
                displayButton.setCustomValidity('Enter a valid date and time, such as 09/23/2026 2:30 PM.');
                return;
            }
            displayButton.setCustomValidity('');
            input.value = formatForInput(typedDate);
            widget.__dateTimeState.selectedDate = typedDate;
            widget.__dateTimeState.activeMonth = new Date(typedDate.getFullYear(), typedDate.getMonth(), 1);
            renderPicker(widget);
        }

        displayButton.addEventListener('blur', commitTypedDate);
        displayButton.addEventListener('keydown', function (event) {
            if (event.key === 'Enter') {
                event.preventDefault();
                commitTypedDate();
                widget.classList.remove('open');
            }
        });

        document.addEventListener('click', function (event) {
            if (!widget.contains(event.target)) {
                widget.classList.remove('open');
            }
        });
    }

    document.querySelectorAll('[data-datetime-picker]').forEach(initializeDateTimePicker);

})();

document.querySelectorAll('[data-filter-table]').forEach(function (table) {
    var search = table.closest('.card') ? table.closest('.card').querySelector('[data-table-search]') : null;
    var rows = Array.from(table.querySelectorAll('tbody tr')).filter(function (row) {
        return !row.querySelector('.empty');
    });

    function applySearch() {
        var term = (search ? search.value : '').trim().toLowerCase();
        rows.forEach(function (row) {
            row.hidden = term !== '' && !row.textContent.toLowerCase().includes(term);
        });
    }

    if (search) {
        search.addEventListener('input', applySearch);
    }
});

document.querySelectorAll('.flash-toast').forEach(function (toast) {
    var dismiss = function () {
        toast.style.opacity = '0';
        toast.style.transform = 'translate3d(20px,-10px,0)';
        setTimeout(function () { toast.remove(); }, 220);
    };
    setTimeout(dismiss, 4000);
});

document.querySelectorAll('.upload-form').forEach(function (form) {
    var fileInput = form.querySelector('input[type="file"]');
    var submit = form.querySelector('.upload-submit');
    if (!fileInput || !submit) return;
    fileInput.addEventListener('change', function () {
        submit.disabled = !fileInput.files.length;
        submit.textContent = fileInput.files.length ? 'Import ' + fileInput.files[0].name : 'Import';
    });
    submit.disabled = !fileInput.files.length;
});

var monthlyTrendChartData = <?= json_encode(array_map(static function (array $row): array {
    return ['label' => date('M Y', strtotime($row['period'] . '-01')), 'value' => (int)$row['outage_count']];
}, $monthlyTrendData), JSON_HEX_TAG | JSON_HEX_AMP | JSON_HEX_APOS | JSON_HEX_QUOT) ?>;
var weeklyTrendChartData = <?= json_encode(array_map(static function (array $row): array {
    return ['label' => date('M d', strtotime($row['week_start'])), 'value' => (int)$row['outage_count']];
}, $weeklyTrendData), JSON_HEX_TAG | JSON_HEX_AMP | JSON_HEX_APOS | JSON_HEX_QUOT) ?>;
var regionChartData = <?= json_encode(array_map(static function (array $row): array {
    return [
        'label' => $row['region'] ?: 'Unassigned',
        'value' => (int)$row['outage_count'],
        'color' => '#' . strtolower(outage_region_color((string)$row['region'])),
    ];
}, $regionAnalysis), JSON_HEX_TAG | JSON_HEX_AMP | JSON_HEX_APOS | JSON_HEX_QUOT) ?>;
var siteChartData = <?= json_encode(array_map(static function (array $row): array {
    return [
        'label' => $row['project_name'],
        'value' => (int)$row['outage_count'],
        'color' => '#' . strtolower(outage_region_color((string)$row['region'])),
    ];
}, array_filter($officeAnalysis, static fn(array $row): bool => (int)$row['outage_count'] > 0)), JSON_HEX_TAG | JSON_HEX_AMP | JSON_HEX_APOS | JSON_HEX_QUOT) ?>;

function chartColor(item) {
    return item.color;
}

function drawRegionPieChart(animate) {
    var canvas = document.getElementById('region-pie-chart');
    var legend = document.getElementById('region-pie-legend');
    if (!canvas) return;
    var context = canvas.getContext('2d');
    var width = canvas.clientWidth || 360;
    var height = 260;
    var scale = window.devicePixelRatio || 1;
    canvas.width = width * scale;
    canvas.height = height * scale;
    context.setTransform(scale, 0, 0, scale, 0, 0);
    var total = regionChartData.reduce(function (sum, item) { return sum + item.value; }, 0);
    context.clearRect(0, 0, width, height);
    if (!total) {
        context.fillStyle = '#6d7a86';
        context.font = '14px Segoe UI';
        context.textAlign = 'center';
        context.fillText('No outage data yet.', width / 2, height / 2);
        if (legend) legend.innerHTML = '';
        return;
    }
    var centerX = width / 2;
    var centerY = height / 2;
    var radius = Math.min(width, height) * .36;
    var startAngle = -Math.PI / 2;
    var progress = animate && !window.matchMedia('(prefers-reduced-motion: reduce)').matches ? 0 : 1;
    function renderPie() {
        context.clearRect(0, 0, width, height);
        var currentAngle = startAngle;
        regionChartData.forEach(function (item) {
            var slice = item.value / total * Math.PI * 2 * progress;
            if (slice > 0) {
                context.beginPath();
                context.moveTo(centerX, centerY);
                context.arc(centerX, centerY, radius, currentAngle, currentAngle + slice);
                context.closePath();
                context.fillStyle = chartColor(item);
                context.fill();
                context.strokeStyle = '#fff';
                context.lineWidth = 2;
                context.stroke();

                if (progress === 1 && slice >= .4) {
                    var labelAngle = currentAngle + slice / 2;
                    var labelRadius = radius * .65;
                    var percent = (item.value / total * 100).toFixed(1) + '%';
                    context.save();
                    context.translate(
                        centerX + Math.cos(labelAngle) * labelRadius,
                        centerY + Math.sin(labelAngle) * labelRadius
                    );
                    context.fillStyle = '#fff';
                    context.strokeStyle = 'rgba(20, 35, 50, 0.5)';
                    context.lineWidth = 3;
                    context.lineJoin = 'round';
                    context.font = '700 12px Segoe UI, Arial, sans-serif';
                    context.textAlign = 'center';
                    context.textBaseline = 'middle';
                    context.strokeText(percent, 0, 0);
                    context.fillText(percent, 0, 0);
                    context.restore();
                }
            }
            currentAngle += slice;
        });
    }
    if (progress < 1) {
        var pieStarted = performance.now();
        function animatePie(now) {
            progress = Math.min(1, (now - pieStarted) / 850);
            progress = 1 - Math.pow(1 - progress, 3);
            renderPie();
            if (now - pieStarted < 850) window.requestAnimationFrame(animatePie);
        }
        window.requestAnimationFrame(animatePie);
    } else {
        renderPie();
    }
    if (legend) {
        legend.innerHTML = regionChartData.map(function (item) {
            var share = (item.value / total * 100).toFixed(1);
            return '<span><i style="background:' + chartColor(item) + '"></i><b>' + escapeHtml(item.label) + '</b><small>' + item.value.toLocaleString() + ' outages · ' + share + '%</small></span>';
        }).join('');
    }
}

function drawSiteBarChart(animate) {
    var canvas = document.getElementById('site-bar-chart');
    if (!canvas) return;
    var context = canvas.getContext('2d');
    var width = canvas.clientWidth || 360;
    var chartItems = siteChartData;
    var rowHeight = 28;
    var top = 18;
    var bottom = 28;
    var height = Math.max(300, chartItems.length * rowHeight + top + bottom);
    canvas.style.height = height + 'px';
    var scale = window.devicePixelRatio || 1;
    canvas.width = width * scale;
    canvas.height = height * scale;
    context.scale(scale, scale);
    context.clearRect(0, 0, width, height);
    if (!siteChartData.length) { context.fillStyle = '#6d7a86'; context.font = '14px Segoe UI'; context.fillText('No outage data yet.', 20, 40); return; }
    var maximum = Math.max.apply(null, chartItems.map(function (item) { return item.value; }));
    var scaleMaximum = Math.max(10, maximum);
    var left = Math.min(420, Math.max(260, width * .42));
    var right = 35;
    var chartWidth = width - left - right;
    context.strokeStyle = '#dbe2e8';
    context.fillStyle = '#6d7a86';
    context.font = '10px Segoe UI';
    for (var tick = 0; tick <= scaleMaximum; tick += scaleMaximum / 2) {
        var tickX = left + (tick / scaleMaximum * chartWidth);
        context.beginPath();
        context.moveTo(tickX, top);
        context.lineTo(tickX, height - bottom);
        context.stroke();
        context.textAlign = 'center';
        context.fillText(Math.round(tick), tickX, height - 8);
    }
    context.beginPath();
    context.moveTo(left, top);
    context.lineTo(width - right, top);
    context.stroke();
    var progress = animate && !window.matchMedia('(prefers-reduced-motion: reduce)').matches ? 0 : 1;
    function renderBars() {
        context.clearRect(0, 0, width, height);
        context.strokeStyle = '#dbe2e8';
        context.fillStyle = '#6d7a86';
        context.font = '10px Segoe UI';
        for (var tick = 0; tick <= scaleMaximum; tick += scaleMaximum / 2) {
            var tickX = left + (tick / scaleMaximum * chartWidth);
            context.beginPath();
            context.moveTo(tickX, top);
            context.lineTo(tickX, height - bottom);
            context.stroke();
            context.textAlign = 'center';
            context.fillText(Math.round(tick), tickX, height - 8);
        }
        context.beginPath();
        context.moveTo(left, top);
        context.lineTo(width - right, top);
        context.stroke();
        chartItems.forEach(function (item, index) {
            var y = top + index * rowHeight + 5;
            var barWidth = item.value / scaleMaximum * chartWidth * progress;
            context.fillStyle = chartColor(item);
            context.fillRect(left, y, barWidth, rowHeight - 10);
            context.fillStyle = '#24313d';
            context.font = '11px Segoe UI';
            context.textAlign = 'right';
            context.fillText(item.label, left - 10, y + 12);
            context.textAlign = 'left';
            context.fillText(item.value, left + barWidth + 6, y + 12);
        });
    }
    if (progress < 1) {
        var barsStarted = performance.now();
        function animateBars(now) {
            progress = Math.min(1, (now - barsStarted) / 800);
            renderBars();
            if (progress < 1) window.requestAnimationFrame(animateBars);
        }
        window.requestAnimationFrame(animateBars);
    } else {
        renderBars();
    }
}

function drawMonthlyLineChart(canvasId, data, options) {
    var canvas = document.getElementById(canvasId);
    if (!canvas || !data || !data.length) return;
    var context = canvas.getContext('2d');
    var width = canvas.clientWidth || 360;
    var height = 260;
    var padding = { top: 18, right: 18, bottom: 36, left: 42 };
    var scale = window.devicePixelRatio || 1;
    canvas.width = width * scale;
    canvas.height = height * scale;
    context.setTransform(scale, 0, 0, scale, 0, 0);
    context.clearRect(0, 0, width, height);

    var maxValue = Math.max.apply(null, data.map(function (item) { return item.value; })) || 1;
    var paddedMax = Math.max(maxValue + 1, maxValue * 1.25, 5);
    var chartWidth = width - padding.left - padding.right;
    var chartHeight = height - padding.top - padding.bottom;
    var lineColor = options && options.color ? options.color : '#3a6ea5';
    var fillColor = options && options.fill ? options.fill : 'rgba(58, 110, 165, 0.12)';
    var gridColor = '#dfe6ee';
    var textColor = '#425a72';

    context.strokeStyle = gridColor;
    context.lineWidth = 1;
    context.font = '10px Segoe UI';
    context.fillStyle = textColor;
    context.textAlign = 'right';

    var tickStep = paddedMax / 4;
    for (var tick = 0; tick <= paddedMax + 0.01; tick += tickStep) {
        var y = height - padding.bottom - (tick / paddedMax) * chartHeight;
        context.beginPath();
        context.moveTo(padding.left, y);
        context.lineTo(width - padding.right, y);
        context.stroke();
        context.fillText(String(Math.round(tick)), padding.left - 8, y + 3);
    }

    var points = data.map(function (item, index) {
        var x = padding.left + (index / Math.max(1, data.length - 1)) * chartWidth;
        var y = height - padding.bottom - (item.value / paddedMax) * chartHeight;
        return { x: x, y: y, label: item.label, value: item.value };
    });

    context.beginPath();
    context.moveTo(points[0].x, points[0].y);
    for (var i = 1; i < points.length; i++) {
        context.lineTo(points[i].x, points[i].y);
    }
    context.lineTo(points[points.length - 1].x, height - padding.bottom);
    context.lineTo(points[0].x, height - padding.bottom);
    context.closePath();
    context.fillStyle = fillColor;
    context.fill();

    context.beginPath();
    context.moveTo(points[0].x, points[0].y);
    for (var i = 1; i < points.length; i++) {
        context.lineTo(points[i].x, points[i].y);
    }
    context.strokeStyle = lineColor;
    context.lineWidth = 2.5;
    context.stroke();

    points.forEach(function (point) {
        context.beginPath();
        context.fillStyle = lineColor;
        context.arc(point.x, point.y, 4, 0, Math.PI * 2);
        context.fill();
        context.fillStyle = textColor;
        context.textAlign = 'center';
        context.fillText(point.label, point.x, height - 10);
    });
}

function drawWeeklyBarChart(canvasId, data, options) {
    var canvas = document.getElementById(canvasId);
    if (!canvas || !data || !data.length) return;
    var context = canvas.getContext('2d');
    var width = canvas.clientWidth || 360;
    var height = 260;
    var padding = { top: 18, right: 18, bottom: 40, left: 42 };
    var scale = window.devicePixelRatio || 1;
    canvas.width = width * scale;
    canvas.height = height * scale;
    context.setTransform(scale, 0, 0, scale, 0, 0);
    context.clearRect(0, 0, width, height);

    var maxValue = Math.max.apply(null, data.map(function (item) { return item.value; })) || 1;
    var paddedMax = Math.max(maxValue + 1, maxValue * 1.25, 5);
    var chartWidth = width - padding.left - padding.right;
    var chartHeight = height - padding.top - padding.bottom;
    var barColor = options && options.color ? options.color : '#f28c28';
    var axisColor = '#dfe6ee';
    var textColor = '#425a72';

    context.strokeStyle = axisColor;
    context.lineWidth = 1;
    context.font = '10px Segoe UI';
    context.fillStyle = textColor;
    context.textAlign = 'right';

    var tickStep = paddedMax / 4;
    for (var tick = 0; tick <= paddedMax + 0.01; tick += tickStep) {
        var y = height - padding.bottom - (tick / paddedMax) * chartHeight;
        context.beginPath();
        context.moveTo(padding.left, y);
        context.lineTo(width - padding.right, y);
        context.stroke();
        context.fillText(String(Math.round(tick)), padding.left - 8, y + 3);
    }

    var groupWidth = chartWidth / data.length;
    var barWidth = Math.max(18, groupWidth * 0.68);

    data.forEach(function (item, index) {
        var x = padding.left + index * groupWidth + (groupWidth - barWidth) / 2;
        var barHeight = (item.value / paddedMax) * chartHeight;
        var y = height - padding.bottom - barHeight;
        context.fillStyle = barColor;
        context.fillRect(x, y, barWidth, barHeight);
        context.fillStyle = textColor;
        context.textAlign = 'center';
        context.fillText(item.label, x + barWidth / 2, height - 10);
    });
}

function escapeHtml(value) {
    return String(value).replace(/[&<>"']/g, function (character) {
        return {'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;'}[character];
    });
}

drawRegionPieChart(true);
drawSiteBarChart(true);
drawMonthlyLineChart('monthly-trend-chart', monthlyTrendChartData, { color: '#3a6ea5', fill: 'rgba(58, 110, 165, 0.12)' });
drawWeeklyBarChart('weekly-trend-chart', weeklyTrendChartData, { color: '#f28c28' });
window.addEventListener('resize', function () {
    drawRegionPieChart(false);
    drawSiteBarChart(false);
    drawMonthlyLineChart('monthly-trend-chart', monthlyTrendChartData, { color: '#3a6ea5', fill: 'rgba(58, 110, 165, 0.12)' });
    drawWeeklyBarChart('weekly-trend-chart', weeklyTrendChartData, { color: '#f28c28' });
});
</script>
</body>
</html>
