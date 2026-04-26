<?php
require_once __DIR__ . '/../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;

header('Content-Type: application/json');

try {
    $filename = $_GET['file'] ?? '';
    $sheetIndex = (int)($_GET['sheet'] ?? 0);

    if (empty($filename)) {
        throw new Exception('Parameter file diperlukan');
    }

    // Sanitize filename
    $filename = basename($filename);
    $filepath = __DIR__ . '/../uploads/' . $filename;

    if (!file_exists($filepath)) {
        throw new Exception('File tidak ditemukan');
    }

    $spreadsheet = IOFactory::load($filepath);

    if ($sheetIndex >= $spreadsheet->getSheetCount()) {
        throw new Exception('Sheet index tidak valid');
    }

    $sheet = $spreadsheet->getSheet($sheetIndex);
    $data = [];

    $highestRow = $sheet->getHighestRow();
    $highestCol = $sheet->getHighestColumn();
    $highestColIndex = Coordinate::columnIndexFromString($highestCol);

    for ($row = 1; $row <= $highestRow; $row++) {
        $rowData = [];
        for ($col = 1; $col <= $highestColIndex; $col++) {
            $colLetter = Coordinate::stringFromColumnIndex($col);
            $cell = $sheet->getCell($colLetter . $row);
            $value = $cell->getFormattedValue();
            $rowData[] = $value;
        }
        $data[] = $rowData;
    }

    echo json_encode([
        'success' => true,
        'data' => $data,
        'rows' => $highestRow,
        'cols' => $highestColIndex
    ]);

} catch (Exception $e) {
    http_response_code(400);
    echo json_encode([
        'success' => false,
        'message' => $e->getMessage()
    ]);
}
