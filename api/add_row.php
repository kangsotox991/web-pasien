<?php
require_once __DIR__ . '/../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;

header('Content-Type: application/json');

try {
    $input = json_decode(file_get_contents('php://input'), true);

    if (!$input || !isset($input['file'], $input['sheet'])) {
        throw new Exception('Parameter tidak lengkap');
    }

    $filename = basename($input['file']);
    $sheetIndex = (int)$input['sheet'];
    $filepath = __DIR__ . '/../uploads/' . $filename;

    if (!file_exists($filepath)) {
        throw new Exception('File tidak ditemukan');
    }

    $spreadsheet = IOFactory::load($filepath);
    $sheet = $spreadsheet->getSheet($sheetIndex);

    $highestRow = $sheet->getHighestRow();
    $highestCol = $sheet->getHighestColumn();
    $highestColIndex = Coordinate::columnIndexFromString($highestCol);

    // Insert empty row at the end
    $newRow = $highestRow + 1;
    for ($col = 1; $col <= $highestColIndex; $col++) {
        $colLetter = Coordinate::stringFromColumnIndex($col);
        $sheet->setCellValue($colLetter . $newRow, '');
    }

    $ext = strtolower(pathinfo($filename, PATHINFO_EXTENSION));
    $writerType = match ($ext) {
        'xlsx' => 'Xlsx',
        'xls' => 'Xls',
        'csv' => 'Csv',
        default => 'Xlsx'
    };

    $writer = IOFactory::createWriter($spreadsheet, $writerType);
    $writer->save($filepath);

    echo json_encode(['success' => true]);

} catch (Exception $e) {
    http_response_code(400);
    echo json_encode([
        'success' => false,
        'message' => $e->getMessage()
    ]);
}
