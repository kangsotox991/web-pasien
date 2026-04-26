<?php
require_once __DIR__ . '/../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;

header('Content-Type: application/json');

try {
    $input = json_decode(file_get_contents('php://input'), true);

    if (!$input || !isset($input['file'], $input['sheet'], $input['row'], $input['col'])) {
        throw new Exception('Parameter tidak lengkap');
    }

    $filename = basename($input['file']);
    $sheetIndex = (int)$input['sheet'];
    $row = (int)$input['row'] + 1; // Convert 0-based to 1-based
    $col = (int)$input['col'] + 1; // Convert 0-based to 1-based
    $value = $input['value'] ?? '';

    $filepath = __DIR__ . '/../uploads/' . $filename;

    if (!file_exists($filepath)) {
        throw new Exception('File tidak ditemukan');
    }

    $spreadsheet = IOFactory::load($filepath);
    $sheet = $spreadsheet->getSheet($sheetIndex);

    $colLetter = Coordinate::stringFromColumnIndex($col);
    $sheet->setCellValue($colLetter . $row, $value);

    // Determine writer type from extension
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
