<?php
require_once __DIR__ . '/../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;

header('Content-Type: application/json');

try {
    $input = json_decode(file_get_contents('php://input'), true);

    if (!$input || !isset($input['file'], $input['sheet'], $input['row'])) {
        throw new Exception('Parameter tidak lengkap');
    }

    $filename = basename($input['file']);
    $sheetIndex = (int)$input['sheet'];
    $rowIndex = (int)$input['row'] + 1; // Convert 0-based to 1-based
    $filepath = __DIR__ . '/../uploads/' . $filename;

    if (!file_exists($filepath)) {
        throw new Exception('File tidak ditemukan');
    }

    $spreadsheet = IOFactory::load($filepath);
    $sheet = $spreadsheet->getSheet($sheetIndex);

    $sheet->removeRow($rowIndex, 1);

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
