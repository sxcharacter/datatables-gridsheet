<?php
/**
 * GridSheet - Save Endpoint for Readonly Demo
 * 
 * Demonstrates server-field auto update feature.
 * Returns generated values that will auto-populate readonly columns.
 */

header('Content-Type: application/json');

// Get JSON payload
$payload = json_decode(file_get_contents('php://input'), true);

if (empty($payload) || !is_array($payload)) {
    http_response_code(400);
    echo json_encode([
        'status' => 'error',
        'message' => 'Invalid JSON payload'
    ]);
    exit;
}

$operation = $payload['operation'] ?? '';
$data = $payload['data'] ?? [];
$meta = $payload['meta'] ?? [];

switch ($operation) {
    case 'insert':
        // Single row insert
        $row = $data['row'] ?? [];
        $tempId = $meta['tempId'] ?? null;

        // Generate new ID (in real app, this comes from database)
        $newId = rand(1000, 99999);

        // Generate employee ID code
        $employeeId = 'EMP-' . str_pad($newId, 3, '0', STR_PAD_LEFT);

        // Generate invoice code
        $invoiceCode = 'INV-' . date('Y') . '-' . str_pad($newId, 5, '0', STR_PAD_LEFT);

        // Get current timestamp and user
        $createdAt = date('Y-m-d H:i:s');
        $createdBy = 'admin'; // In real app: $_SESSION['username'] ?? 'system'

        echo json_encode([
            'status' => 'ok',
            'data' => [
                'id' => $newId,
                'tempId' => $tempId,

                // Server-generated fields for data-server-field auto-update
                'employee_id' => $employeeId,
                'invoice_code' => $invoiceCode,
                'created_at' => $createdAt,
                'created_by' => $createdBy
            ]
        ]);
        break;

    case 'insert_batch':
        // Multiple rows insert (from paste)
        $rows = $data['rows'] ?? [];
        $insertedRows = [];

        foreach ($rows as $index => $row) {
            $tempId = $row['_tempId'] ?? null;
            unset($row['_tempId']);

            // Generate new ID
            $newId = rand(1000, 99999) + $index;

            // Generate server-side fields
            $employeeId = 'EMP-' . str_pad($newId, 3, '0', STR_PAD_LEFT);
            $invoiceCode = 'INV-' . date('Y') . '-' . str_pad($newId, 5, '0', STR_PAD_LEFT);
            $createdAt = date('Y-m-d H:i:s');
            $createdBy = 'admin';

            $insertedRows[] = [
                'tempId' => $tempId,
                'id' => $newId,
                // Server-generated fields
                'employee_id' => $employeeId,
                'invoice_code' => $invoiceCode,
                'created_at' => $createdAt,
                'created_by' => $createdBy
            ];
        }

        echo json_encode([
            'status' => 'ok',
            'data' => [
                'inserted' => $insertedRows
            ]
        ]);
        break;

    default:
        // Fallback for legacy format
        if (!empty($data)) {
            $newId = rand(1000, 99999);
            $employeeId = 'EMP-' . str_pad($newId, 3, '0', STR_PAD_LEFT);
            $invoiceCode = 'INV-' . date('Y') . '-' . str_pad($newId, 5, '0', STR_PAD_LEFT);

            echo json_encode([
                'status' => 'ok',
                'data' => [
                    'id' => $newId,
                    'employee_id' => $employeeId,
                    'invoice_code' => $invoiceCode,
                    'created_at' => date('Y-m-d H:i:s'),
                    'created_by' => 'admin'
                ],
                'id' => $newId  // Backward compatibility
            ]);
        } else {
            http_response_code(400);
            echo json_encode([
                'status' => 'error',
                'message' => 'Unknown operation: ' . $operation
            ]);
        }
}
