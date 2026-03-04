<?php
/**
 * GridSheet - Save Endpoint (Direct Mode)
 * 
 * Handles two operations:
 * 1. insert - Single new row
 * 2. insert_batch - Multiple new rows (from paste)
 * 
 * Payload structure:
 * - insert: { operation: 'insert', data: { row: {...} }, meta: { tempId, timestamp } }
 * - insert_batch: { operation: 'insert_batch', data: { rows: [...] }, meta: { rowCount, timestamp } }
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

        // TODO: Insert into database
        // Example:
        // $sql = "INSERT INTO employees (name, email, salary) VALUES (?, ?, ?)";
        // $stmt = $pdo->prepare($sql);
        // $stmt->execute([$row['name'], $row['email'], $row['salary']]);
        // $newId = $pdo->lastInsertId();

        // Demo: Generate random ID
        $newId = rand(1000, 99999);

        // Generate server-side fields for data-server-field auto update
        $invoiceCode = 'INV-' . date('Y') . '-' . str_pad($newId, 5, '0', STR_PAD_LEFT);
        $createdAt = date('Y-m-d H:i:s');
        $createdBy = 'admin'; // In real app: $_SESSION['username'] ?? 'system'

        echo json_encode([
            'status' => 'ok',
            'data' => [
                'id' => $newId,
                'tempId' => $tempId,
                // Server-generated fields (will be mapped to data-server-field columns)
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

            // TODO: Insert into database
            // Example:
            // $cols = implode(', ', array_keys($row));
            // $placeholders = implode(', ', array_fill(0, count($row), '?'));
            // $stmt = $pdo->prepare("INSERT INTO employees ($cols) VALUES ($placeholders)");
            // $stmt->execute(array_values($row));
            // $newId = $pdo->lastInsertId();

            // Demo: Generate random ID
            $newId = rand(1000, 99999) + $index;

            // Generate server-side fields for data-server-field auto update
            $invoiceCode = 'INV-' . date('Y') . '-' . str_pad($newId, 5, '0', STR_PAD_LEFT);
            $createdAt = date('Y-m-d H:i:s');
            $createdBy = 'admin';

            $insertedRows[] = [
                'tempId' => $tempId,
                'id' => $newId,
                // Server-generated fields
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
        // Legacy format support (no operation field)
        // Treat as single insert
        if (!empty($data)) {
            $row = $data;
            $newId = rand(1000, 99999);

            echo json_encode([
                'status' => 'ok',
                'data' => ['id' => $newId],
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
