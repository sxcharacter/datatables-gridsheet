<?php
/**
 * GridSheet - Batch Endpoint (Batch Mode)
 * 
 * Handles combined operations when "Save All" is clicked:
 * - batch: { operation: 'batch', data: { insert: [...], update: [...], delete: [...] }, meta: {...} }
 * 
 * Response structure:
 * - { status: 'ok', data: { inserted: [...], updated: N, deleted: N } }
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

// Extract arrays
$inserts = $data['insert'] ?? [];
$updates = $data['update'] ?? [];
$deletes = $data['delete'] ?? [];

// Results
$insertedRows = [];
$updatedCount = 0;
$deletedCount = 0;

// ========== PROCESS INSERTS ==========
foreach ($inserts as $index => $row) {
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
    $createdBy = 'admin'; // In real app: $_SESSION['username'] ?? 'system'

    $insertedRows[] = [
        'tempId' => $tempId,
        'id' => $newId,
        // Server-generated fields (will be mapped to data-server-field columns)
        'invoice_code' => $invoiceCode,
        'created_at' => $createdAt,
        'created_by' => $createdBy
    ];
}

// ========== PROCESS UPDATES ==========
foreach ($updates as $row) {
    $id = $row['id'] ?? null;
    $fields = $row['fields'] ?? [];

    if (!$id)
        continue;

    // TODO: Update database
    // Example:
    // $sets = [];
    // $params = [];
    // foreach ($fields as $key => $value) {
    //     $sets[] = "`$key` = ?";
    //     $params[] = $value;
    // }
    // $params[] = $id;
    // $sql = "UPDATE employees SET " . implode(', ', $sets) . " WHERE id = ?";
    // $stmt = $pdo->prepare($sql);
    // $stmt->execute($params);

    $updatedCount++;
}

// ========== PROCESS DELETES ==========
if (!empty($deletes)) {
    // TODO: Delete from database
    // Example:
    // $placeholders = implode(',', array_fill(0, count($deletes), '?'));
    // $stmt = $pdo->prepare("DELETE FROM employees WHERE id IN ($placeholders)");
    // $stmt->execute($deletes);

    $deletedCount = count($deletes);
}

// Return response
echo json_encode([
    'status' => 'ok',
    'data' => [
        'inserted' => $insertedRows,
        'updated' => $updatedCount,
        'deleted' => $deletedCount
    ],
    'meta' => [
        'insertedCount' => count($insertedRows),
        'updatedCount' => $updatedCount,
        'deletedCount' => $deletedCount,
        'timestamp' => time() * 1000
    ]
]);
