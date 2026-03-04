<?php
/**
 * GridSheet - Delete Endpoint (Direct Mode)
 * 
 * Handles delete operation:
 * - delete: { operation: 'delete', data: { id }, meta: { timestamp } }
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

// Get ID from new structure or legacy structure
$id = $data['id'] ?? $payload['id'] ?? null;

if (!$id) {
    http_response_code(400);
    echo json_encode([
        'status' => 'error',
        'message' => 'Row ID is required'
    ]);
    exit;
}

// TODO: Delete from database
// Example:
// $stmt = $pdo->prepare("DELETE FROM employees WHERE id = ?");
// $stmt->execute([$id]);

echo json_encode([
    'status' => 'ok',
    'data' => [
        'id' => $id,
        'deleted' => 1
    ],
    // Legacy support
    'success' => true,
    'message' => 'Row deleted successfully'
]);
