<?php
/**
 * GridSheet - Update Endpoint (Direct Mode)
 * 
 * Handles two operations:
 * 1. update - Single row update
 * 2. update_batch - Multiple rows update (from paste/fill handle)
 * 
 * Payload structure:
 * - update: { operation: 'update', data: { id, fields: {...} }, meta: { fieldCount, timestamp } }
 * - update_batch: { operation: 'update_batch', data: { rows: [...] }, meta: { rowCount, timestamp } }
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
    case 'update':
        // Single row update
        $id = $data['id'] ?? null;
        $fields = $data['fields'] ?? [];

        if (!$id) {
            http_response_code(400);
            echo json_encode([
                'status' => 'error',
                'message' => 'Row ID is required'
            ]);
            exit;
        }

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

        echo json_encode([
            'status' => 'ok',
            'data' => [
                'id' => $id,
                'updated' => 1,
                'fields' => $fields
            ]
        ]);
        break;

    case 'update_batch':
        // Multiple rows update (from paste/fill handle)
        $rows = $data['rows'] ?? [];
        $updatedCount = 0;

        foreach ($rows as $row) {
            $id = $row['id'] ?? null;
            $fields = $row['fields'] ?? [];

            if (!$id)
                continue;

            // TODO: Update database for each row
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

        echo json_encode([
            'status' => 'ok',
            'data' => [
                'updated' => $updatedCount
            ]
        ]);
        break;

    default:
        // Legacy format support (check for 'rows' or 'fields' directly)
        if (isset($payload['rows']) && is_array($payload['rows'])) {
            // Legacy batch update
            $rows = $payload['rows'];
            $updatedCount = count($rows);

            echo json_encode([
                'status' => 'ok',
                'data' => ['updated' => $updatedCount],
                'updated' => $updatedCount
            ]);
        } elseif (isset($payload['id']) && isset($payload['fields'])) {
            // Legacy single update with fields
            echo json_encode([
                'status' => 'ok',
                'data' => [
                    'id' => $payload['id'],
                    'updated' => 1
                ]
            ]);
        } elseif (isset($payload['id']) && isset($payload['field']) && isset($payload['value'])) {
            // Legacy single field update
            echo json_encode([
                'status' => 'ok',
                'id' => $payload['id'],
                'field' => $payload['field'],
                'value' => $payload['value']
            ]);
        } else {
            http_response_code(400);
            echo json_encode([
                'status' => 'error',
                'message' => 'Unknown operation or invalid payload'
            ]);
        }
}
