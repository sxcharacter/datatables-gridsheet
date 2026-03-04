<?php
// datatable-excel/save-batch.php
// Untuk save BATCH (newRows + editedRows + deletedRows)
header('Content-Type: application/json');

// Tangkap input POST JSON
$input = json_decode(file_get_contents('php://input'), true);

if (empty($input) || !is_array($input)) {
    http_response_code(400);
    echo json_encode([
        'status' => 'error',
        'message' => 'Data harus dikirim dalam JSON'
    ]);
    exit;
}

// Pastikan ini adalah batch request
if (empty($input['batch'])) {
    http_response_code(400);
    echo json_encode([
        'status' => 'error',
        'message' => 'Batch flag not set'
    ]);
    exit;
}

$newRows = $input['newRows'] ?? [];
$editedRows = $input['editedRows'] ?? [];
$deletedRows = $input['deletedRows'] ?? [];

$insertedIds = [];
$updatedIds = [];
$deletedIds = [];

// =====================================================
// INSERT NEW ROWS
// =====================================================
// foreach ($newRows as $row) {
//     // Contoh: INSERT INTO employees
//     // $sql = "INSERT INTO employees (name, email, office, age, join_date, salary) VALUES (?, ?, ?, ?, ?, ?)";
//     // $stmt = $pdo->prepare($sql);
//     // $stmt->execute([
//     //     $row['name'] ?? '',
//     //     $row['email'] ?? '',
//     //     $row['office'] ?? '',
//     //     $row['age'] ?? null,
//     //     $row['join_date'] ?? null,
//     //     $row['salary'] ?? null
//     // ]);
//     // $insertedIds[] = $pdo->lastInsertId();
// }

// =====================================================
// UPDATE EDITED ROWS
// =====================================================
// foreach ($editedRows as $edited) {
//     $rowId = $edited['id'];
//     $fields = $edited['fields']; // { "name": "new name", "email": "new@email.com", ... }
//     
//     // Build dynamic UPDATE query
//     $setClauses = [];
//     $values = [];
//     
//     // IMPORTANT: Whitelist allowed fields to prevent SQL injection
//     $allowedFields = ['name', 'email', 'office', 'age', 'join_date', 'salary'];
//     
//     foreach ($fields as $fieldName => $fieldValue) {
//         if (in_array($fieldName, $allowedFields)) {
//             $setClauses[] = "`$fieldName` = ?";
//             $values[] = $fieldValue;
//         }
//     }
//     
//     if (!empty($setClauses)) {
//         $values[] = $rowId; // WHERE id = ?
//         $sql = "UPDATE employees SET " . implode(", ", $setClauses) . " WHERE id = ?";
//         $stmt = $pdo->prepare($sql);
//         $stmt->execute($values);
//         $updatedIds[] = $rowId;
//     }
// }

// =====================================================
// DELETE ROWS
// =====================================================
// foreach ($deletedRows as $rowId) {
//     // Contoh: DELETE FROM employees WHERE id = ?
//     // $sql = "DELETE FROM employees WHERE id = ?";
//     // $stmt = $pdo->prepare($sql);
//     // $stmt->execute([$rowId]);
//     // $deletedIds[] = $rowId;
// }

// Demo response - simulate success
echo json_encode([
    'status' => 'ok',
    'message' => 'Batch save berhasil',
    'inserted' => count($newRows),
    'updated' => count($editedRows),
    'deleted' => count($deletedRows),
    // 'insertedIds' => $insertedIds,  // ID baru untuk rows yang di-insert
    // 'updatedIds' => $updatedIds,    // ID rows yang di-update
    // 'deletedIds' => $deletedIds     // ID rows yang di-delete
]);
