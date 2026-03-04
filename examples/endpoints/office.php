<?php
// datatable-excel/office-list-s2.php
header('Content-Type: application/json');

// Data array
$options = [
    ['id' => 101, 'text' => 'New York'],
    ['id' => 103, 'text' => 'Los Angeles'],
    ['id' => 104, 'text' => 'Chicago'],
    ['id' => 107, 'text' => 'Houston'],
    ['id' => 109, 'text' => 'Miami'],
    ['id' => 115, 'text' => 'Seattle'],
    ['id' => 120, 'text' => 'Boston'],
    ['id' => 130, 'text' => 'Austin'],
];

echo json_encode($options);
