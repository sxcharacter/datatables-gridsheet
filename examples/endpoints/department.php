<?php
header('Content-Type: application/json');
header('Access-Control-Allow-Origin: *');

// Return list of departments
echo json_encode([
    'Engineering',
    'Marketing',
    'Finance',
    'Human Resources',
    'Operations',
    'Sales',
    'IT Support',
    'Legal'
]);
