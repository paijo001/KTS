<?php
// Copy this file to config.php and set the environment variables for your server.
define('DB_HOST', getenv('DB_HOST') ?: 'localhost');
define('DB_USER', getenv('DB_USER') ?: 'root');
define('DB_PASS', getenv('DB_PASS') ?: '');
define('DB_NAME', getenv('DB_NAME') ?: 'data_sblw');

$conn = mysqli_connect(DB_HOST, DB_USER, DB_PASS, DB_NAME);
if (!$conn) {
    error_log('Database connection failed: ' . mysqli_connect_error());
    http_response_code(500);
    exit('Database connection failed. Check the server configuration.');
}

define('APP_NAME', 'Aplikasi KTS');
define('APP_VERSION', '1.0');
define('MAYAR_API_KEY', getenv('MAYAR_API_KEY') ?: '');
define('MAYAR_BASE_URL', getenv('MAYAR_BASE_URL') ?: 'https://api.mayar.club/hl/v1');
define('MAYAR_WEBHOOK_SECRET', getenv('MAYAR_WEBHOOK_SECRET') ?: '');
