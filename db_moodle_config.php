<?php
// db_moodle_config.php
define('DB_HOST', 'localhost');
define('DB_NAME', 'moodle');
define('DB_USER', 'moodle');
define('DB_PASS', 'M00dl3');
define('DB_PORT', '5432');

// Protección contra acceso directo
if(basename($_SERVER['PHP_SELF']) == "db_moodle_config.php") {
    die('Acceso no autorizado');
}
