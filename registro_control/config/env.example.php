<?php
// Ejemplo de configuración - Renombrar a env.php y completar con valores reales
return [
    'ORACLE_DB' => [
        'host' => 'host_oracle',
        'port' => '1452',
        'service_name' => 'service_name',
        'username' => 'usuario_oracle',
        'password' => 'contraseña_oracle'
    ],
    'MOODLE_DB' => [
        'host' => 'localhost',
        'dbname' => 'moodle',
        'username' => 'usuario_moodle',
        'password' => 'contraseña_moodle'
    ],
    'EMAIL' => [
        'from' => 'moodle@utp.edu.co',
        'to' => 'soporteunivirtual@utp.edu.co'
    ],
    'PATHS' => [
        'moodle_config' => '/ruta/a/moodle/config.php'
    ]
];