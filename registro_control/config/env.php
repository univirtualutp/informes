<?php
// Ejemplo de configuración - Renombrar a env.php y completar con valores reales
return [
    'ORACLE_DB' => [
		'driver' => 'oracle',
		'host' => 'clusteroracle.utp.edu.co',
		'port' => '1452',
		'database' => 'PRODUCT',
		'service_name' => 'PRODUCT',
		'username' => 'CONSULTA_UNIVIRTUAL',
		'password' => 'x10v$rta!5key',
		'charset' => 'AL32UTF8',
		'prefix' => '',
    ],
    'MOODLE_DB' => [
        'host' => 'localhost',
        'dbname' => 'moodle',
        'username' => 'moodle',
        'password' => 'M00dl3'
    ],
    'EMAIL' => [
        'from' => 'univirtual@utp.edu.co',
        'to' => 'soporteunivirtual@utp.edu.co'
    ],
    'PATHS' => [
        'moodle_config' => '/data/htdocs/campusunivirtual/moodle/config.php'
    ]
];