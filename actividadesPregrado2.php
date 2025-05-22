<?php
// Datos de conexión a PostgreSQL
$host = 'localhost';       // Servidor de la base de datos
$dbname = 'moodle';        // Nombre de la base de datos
$user = 'moodle';          // Usuario de la base de datos
$pass = 'M00dl3';          // Contraseña del usuario
$port = '5432';            // Puerto de PostgreSQL (por defecto es 5432)

try {
    // Conexión a PostgreSQL con PDO
    $dsn = "pgsql:host=$host;port=$port;dbname=$dbname;user=$user;password=$pass";
    $pdo = new PDO($dsn);
    $pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);

    // Si la conexión es exitosa, mostrar un mensaje
    echo "¡Conexión a la base de datos PostgreSQL exitosa!";
} catch (PDOException $e) {
    // Si hay un error, mostrar el mensaje de error
    echo "Error al conectar a la base de datos PostgreSQL: " . $e->getMessage();
}
?>
