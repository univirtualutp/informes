<?php
// Configuración inicial
error_reporting(E_ALL);
ini_set('display_errors', 1);
date_default_timezone_set('America/Bogota');

// Configuración de correo
$destinatario = "daniel.pardo@utp.edu.co";
$asunto = "Resumen de cambios de roles en Moodle - " . date('Y-m-d H:i:s');

// Configuración de bases de datos
$config_oracle = [
    'driver' => 'oracle',
    'host' => 'clusteroracle.utp.edu.co',
    'port' => '1452',
    'database' => 'PRODUCT',
    'service_name' => 'PRODUCT',
    'username' => 'CONSULTA_UNIVIRTUAL',
    'password' => 'x10v$rta!5key',
    'charset' => 'AL32UTF8',
    'prefix' => ''
];

$config_moodle = [
    'host' => 'localhost',
    'port' => '5432',
    'dbname' => 'moodle',
    'username' => 'moodle',
    'password' => 'M00dl3'
];

// Modo simulación (true para solo mostrar cambios, false para ejecutar)
$modo_simulacion = true;

// Archivo de registro para control de ejecuciones
$archivo_registro = __DIR__ . '/registro_cancelaciones_' . date('Y-m-d') . '.log';

// Función para registrar procesamiento
function registrarProcesamiento($archivo, $documento, $idgrupo, $accion) {
    $linea = date('Y-m-d H:i:s') . "|" . $documento . "|" . $idgrupo . "|" . $accion . "\n";
    file_put_contents($archivo, $linea, FILE_APPEND);
}

// Función para verificar si ya fue procesado
function yaProcesado($archivo, $documento, $idgrupo) {
    if (!file_exists($archivo)) return false;
    
    $contenido = file_get_contents($archivo);
    $lineas = explode("\n", $contenido);
    
    foreach ($lineas as $linea) {
        $partes = explode("|", $linea);
        if (count($partes) >= 3 && $partes[1] == $documento && $partes[2] == $idgrupo) {
            return true;
        }
    }
    
    return false;
}

// Función para conectar a Oracle
function conectarOracle($config) {
    $tns = "(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(Host=" . $config['host'] . ")(Port=" . $config['port'] . "))(CONNECT_DATA=(SID=" . $config['service_name'] . ")))";
    
    try {
        $conn = oci_connect($config['username'], $config['password'], $tns);
        if (!$conn) {
            $e = oci_error();
            throw new Exception("Error de conexión a Oracle: " . $e['message']);
        }
        return $conn;
    } catch (Exception $e) {
        die($e->getMessage());
    }
}

// Función para conectar a Moodle
function conectarMoodle($config) {
    try {
        $conn = new PDO(
            "mysql:host=" . $config['host'] . ";dbname=" . $config['dbname'],
            $config['username'],
            $config['password']
        );
        $conn->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
        return $conn;
    } catch (PDOException $e) {
        die("Error de conexión a Moodle: " . $e->getMessage());
    }
}

// Función para obtener cancelaciones de Oracle del día actual
function obtenerCancelacionesOracle($conn) {
    $query = "SELECT IDGRUPO, NUMERODOCUMENTO, VALORTIPO, FECHACREACION
              FROM VI_RYC_UNIVIRTUALASIGNCANCELCO 
              WHERE VALORTIPO = 'cancelacion' 
              AND TRUNC(FECHACREACION) = TRUNC(SYSDATE)";
    
    $stid = oci_parse($conn, $query);
    oci_execute($stid);
    
    $cancelaciones = [];
    while ($row = oci_fetch_assoc($stid)) {
        $cancelaciones[] = [
            'idgrupo' => $row['IDGRUPO'],
            'documento' => $row['NUMERODOCUMENTO'],
            'tipo' => $row['VALORTIPO'],
            'fecha' => $row['FECHACREACION']
        ];
    }
    
    oci_free_statement($stid);
    return $cancelaciones;
}

// Función para buscar usuarios en Moodle (optimizada)
function buscarUsuariosMoodle($conn, $documento, $idgrupo) {
    $query = "SELECT 
                u.id AS id_usuario,
                u.username as codigo,
                CONCAT(u.firstname, ' ', u.lastname) as nombre_completo,
                u.email as correo,
                c.id AS id_curso,
                c.fullname AS curso_nombre,
                r.id AS id_rol,
                r.name AS rol_nombre,
                ra.id AS id_asignacion_rol
              FROM mdl_user u
              JOIN mdl_role_assignments ra ON ra.userid = u.id
              JOIN mdl_role r ON r.id = ra.roleid
              JOIN mdl_context ctx ON ctx.id = ra.contextid AND ctx.contextlevel = 50
              JOIN mdl_course c ON c.id = ctx.instanceid
              JOIN mdl_customfield_data cfd ON cfd.instanceid = c.id
              JOIN mdl_customfield_field cff ON cff.id = cfd.fieldid AND cff.shortname = 'id_sistemas'
              WHERE u.username = :documento
              AND cfd.value = :idgrupo
              AND ra.roleid IN (5, 16, 17)";
    
    $stmt = $conn->prepare($query);
    $stmt->bindParam(':documento', $documento);
    $stmt->bindParam(':idgrupo', $idgrupo);
    $stmt->execute();
    
    return $stmt->fetchAll(PDO::FETCH_ASSOC);
}

// Función para cambiar el rol en Moodle
function cambiarRolMoodle($conn, $id_asignacion_rol, $nuevo_rol_id = 9) {
    $query = "UPDATE mdl_role_assignments SET roleid = :nuevo_rol WHERE id = :id_asignacion";
    $stmt = $conn->prepare($query);
    $stmt->bindParam(':nuevo_rol', $nuevo_rol_id);
    $stmt->bindParam(':id_asignacion', $id_asignacion_rol);
    return $stmt->execute();
}

// Función para enviar correo
function enviarCorreo($destinatario, $asunto, $cuerpo) {
    $headers = "From: no-reply@utp.edu.co\r\n";
    $headers .= "Content-Type: text/html; charset=UTF-8\r\n";
    
    return mail($destinatario, $asunto, $cuerpo, $headers);
}

// Proceso principal
try {
    // Conectar a bases de datos
    $oracle_conn = conectarOracle($config_oracle);
    $moodle_conn = conectarMoodle($config_moodle);
    
    // Obtener cancelaciones del día
    $cancelaciones = obtenerCancelacionesOracle($oracle_conn);
    
    if (empty($cancelaciones)) {
        $mensaje = "No hay cancelaciones para procesar hoy " . date('Y-m-d') . ".\n";
        echo $mensaje;
        registrarProcesamiento($archivo_registro, 'SISTEMA', '0', 'No hay cancelaciones');
        exit;
    }
    
    $cambios_propuestos = [];
    $cambios_realizados = [];
    $procesados = 0;
    $omitidos = 0;
    
    foreach ($cancelaciones as $cancelacion) {
        // Verificar si ya fue procesado
        if (yaProcesado($archivo_registro, $cancelacion['documento'], $cancelacion['idgrupo'])) {
            $omitidos++;
            continue;
        }
        
        $usuarios_moodle = buscarUsuariosMoodle(
            $moodle_conn, 
            $cancelacion['documento'], 
            $cancelacion['idgrupo']
        );
        
        if (!empty($usuarios_moodle)) {
            foreach ($usuarios_moodle as $usuario) {
                $cambio = [
                    'documento' => $cancelacion['documento'],
                    'idgrupo' => $cancelacion['idgrupo'],
                    'usuario_id' => $usuario['id_usuario'],
                    'nombre_completo' => $usuario['nombre_completo'],
                    'curso_id' => $usuario['id_curso'],
                    'curso_nombre' => $usuario['curso_nombre'],
                    'rol_actual' => $usuario['rol_nombre'] . ' ('.$usuario['id_rol'].')',
                    'id_asignacion_rol' => $usuario['id_asignacion_rol']
                ];
                
                $cambios_propuestos[] = $cambio;
                
                if (!$modo_simulacion) {
                    $resultado = cambiarRolMoodle($moodle_conn, $usuario['id_asignacion_rol']);
                    if ($resultado) {
                        $cambio['resultado'] = 'Éxito';
                        $cambios_realizados[] = $cambio;
                        registrarProcesamiento($archivo_registro, $cancelacion['documento'], $cancelacion['idgrupo'], 'ROL_CAMBIADO');
                        $procesados++;
                    } else {
                        $cambio['resultado'] = 'Error';
                        $cambios_realizados[] = $cambio;
                        registrarProcesamiento($archivo_registro, $cancelacion['documento'], $cancelacion['idgrupo'], 'ERROR');
                    }
                } else {
                    registrarProcesamiento($archivo_registro, $cancelacion['documento'], $cancelacion['idgrupo'], 'SIMULACION');
                    $procesados++;
                }
            }
        } else {
            registrarProcesamiento($archivo_registro, $cancelacion['documento'], $cancelacion['idgrupo'], 'NO_ENCONTRADO');
        }
    }
    
    // Mostrar resultados
    $mensaje_correo = "<h1>Resumen de cambios de roles en Moodle</h1>";
    $mensaje_correo .= "<p>Fecha: " . date('Y-m-d H:i:s') . "</p>";
    $mensaje_correo .= "<p>Modo simulación: " . ($modo_simulacion ? 'Sí' : 'No') . "</p>";
    $mensaje_correo .= "<p>Total cancelaciones encontradas: " . count($cancelaciones) . "</p>";
    $mensaje_correo .= "<p>Registros procesados: " . $procesados . "</p>";
    $mensaje_correo .= "<p>Registros omitidos (ya procesados): " . $omitidos . "</p>";
    
    if (!empty($cambios_propuestos)) {
        $mensaje_correo .= "<h2>Cambios " . ($modo_simulacion ? "Propuestos" : "Realizados") . ":</h2>";
        $mensaje_correo .= "<table border='1' style='border-collapse: collapse;'>";
        $mensaje_correo .= "<tr><th>Documento</th><th>ID Grupo</th><th>Usuario</th><th>Curso</th><th>Rol Actual</th>";
        if (!$modo_simulacion) $mensaje_correo .= "<th>Resultado</th>";
        $mensaje_correo .= "</tr>";
        
        $lista = $modo_simulacion ? $cambios_propuestos : $cambios_realizados;
        foreach ($lista as $cambio) {
            $mensaje_correo .= "<tr>";
            $mensaje_correo .= "<td>" . $cambio['documento'] . "</td>";
            $mensaje_correo .= "<td>" . $cambio['idgrupo'] . "</td>";
            $mensaje_correo .= "<td>" . $cambio['nombre_completo'] . "</td>";
            $mensaje_correo .= "<td>" . $cambio['curso_nombre'] . "</td>";
            $mensaje_correo .= "<td>" . $cambio['rol_actual'] . "</td>";
            if (!$modo_simulacion) $mensaje_correo .= "<td>" . $cambio['resultado'] . "</td>";
            $mensaje_correo .= "</tr>";
        }
        $mensaje_correo .= "</table>";
    } else {
        $mensaje_correo .= "<p>No se encontraron cambios para realizar.</p>";
    }
    
    // Enviar correo
    enviarCorreo($destinatario, $asunto, $mensaje_correo);
    
    // Mostrar en consola (para cron)
    echo $mensaje_correo;
    
    // Cerrar conexiones
    oci_close($oracle_conn);
    $moodle_conn = null;
    
} catch (Exception $e) {
    $error_msg = "Error: " . $e->getMessage();
    echo "<p style='color: red;'>$error_msg</p>";
    registrarProcesamiento($archivo_registro, 'SISTEMA', '0', 'ERROR: ' . $e->getMessage());
    enviarCorreo($destinatario, "Error en script de cancelaciones", $error_msg);
}

// Instrucciones para cron
if (php_sapi_name() === 'cli') {
    echo "\n\nInstrucciones para CRON:\n";
    echo "Para ejecutar este script cada 4 horas (por ejemplo), agregar al crontab:\n";
    echo "0 */4 * * * /usr/bin/php " . __FILE__ . " > /var/log/moodle_cancelaciones.log\n";
    echo "\nPara ejecutar en producción (cambios reales), modificar \$modo_simulacion a false\n";
}
?>
