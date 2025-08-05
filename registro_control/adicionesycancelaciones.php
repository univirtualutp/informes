<?php
/**
 * SCRIPT PARA MANEJO DE ADICIONES Y CANCELACIONES EN MOODLE
 * 
 * Versión completa corregida:
 * 1. Incluye nombre de asignatura en CSV
 * 2. Ejecución horaria compatible
 * 3. Control por fecha/hora exacta
 */

// =============================================================================
// CONFIGURACIÓN INICIAL
// =============================================================================

define('CLI_SCRIPT', true);
error_reporting(E_ALL);
ini_set('display_errors', 1);
set_time_limit(0);

// Verificar modo prueba
$modoPrueba = in_array('--dry-run', $GLOBALS['argv']);
if ($modoPrueba) {
    echo "=== MODO PRUEBA ACTIVADO ===\n";
    echo "Validando datos SIN realizar cambios reales\n\n";
}

// =============================================================================
// CONFIGURACIÓN DE RUTAS
// =============================================================================

require_once '/root/scripts/informes/vendor/autoload.php';
$config = require __DIR__.'/config/env.php';
require_once $config['PATHS']['moodle_config'];
global $DB, $CFG;

// =============================================================================
// CONSTANTES
// =============================================================================

define('EMAIL_SOPORTE', 'soporteunivirtual@utp.edu.co');
define('EMAIL_SOPORTE_ADICIONAL', 'univirtual-utp@utp.edu.co');
define('FECHA_INICIO', strtotime('2025-08-04 00:00:00'));
define('FROM_EMAIL', 'noreply@utp.edu.co');
define('REGISTRO_PROCESADOS_FILE', __DIR__.'/procesados.log');

// Tipos de operación
define('TIPO_CANCELACION', 'CANCELACIÓN');
define('TIPO_ADICION', 'ADICIÓN');
define('TIPO_CAMBIO_GRUPO', 'CAMBIO DE GRUPO');

// Roles de Moodle
define('ROL_ESTUDIANTE', 5);
define('ROL_DESMATRICULADO', 9);

// =============================================================================
// FUNCIONES AUXILIARES
// =============================================================================

function enviarCorreo($to, $subject, $message) {
    $headers = "From: ".FROM_EMAIL."\r\n".
               "MIME-Version: 1.0\r\n".
               "Content-Type: text/plain; charset=UTF-8\r\n";
    
    return mail($to, $subject, $message, $headers);
}

function enviarCorreoConAdjunto($to, $subject, $message, $filePath, $fileName) {
    $boundary = uniqid('np');
    $headers = "From: ".FROM_EMAIL."\r\n".
               "MIME-Version: 1.0\r\n".
               "Content-Type: multipart/mixed; boundary=$boundary\r\n";
    
    $body = "--$boundary\r\n".
            "Content-Type: text/plain; charset=UTF-8\r\n".
            "Content-Transfer-Encoding: 7bit\r\n\r\n".
            $message."\r\n";
    
    $fileContent = chunk_split(base64_encode(file_get_contents($filePath)));
    $body .= "--$boundary\r\n".
             "Content-Type: application/octet-stream; name=\"$fileName\"\r\n".
             "Content-Transfer-Encoding: base64\r\n".
             "Content-Disposition: attachment; filename=\"$fileName\"\r\n\r\n".
             $fileContent."\r\n--$boundary--";
    
    return mail($to, $subject, $body, $headers);
}

function registrarProcesado($idgrupo, $usuario, $tipo, $fechaOracle) {
    $linea = sprintf("%s|%s|%s|%s|%s\n",
        $idgrupo,
        $usuario,
        $tipo,
        date('Y-m-d H:i:s', strtotime($fechaOracle)),
        date('Y-m-d H:i:s')
    );
    file_put_contents(REGISTRO_PROCESADOS_FILE, $linea, FILE_APPEND);
}

function yaProcesado($idgrupo, $usuario, $tipo, $fechaOracle) {
    if (!file_exists(REGISTRO_PROCESADOS_FILE)) return false;
    
    $busqueda = sprintf("%s|%s|%s|%s|",
        $idgrupo,
        $usuario,
        $tipo,
        date('Y-m-d H:i:s', strtotime($fechaOracle))
    );
    
    return strpos(file_get_contents(REGISTRO_PROCESADOS_FILE), $busqueda) !== false;
}

function limpiarRegistrosAntiguos() {
    if (!file_exists(REGISTRO_PROCESADOS_FILE)) return;
    
    $limite = strtotime('-7 days');
    $nuevoContenido = '';
    
    foreach (file(REGISTRO_PROCESADOS_FILE, FILE_IGNORE_NEW_LINES) as $linea) {
        $partes = explode('|', $linea);
        if (count($partes) >= 5 && strtotime($partes[4]) >= $limite) {
            $nuevoContenido .= $linea."\n";
        }
    }
    
    file_put_contents(REGISTRO_PROCESADOS_FILE, $nuevoContenido);
}

// =============================================================================
// FUNCIONES DE CORREO
// =============================================================================

function enviarCorreoCancelacionEstudiante($user, $curso) {
    $subject = "Cancelación de asignatura - {$curso->fullname}";
    $message = "Estimado/a {$user->firstname} {$user->lastname},\n\n".
               "Te informamos que tu matrícula en {$curso->fullname} ha sido cancelada.\n\n".
               "Si es un error, contacta a soporte.\n\n".
               "Atentamente,\nUnivirtual UTP";
    
    enviarCorreo($user->email, $subject, $message);
}

function enviarCorreoCancelacionDocente($user, $curso) {
    global $DB;
    
    $profesores = $DB->get_records_sql("
        SELECT u.* FROM {user} u
        JOIN {role_assignments} ra ON ra.userid = u.id
        JOIN {context} ctx ON ctx.id = ra.contextid
        WHERE ctx.instanceid = ? AND ra.roleid = 3", [$curso->id]);
    
    $subject = "Estudiante cancelado - {$curso->fullname}";
    $message = "El estudiante {$user->firstname} {$user->lastname} ha sido cancelado.\n\n".
               "Atentamente,\nUnivirtual UTP";
    
    foreach ($profesores as $profesor) {
        enviarCorreo($profesor->email, $subject, $message);
    }
}

function enviarCorreoAdicionEstudiante($user, $curso) {
    $subject = "Matrícula en {$curso->fullname}";
    $message = "Estimado/a {$user->firstname} {$user->lastname},\n\n".
               "Has sido matriculado/a en {$curso->fullname}.\n\n".
               "Accede al campus virtual con tu documento.\n\n".
               "Atentamente,\nUnivirtual UTP";
    
    enviarCorreo($user->email, $subject, $message);
}

function enviarCorreoCambioGrupo($registro) {
    $subject = "Cambio de grupo requerido - {$registro['NUMERODOCUMENTO']}";
    $message = "Se requiere cambio de grupo para:\n\n".
               "Estudiante: {$registro['NOMBRES']} {$registro['APELLIDOS']}\n".
               "Documento: {$registro['NUMERODOCUMENTO']}\n".
               "Grupo: {$registro['IDGRUPO']}\n".
               "Asignatura: {$registro['NOMBREASIGNATURA']}\n".
               "Fecha: {$registro['FECHACREACION']}\n\n".
               "Acción requerida: Cambio manual en Moodle";
    
    enviarCorreo(EMAIL_SOPORTE, $subject, $message);
    enviarCorreo(EMAIL_SOPORTE_ADICIONAL, $subject, $message);
}

// =============================================================================
// FUNCIONES PRINCIPALES
// =============================================================================

function conectarOracle() {
    $config = $GLOBALS['config']['ORACLE_DB'];
    $tns = "(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(Host={$config['host']})(Port={$config['port']}))".
           "(CONNECT_DATA=(SID={$config['service_name']})))";
    
    $conn = oci_connect($config['username'], $config['password'], $tns, 'AL32UTF8');
    if (!$conn) throw new Exception("Error Oracle: ".oci_error()['message']);
    return $conn;
}

function obtenerRegistrosProcesar($conn) {
    $sql = "SELECT * FROM REGISTRO.VI_RYC_NOUNIVIRTUALASIGNCANCEL 
            WHERE regexp_like(PERIODOACADEMICO,'^20252(.)*[Pp][Rr][Ee][Gg][Rr][Aa][Dd][Oo].*')
            AND IDPERIODOACADEMICO <> '21896'
            AND FECHACREACION >= TO_DATE('".date('Y-m-d H:i:s', FECHA_INICIO)."', 'YYYY-MM-DD HH24:MI:SS')
            ORDER BY FECHACREACION DESC";
    
    $stid = oci_parse($conn, $sql);
    oci_execute($stid);
    
    $registros = [];
    while ($row = oci_fetch_array($stid, OCI_ASSOC+OCI_RETURN_NULLS)) {
        $registros[] = $row;
    }
    
    oci_free_statement($stid);
    return $registros;
}

function procesarCancelacion($registro, $modoPrueba, &$resumen) {
    global $DB;
    
    $idgrupo = $registro['IDGRUPO'];
    $username = $registro['NUMERODOCUMENTO'];
    $fechaOracle = $registro['FECHACREACION'];
    
    if (yaProcesado($idgrupo, $username, TIPO_CANCELACION, $fechaOracle)) {
        echo "[INFO] Cancelación ya procesada para $username en $idgrupo\n";
        return false;
    }
    
    $curso = $DB->get_record('course', ['summary' => $idgrupo], 'id,fullname,shortname');
    if (!$curso) {
        $resumen[] = "[ERROR] Curso no encontrado - IDGRUPO: $idgrupo, Username: $username";
        return false;
    }
    
    $user = $DB->get_record('user', ['username' => $username], 'id,email,firstname,lastname');
    if (!$user) {
        $resumen[] = "[ERROR] Usuario no encontrado - Username: $username";
        return false;
    }
    
    if ($modoPrueba) {
        $resumen[] = "[SIMULACIÓN] Cancelación - Username: $username, Estudiante: {$user->firstname} {$user->lastname}, Curso: {$curso->fullname}, Asignatura: {$curso->fullname}, IDGRUPO: $idgrupo, Fecha Oracle: $fechaOracle";
        return true;
    }
    
    $context = context_course::instance($curso->id);
    $DB->delete_records('role_assignments', [
        'userid' => $user->id,
        'contextid' => $context->id,
        'roleid' => ROL_ESTUDIANTE
    ]);
    
    $role_assign = (object)[
        'roleid' => ROL_DESMATRICULADO,
        'contextid' => $context->id,
        'userid' => $user->id,
        'timemodified' => time(),
        'modifierid' => 2
    ];
    $DB->insert_record('role_assignments', $role_assign);
    
    registrarProcesado($idgrupo, $username, TIPO_CANCELACION, $fechaOracle);
    $resumen[] = "Cancelación procesada - Username: $username, Estudiante: {$user->firstname} {$user->lastname}, Curso: {$curso->fullname}, Asignatura: {$curso->fullname}, IDGRUPO: $idgrupo, Fecha Oracle: $fechaOracle, Fecha Proceso: ".date('Y-m-d H:i:s');
    
    enviarCorreoCancelacionEstudiante($user, $curso);
    enviarCorreoCancelacionDocente($user, $curso);
    
    return true;
}

function procesarAdicion($registro, $modoPrueba, &$resumen) {
    global $DB;
    
    $idgrupo = $registro['IDGRUPO'];
    $username = $registro['NUMERODOCUMENTO'];
    $fechaOracle = $registro['FECHACREACION'];
    
    if (yaProcesado($idgrupo, $username, TIPO_ADICION, $fechaOracle)) {
        echo "[INFO] Adición ya procesada para $username en $idgrupo\n";
        return false;
    }
    
    $curso = $DB->get_record('course', ['summary' => $idgrupo], 'id,fullname,shortname');
    if (!$curso) {
        $resumen[] = "[ERROR] Curso no encontrado - IDGRUPO: $idgrupo, Username: $username";
        return false;
    }
    
    $user = $DB->get_record('user', ['username' => $username], 'id,email,firstname,lastname');
    if (!$user) {
        $resumen[] = "[ERROR] Usuario no encontrado - Username: $username, IDGRUPO: $idgrupo";
        enviarCorreoUsuarioNoExiste($username, $idgrupo);
        return false;
    }
    
    if ($modoPrueba) {
        $resumen[] = "[SIMULACIÓN] Adición - Username: $username, Estudiante: {$user->firstname} {$user->lastname}, Curso: {$curso->fullname}, Asignatura: {$curso->fullname}, IDGRUPO: $idgrupo, Fecha Oracle: $fechaOracle";
        return true;
    }
    
    $enrol = $DB->get_record('enrol', ['courseid' => $curso->id, 'enrol' => 'manual'], '*', MUST_EXIST);
    
    if ($DB->record_exists('user_enrolments', ['userid' => $user->id, 'enrolid' => $enrol->id])) {
        $resumen[] = "[INFO] Usuario ya matriculado - Username: $username, Curso: {$curso->fullname}";
        return false;
    }
    
    $context = context_course::instance($curso->id);
    $DB->insert_record('role_assignments', (object)[
        'roleid' => ROL_ESTUDIANTE,
        'contextid' => $context->id,
        'userid' => $user->id,
        'timemodified' => time(),
        'modifierid' => 2
    ]);
    
    $DB->insert_record('user_enrolments', (object)[
        'status' => 0,
        'enrolid' => $enrol->id,
        'userid' => $user->id,
        'timestart' => time(),
        'timecreated' => time(),
        'timemodified' => time()
    ]);
    
    registrarProcesado($idgrupo, $username, TIPO_ADICION, $fechaOracle);
    $resumen[] = "Adición procesada - Username: $username, Estudiante: {$user->firstname} {$user->lastname}, Curso: {$curso->fullname}, Asignatura: {$curso->fullname}, IDGRUPO: $idgrupo, Fecha Oracle: $fechaOracle, Fecha Proceso: ".date('Y-m-d H:i:s');
    
    enviarCorreoAdicionEstudiante($user, $curso);
    return true;
}

function procesarCambioGrupo($registro, $modoPrueba, &$resumen) {
    $idgrupo = $registro['IDGRUPO'];
    $username = $registro['NUMERODOCUMENTO'];
    $fechaOracle = $registro['FECHACREACION'];
    $asignatura = $registro['NOMBREASIGNATURA'];
    
    if (yaProcesado($idgrupo, $username, TIPO_CAMBIO_GRUPO, $fechaOracle)) {
        echo "[INFO] Cambio de grupo ya procesado para $username\n";
        return false;
    }
    
    if ($modoPrueba) {
        $resumen[] = "[SIMULACIÓN] Cambio de grupo - Username: $username, Documento: {$registro['NUMERODOCUMENTO']}, Nombre: {$registro['NOMBRES']} {$registro['APELLIDOS']}, ID Grupo: $idgrupo, Asignatura: $asignatura, Fecha Oracle: $fechaOracle";
        return true;
    }
    
    registrarProcesado($idgrupo, $username, TIPO_CAMBIO_GRUPO, $fechaOracle);
    $resumen[] = "Cambio de grupo reportado - Username: $username, Documento: {$registro['NUMERODOCUMENTO']}, Nombre: {$registro['NOMBRES']} {$registro['APELLIDOS']}, ID Grupo: $idgrupo, Asignatura: $asignatura, Fecha Oracle: $fechaOracle, Fecha Proceso: ".date('Y-m-d H:i:s');
    
    enviarCorreoCambioGrupo($registro);
    return true;
}

function enviarReporteFinal($resultados, $modoPrueba, $resumen) {
    global $DB;
    
    $subject = ($modoPrueba ? "[PRUEBA] " : "")."Reporte de Adiciones/Cancelaciones - ".date('Y-m-d H:i:s');
    $message = "Reporte de ejecución ".($modoPrueba ? "en modo prueba" : "en producción")."\n\n".
               "Total registros: {$resultados['total']}\n".
               "Cancelaciones: {$resultados['cancelaciones']}\n".
               "Adiciones: {$resultados['adiciones']}\n".
               "Cambios de grupo: {$resultados['cambios_grupo']}\n".
               "Errores: {$resultados['errores']}\n\n".
               "Detalles:\n".implode("\n", $resumen);
    
    $csvFileName = tempnam(sys_get_temp_dir(), 'reporte_').'.csv';
    $csvFile = fopen($csvFileName, 'w');
    
    // Encabezados CSV
    fputcsv($csvFile, [
        'Tipo Operación',
        'Username',
        'Nombre Completo',
        'Email',
        'ID Grupo',
        'Asignatura',
        'Fecha Oracle',
        'Fecha Proceso',
        'Estado'
    ], ';');
    
    // Procesar cada línea del resumen
    foreach ($resumen as $linea) {
        // Procesar Cancelaciones/Adiciones
        if (preg_match('/(Cancelación|Adición) procesada - Username: (.*?), Estudiante: (.*?), Curso: (.*?), Asignatura: (.*?), IDGRUPO: (.*?), Fecha Oracle: (.*?), Fecha Proceso: (.*)/', $linea, $matches)) {
            fputcsv($csvFile, [
                $matches[1],
                $matches[2],
                $matches[3],
                '', // Email se obtendrá después
                $matches[6],
                $matches[5], // Asignatura
                $matches[7],
                $matches[8],
                'Completado'
            ], ';');
        }
        // Procesar Cambios de Grupo
        elseif (preg_match('/Cambio de grupo reportado - Username: (.*?), Documento: (.*?), Nombre: (.*?), ID Grupo: (.*?), Asignatura: (.*?), Fecha Oracle: (.*?), Fecha Proceso: (.*)/', $linea, $matches)) {
            fputcsv($csvFile, [
                'Cambio de grupo',
                trim($matches[1]),
                trim($matches[3]),
                '',
                trim($matches[4]),
                trim($matches[5]),
                trim($matches[6]),
                trim($matches[7]),
                'Reportado'
            ], ';');
        }
        // Procesar Errores
        elseif (preg_match('/\[ERROR\] (.*?) - (.*)/', $linea, $matches)) {
            fputcsv($csvFile, [
                'Error',
                '',
                '',
                '',
                '',
                '',
                '',
                date('Y-m-d H:i:s'),
                $matches[1].' - '.$matches[2]
            ], ';');
        }
    }
    
    fclose($csvFile);
    
    enviarCorreoConAdjunto(EMAIL_SOPORTE, $subject, $message, $csvFileName, 'reporte_'.date('Ymd_His').'.csv');
    enviarCorreoConAdjunto(EMAIL_SOPORTE_ADICIONAL, $subject, $message, $csvFileName, 'reporte_'.date('Ymd_His').'.csv');
    
    unlink($csvFileName);
}

// =============================================================================
// EJECUCIÓN PRINCIPAL
// =============================================================================

try {
    limpiarRegistrosAntiguos();
    
    echo "Iniciando proceso...\n";
    echo "Filtrando registros desde ".date('Y-m-d H:i:s', FECHA_INICIO)."\n";
    
    $connOracle = conectarOracle();
    $registros = obtenerRegistrosProcesar($connOracle);
    oci_close($connOracle);
    
    echo "Registros a procesar: ".count($registros)."\n";
    
    $resultados = [
        'total' => 0,
        'cancelaciones' => 0,
        'adiciones' => 0,
        'cambios_grupo' => 0,
        'errores' => 0
    ];
    
    $resumen = [];
    
    foreach ($registros as $registro) {
        $resultados['total']++;
        
        try {
            switch ($registro['VALORTIPO']) {
                case TIPO_CANCELACION:
                    if (procesarCancelacion($registro, $modoPrueba, $resumen)) {
                        $resultados['cancelaciones']++;
                    }
                    break;
                    
                case TIPO_ADICION:
                    if (procesarAdicion($registro, $modoPrueba, $resumen)) {
                        $resultados['adiciones']++;
                    }
                    break;
                    
                case TIPO_CAMBIO_GRUPO:
                    if (procesarCambioGrupo($registro, $modoPrueba, $resumen)) {
                        $resultados['cambios_grupo']++;
                    }
                    break;
                    
                default:
                    $resumen[] = "[ERROR] Tipo no reconocido: {$registro['VALORTIPO']}, Username: {$registro['NUMERODOCUMENTO']}";
                    $resultados['errores']++;
            }
        } catch (Exception $e) {
            $resumen[] = "[ERROR] Procesamiento fallido: ".$e->getMessage().", Username: {$registro['NUMERODOCUMENTO']}";
            $resultados['errores']++;
        }
    }
    
    enviarReporteFinal($resultados, $modoPrueba, $resumen);
    
    echo "\nRESUMEN FINAL:\n";
    echo "Total: {$resultados['total']}\n";
    echo "Cancelaciones: {$resultados['cancelaciones']}\n";
    echo "Adiciones: {$resultados['adiciones']}\n";
    echo "Cambios grupo: {$resultados['cambios_grupo']}\n";
    echo "Errores: {$resultados['errores']}\n";

} catch (Exception $e) {
    echo "ERROR: ".$e->getMessage()."\n";
    exit(1);
}

exit(0);
?>