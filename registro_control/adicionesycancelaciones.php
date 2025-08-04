<?php
/**
 * SCRIPT PARA MANEJO DE ADICIONES Y CANCELACIONES EN MOODLE
 * 
 * Modificado para:
 * 1. Control exacto por fecha y hora para evitar reprocesamiento
 * 2. Generar reporte CSV adjunto
 * 3. Mantener resumen textual
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

// Ruta al autoload.php
$autoloadPath = '/root/scripts/informes/vendor/autoload.php';
if (!file_exists($autoloadPath)) {
    die("ERROR: No se encontró autoload.php en $autoloadPath\n");
}
require_once $autoloadPath;

// Cargar configuración
$config = require __DIR__.'/config/env.php';

// Cargar Moodle
$moodleConfigPath = $config['PATHS']['moodle_config'];
if (!file_exists($moodleConfigPath)) {
    die("ERROR: No se encontró config.php de Moodle en $moodleConfigPath\n");
}
require_once $moodleConfigPath;
global $DB, $CFG;

// =============================================================================
// CONSTANTES Y CONFIGURACIÓN
// =============================================================================

define('EMAIL_SOPORTE', 'soporteunivirtual@utp.edu.co');
define('EMAIL_SOPORTE_ADICIONAL', 'univirtual-utp@utp.edu.co');
define('FECHA_INICIO', strtotime('2025-08-04 00:00:00')); // Fecha inicial para procesar registros
define('FROM_EMAIL', 'noreply@utp.edu.co'); // Dirección de correo para el remitente
define('REGISTRO_PROCESADOS_FILE', __DIR__.'/procesados.log'); // Archivo único de registros procesados

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

/**
 * Envía un correo usando la función mail de PHP
 */
function enviarCorreo($to, $subject, $message) {
    $headers = "From: " . FROM_EMAIL . "\r\n";
    $headers .= "MIME-Version: 1.0\r\n";
    $headers .= "Content-Type: text/plain; charset=UTF-8\r\n";
    
    if (mail($to, $subject, $message, $headers)) {
        echo "[INFO] Correo enviado a $to\n";
        return true;
    } else {
        echo "[ERROR] Falló el envío de correo a $to\n";
        return false;
    }
}

/**
 * Envía correo con archivo adjunto
 */
function enviarCorreoConAdjunto($to, $subject, $message, $filePath, $fileName) {
    $boundary = uniqid('np');
    
    $headers = "MIME-Version: 1.0\r\n";
    $headers .= "From: " . FROM_EMAIL . "\r\n";
    $headers .= "Content-Type: multipart/mixed; boundary=$boundary\r\n";
    
    // Mensaje en texto plano
    $body = "--$boundary\r\n";
    $body .= "Content-Type: text/plain; charset=UTF-8\r\n";
    $body .= "Content-Transfer-Encoding: 7bit\r\n\r\n";
    $body .= $message . "\r\n";
    
    // Adjuntar archivo
    $fileContent = file_get_contents($filePath);
    $fileContent = chunk_split(base64_encode($fileContent));
    
    $body .= "--$boundary\r\n";
    $body .= "Content-Type: application/octet-stream; name=\"$fileName\"\r\n";
    $body .= "Content-Transfer-Encoding: base64\r\n";
    $body .= "Content-Disposition: attachment; filename=\"$fileName\"\r\n\r\n";
    $body .= $fileContent . "\r\n";
    $body .= "--$boundary--";
    
    if (mail($to, $subject, $body, $headers)) {
        echo "[INFO] Correo con adjunto enviado a $to\n";
        return true;
    } else {
        echo "[ERROR] Falló el envío de correo con adjunto a $to\n";
        return false;
    }
}

/**
 * Registra los IDs de grupo + usuario ya procesados con fecha y hora exacta
 */
function registrarProcesado($idgrupo, $usuario, $tipo, $fechaOracle) {
    $fechaFormateada = date('Y-m-d H:i:s', strtotime($fechaOracle));
    $linea = "$idgrupo|$usuario|$tipo|$fechaFormateada|".date('Y-m-d H:i:s')."\n";
    file_put_contents(REGISTRO_PROCESADOS_FILE, $linea, FILE_APPEND);
}

/**
 * Verifica si ya fue procesado considerando fecha y hora exacta de Oracle
 */
function yaProcesado($idgrupo, $usuario, $tipo, $fechaOracle) {
    if (!file_exists(REGISTRO_PROCESADOS_FILE)) return false;
    
    $fechaBusqueda = date('Y-m-d H:i:s', strtotime($fechaOracle));
    $contenido = file_get_contents(REGISTRO_PROCESADOS_FILE);
    
    // Buscar coincidencia exacta con fecha de Oracle
    $busqueda = "$idgrupo|$usuario|$tipo|$fechaBusqueda|";
    return strpos($contenido, $busqueda) !== false;
}

/**
 * Limpia registros antiguos (conserva solo los de los últimos 7 días)
 */
function limpiarRegistrosAntiguos() {
    if (!file_exists(REGISTRO_PROCESADOS_FILE)) return;
    
    $lineas = file(REGISTRO_PROCESADOS_FILE, FILE_IGNORE_NEW_LINES | FILE_SKIP_EMPTY_LINES);
    $nuevoContenido = '';
    $limiteTiempo = strtotime('-7 days');
    
    foreach ($lineas as $linea) {
        $partes = explode('|', $linea);
        if (count($partes) >= 5) {
            $fechaRegistro = strtotime($partes[4]);
            if ($fechaRegistro >= $limiteTiempo) {
                $nuevoContenido .= $linea . "\n";
            }
        }
    }
    
    file_put_contents(REGISTRO_PROCESADOS_FILE, $nuevoContenido);
}

// =============================================================================
// FUNCIONES PRINCIPALES
// =============================================================================

/**
 * Conecta a la base de datos Oracle
 */
function conectarOracle() {
    $config = $GLOBALS['config']['ORACLE_DB'];
    $tns = "(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(Host={$config['host']})(Port={$config['port']}))".
           "(CONNECT_DATA=(SID={$config['service_name']})))";
    
    $conn = oci_connect(
        $config['username'], 
        $config['password'], 
        $tns, 
        'AL32UTF8'
    );
    
    if (!$conn) {
        $e = oci_error();
        throw new Exception("Error Oracle: " . $e['message']);
    }
    
    return $conn;
}

/**
 * Obtiene registros a procesar desde Oracle (fecha >= FECHA_INICIO)
 */
function obtenerRegistrosProcesar($conn) {
    $sql = "SELECT * 
            FROM REGISTRO.VI_RYC_NOUNIVIRTUALASIGNCANCEL u
            WHERE (regexp_like(u.PERIODOACADEMICO,'^'||'20252'||'(.)*[Pp][Rr][Ee][Gg][Rr][Aa][Dd][Oo].*'))
            AND u.IDPERIODOACADEMICO <> '21896'
            AND u.FECHACREACION >= TO_DATE('".date('Y-m-d H:i:s', FECHA_INICIO)."', 'YYYY-MM-DD HH24:MI:SS')
            ORDER by u.FECHACREACION DESC";
    
    $stid = oci_parse($conn, $sql);
    oci_execute($stid);
    
    $registros = [];
    while ($row = oci_fetch_array($stid, OCI_ASSOC+OCI_RETURN_NULLS)) {
        $registros[] = $row;
    }
    
    oci_free_statement($stid);
    return $registros;
}

/**
 * Procesa cancelación de asignatura
 */
function procesarCancelacion($registro, $modoPrueba, &$resumen) {
    global $DB;
    
    $idgrupo = $registro['IDGRUPO'];
    $username = $registro['NUMERODOCUMENTO'];
    $fechaOracle = $registro['FECHACREACION'];
    
    // Verificar si ya fue procesado con la misma fecha de Oracle
    if (yaProcesado($idgrupo, $username, TIPO_CANCELACION, $fechaOracle)) {
        echo "[INFO] Cancelación ya procesada anteriormente para $username en grupo $idgrupo con fecha $fechaOracle\n";
        return false;
    }
    
    // Buscar curso en Moodle
    $curso = $DB->get_record('course', ['summary' => $idgrupo], 'id,fullname,shortname');
    if (!$curso) {
        echo "[CANCELACIÓN] Curso no encontrado para IDGRUPO: $idgrupo\n";
        $resumen[] = "[ERROR] Cancelación fallida - Curso no encontrado para IDGRUPO: $idgrupo, Username: $username";
        return false;
    }
    
    // Buscar usuario en Moodle
    $user = $DB->get_record('user', ['username' => $username], 'id,email,firstname,lastname');
    if (!$user) {
        echo "[CANCELACIÓN] Usuario no encontrado: $username\n";
        $resumen[] = "[ERROR] Cancelación fallida - Usuario no encontrado: $username";
        return false;
    }
    
    if ($modoPrueba) {
        echo "[SIMULACIÓN] Se cambiaría rol de $username a DESMATRICULADO en curso {$curso->fullname}\n";
        $resumen[] = "[SIMULACIÓN] Cancelación - Username: $username, Estudiante: {$user->firstname} {$user->lastname} ({$user->email}), Curso: {$curso->fullname} (IDGRUPO: $idgrupo), Fecha Oracle: $fechaOracle";
        return true;
    }
    
    // Cambiar rol de estudiante a desmatriculado
    $context = context_course::instance($curso->id);
    
    // Eliminar asignación de rol de estudiante
    $DB->delete_records('role_assignments', [
        'userid' => $user->id,
        'contextid' => $context->id,
        'roleid' => ROL_ESTUDIANTE
    ]);
    
    // Asignar rol de desmatriculado
    $role_assign = new stdClass();
    $role_assign->roleid = ROL_DESMATRICULADO;
    $role_assign->contextid = $context->id;
    $role_assign->userid = $user->id;
    $role_assign->timemodified = time();
    $role_assign->modifierid = 2; // Usuario admin
    
    $DB->insert_record('role_assignments', $role_assign);
    
    // Registrar como procesado con fecha exacta de Oracle
    registrarProcesado($idgrupo, $username, TIPO_CANCELACION, $fechaOracle);
    
    // Agregar al resumen
    $resumen[] = "Cancelación procesada - Username: $username, Estudiante: {$user->firstname} {$user->lastname} ({$user->email}), Curso: {$curso->fullname} (IDGRUPO: $idgrupo), Fecha Oracle: $fechaOracle, Fecha Proceso: " . date('Y-m-d H:i:s');
    
    // Enviar correos de notificación
    enviarCorreoCancelacionEstudiante($user, $curso);
    enviarCorreoCancelacionDocente($user, $curso);
    
    return true;
}

/**
 * Procesa adición de asignatura
 */
function procesarAdicion($registro, $modoPrueba, &$resumen) {
    global $DB;
    
    $idgrupo = $registro['IDGRUPO'];
    $username = $registro['NUMERODOCUMENTO'];
    $fechaOracle = $registro['FECHACREACION'];
    
    // Verificar si ya fue procesado con la misma fecha de Oracle
    if (yaProcesado($idgrupo, $username, TIPO_ADICION, $fechaOracle)) {
        echo "[INFO] Adición ya procesada anteriormente para $username en grupo $idgrupo con fecha $fechaOracle\n";
        return false;
    }
    
    // Buscar curso en Moodle
    $curso = $DB->get_record('course', ['summary' => $idgrupo], 'id,fullname,shortname');
    if (!$curso) {
        echo "[ADICIÓN] Curso no encontrado para IDGRUPO: $idgrupo\n";
        $resumen[] = "[ERROR] Adición fallida - Curso no encontrado para IDGRUPO: $idgrupo, Username: $username";
        return false;
    }
    
    // Buscar usuario en Moodle
    $user = $DB->get_record('user', ['username' => $username], 'id,email,firstname,lastname');
    if (!$user) {
        echo "[ADICIÓN] Usuario no encontrado: $username - Informando a soporte\n";
        $resumen[] = "[ERROR] Adición fallida - Usuario no encontrado: $username, IDGRUPO: $idgrupo";
        enviarCorreoUsuarioNoExiste($username, $idgrupo);
        return false;
    }
    
    if ($modoPrueba) {
        echo "[SIMULACIÓN] Se matricularía a $username en curso {$curso->fullname}\n";
        $resumen[] = "[SIMULACIÓN] Adición - Username: $username, Estudiante: {$user->firstname} {$user->lastname} ({$user->email}), Curso: {$curso->fullname} (IDGRUPO: $idgrupo), Fecha Oracle: $fechaOracle";
        return true;
    }
    
    // Matricular usuario con rol de estudiante
    $enrol = $DB->get_record('enrol', ['courseid' => $curso->id, 'enrol' => 'manual'], '*', MUST_EXIST);
    
    // Verificar si ya está matriculado
    if ($DB->record_exists('user_enrolments', ['userid' => $user->id, 'enrolid' => $enrol->id])) {
        echo "[ADICIÓN] Usuario $username ya está matriculado en el curso\n";
        $resumen[] = "[INFO] Adición omitida - Usuario $username ya está matriculado en el curso {$curso->fullname} (IDGRUPO: $idgrupo)";
        return false;
    }
    
    // Asignar rol de estudiante
    $context = context_course::instance($curso->id);
    $role_assign = [
        'roleid' => ROL_ESTUDIANTE,
        'contextid' => $context->id,
        'userid' => $user->id,
        'timemodified' => time(),
        'modifierid' => 2
    ];
    $DB->insert_record('role_assignments', (object)$role_assign);
    
    // Crear matrícula
    $user_enrolment = [
        'status' => 0,
        'enrolid' => $enrol->id,
        'userid' => $user->id,
        'timestart' => time(),
        'timecreated' => time(),
        'timemodified' => time()
    ];
    $DB->insert_record('user_enrolments', (object)$user_enrolment);
    
    // Registrar como procesado con fecha exacta de Oracle
    registrarProcesado($idgrupo, $username, TIPO_ADICION, $fechaOracle);
    
    // Agregar al resumen
    $resumen[] = "Adición procesada - Username: $username, Estudiante: {$user->firstname} {$user->lastname} ({$user->email}), Curso: {$curso->fullname} (IDGRUPO: $idgrupo), Fecha Oracle: $fechaOracle, Fecha Proceso: " . date('Y-m-d H:i:s');
    
    // Enviar correo al estudiante
    enviarCorreoAdicionEstudiante($user, $curso);
    
    return true;
}

/**
 * Procesa cambio de grupo
 */
function procesarCambioGrupo($registro, $modoPrueba, &$resumen) {
    $idgrupo = $registro['IDGRUPO'];
    $username = $registro['NUMERODOCUMENTO'];
    $fechaOracle = $registro['FECHACREACION'];
    
    // Verificar si ya fue procesado con la misma fecha de Oracle
    if (yaProcesado($idgrupo, $username, TIPO_CAMBIO_GRUPO, $fechaOracle)) {
        echo "[INFO] Cambio de grupo ya procesado anteriormente para $username en grupo $idgrupo con fecha $fechaOracle\n";
        return false;
    }
    
    if ($modoPrueba) {
        echo "[SIMULACIÓN] Se reportaría cambio de grupo para {$registro['NUMERODOCUMENTO']}\n";
        $resumen[] = "[SIMULACIÓN] Cambio de grupo - Username: $username, Documento: {$registro['NUMERODOCUMENTO']}, Nombre: {$registro['NOMBRES']} {$registro['APELLIDOS']}, ID Grupo: {$registro['IDGRUPO']}, Asignatura: {$registro['NOMBREASIGNATURA']}, Fecha Oracle: $fechaOracle";
    } else {
        // Registrar como procesado con fecha exacta de Oracle
        registrarProcesado($idgrupo, $username, TIPO_CAMBIO_GRUPO, $fechaOracle);
        
        $resumen[] = "Cambio de grupo reportado - Username: $username, Documento: {$registro['NUMERODOCUMENTO']}, Nombre: {$registro['NOMBRES']} {$registro['APELLIDOS']}, ID Grupo: {$registro['IDGRUPO']}, Asignatura: {$registro['NOMBREASIGNATURA']}, Fecha Oracle: $fechaOracle, Fecha Proceso: " . date('Y-m-d H:i:s');
    }
    
    if (!$modoPrueba) {
        enviarCorreoCambioGrupo($registro);
    }
    return true;
}

// ... (las funciones de envío de correos se mantienen igual que en la versión anterior)

/**
 * Envía reporte final de ejecución con CSV adjunto
 */
function enviarReporteFinal($resultados, $modoPrueba, $resumen) {
    $subject = ($modoPrueba ? "[PRUEBA] " : "") . "Reporte de Adiciones/Cancelaciones - " . date('Y-m-d H:i:s');
    
    // Crear contenido del correo (texto plano)
    $message = "Reporte de ejecución " . ($modoPrueba ? "en modo prueba" : "en producción") . "\n\n";
    $message .= "Total registros procesados: {$resultados['total']}\n";
    $message .= "Cancelaciones procesadas: {$resultados['cancelaciones']}\n";
    $message .= "Adiciones procesadas: {$resultados['adiciones']}\n";
    $message .= "Cambios de grupo reportados: {$resultados['cambios_grupo']}\n";
    $message .= "Errores encontrados: {$resultados['errores']}\n\n";
    $message .= "Fecha de inicio del filtro: " . date('Y-m-d H:i:s', FECHA_INICIO) . "\n\n";
    $message .= "Detalles de las operaciones:\n";
    $message .= "----------------------------------------\n";
    $message .= implode("\n", $resumen) . "\n";
    
    // Crear archivo CSV temporal
    $csvFileName = tempnam(sys_get_temp_dir(), 'reporte_') . '.csv';
    $csvFile = fopen($csvFileName, 'w');
    
    // Escribir encabezados CSV
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
    
    // Procesar el resumen para extraer datos para CSV
    foreach ($resumen as $linea) {
        // Procesar operaciones exitosas
        if (preg_match('/(Cancelación|Adición|Cambio de grupo).*Username: (.*?),.*Estudiante: (.*?) (.*?) \((.*?)\).*IDGRUPO: (\d+).*Fecha Oracle: (.*?), Fecha Proceso: (.*)/', $linea, $matches)) {
            fputcsv($csvFile, [
                $matches[1], // Tipo operación
                $matches[2], // Username
                $matches[3].' '.$matches[4], // Nombre completo
                $matches[5], // Email
                $matches[6], // ID Grupo
                '', // Asignatura (se puede extraer si está disponible)
                $matches[7], // Fecha Oracle
                $matches[8], // Fecha Proceso
                'Completado'
            ], ';');
        }
        // Procesar errores
        elseif (preg_match('/\[ERROR\].*Username: (.*?),/', $linea, $matches)) {
            fputcsv($csvFile, [
                'Error',
                $matches[1],
                '',
                '',
                '',
                '',
                '',
                date('Y-m-d H:i:s'),
                'Fallido'
            ], ';');
        }
    }
    
    fclose($csvFile);
    
    // Enviar correo con adjunto
    enviarCorreoConAdjunto(
        EMAIL_SOPORTE, 
        $subject, 
        $message, 
        $csvFileName,
        'reporte_adiciones_cancelaciones_' . date('Y-m-d_His') . '.csv'
    );
    
    enviarCorreoConAdjunto(
        EMAIL_SOPORTE_ADICIONAL, 
        $subject, 
        $message, 
        $csvFileName,
        'reporte_adiciones_cancelaciones_' . date('Y-m-d_His') . '.csv'
    );
    
    // Eliminar archivo temporal
    unlink($csvFileName);
}

// =============================================================================
// EJECUCIÓN PRINCIPAL
// =============================================================================

try {
    // Limpiar registros antiguos (más de 7 días)
    limpiarRegistrosAntiguos();
    
    echo "Iniciando proceso de adiciones y cancelaciones " . ($modoPrueba ? "en MODO PRUEBA" : "") . "...\n";
    echo "Filtrando registros con fecha CREACIÓN >= ".date('Y-m-d H:i:s', FECHA_INICIO)."\n";
    
    // 1. Conectar a Oracle y obtener registros
    $connOracle = conectarOracle();
    $registros = obtenerRegistrosProcesar($connOracle);
    oci_close($connOracle);
    
    echo "Registros a procesar: " . count($registros) . "\n";
    
    $resultados = [
        'total' => 0,
        'cancelaciones' => 0,
        'adiciones' => 0,
        'cambios_grupo' => 0,
        'errores' => 0
    ];
    
    $resumen = [];
    
    // 2. Procesar cada registro
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
                    echo "Tipo de operación no reconocido: {$registro['VALORTIPO']}\n";
                    $resumen[] = "[ERROR] Tipo de operación no reconocido: {$registro['VALORTIPO']}, Username: {$registro['NUMERODOCUMENTO']}";
                    $resultados['errores']++;
            }
        } catch (Exception $e) {
            echo "ERROR procesando registro: " . $e->getMessage() . "\n";
            $resumen[] = "[ERROR] Procesamiento fallido - " . $e->getMessage() . ", Username: {$registro['NUMERODOCUMENTO']}";
            $resultados['errores']++;
        }
    }
    
    // 3. Generar y enviar reporte final
    enviarReporteFinal($resultados, $modoPrueba, $resumen);
    
    // 4. Mostrar resumen
    echo "\n=== RESUMEN FINAL ===\n";
    echo "Total registros procesados: {$resultados['total']}\n";
    echo "Cancelaciones procesadas: {$resultados['cancelaciones']}\n";
    echo "Adiciones procesadas: {$resultados['adiciones']}\n";
    echo "Cambios de grupo reportados: {$resultados['cambios_grupo']}\n";
    echo "Errores encontrados: {$resultados['errores']}\n";

} catch (Exception $e) {
    echo "ERROR: " . $e->getMessage() . "\n";
    exit(1);
}

exit(0);
?>