<?php
/**
 * SCRIPT PARA MANEJO DE ADICIONES Y CANCELACIONES EN MOODLE
 * 
 * Uso normal: php adicionesycancelaciones.php
 * Modo prueba: php adicionesycancelaciones.php --dry-run
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
define('EMAIL_UNIVIRTUAL', 'univirtual-utp@utp.edu.co');
define('FECHA_INICIO', strtotime('2025-08-04 00:00:00')); // Fecha inicial para procesar registros

// Tipos de operación
define('TIPO_CANCELACION', 'CANCELACIÓN');
define('TIPO_ADICION', 'ADICIÓN');
define('TIPO_CAMBIO_GRUPO', 'CAMBIO DE GRUPO');

// Roles de Moodle
define('ROL_ESTUDIANTE', 5);
define('ROL_DESMATRICULADO', 9);

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
function procesarCancelacion($registro, $modoPrueba) {
    global $DB;
    
    $idgrupo = $registro['IDGRUPO'];
    $username = $registro['NUMERODOCUMENTO'];
    
    // Buscar curso en Moodle
    $curso = $DB->get_record('course', ['summary' => $idgrupo], 'id,fullname,shortname');
    if (!$curso) {
        echo "[CANCELACIÓN] Curso no encontrado para IDGRUPO: $idgrupo\n";
        return false;
    }
    
    // Buscar usuario en Moodle
    $user = $DB->get_record('user', ['username' => $username], 'id,email,firstname,lastname');
    if (!$user) {
        echo "[CANCELACIÓN] Usuario no encontrado: $username\n";
        return false;
    }
    
    if ($modoPrueba) {
        echo "[SIMULACIÓN] Se cambiaría rol de $username a DESMATRICULADO en curso {$curso->fullname}\n";
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
    
    // Enviar correos de notificación
    enviarCorreoCancelacionEstudiante($user, $curso);
    enviarCorreoCancelacionDocente($user, $curso);
    enviarCorreoNotificacionSoporte($user, $curso, 'cancelación');
    
    return true;
}

/**
 * Procesa adición de asignatura
 */
function procesarAdicion($registro, $modoPrueba) {
    global $DB;
    
    $idgrupo = $registro['IDGRUPO'];
    $username = $registro['NUMERODOCUMENTO'];
    
    // Buscar curso en Moodle
    $curso = $DB->get_record('course', ['summary' => $idgrupo], 'id,fullname,shortname');
    if (!$curso) {
        echo "[ADICIÓN] Curso no encontrado para IDGRUPO: $idgrupo\n";
        return false;
    }
    
    // Buscar usuario en Moodle
    $user = $DB->get_record('user', ['username' => $username], 'id,email,firstname,lastname');
    if (!$user) {
        echo "[ADICIÓN] Usuario no encontrado: $username - Informando a soporte\n";
        enviarCorreoUsuarioNoExiste($username, $idgrupo);
        return false;
    }
    
    if ($modoPrueba) {
        echo "[SIMULACIÓN] Se matricularía a $username en curso {$curso->fullname}\n";
        return true;
    }
    
    // Matricular usuario con rol de estudiante
    $enrol = $DB->get_record('enrol', ['courseid' => $curso->id, 'enrol' => 'manual'], '*', MUST_EXIST);
    
    // Verificar si ya está matriculado
    if ($DB->record_exists('user_enrolments', ['userid' => $user->id, 'enrolid' => $enrol->id])) {
        echo "[ADICIÓN] Usuario $username ya está matriculado en el curso\n";
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
    
    // Enviar correo al estudiante
    enviarCorreoAdicionEstudiante($user, $curso);
    enviarCorreoNotificacionSoporte($user, $curso, 'adición');
    
    return true;
}

/**
 * Procesa cambio de grupo
 */
function procesarCambioGrupo($registro, $modoPrueba) {
    if ($modoPrueba) {
        echo "[SIMULACIÓN] Se reportaría cambio de grupo para {$registro['NUMERODOCUMENTO']}\n";
    }
    
    enviarCorreoCambioGrupo($registro);
    return true;
}

/**
 * Envía correo al estudiante cuando se cancela su matrícula
 */
function enviarCorreoCancelacionEstudiante($user, $curso) {
    global $CFG;
    
    $subject = "Cancelación de asignatura - {$curso->fullname}";
    $message = "Estimado/a {$user->firstname} {$user->lastname},\n\n";
    $message .= "Te informamos que tu matrícula en la asignatura {$curso->fullname} ha sido cancelada.\n\n";
    $message .= "Si tienes dudas comunícate a través de WhatsApp: <a href=\"https://api.whatsapp.com/send/?phone=3203921622&text&type=phone_number&app_absent=0\" target=\"_blank\">3203921622</a>\n\n";
    $message .= "Atentamente,\n";
    $message .= "Univirtual UTP";
    
    email_to_user($user, $CFG->noreplyaddress, $subject, $message);
}

/**
 * Envía correo al docente cuando se cancela una matrícula
 */
function enviarCorreoCancelacionDocente($user, $curso) {
    global $DB, $CFG;
    
    // Obtener todos los profesores del curso
    $profesores = $DB->get_records_sql("
        SELECT u.* 
        FROM {user} u
        JOIN {role_assignments} ra ON ra.userid = u.id
        JOIN {context} ctx ON ctx.id = ra.contextid AND ctx.contextlevel = 50
        JOIN {course} c ON c.id = ctx.instanceid
        WHERE c.id = ? AND ra.roleid = 3", [$curso->id]);
    
    $subject = "Estudiante cancelado - {$curso->fullname}";
    $message = "Estimado/a docente,\n\n";
    $message .= "El estudiante {$user->firstname} {$user->lastname} ({$user->email}) ";
    $message .= "ha cancelado la asignatura {$curso->fullname}.\n\n";
    $message .= "Atentamente,\n";
    $message .= "Univirtual UTP";
    
    foreach ($profesores as $profesor) {
        email_to_user($profesor, $CFG->noreplyaddress, $subject, $message);
    }
}

/**
 * Envía correo de notificación a soporte y univirtual sobre cambios realizados
 */
function enviarCorreoNotificacionSoporte($user, $curso, $tipoOperacion) {
    global $CFG;
    
    $subject = "Notificación de $tipoOperacion - {$user->username}";
    $message = "Se ha realizado una $tipoOperacion en el sistema:\n\n";
    $message .= "Estudiante: {$user->firstname} {$user->lastname}\n";
    $message .= "Documento: {$user->username}\n";
    $message .= "Email: {$user->email}\n";
    $message .= "Curso: {$curso->fullname}\n";
    $message .= "ID Curso: {$curso->id}\n";
    $message .= "ID Grupo: {$curso->summary}\n";
    $message .= "Fecha: ".date('Y-m-d H:i:s')."\n\n";
    $message .= "Este correo es informativo, no requiere respuesta.";
    
    // Enviar a soporteunivirtual@utp.edu.co
    $userSoporte = (object)['email' => EMAIL_SOPORTE];
    email_to_user($userSoporte, $CFG->noreplyaddress, $subject, $message);
    
    // Enviar a univirtual-utp@utp.edu.co
    $userUnivirtual = (object)['email' => EMAIL_UNIVIRTUAL];
    email_to_user($userUnivirtual, $CFG->noreplyaddress, $subject, $message);
}

/**
 * Envía correo cuando un usuario no existe en Moodle
 */
function enviarCorreoUsuarioNoExiste($username, $idgrupo) {
    global $CFG;
    
    $subject = "[URGENTE] Usuario no existe en Moodle";
    $message = "Se intentó matricular al usuario $username en el curso con IDGRUPO $idgrupo ";
    $message .= "pero no existe en Moodle.\n\n";
    $message .= "Por favor crear el usuario manualmente.\n\n";
    $message .= "Fecha: ".date('Y-m-d H:i:s')."\n";
    
    // Enviar a ambos correos
    $userSoporte = (object)['email' => EMAIL_SOPORTE];
    $userUnivirtual = (object)['email' => EMAIL_UNIVIRTUAL];
    
    email_to_user($userSoporte, $CFG->noreplyaddress, $subject, $message);
    email_to_user($userUnivirtual, $CFG->noreplyaddress, $subject, $message);
}

/**
 * Envía correo al estudiante cuando se añade su matrícula
 */
function enviarCorreoAdicionEstudiante($user, $curso) {
    global $CFG;
    
    $subject = "Matrícula en asignatura - {$curso->fullname}";
    $message = "Estimado/a {$user->firstname} {$user->lastname},\n\n";
    $message .= "Ha sido matriculado/a en la asignatura {$curso->fullname}.\n\n";
    $message .= "Para acceder al curso, ingrese al campus virtual con su número de documento ";
    $message .= "y su contraseña.\n\n";
    $message .= "Atentamente,\n";
    $message .= "Univirtual UTP";
    
    email_to_user($user, $CFG->noreplyaddress, $subject, $message);
}

/**
 * Envía correo sobre cambio de grupo a soporte
 */
function enviarCorreoCambioGrupo($registro) {
    global $CFG;
    
    $subject = "Cambio de grupo requerido - {$registro['NUMERODOCUMENTO']}";
    $message = "Se requiere cambio de grupo para el estudiante:\n\n";
    $message .= "Documento: {$registro['NUMERODOCUMENTO']}\n";
    $message .= "Nombre: {$registro['NOMBRES']} {$registro['APELLIDOS']}\n";
    $message .= "ID Grupo actual: {$registro['IDGRUPO']}\n";
    $message .= "Asignatura: {$registro['NOMBREASIGNATURA']}\n";
    $message .= "Fecha registro: {$registro['FECHACREACION']}\n\n";
    $message .= "Este cambio debe realizarse manualmente en Moodle.";
    
    // Enviar a ambos correos
    $userSoporte = (object)['email' => EMAIL_SOPORTE];
    $userUnivirtual = (object)['email' => EMAIL_UNIVIRTUAL];
    
    email_to_user($userSoporte, $CFG->noreplyaddress, $subject, $message);
    email_to_user($userUnivirtual, $CFG->noreplyaddress, $subject, $message);
}

/**
 * Envía reporte final de ejecución
 */
function enviarReporteFinal($resultados, $modoPrueba) {
    global $CFG;
    
    $subject = ($modoPrueba ? "[PRUEBA] " : "") . "Reporte de Adiciones/Cancelaciones - " . date('Y-m-d H:i:s');
    
    $message = "Reporte de ejecución " . ($modoPrueba ? "en modo prueba" : "en producción") . "\n\n";
    $message .= "Total registros procesados: {$resultados['total']}\n";
    $message .= "Cancelaciones procesadas: {$resultados['cancelaciones']}\n";
    $message .= "Adiciones procesadas: {$resultados['adiciones']}\n";
    $message .= "Cambios de grupo reportados: {$resultados['cambios_grupo']}\n";
    $message .= "Errores encontrados: {$resultados['errores']}\n\n";
    $message .= "Fecha de inicio del filtro: ".date('Y-m-d H:i:s', FECHA_INICIO)."\n";
    
    // Enviar a ambos correos
    $userSoporte = (object)['email' => EMAIL_SOPORTE];
    $userUnivirtual = (object)['email' => EMAIL_UNIVIRTUAL];
    
    email_to_user($userSoporte, $CFG->noreplyaddress, $subject, $message);
    email_to_user($userUnivirtual, $CFG->noreplyaddress, $subject, $message);
}

// =============================================================================
// EJECUCIÓN PRINCIPAL
// =============================================================================

try {
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
    
    // 2. Procesar cada registro
    foreach ($registros as $registro) {
        $resultados['total']++;
        
        try {
            switch ($registro['VALORTIPO']) {
                case TIPO_CANCELACION:
                    if (procesarCancelacion($registro, $modoPrueba)) {
                        $resultados['cancelaciones']++;
                    }
                    break;
                    
                case TIPO_ADICION:
                    if (procesarAdicion($registro, $modoPrueba)) {
                        $resultados['adiciones']++;
                    }
                    break;
                    
                case TIPO_CAMBIO_GRUPO:
                    procesarCambioGrupo($registro, $modoPrueba);
                    $resultados['cambios_grupo']++;
                    break;
                    
                default:
                    echo "Tipo de operación no reconocido: {$registro['VALORTIPO']}\n";
                    $resultados['errores']++;
            }
        } catch (Exception $e) {
            echo "ERROR procesando registro: " . $e->getMessage() . "\n";
            $resultados['errores']++;
        }
    }
    
    // 3. Generar y enviar reporte final
    enviarReporteFinal($resultados, $modoPrueba);
    
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