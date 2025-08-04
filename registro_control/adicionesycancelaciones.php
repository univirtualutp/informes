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

// Configuración de correo adicional
$CFG->smtphosts = 'mail.utp.edu.co'; // Asegurar que usa el servidor correcto
$CFG->smtpsecure = 'tls';
$CFG->smtpauthtype = 'LOGIN';
$CFG->smtpuser = 'tu_correo@utp.edu.co'; // Reemplazar con credenciales válidas
$CFG->smtppass = 'tu_contraseña';
$CFG->smtpmaxbulk = 1; // Enviar correos de uno en uno

// =============================================================================
// FUNCIONES PRINCIPALES
// =============================================================================

// [Las funciones conectarOracle() y obtenerRegistrosProcesar() permanecen igual...]

/**
 * Función mejorada para enviar correos con logs detallados
 */
function enviarCorreoMejorado($destinatarios, $subject, $message, $modoPrueba = false) {
    global $CFG;
    
    if (!is_array($destinatarios)) {
        $destinatarios = [$destinatarios];
    }
    
    $resultados = [];
    $from = $CFG->noreplyaddress;
    
    foreach ($destinatarios as $destinatario) {
        $emailuser = new stdClass();
        $emailuser->email = $destinatario;
        
        if ($modoPrueba) {
            echo "[SIMULACIÓN] Correo a enviar:\n";
            echo "Para: $destinatario\n";
            echo "Asunto: $subject\n";
            echo "Mensaje: $message\n\n";
            $resultados[$destinatario] = true;
            continue;
        }
        
        try {
            // Registrar intento de envío
            error_log("Intentando enviar correo a: $destinatario - Asunto: $subject");
            
            // Usar la función de Moodle para enviar
            $result = email_to_user($emailuser, $from, $subject, $message);
            
            if ($result) {
                error_log("Correo enviado exitosamente a $destinatario");
                $resultados[$destinatario] = true;
            } else {
                error_log("Fallo al enviar correo a $destinatario (email_to_user devolvió false)");
                $resultados[$destinatario] = false;
                
                // Intentar método alternativo
                $resultados[$destinatario] = enviarCorreoAlternativo($destinatario, $subject, $message);
            }
        } catch (Exception $e) {
            error_log("Error al enviar correo a $destinatario: " . $e->getMessage());
            $resultados[$destinatario] = false;
        }
    }
    
    return $resultados;
}

/**
 * Método alternativo para enviar correos si falla email_to_user
 */
function enviarCorreoAlternativo($destinatario, $subject, $message) {
    global $CFG;
    
    try {
        $headers = "From: $CFG->noreplyaddress\r\n";
        $headers .= "Reply-To: $CFG->noreplyaddress\r\n";
        $headers .= "MIME-Version: 1.0\r\n";
        $headers .= "Content-Type: text/plain; charset=UTF-8\r\n";
        
        $result = mail($destinatario, $subject, $message, $headers);
        
        if ($result) {
            error_log("Correo enviado exitosamente (método alternativo) a $destinatario");
            return true;
        } else {
            error_log("Fallo al enviar correo (método alternativo) a $destinatario");
            return false;
        }
    } catch (Exception $e) {
        error_log("Error en método alternativo para $destinatario: " . $e->getMessage());
        return false;
    }
}

/**
 * Procesa cancelación de asignatura (versión mejorada)
 */
function procesarCancelacion($registro, $modoPrueba) {
    global $DB;
    
    $idgrupo = $registro['IDGRUPO'];
    $username = $registro['NUMERODOCUMENTO'];
    
    // Buscar curso en Moodle
    $curso = $DB->get_record('course', ['summary' => $idgrupo], 'id,fullname,shortname');
    if (!$curso) {
        error_log("[CANCELACIÓN] Curso no encontrado para IDGRUPO: $idgrupo");
        return false;
    }
    
    // Buscar usuario en Moodle
    $user = $DB->get_record('user', ['username' => $username], 'id,email,firstname,lastname');
    if (!$user) {
        error_log("[CANCELACIÓN] Usuario no encontrado: $username");
        return false;
    }
    
    if ($modoPrueba) {
        error_log("[SIMULACIÓN] Se cambiaría rol de $username a DESMATRICULADO en curso {$curso->fullname}");
        return true;
    }
    
    // [Resto del código de procesarCancelacion permanece igual...]
    
    // Enviar correos de notificación con la nueva función
    $resultadosCorreos = [];
    
    // Correo al estudiante
    $subjectEst = "Cancelación de asignatura - {$curso->fullname}";
    $messageEst = "Estimado/a {$user->firstname} {$user->lastname},\n\n";
    $messageEst .= "Te informamos que tu matrícula en la asignatura {$curso->fullname} ha sido cancelada.\n\n";
    $messageEst .= "Si tienes dudas o es un error, comunícate a través de WhatsApp: <a href=\"https://api.whatsapp.com/send/?phone=3203921622&text&type=phone_number&app_absent=0\" target=\"_blank\">3203921622</a>\n\n";
    $messageEst .= "Atentamente,\n";
    $messageEst .= "Univirtual UTP";
    
    $resultadosCorreos['estudiante'] = enviarCorreoMejorado($user->email, $subjectEst, $messageEst, $modoPrueba);
    
    // Correo a docentes
    $profesores = $DB->get_records_sql("
        SELECT u.* 
        FROM {user} u
        JOIN {role_assignments} ra ON ra.userid = u.id
        JOIN {context} ctx ON ctx.id = ra.contextid AND ctx.contextlevel = 50
        JOIN {course} c ON c.id = ctx.instanceid
        WHERE c.id = ? AND ra.roleid = 3", [$curso->id]);
    
    $subjectDoc = "Estudiante cancelado - {$curso->fullname}";
    $messageDoc = "Estimado/a docente,\n\n";
    $messageDoc .= "El estudiante {$user->firstname} {$user->lastname} ({$user->email}) ";
    $messageDoc .= "ha cancelado la asignatura {$curso->fullname}.\n\n";
    $messageDoc .= "Atentamente,\n";
    $messageDoc .= "Univirtual UTP";
    
    foreach ($profesores as $profesor) {
        $resultadosCorreos['docente_'.$profesor->id] = enviarCorreoMejorado($profesor->email, $subjectDoc, $messageDoc, $modoPrueba);
    }
    
    // Correo a soporte y univirtual
    $subjectSoporte = "Notificación de cancelación - {$user->username}";
    $messageSoporte = "Se ha realizado una cancelación en el sistema:\n\n";
    $messageSoporte .= "Estudiante: {$user->firstname} {$user->lastname}\n";
    $messageSoporte .= "Documento: {$user->username}\n";
    $messageSoporte .= "Email: {$user->email}\n";
    $messageSoporte .= "Curso: {$curso->fullname}\n";
    $messageSoporte .= "ID Curso: {$curso->id}\n";
    $messageSoporte .= "ID Grupo: {$curso->summary}\n";
    $messageSoporte .= "Fecha: ".date('Y-m-d H:i:s')."\n\n";
    $messageSoporte .= "Resultados envíos:\n".print_r($resultadosCorreos, true)."\n\n";
    $messageSoporte .= "Este correo es informativo, no requiere respuesta.";
    
    $resultadosCorreos['soporte'] = enviarCorreoMejorado(
        [EMAIL_SOPORTE, EMAIL_UNIVIRTUAL],
        $subjectSoporte,
        $messageSoporte,
        $modoPrueba
    );
    
    // Registrar todos los resultados
    error_log("Resultados de envíos de correo: ".print_r($resultadosCorreos, true));
    
    return true;
}

// [Las demás funciones se modificarían de manera similar...]

// =============================================================================
// EJECUCIÓN PRINCIPAL
// =============================================================================

try {
    error_log("Iniciando proceso de adiciones y cancelaciones " . ($modoPrueba ? "en MODO PRUEBA" : ""));
    error_log("Filtrando registros con fecha CREACIÓN >= ".date('Y-m-d H:i:s', FECHA_INICIO));
    
    // 1. Conectar a Oracle y obtener registros
    $connOracle = conectarOracle();
    $registros = obtenerRegistrosProcesar($connOracle);
    oci_close($connOracle);
    
    error_log("Registros a procesar: " . count($registros));
    
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
                    error_log("Tipo de operación no reconocido: {$registro['VALORTIPO']}");
                    $resultados['errores']++;
            }
        } catch (Exception $e) {
            error_log("ERROR procesando registro: " . $e->getMessage());
            $resultados['errores']++;
        }
    }
    
    // 3. Generar y enviar reporte final
    enviarReporteFinal($resultados, $modoPrueba);
    
    // 4. Mostrar resumen
    error_log("\n=== RESUMEN FINAL ===");
    error_log("Total registros procesados: {$resultados['total']}");
    error_log("Cancelaciones procesadas: {$resultados['cancelaciones']}");
    error_log("Adiciones procesadas: {$resultados['adiciones']}");
    error_log("Cambios de grupo reportados: {$resultados['cambios_grupo']}");
    error_log("Errores encontrados: {$resultados['errores']}");

} catch (Exception $e) {
    error_log("ERROR: " . $e->getMessage());
    
    // Enviar correo de error
    $subjectError = "ERROR en script de adiciones/cancelaciones";
    $messageError = "Ocurrió un error grave en el script:\n\n";
    $messageError .= $e->getMessage()."\n\n";
    $messageError .= "Fecha: ".date('Y-m-d H:i:s')."\n";
    
    enviarCorreoMejorado([EMAIL_SOPORTE, EMAIL_UNIVIRTUAL], $subjectError, $messageError, $modoPrueba);
    
    exit(1);
}

exit(0);