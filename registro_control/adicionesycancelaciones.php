<?php
/**
 * SCRIPT PARA MANEJO DE ADICIONES Y CANCELACIONES EN MOODLE
 * Optimizado con fecha incremental (last_run.txt)
 */

define('CLI_SCRIPT', true);
error_reporting(E_ALL);
ini_set('display_errors', 1);
set_time_limit(0);

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

define('EMAIL_SOPORTE', 'juapabgonzalez@utp.edu.co');
define('EMAIL_SOPORTE_RAMIREZ', 's.ramirez9@utp.edu.co');
define('EMAIL_SOPORTE_ADICIONAL', 'univirtual@utp.edu.co');
define('FROM_EMAIL', 'noreply@utp.edu.co');
define('REGISTRO_PROCESADOS_FILE', __DIR__.'/procesados.log');
define('LAST_RUN_FILE', __DIR__.'/last_run.txt');

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
               "se ha procesado la accion: {$curso->fullname} ha sido cancelada por usted.\n\n".
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
    $subject = "Solicitud de Cancelación - {$curso->fullname}";
    $message = "El estudiante {$user->firstname} {$user->lastname} ha solicitado cancelación ante Registro y Control.\n\n".
               "Atentamente,\nUnivirtual UTP";
    foreach ($profesores as $profesor) {
        enviarCorreo($profesor->email, $subject, $message);
    }
}

function enviarCorreoUsuarioNoExiste($username, $idgrupo) {
    $subject = "[URGENTE] Usuario no existe en Moodle";
    $message = "Se intentó matricular al usuario $username en el curso con IDGRUPO $idgrupo pero no existe en Moodle.\n\n".
               "Por favor crear el usuario manualmente.\n\n".
               "Fecha: ".date('Y-m-d H:i:s')."\n";
    enviarCorreo(EMAIL_SOPORTE, $subject, $message);
    enviarCorreo(EMAIL_SOPORTE_ADICIONAL, $subject, $message);
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
    // Determinar rango de fechas
    if (file_exists(LAST_RUN_FILE)) {
        $fechaInicio = trim(file_get_contents(LAST_RUN_FILE));
    } else {
        // primera vez: desde el día anterior 00:00:00
        $fechaInicio = date('Y-m-d 00:00:00', strtotime('-1 day'));
    }
    $fechaFin = date('Y-m-d H:i:s');

    echo "Filtrando registros entre $fechaInicio y $fechaFin\n";

    $sql = "SELECT * FROM REGISTRO.VI_RYC_NOUNIVIRTUALASIGNCANCEL 
            WHERE regexp_like(PERIODOACADEMICO,'^20252(.)*[Pp][Rr][Ee][Gg][Rr][Aa][Dd][Oo].*')
            AND IDPERIODOACADEMICO <> '21896'
            AND FECHACREACION BETWEEN TO_DATE('$fechaInicio', 'YYYY-MM-DD HH24:MI:SS')
                                   AND TO_DATE('$fechaFin', 'YYYY-MM-DD HH24:MI:SS')
            ORDER BY FECHACREACION ASC";

    $stid = oci_parse($conn, $sql);
    oci_execute($stid);

    $registros = [];
    while ($row = oci_fetch_array($stid, OCI_ASSOC+OCI_RETURN_NULLS)) {
        $registros[] = $row;
    }
    oci_free_statement($stid);
    return $registros;
}

// =============================================================================
// FUNCIONES DE PROCESO
// =============================================================================

function procesarCancelacion($registro, $modoPrueba, &$resumen) {
    global $DB;

    $idgrupo = $registro['IDGRUPO'];
    $username = $registro['NUMERODOCUMENTO'];
    $fechaOracle = $registro['FECHACREACION'];

    // Evitar reprocesar
    if (yaProcesado($idgrupo, $username, TIPO_CANCELACION, $fechaOracle)) {
        $resumen[] = [
            'tipo' => 'Cancelación',
            'username' => $username,
            'nombre' => '-',
            'email' => '-',
            'idgrupo' => $idgrupo,
            'curso' => '-',
            'fecha_oracle' => $fechaOracle,
            'fecha_proceso' => date('Y-m-d H:i:s'),
            'estado' => 'Ya procesado'
        ];
        return false;
    }

    // Buscar curso por IDGRUPO en summary
    $curso = $DB->get_record('course', ['summary' => $idgrupo], 'id,fullname');
    if (!$curso) {
        $resumen[] = [
            'tipo' => 'Cancelación',
            'username' => $username,
            'nombre' => '-',
            'email' => '-',
            'idgrupo' => $idgrupo,
            'curso' => '(Curso no encontrado)',
            'fecha_oracle' => $fechaOracle,
            'fecha_proceso' => date('Y-m-d H:i:s'),
            'estado' => 'Error: curso no encontrado'
        ];
        return false;
    }

    // Buscar usuario por documento
    $user = $DB->get_record('user', ['username' => $username]);
    if (!$user) {
        $resumen[] = [
            'tipo' => 'Cancelación',
            'username' => $username,
            'nombre' => '(Usuario no encontrado)',
            'email' => '-',
            'idgrupo' => $idgrupo,
            'curso' => $curso->fullname,
            'fecha_oracle' => $fechaOracle,
            'fecha_proceso' => date('Y-m-d H:i:s'),
            'estado' => 'Error: usuario no encontrado'
        ];
        return false;
    }

    // Si no es modo prueba, aplicar cambios en Moodle
    if (!$modoPrueba) {
        $context = context_course::instance($curso->id);

        // Quitar rol de estudiante
        $DB->delete_records('role_assignments', [
            'userid' => $user->id,
            'contextid' => $context->id,
            'roleid' => ROL_ESTUDIANTE
        ]);

        // Asignar rol "desmatriculado"
        $role_assign = (object)[
            'roleid' => ROL_DESMATRICULADO,
            'contextid' => $context->id,
            'userid' => $user->id,
            'timemodified' => time(),
            'modifierid' => 2
        ];
        $DB->insert_record('role_assignments', $role_assign);

        // Registrar en log interno
        registrarProcesado($idgrupo, $username, TIPO_CANCELACION, $fechaOracle);

        // Notificaciones
        enviarCorreoCancelacionEstudiante($user, $curso);
        enviarCorreoCancelacionDocente($user, $curso);
    }

    // Agregar al resumen para reporte
    $resumen[] = [
        'tipo' => 'Cancelación',
        'username' => $username,
        'nombre' => $user->firstname . ' ' . $user->lastname,
        'email' => $user->email,
        'idgrupo' => $idgrupo,
        'curso' => $curso->fullname,
        'fecha_oracle' => $fechaOracle,
        'fecha_proceso' => date('Y-m-d H:i:s'),
        'estado' => $modoPrueba ? 'Simulado (modo prueba)' : 'Procesado OK'
    ];

    return true;
}
function procesarAdicion($registro, $modoPrueba, &$resumen) {
    global $DB;

    $idgrupo = $registro['IDGRUPO'];
    $username = $registro['NUMERODOCUMENTO'];
    $fechaOracle = $registro['FECHACREACION'];

    if (yaProcesado($idgrupo, $username, TIPO_ADICION, $fechaOracle)) {
        $resumen[] = [
            'tipo' => 'Adición',
            'username' => $username,
            'nombre' => '-',
            'email' => '-',
            'idgrupo' => $idgrupo,
            'curso' => '-',
            'fecha_oracle' => $fechaOracle,
            'fecha_proceso' => date('Y-m-d H:i:s'),
            'estado' => 'Ya procesado'
        ];
        return false;
    }

    $curso = $DB->get_record('course', ['summary' => $idgrupo], 'id,fullname');
    if (!$curso) {
        $resumen[] = [
            'tipo' => 'Adición',
            'username' => $username,
            'nombre' => '-',
            'email' => '-',
            'idgrupo' => $idgrupo,
            'curso' => '(Curso no encontrado)',
            'fecha_oracle' => $fechaOracle,
            'fecha_proceso' => date('Y-m-d H:i:s'),
            'estado' => 'Error: curso no encontrado'
        ];
        return false;
    }

    $user = $DB->get_record('user', ['username' => $username]);
    if (!$user) {
        enviarCorreoUsuarioNoExiste($username, $idgrupo);
        $resumen[] = [
            'tipo' => 'Adición',
            'username' => $username,
            'nombre' => '(Usuario no encontrado)',
            'email' => '-',
            'idgrupo' => $idgrupo,
            'curso' => $curso->fullname,
            'fecha_oracle' => $fechaOracle,
            'fecha_proceso' => date('Y-m-d H:i:s'),
            'estado' => 'Error: usuario no encontrado'
        ];
        return false;
    }

    if (!$modoPrueba) {
        $enrol = $DB->get_record('enrol', ['courseid' => $curso->id, 'enrol' => 'manual'], '*', MUST_EXIST);

        if (!$DB->record_exists('user_enrolments', ['userid' => $user->id, 'enrolid' => $enrol->id])) {
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
        }

        registrarProcesado($idgrupo, $username, TIPO_ADICION, $fechaOracle);
        enviarCorreoAdicionEstudiante($user, $curso);
    }

    $resumen[] = [
        'tipo' => 'Adición',
        'username' => $username,
        'nombre' => $user->firstname . ' ' . $user->lastname,
        'email' => $user->email,
        'idgrupo' => $idgrupo,
        'curso' => $curso->fullname,
        'fecha_oracle' => $fechaOracle,
        'fecha_proceso' => date('Y-m-d H:i:s'),
        'estado' => $modoPrueba ? 'Simulado (modo prueba)' : 'Procesado OK'
    ];

    return true;
}

function procesarCambioGrupo($registro, $modoPrueba, &$resumen) {
    $idgrupo = $registro['IDGRUPO'];
    $username = $registro['NUMERODOCUMENTO'];
    $fechaOracle = $registro['FECHACREACION'];

    if (yaProcesado($idgrupo, $username, TIPO_CAMBIO_GRUPO, $fechaOracle)) {
        $resumen[] = [
            'tipo' => 'Cambio de grupo',
            'username' => $username,
            'nombre' => $registro['NOMBRES'] . ' ' . $registro['APELLIDOS'],
            'email' => '-',
            'idgrupo' => $idgrupo,
            'curso' => $registro['NOMBREASIGNATURA'] ?? '-',
            'fecha_oracle' => $fechaOracle,
            'fecha_proceso' => date('Y-m-d H:i:s'),
            'estado' => 'Ya procesado'
        ];
        return false;
    }

    if (!$modoPrueba) {
        registrarProcesado($idgrupo, $username, TIPO_CAMBIO_GRUPO, $fechaOracle);
        enviarCorreoCambioGrupo($registro);
    }

    $resumen[] = [
        'tipo' => 'Cambio de grupo',
        'username' => $username,
        'nombre' => $registro['NOMBRES'] . ' ' . $registro['APELLIDOS'],
        'email' => '-', // no siempre disponible en registro Oracle
        'idgrupo' => $idgrupo,
        'curso' => $registro['NOMBREASIGNATURA'] ?? '-',
        'fecha_oracle' => $fechaOracle,
        'fecha_proceso' => date('Y-m-d H:i:s'),
        'estado' => $modoPrueba ? 'Simulado (modo prueba)' : 'Reportado'
    ];

    return true;
}

use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\Exception;

function enviarCorreoConAdjunto($to, $subject, $message, $filePath, $fileName) {
    $mail = new PHPMailer(true);

    try {
        // Configuración del servidor SMTP (Gmail)
        $mail->isSMTP();
        $mail->Host = 'smtp.gmail.com';
        $mail->SMTPAuth = true;
        $mail->Username = 'univirtual@utp.edu.co';     // ⚠️ tu correo Gmail
        $mail->Password = 'fzti zzyv xhma melo';      // ⚠️ tu contraseña de aplicación
        $mail->SMTPSecure = 'tls';
        $mail->Port = 587;

        // Remitente
        $mail->setFrom('univirtual@utp.edu.co', 'Univirtual UTP');

        // Destinatario
        $mail->addAddress($to);

        // Contenido
        $mail->isHTML(false);
        $mail->Subject = $subject;
        $mail->Body    = $message;

        // Adjuntar archivo si existe
        if ($filePath && file_exists($filePath)) {
            $mail->addAttachment($filePath, $fileName);
        }

        $mail->send();
        echo "Correo enviado a $to\n";
        return true;
    } catch (Exception $e) {
        echo "Error enviando correo a $to: {$mail->ErrorInfo}\n";
        return false;
    }
}
function enviarReporteFinal($resultados, $modoPrueba, $resumen) {
    // Si no hubo registros, añadimos una línea mínima al resumen
    if (empty($resumen)) {
        $resumen[] = [
            'tipo' => '-',
            'username' => '-',
            'nombre' => '-',
            'email' => '-',
            'idgrupo' => '-',
            'curso' => '-',
            'fecha_oracle' => '-',
            'fecha_proceso' => date('Y-m-d H:i:s'),
            'estado' => 'No se encontraron registros en este rango.'
        ];
    }

    $subject = ($modoPrueba ? "[PRUEBA] " : "")."Reporte de Adiciones/Cancelaciones - ".date('Y-m-d H:i:s');
    $message = "Reporte de ejecución ".($modoPrueba ? "en modo prueba" : "en producción")."\n\n".
               "Total registros: {$resultados['total']}\n".
               "Cancelaciones: {$resultados['cancelaciones']}\n".
               "Adiciones: {$resultados['adiciones']}\n".
               "Cambios de grupo: {$resultados['cambios_grupo']}\n".
               "Errores: {$resultados['errores']}\n\n".
               "Detalles:\n";

    // Construir el detalle para el cuerpo del correo
    foreach ($resumen as $r) {
        $message .= "- {$r['tipo']} {$r['username']} en {$r['idgrupo']} ({$r['curso']}) Estado: {$r['estado']}\n";
    }

    // archivo CSV temporal
    $csvFileName = tempnam(sys_get_temp_dir(), 'reporte_').'.csv';
    $csvFile = fopen($csvFileName, 'w');

    // encabezados CSV
    fputcsv($csvFile, [
        'Tipo Operación', 'Username', 'Nombre Completo', 'Email',
        'ID Grupo', 'Asignatura', 'Fecha Oracle', 'Fecha Proceso', 'Estado'
    ], ';');

    // recorrer los detalles del resumen y escribir en CSV
    foreach ($resumen as $r) {
        fputcsv($csvFile, [
            $r['tipo'],
            $r['username'],
            $r['nombre'],
            $r['email'],
            $r['idgrupo'],
            $r['curso'],
            $r['fecha_oracle'],
            $r['fecha_proceso'],
            $r['estado']
        ], ';');
    }

    fclose($csvFile);

    // enviar correo con adjunto a los destinatarios definidos
    enviarCorreoConAdjunto(EMAIL_SOPORTE, $subject, $message, $csvFileName, 'reporte.csv');
    enviarCorreoConAdjunto(EMAIL_SOPORTE_ADICIONAL, $subject, $message, $csvFileName, 'reporte.csv');
    enviarCorreoConAdjunto(EMAIL_SOPORTE_RAMIREZ, $subject, $message, $csvFileName, 'reporte.csv'); // 👈 nuevo

    // borrar el archivo temporal
    unlink($csvFileName);
}

//===========================================================
// EJECUCIÓN PRINCIPAL
// =============================================================================

try {
    limpiarRegistrosAntiguos();
    echo "Iniciando proceso...\n";

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

        // Enviar reporte final al terminar
    enviarReporteFinal($resultados, $modoPrueba, $resumen);


    // Guardar última fecha procesada
    file_put_contents(LAST_RUN_FILE, date('Y-m-d H:i:s'));
    echo "Fecha de última ejecución registrada: ".date('Y-m-d H:i:s')."\n";

} catch (Exception $e) {
    echo "ERROR: ".$e->getMessage()."\n";
    exit(1);
}

exit(0);
?>
