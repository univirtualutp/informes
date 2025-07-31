<?php
/**
 * SCRIPT DE MATRÍCULA AUTOMÁTICA PARA MOODLE 4.2.1
 * 
 * Uso normal: php matricula_moodle.php
 * Modo prueba: php matricula_moodle.php --dry-run
 * 
 * Características:
 * - Consulta estudiantes en Oracle -> Sistemas UTP
 * - Verifica existencia en Moodle
 * - Valida coincidencia de cursos por IDGRUPO
 * - Modo prueba realiza todas las validaciones sin modificar datos
 * - Genera reportes Excel y envía correos en ambos modos
 */

// =============================================================================
// CONFIGURACIÓN INICIAL (IMPORTANTE: CLI_SCRIPT debe ir primero)
// =============================================================================

define('CLI_SCRIPT', true); // Definición esencial para scripts CLI de Moodle

error_reporting(E_ALL);
ini_set('display_errors', 1);
set_time_limit(0);

// Verificar modo prueba
$modoPrueba = in_array('--dry-run', $GLOBALS['argv']);
if ($modoPrueba) {
    echo "=== MODO PRUEBA ACTIVADO ===\n";
    echo "Validando datos SIN realizar matrículas reales\n\n";
}

// Cargar configuración
$config = require __DIR__.'/config/env.php';

// Cargar Moodle (requiere CLI_SCRIPT definido primero)
require_once($config['PATHS']['moodle_config']);
global $DB, $CFG;

// Cargar PHPExcel
require_once __DIR__.'/../../vendor/autoload.php';

// =============================================================================
// CONSTANTES
// =============================================================================

define('ROLE_ESTUDIANTE', 5); // Rol de estudiante en Moodle
define('EMAIL_SOPORTE', 'soporteunivirtual@utp.edu.co');
define('EMAIL_ASUNTO', 'Reporte de Matrículas Automáticas Moodle');

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
 * Obtiene estudiantes a matricular desde Oracle
 */
function obtenerEstudiantesMatricular($conn) {
    $sql = "SELECT NUMERODOCUMENTO, INITCAP(NOMBRES) AS NOMBRES, INITCAP(APELLIDOS) AS APELLIDOS, 
                   EMAIL, CODIGOASIGNATURA, NOMBREASIGNATURA, NUMEROGRUPO, IDGRUPO, 
                   IDESTRUCFACULTAD, FACULTAD, CODIGOPROGRAMA, PROGRAMA, JORNADA, 
                   FECHANACIMIENTO, SEXO, TELEFONO, ESTRATOSOCIAL 
            FROM REGISTRO.VI_RYC_UNIVIRTUALESTUDMATRIC 
            WHERE regexp_like(PERIODOACADEMICO,'^'||'20252'||'(.)*[Pp][Rr][Ee][Gg][Rr][Aa][Dd][Oo].*')
            AND CODIGOPROGRAMA NOT IN ('TR')
            ORDER BY NOMBREASIGNATURA, NUMEROGRUPO, INITCAP(NOMBRES) ASC";
    
    $stid = oci_parse($conn, $sql);
    oci_execute($stid);
    
    $estudiantes = [];
    while ($row = oci_fetch_array($stid, OCI_ASSOC+OCI_RETURN_NULLS)) {
        $estudiantes[] = $row;
    }
    
    oci_free_statement($stid);
    return $estudiantes;
}

/**
 * Verifica si usuario existe en Moodle (consulta REAL en ambos modos)
 */
function usuarioExisteEnMoodle($username) {
    global $DB;
    $existe = $DB->record_exists('user', ['username' => $username, 'deleted' => 0]);
    
    if ($GLOBALS['modoPrueba']) {
        echo "[VALIDACIÓN] Usuario $username " . ($existe ? "EXISTE" : "NO EXISTE") . " en Moodle\n";
    }
    
    return $existe;
}

/**
 * Busca curso por IDGRUPO (summary) - Consulta REAL en ambos modos
 */
function obtenerCursoPorSummary($summary) {
    global $DB;
    $curso = $DB->get_record('course', ['summary' => $summary], 'id,fullname,shortname');
    
    if ($GLOBALS['modoPrueba']) {
        if ($curso) {
            echo "[VALIDACIÓN] Curso con summary '$summary' ENCONTRADO: {$curso->fullname} (ID: {$curso->id})\n";
        } else {
            echo "[VALIDACIÓN] ERROR: No existe curso con summary '$summary' en Moodle\n";
        }
    }
    
    return $curso;
}

/**
 * Matricula usuario (solo ejecuta en modo producción)
 */
function matricularUsuario($userid, $courseid) {
    if ($GLOBALS['modoPrueba']) {
        echo "[SIMULACIÓN] Se MATRICULARÍA al usuario $userid en el curso $courseid\n";
        return true;
    }
    
    global $DB;
    
    // Verificar si ya está matriculado
    if ($DB->record_exists('user_enrolments', ['userid' => $userid, 'enrolid' => $courseid])) {
        return false;
    }
    
    // Obtener instancia de matriculación manual
    $enrol = $DB->get_record('enrol', ['courseid' => $courseid, 'enrol' => 'manual'], '*', MUST_EXIST);
    
    // Asignar rol de estudiante
    $context = context_course::instance($courseid);
    $role_assign = [
        'roleid' => ROLE_ESTUDIANTE,
        'contextid' => $context->id,
        'userid' => $userid,
        'timemodified' => time(),
        'modifierid' => 2
    ];
    $DB->insert_record('role_assignments', (object)$role_assign);
    
    // Crear matrícula
    $user_enrolment = [
        'status' => 0,
        'enrolid' => $enrol->id,
        'userid' => $userid,
        'timestart' => time(),
        'timecreated' => time(),
        'timemodified' => time()
    ];
    $DB->insert_record('user_enrolments', (object)$user_enrolment);
    
    return true;
}

/**
 * Genera reporte Excel con los resultados
 */
function generarReporteExcel($datos, $rutaArchivo) {
    $objPHPExcel = new PHPExcel();
    $objPHPExcel->setActiveSheetIndex(0);
    $sheet = $objPHPExcel->getActiveSheet();
    
    // Encabezados
    $sheet->setCellValue('A1', 'Username');
    $sheet->setCellValue('B1', 'Nombres');
    $sheet->setCellValue('C1', 'Apellidos');
    $sheet->setCellValue('D1', 'Email');
    $sheet->setCellValue('E1', 'ID Curso');
    $sheet->setCellValue('F1', 'Nombre Curso');
    $sheet->setCellValue('G1', 'ID Grupo');
    $sheet->setCellValue('H1', 'Resultado');
    
    // Datos
    $row = 2;
    foreach ($datos as $item) {
        $sheet->setCellValue('A'.$row, $item['username']);
        $sheet->setCellValue('B'.$row, $item['nombres']);
        $sheet->setCellValue('C'.$row, $item['apellidos']);
        $sheet->setCellValue('D'.$row, $item['email']);
        $sheet->setCellValue('E'.$row, $item['courseid']);
        $sheet->setCellValue('F'.$row, $item['coursename']);
        $sheet->setCellValue('G'.$row, $item['idgrupo']);
        $sheet->setCellValue('H'.$row, $item['resultado']);
        $row++;
    }
    
    // Autoajustar columnas
    foreach (range('A', 'H') as $column) {
        $sheet->getColumnDimension($column)->setAutoSize(true);
    }
    
    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
    $objWriter->save($rutaArchivo);
}

/**
 * Envía correo con el reporte adjunto
 */
function enviarCorreoConReporte($archivo, $conteo, $asunto) {
    $to = EMAIL_SOPORTE;
    $subject = $asunto;
    
    $message = "Reporte de matrículas " . (strpos($asunto, '[PRUEBA]') !== false ? "de prueba" : "") . "\n\n";
    $message .= "Total procesados: {$conteo['total']}\n";
    $message .= ($GLOBALS['modoPrueba'] ? "Estudiantes validados para matrícula" : "Estudiantes matriculados") . ": {$conteo['exitosos']}\n";
    $message .= "Usuarios no existentes en Moodle: {$conteo['no_existen']}\n";
    $message .= "Errores (cursos no encontrados): {$conteo['errores']}\n\n";
    
    foreach ($conteo['por_curso'] as $curso => $cantidad) {
        $message .= "Curso {$curso}: {$cantidad} estudiantes\n";
    }
    
    $headers = "From: {$GLOBALS['config']['EMAIL']['from']}\r\n";
    $headers .= "MIME-Version: 1.0\r\n";
    $headers .= "Content-Type: multipart/mixed; boundary=\"boundary\"\r\n";
    
    $body = "--boundary\r\n";
    $body .= "Content-Type: text/plain; charset=\"utf-8\"\r\n";
    $body .= "Content-Transfer-Encoding: 7bit\r\n\r\n";
    $body .= $message."\r\n\r\n";
    $body .= "--boundary\r\n";
    $body .= "Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet; name=\"".basename($archivo)."\"\r\n";
    $body .= "Content-Transfer-Encoding: base64\r\n";
    $body .= "Content-Disposition: attachment\r\n\r\n";
    $body .= chunk_split(base64_encode(file_get_contents($archivo)))."\r\n";
    $body .= "--boundary--";
    
    return mail($to, $subject, $body, $headers);
}

// =============================================================================
// EJECUCIÓN PRINCIPAL
// =============================================================================

try {
    echo "Iniciando proceso de matrícula " . ($modoPrueba ? "en MODO PRUEBA" : "en MODO PRODUCCIÓN") . "...\n";
    
    // 1. Conectar a Oracle y obtener estudiantes
    $connOracle = conectarOracle();
    $estudiantes = obtenerEstudiantesMatricular($connOracle);
    oci_close($connOracle);
    
    echo "Estudiantes a procesar: " . count($estudiantes) . "\n";
    
    $resultados = [];
    $conteo = [
        'total' => 0,
        'exitosos' => 0,
        'no_existen' => 0,
        'errores' => 0,
        'por_curso' => []
    ];
    
    // 2. Procesar cada estudiante
    foreach ($estudiantes as $est) {
        $conteo['total']++;
        $username = $est['NUMERODOCUMENTO'];
        $idgrupo = $est['IDGRUPO'];
        
        $resultado = [
            'username' => $username,
            'nombres' => $est['NOMBRES'],
            'apellidos' => $est['APELLIDOS'],
            'email' => $est['EMAIL'],
            'idgrupo' => $idgrupo,
            'courseid' => null,
            'coursename' => null,
            'resultado' => null
        ];
        
        // Verificar usuario
        if (!usuarioExisteEnMoodle($username)) {
            $resultado['resultado'] = 'Usuario no existe en Moodle';
            $conteo['no_existen']++;
            $resultados[] = $resultado;
            continue;
        }
        
        // Buscar curso
        $curso = obtenerCursoPorSummary($idgrupo);
        if (!$curso) {
            $resultado['resultado'] = 'Curso no encontrado (IDGRUPO no coincide)';
            $conteo['errores']++;
            $resultados[] = $resultado;
            continue;
        }
        
        // Obtener ID de usuario
        $user = $DB->get_record('user', ['username' => $username], 'id');
        
        // Matricular (o simular en modo prueba)
        if (matricularUsuario($user->id, $curso->id)) {
            $resultado['courseid'] = $curso->id;
            $resultado['coursename'] = $curso->fullname;
            $resultado['resultado'] = $modoPrueba ? 'Validación exitosa (se matricularía)' : 'Matriculado exitosamente';
            $conteo['exitosos']++;
            
            if (!isset($conteo['por_curso'][$curso->fullname])) {
                $conteo['por_curso'][$curso->fullname] = 0;
            }
            $conteo['por_curso'][$curso->fullname]++;
        } else {
            $resultado['resultado'] = 'El usuario ya estaba matriculado';
        }
        
        $resultados[] = $resultado;
    }
    
    // 3. Generar reporte Excel
    $nombreArchivo = '/tmp/reporte_matriculas_' . date('Ymd_His') . '.xlsx';
    generarReporteExcel($resultados, $nombreArchivo);
    echo "\nReporte generado: $nombreArchivo\n";
    
    // 4. Enviar correo con reporte
    $asunto = ($modoPrueba ? '[PRUEBA] ' : '') . EMAIL_ASUNTO . ' - ' . date('Y-m-d H:i:s');
    if (enviarCorreoConReporte($nombreArchivo, $conteo, $asunto)) {
        echo "Correo enviado a " . EMAIL_SOPORTE . "\n";
    }
    
    // 5. Mostrar resumen final
    echo "\n=== RESUMEN FINAL ===\n";
    echo "Total procesados: {$conteo['total']}\n";
    echo ($modoPrueba ? "Estudiantes validados para matrícula" : "Estudiantes matriculados") . ": {$conteo['exitosos']}\n";
    echo "Usuarios no existentes en Moodle: {$conteo['no_existen']}\n";
    echo "Errores (cursos no encontrados): {$conteo['errores']}\n";
    
    if ($modoPrueba) {
        echo "\n=== MODO PRUEBA COMPLETADO ===\n";
        echo "Revisar el reporte generado para verificar los resultados\n";
        echo "Ejecutar sin '--dry-run' para realizar las matrículas reales\n";
    }

} catch (Exception $e) {
    echo "ERROR: " . $e->getMessage() . "\n";
    exit(1);
}

exit(0);