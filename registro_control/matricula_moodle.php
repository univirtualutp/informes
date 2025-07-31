<?php
/**
 * SCRIPT DE MATRÍCULA AUTOMÁTICA - MODO PRUEBA MEJORADO
 * 
 * Modo prueba: Verifica todo pero no realiza cambios
 * php matricula_moodle.php --dry-run
 * 
 * Modo producción: Ejecuta el proceso completo
 * php matricula_moodle.php
 */

// =============================================================================
// CONFIGURACIÓN INICIAL
// =============================================================================

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

// Cargar Moodle (siempre necesario para verificaciones reales)
require_once($config['PATHS']['moodle_config']);
global $DB, $CFG;

// Cargar PHPExcel
require_once __DIR__.'/../../vendor/autoload.php';

// =============================================================================
// CONSTANTES
// =============================================================================

define('ROLE_ESTUDIANTE', 5);
define('EMAIL_SOPORTE', 'soporteunivirtual@utp.edu.co');
define('EMAIL_ASUNTO', 'Reporte de Matrículas Automáticas Moodle');

// =============================================================================
// FUNCIONES PRINCIPALES (ACTUALIZADAS PARA VALIDACIÓN REAL EN MODO PRUEBA)
// =============================================================================

/**
 * Obtiene estudiantes desde Oracle (igual para ambos modos)
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
 * Verifica existencia de usuario (consulta REAL incluso en modo prueba)
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
 * Busca curso por IDGRUPO (consulta REAL en ambos modos)
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
 * Matricula usuario (solo en modo producción)
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
    
    // Asignar rol
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

// ... [funciones generarReporteExcel y enviarCorreoConReporte permanecen iguales] ...

// =============================================================================
// EJECUCIÓN PRINCIPAL
// =============================================================================

try {
    echo "Iniciando proceso de matrícula " . ($modoPrueba ? "en MODO PRUEBA" : "en MODO PRODUCCIÓN") . "...\n";
    
    // 1. Conectar a Oracle
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
    
    // 3. Generar reporte
    $nombreArchivo = '/tmp/reporte_matriculas_' . date('Ymd_His') . '.xlsx';
    generarReporteExcel($resultados, $nombreArchivo);
    echo "\nReporte generado: $nombreArchivo\n";
    
    // 4. Enviar correo
    $asunto = ($modoPrueba ? '[PRUEBA] ' : '') . EMAIL_ASUNTO . ' - ' . date('Y-m-d H:i:s');
    if (enviarCorreoConReporte($nombreArchivo, $conteo, $asunto)) {
        echo "Correo enviado a " . EMAIL_SOPORTE . "\n";
    }
    
    // 5. Resumen final
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