<?php
// Cargar dependencias
require 'vendor/autoload.php'; 
require_once __DIR__ . '/db_moodle_config.php';

use PHPMailer\PHPMailer\PHPMailer;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;

// Configurar zona horaria de Bogotá
date_default_timezone_set('America/Bogota');

// Configuración de correos
$modo_prueba = true; // Cambiar a false para enviar a los destinatarios reales

$correos_produccion = ['soporteunivirtual@utp.edu.co', 'univirtual-utp@utp.edu.co'];
$correo_notificacion_produccion = 'soporteunivirtual@utp.edu.co';
$correos_pruebas = ['daniel.pardo@utp.edu.co']; // Agrega tu correo de prueba aquí
$correo_notificacion_pruebas = 'daniel.pardo@utp.edu.co'; // Correo para notificaciones de prueba

// Seleccionar correos según modo
$correos_destino = $modo_prueba ? $correos_pruebas : $correos_produccion;
$correo_notificacion = $modo_prueba ? $correo_notificacion_pruebas : $correo_notificacion_produccion;

// Lista centralizada de cursos
$cursos_pregrado = [
    '494', '415', '507', '481', '508', '482', '509', '485', '526', '510',
    '511', '486', '490', '416', '503', '504', '527', '417', '496', '497',
    '418', '498', '419', '475', '421', '420', '422', '423', '512', '513',
    '515', '488', '489', '424', '516', '517', '491', '518', '492', '519',
    '493', '520', '425', '476', '426', '505', '506', '479', '521', '428',
    '430', '522', '495', '499', '431', '453', '500', '523', '434', '524',
    '435', '436', '437', '438', '440', '502', '439', '452', '525', '442'
];
$lista_cursos_sql = "'" . implode("','", $cursos_pregrado) . "'";

// Configuración de fechas dinámicas
$fecha_inicio_consulta = '2025-02-03 00:00:00'; // Fecha inicial (semana 1)
$fecha_fin_consulta = calcularUltimoLunesSemana16($fecha_inicio_consulta); // Semana 16
$fecha_inicio_semana_extra = calcularInicioSemanaExtra($fecha_inicio_consulta); // Calcula dinámicamente

try {
    // 1. Generar las semanas académicas automáticamente
    $semanas_academicas = generarSemanasAcademicas($fecha_inicio_consulta, $fecha_inicio_semana_extra);

    // 2. Conectar a la base de datos Moodle
    $dsn = "pgsql:host=" . DB_HOST . ";port=" . DB_PORT . ";dbname=" . DB_NAME;
    $dbh = new PDO($dsn, DB_USER, DB_PASS);
    $dbh->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);

    // 3. Construir la consulta SQL con las semanas generadas
    $sql = construirConsultaSQL($semanas_academicas, $lista_cursos_sql, $fecha_inicio_consulta, $fecha_fin_consulta, $fecha_inicio_semana_extra);

    // 4. Ejecutar la consulta
    $stmt = $dbh->query($sql);
    $resultados = $stmt->fetchAll(PDO::FETCH_ASSOC); // Corrección del error 'flaps'

    // 5. Generar el archivo Excel
    $archivo_excel = generarExcel($resultados, $semanas_academicas);

    // 6. Enviar correos electrónicos
    enviarCorreos($archivo_excel, $correos_destino, $correo_notificacion, $modo_prueba);

    // 7. Limpiar archivo temporal
    if (file_exists($archivo_excel)) {
        unlink($archivo_excel);
    }

} catch (Exception $e) {
    $mensaje_error = 'Error: ' . $e->getMessage() . "\nModo prueba: " . ($modo_prueba ? 'SI' : 'NO');
    mail($correo_notificacion, 'Error Reporte Actividades Pregrado', $mensaje_error);
    exit("Error: " . $e->getMessage());
}

/**
 * Calcula el último lunes (semana 16) a partir de la fecha de inicio
 */
function calcularUltimoLunesSemana16($fecha_inicio) {
    $inicio = new DateTime($fecha_inicio);
    
    // Desde 2025-02-03 hasta 2025-05-26 (fin de semana 16) son 112 días
    $intervalo = new DateInterval('P111D'); // 16 semanas * 7 días - 1
    $fin = clone $inicio;
    $fin->add($intervalo);
    
    // Ajustar para que sea exactamente a las 23:59:59 del lunes
    $fin->setTime(23, 59, 59);
    
    return $fin->format('Y-m-d H:i:s'); // Debería devolver 2025-05-26 23:59:59
}

/**
 * Calcula dinámicamente el inicio de la semana extra (martes siguiente a la semana 16)
 */
function calcularInicioSemanaExtra($fecha_inicio) {
    $ultimo_lunes_semana16 = new DateTime(calcularUltimoLunesSemana16($fecha_inicio));
    
    // La semana extra comienza el martes siguiente al último lunes de la semana 16
    $inicio_semana_extra = clone $ultimo_lunes_semana16;
    $inicio_semana_extra->add(new DateInterval('P1D')); // Martes 00:00:00
    
    return $inicio_semana_extra->format('Y-m-d H:i:s'); // Debería devolver 2025-05-27 00:00:00
}

/**
 * Genera el array de semanas académicas
 */
function generarSemanasAcademicas($fechaInicioSemana1, $fechaInicioSemanaExtra) {
    $semanas = [];
    $inicio = new DateTime($fechaInicioSemana1);
    
    // Semana 1 (7 días: 2025-02-03 al 2025-02-09)
    $fin_semana = clone $inicio;
    $fin_semana->add(new DateInterval('P6DT23H59M59S')); // Lunes a Domingo
    $semanas[] = [
        'semana_numero' => 1,
        'semana_nombre' => 'SEMANA 1',
        'inicio_semana' => $inicio->format('Y-m-d H:i:s'), // 2025-02-03 00:00:00
        'fin_semana' => $fin_semana->format('Y-m-d H:i:s') // 2025-02-09 23:59:59
    ];

    // Semanas 2 a 16 (7 días cada una: martes a lunes)
    for ($i = 2; $i <= 16; $i++) {
        $inicio = clone $fin_semana;
        $inicio->add(new DateInterval('PT1S')); // Comienza el martes
        $fin_semana = clone $inicio;
        $fin_semana->add(new DateInterval('P6DT23H59M59S')); // Hasta el lunes siguiente
        
        $semanas[] = [
            'semana_numero' => $i,
            'semana_nombre' => 'SEMANA ' . $i,
            'inicio_semana' => $inicio->format('Y-m-d H:i:s'),
            'fin_semana' => $fin_semana->format('Y-m-d H:i:s')
        ];
    }

    // Semana extra (17): 2025-05-27 al 2025-06-02
    $inicio_semana_extra = new DateTime($fechaInicioSemanaExtra);
    $fin_semana_extra = clone $inicio_semana_extra;
    $fin_semana_extra->add(new DateInterval('P6DT23H59M59S')); // 7 días
    
    $semanas[] = [
        'semana_numero' => 17,
        'semana_nombre' => 'SEMANA EXTRA',
        'inicio_semana' => $inicio_semana_extra->format('Y-m-d H:i:s'), // 2025-05-27 00:00:00
        'fin_semana' => $fin_semana_extra->format('Y-m-d H:i:s') // 2025-06-02 23:59:59
    ];

    return $semanas;
}

/**
 * Construye la consulta SQL con parámetros dinámicos
 */
function construirConsultaSQL($semanas, $lista_cursos, $fecha_inicio, $fecha_fin, $fecha_inicio_semana_extra) {
    $timestamp_inicio = strtotime($fecha_inicio);
    $timestamp_fin = strtotime($fecha_fin);
    
    $sql_semanas = "";
    $first = true;
    foreach ($semanas as $semana) {
        if (!$first) {
            $sql_semanas .= "UNION ALL\n";
        }
        $sql_semanas .= "SELECT {$semana['semana_numero']} AS semana_numero, '{$semana['semana_nombre']}' AS semana_nombre, 
                        '{$semana['inicio_semana']}'::timestamp AS inicio_semana, 
                        '{$semana['fin_semana']}'::timestamp AS fin_semana\n";
        $first = false;
    }

    $sql = "WITH semanas_academicas AS (
        $sql_semanas
    ),
    all_users AS (
        SELECT DISTINCT 
            u.id AS userid, 
            c.id AS courseid,
            mc.id AS contextid
        FROM mdl_user u
        JOIN mdl_role_assignments ra ON ra.userid = u.id
        JOIN mdl_context mc ON ra.contextid = mc.id AND mc.contextlevel = 50
        JOIN mdl_course c ON mc.instanceid = c.id 
        WHERE c.id IN ($lista_cursos)
          AND u.username <> '12345678'
          AND ra.roleid IN ('5', '9', '16', '17')
    ),
    all_activities AS (
        SELECT 
            f.course AS courseid,
            f.id AS activity_id,
            'Foro' AS activity_type,
            f.name AS activity_name,
            CASE 
                WHEN f.duedate > 0 THEN TO_TIMESTAMP(f.duedate)
                ELSE TO_TIMESTAMP(c.startdate)
            END AS fecha_entrega,
            CASE 
                WHEN f.cutoffdate > 0 THEN TO_TIMESTAMP(f.cutoffdate)
                ELSE TO_TIMESTAMP(c.enddate)
            END AS fecha_limite
        FROM mdl_forum f
        JOIN mdl_course c ON f.course = c.id
        WHERE f.course IN ($lista_cursos)
          AND f.name != 'Avisos'
        
        UNION ALL
        
        SELECT 
            q.course AS courseid,
            q.id AS activity_id,
            'Quiz' AS activity_type,
            q.name AS activity_name,
            CASE 
                WHEN q.timeopen > 0 THEN TO_TIMESTAMP(q.timeopen)
                ELSE TO_TIMESTAMP(c.startdate)
            END AS fecha_entrega,
            CASE 
                WHEN q.timeclose > 0 THEN TO_TIMESTAMP(q.timeclose)
                ELSE TO_TIMESTAMP(c.enddate)
            END AS fecha_limite
        FROM mdl_quiz q
        JOIN mdl_course c ON q.course = c.id
        WHERE q.course IN ($lista_cursos)
        
        UNION ALL
        
        SELECT 
            a.course AS courseid,
            a.id AS activity_id,
            'Tarea' AS activity_type,
            a.name AS activity_name,
            CASE 
                WHEN a.allowsubmissionsfromdate > 0 THEN TO_TIMESTAMP(a.allowsubmissionsfromdate)
                ELSE TO_TIMESTAMP(c.startdate)
            END AS fecha_entrega,
            CASE 
                WHEN a.duedate > 0 THEN TO_TIMESTAMP(a.duedate)
                ELSE TO_TIMESTAMP(c.enddate)
            END AS fecha_limite
        FROM mdl_assign a
        JOIN mdl_course c ON a.course = c.id
        WHERE a.course IN ($lista_cursos)
        
        UNION ALL
        
        SELECT 
            gi.courseid,
            gi.id AS activity_id,
            'Calificación Manual' AS activity_type,
            gi.itemname AS activity_name,
            TO_TIMESTAMP(c.startdate) AS fecha_entrega,
            TO_TIMESTAMP(c.enddate) AS fecha_limite
        FROM mdl_grade_items gi
        JOIN mdl_course c ON gi.courseid = c.id
        WHERE gi.courseid IN ($lista_cursos)
          AND gi.itemmodule IS NULL
          AND gi.itemtype = 'manual'
          AND gi.itemname IS NOT NULL
        
        UNION ALL
        
        SELECT 
            l.course AS courseid,
            l.id AS activity_id,
            'Lección' AS activity_type,
            l.name AS activity_name,
            CASE 
                WHEN l.available > 0 THEN TO_TIMESTAMP(l.available)
                ELSE TO_TIMESTAMP(c.startdate)
            END AS fecha_entrega,
            CASE 
                WHEN l.deadline > 0 THEN TO_TIMESTAMP(l.deadline)
                ELSE TO_TIMESTAMP(c.enddate)
            END AS fecha_limite
        FROM mdl_lesson l
        JOIN mdl_course c ON l.course = c.id
        WHERE l.course IN ($lista_cursos)
        
        UNION ALL
        
        SELECT 
            g.course AS courseid,
            g.id AS activity_id,
            'Glosario' AS activity_type,
            g.name AS activity_name,
            TO_TIMESTAMP(c.startdate) AS fecha_entrega,
            TO_TIMESTAMP(c.enddate) AS fecha_limite
        FROM mdl_glossary g
        JOIN mdl_course c ON g.course = c.id
        WHERE g.course IN ($lista_cursos)
        
        UNION ALL
        
        SELECT 
            s.course AS courseid,
            s.id AS activity_id,
            'SCORM' AS activity_type,
            s.name AS activity_name,
            CASE 
                WHEN s.timeopen > 0 THEN TO_TIMESTAMP(s.timeopen)
                ELSE TO_TIMESTAMP(c.startdate)
            END AS fecha_entrega,
            CASE 
                WHEN s.timeclose > 0 THEN TO_TIMESTAMP(s.timeclose)
                ELSE TO_TIMESTAMP(c.enddate)
            END AS fecha_limite
        FROM mdl_scorm s
        JOIN mdl_course c ON s.course = c.id
        WHERE s.course IN ($lista_cursos)
        
        UNION ALL
        
        SELECT 
            w.course AS courseid,
            w.id AS activity_id,
            'Taller' AS activity_type,
            w.name AS activity_name,
            CASE 
                WHEN w.submissionstart > 0 THEN TO_TIMESTAMP(w.submissionstart)
                ELSE TO_TIMESTAMP(c.startdate)
            END AS fecha_entrega,
            CASE 
                WHEN w.submissionend > 0 THEN TO_TIMESTAMP(w.submissionend)
                ELSE TO_TIMESTAMP(c.enddate)
            END AS fecha_limite
        FROM mdl_workshop w
        JOIN mdl_course c ON w.course = c.id
        WHERE w.course IN ($lista_cursos)
    ),
    forum_rubrics AS (
        SELECT 
            f.id AS activity_id,
            'Foro' AS activity_type,
            CASE WHEN COUNT(gd.id) > 0 THEN 'SI' ELSE 'NO' END AS tiene_rubrica
        FROM mdl_forum f
        JOIN mdl_course_modules cm ON cm.instance = f.id AND cm.course = f.course
        JOIN mdl_modules m ON m.id = cm.module AND m.name = 'forum'
        JOIN mdl_context ctx ON ctx.instanceid = cm.id AND ctx.contextlevel = 70
        LEFT JOIN mdl_grading_areas ga ON ga.contextid = ctx.id
        LEFT JOIN mdl_grading_definitions gd ON gd.areaid = ga.id AND gd.method = 'rubric'
        WHERE f.course IN ($lista_cursos)
          AND f.name != 'Avisos'
        GROUP BY f.id
    ),
    assign_rubrics AS (
        SELECT 
            a.id AS activity_id,
            'Tarea' AS activity_type,
            CASE WHEN COUNT(gd.id) > 0 THEN 'SI' ELSE 'NO' END AS tiene_rubrica
        FROM mdl_assign a
        JOIN mdl_course_modules cm ON cm.instance = a.id AND cm.course = a.course
        JOIN mdl_modules m ON m.id = cm.module AND m.name = 'assign'
        JOIN mdl_context ctx ON ctx.instanceid = cm.id AND ctx.contextlevel = 70
        LEFT JOIN mdl_grading_areas ga ON ga.contextid = ctx.id
        LEFT JOIN mdl_grading_definitions gd ON gd.areaid = ga.id AND gd.method = 'rubric'
        WHERE a.course IN ($lista_cursos)
        GROUP BY a.id
    ),
    workshop_rubrics AS (
        SELECT 
            w.id AS activity_id,
            'Taller' AS activity_type,
            CASE WHEN COUNT(gd.id) > 0 THEN 'SI' ELSE 'NO' END AS tiene_rubrica
        FROM mdl_workshop w
        JOIN mdl_course_modules cm ON cm.instance = w.id AND cm.course = w.course
        JOIN mdl_modules m ON m.id = cm.module AND m.name = 'workshop'
        JOIN mdl_context ctx ON ctx.instanceid = cm.id AND ctx.contextlevel = 70
        LEFT JOIN mdl_grading_areas ga ON ga.contextid = ctx.id
        LEFT JOIN mdl_grading_definitions gd ON gd.areaid = ga.id AND gd.method = 'rubric'
        WHERE w.course IN ($lista_cursos)
        GROUP BY w.id
    ),
    user_activities AS (
        SELECT 
            au.userid,
            au.courseid,
            aa.activity_id,
            aa.activity_type,
            aa.activity_name,
            aa.fecha_entrega,
            aa.fecha_limite
        FROM all_users au
        JOIN all_activities aa ON au.courseid = aa.courseid
    ),
    quiz_grades AS (
        SELECT 
            gg.userid,
            gi.iteminstance AS quiz_id,
            gg.finalgrade AS final_grade,
            gg.timemodified AS grade_date,
            CASE WHEN gg.feedback IS NOT NULL AND gg.feedback != '' THEN 'SÍ' ELSE 'NO' END AS retroalimentada
        FROM mdl_grade_grades gg
        JOIN mdl_grade_items gi ON gi.id = gg.itemid
        WHERE gi.itemmodule = 'quiz'
          AND gi.courseid IN ($lista_cursos)
    ),
    forum_grades AS (
        SELECT 
            gg.userid,
            gi.iteminstance AS forum_id,
            gg.finalgrade AS final_grade,
            gg.timemodified AS grade_date,
            CASE WHEN gg.feedback IS NOT NULL AND gg.feedback != '' THEN 'SÍ' ELSE 'NO' END AS retroalimentada
        FROM mdl_grade_grades gg
        JOIN mdl_grade_items gi ON gi.id = gg.itemid
        WHERE gi.itemmodule = 'forum'
          AND gi.courseid IN ($lista_cursos)
    ),
    assign_grades AS (
        SELECT 
            gg.userid,
            gi.iteminstance AS assign_id,
            gg.finalgrade AS final_grade,
            gg.timemodified AS grade_date,
            CASE WHEN gg.feedback IS NOT NULL AND gg.feedback != '' THEN 'SÍ' ELSE 'NO' END AS retroalimentada
        FROM mdl_grade_grades gg
        JOIN mdl_grade_items gi ON gi.id = gg.itemid
        WHERE gi.itemmodule = 'assign'
          AND gi.courseid IN ($lista_cursos)
    ),
    lesson_grades AS (
        SELECT 
            gg.userid,
            gi.iteminstance AS lesson_id,
            gg.finalgrade AS final_grade,
            gg.timemodified AS grade_date,
            CASE WHEN gg.feedback IS NOT NULL AND gg.feedback != '' THEN 'SÍ' ELSE 'NO' END AS retroalimentada
        FROM mdl_grade_grades gg
        JOIN mdl_grade_items gi ON gi.id = gg.itemid
        WHERE gi.itemmodule = 'lesson'
          AND gi.courseid IN ($lista_cursos)
    ),
    glossary_grades AS (
        SELECT 
            gg.userid,
            gi.iteminstance AS glossary_id,
            gg.finalgrade AS final_grade,
            gg.timemodified AS grade_date,
            CASE WHEN gg.feedback IS NOT NULL AND gg.feedback != '' THEN 'SÍ' ELSE 'NO' END AS retroalimentada
        FROM mdl_grade_grades gg
        JOIN mdl_grade_items gi ON gi.id = gg.itemid
        WHERE gi.itemmodule = 'glossary'
          AND gi.courseid IN ($lista_cursos)
    ),
    scorm_grades AS (
        SELECT 
            gg.userid,
            gi.iteminstance AS scorm_id,
            gg.finalgrade AS final_grade,
            gg.timemodified AS grade_date,
            CASE WHEN gg.feedback IS NOT NULL AND gg.feedback != '' THEN 'SÍ' ELSE 'NO' END AS retroalimentada
        FROM mdl_grade_grades gg
        JOIN mdl_grade_items gi ON gi.id = gg.itemid
        WHERE gi.itemmodule = 'scorm'
          AND gi.courseid IN ($lista_cursos)
    ),
    workshop_grades AS (
        SELECT 
            gg.userid,
            gi.iteminstance AS workshop_id,
            gg.finalgrade AS final_grade,
            gg.timemodified AS grade_date,
            CASE WHEN gg.feedback IS NOT NULL AND gg.feedback != '' THEN 'SÍ' ELSE 'NO' END AS retroalimentada
        FROM mdl_grade_grades gg
        JOIN mdl_grade_items gi ON gi.id = gg.itemid
        WHERE gi.itemmodule = 'workshop'
          AND gi.courseid IN ($lista_cursos)
    ),
    manual_grades AS (
        SELECT 
            gg.userid,
            gi.id AS grade_item_id,
            gi.courseid,
            gg.finalgrade AS final_grade,
            gg.timemodified AS grade_date,
            CASE WHEN gg.feedback IS NOT NULL AND gg.feedback != '' THEN 'SÍ' ELSE 'NO' END AS retroalimentada
        FROM mdl_grade_grades gg
        JOIN mdl_grade_items gi ON gi.id = gg.itemid
        WHERE gi.itemmodule IS NULL
          AND gi.itemtype = 'manual'
          AND gi.courseid IN ($lista_cursos)
    ),
    interacciones AS (
        SELECT 
            u.id AS userid,
            c.id AS courseid,
            f.id AS activity_id,
            'Foro' AS activity_type,
            COUNT(DISTINCT fp.id) AS visitas,
            COUNT(DISTINCT fp.id) AS envios,
            ROUND(COALESCE(fg.final_grade, 0), 2) AS nota,
            TO_TIMESTAMP(fg.grade_date) AS fecha_calificacion,
            COALESCE(fg.retroalimentada, 'NO') AS retroalimentada,
            TO_TIMESTAMP(MAX(fp.created)) AS fecha_presentacion
        FROM mdl_user u
        JOIN mdl_forum_posts fp ON fp.userid = u.id
        JOIN mdl_forum_discussions fd ON fd.id = fp.discussion
        JOIN mdl_forum f ON f.id = fd.forum
        JOIN mdl_course c ON c.id = f.course
        JOIN mdl_context ctx ON ctx.instanceid = c.id AND ctx.contextlevel = 50
        JOIN mdl_role_assignments ra ON ra.userid = u.id AND ra.contextid = ctx.id
        LEFT JOIN forum_grades fg ON fg.userid = u.id AND fg.forum_id = f.id
        WHERE c.id IN ($lista_cursos)
          AND ra.roleid IN ('5', '9', '16', '17')
          AND f.name != 'Avisos'
          AND fp.created BETWEEN EXTRACT(EPOCH FROM TIMESTAMP '$fecha_inicio') 
                          AND EXTRACT(EPOCH FROM TIMESTAMP '$fecha_fin')
        GROUP BY u.id, c.id, f.id, fg.final_grade, fg.grade_date, fg.retroalimentada
        
        UNION ALL
        
        SELECT 
            qa.userid,
            q.course AS courseid,
            q.id AS activity_id,
            'Quiz' AS activity_type,
            COUNT(DISTINCT qa.id) AS visitas,
            COUNT(DISTINCT qa.id) AS envios,
            ROUND(COALESCE(qg.final_grade, 0), 2) AS nota,
            TO_TIMESTAMP(qg.grade_date) AS fecha_calificacion,
            COALESCE(qg.retroalimentada, 'NO') AS retroalimentada,
            TO_TIMESTAMP(MAX(qa.timefinish)) AS fecha_presentacion
        FROM mdl_quiz q
        LEFT JOIN mdl_quiz_attempts qa ON qa.quiz = q.id AND qa.timefinish > 0
        LEFT JOIN quiz_grades qg ON qg.quiz_id = q.id AND qg.userid = qa.userid
        WHERE q.course IN ($lista_cursos)
          AND (qa.timefinish BETWEEN EXTRACT(EPOCH FROM TIMESTAMP '$fecha_inicio') 
                          AND EXTRACT(EPOCH FROM TIMESTAMP '$fecha_fin') OR qa.timefinish IS NULL)
        GROUP BY qa.userid, q.course, q.id, qg.final_grade, qg.grade_date, qg.retroalimentada
        
        UNION ALL
        
        SELECT 
            sub.userid,
            a.course AS courseid,
            a.id AS activity_id,
            'Tarea' AS activity_type,
            COUNT(DISTINCT sub.id) AS visitas,
            COUNT(DISTINCT sub.id) AS envios,
            ROUND(COALESCE(ag.final_grade, 0), 2) AS nota,
            TO_TIMESTAMP(ag.grade_date) AS fecha_calificacion,
            COALESCE(ag.retroalimentada, 'NO') AS retroalimentada,
            TO_TIMESTAMP(MAX(sub.timemodified)) AS fecha_presentacion
        FROM mdl_assign a
        JOIN mdl_assign_submission sub ON sub.assignment = a.id
        LEFT JOIN assign_grades ag ON ag.userid = sub.userid AND ag.assign_id = a.id
        WHERE a.course IN ($lista_cursos)
          AND (sub.timecreated BETWEEN EXTRACT(EPOCH FROM TIMESTAMP '$fecha_inicio') 
                          AND EXTRACT(EPOCH FROM TIMESTAMP '$fecha_fin') OR sub.timecreated IS NULL)
          AND (sub.status IN ('submitted', 'draft') OR sub.status IS NULL)
        GROUP BY sub.userid, a.course, a.id, ag.final_grade, ag.grade_date, ag.retroalimentada
        
        UNION ALL
        
        SELECT 
            mg.userid,
            mg.courseid,
            gi.id AS activity_id,
            'Calificación Manual' AS activity_type,
            0 AS visitas,
            0 AS envios,
            ROUND(COALESCE(mg.final_grade, 0), 2) AS nota,
            TO_TIMESTAMP(mg.grade_date) AS fecha_calificacion,
            COALESCE(mg.retroalimentada, 'NO') AS retroalimentada,
            TO_TIMESTAMP(mg.grade_date) AS fecha_presentacion
        FROM manual_grades mg
        JOIN mdl_grade_items gi ON gi.id = mg.grade_item_id
        WHERE (mg.grade_date BETWEEN EXTRACT(EPOCH FROM TIMESTAMP '$fecha_inicio') 
                          AND EXTRACT(EPOCH FROM TIMESTAMP '$fecha_fin') OR mg.grade_date IS NULL)
        GROUP BY mg.userid, mg.courseid, gi.id, mg.final_grade, mg.grade_date, mg.retroalimentada
        
        UNION ALL
        
        SELECT 
            la.userid,
            l.course AS courseid,
            l.id AS activity_id,
            'Lección' AS activity_type,
            COUNT(DISTINCT la.id) AS visitas,
            COUNT(DISTINCT la.id) AS envios,
            ROUND(COALESCE(lg.final_grade, 0), 2) AS nota,
            TO_TIMESTAMP(lg.grade_date) AS fecha_calificacion,
            COALESCE(lg.retroalimentada, 'NO') AS retroalimentada,
            TO_TIMESTAMP(MAX(la.timeseen)) AS fecha_presentacion
        FROM mdl_lesson l
        LEFT JOIN mdl_lesson_attempts la ON la.lessonid = l.id
        LEFT JOIN lesson_grades lg ON lg.userid = la.userid AND lg.lesson_id = l.id
        WHERE l.course IN ($lista_cursos)
          AND (la.timeseen BETWEEN EXTRACT(EPOCH FROM TIMESTAMP '$fecha_inicio') 
                          AND EXTRACT(EPOCH FROM TIMESTAMP '$fecha_fin') OR la.timeseen IS NULL)
        GROUP BY la.userid, l.course, l.id, lg.final_grade, lg.grade_date, lg.retroalimentada
        
        UNION ALL
        
        SELECT 
            ge.userid,
            g.course AS courseid,
            g.id AS activity_id,
            'Glosario' AS activity_type,
            COUNT(DISTINCT ge.id) AS visitas,
            COUNT(DISTINCT ge.id) AS envios,
            ROUND(COALESCE(gg.final_grade, 0), 2) AS nota,
            TO_TIMESTAMP(gg.grade_date) AS fecha_calificacion,
            COALESCE(gg.retroalimentada, 'NO') AS retroalimentada,
            TO_TIMESTAMP(MAX(ge.timecreated)) AS fecha_presentacion
        FROM mdl_glossary g
        LEFT JOIN mdl_glossary_entries ge ON ge.glossaryid = g.id AND ge.userid IS NOT NULL
        LEFT JOIN glossary_grades gg ON gg.userid = ge.userid AND gg.glossary_id = g.id
        WHERE g.course IN ($lista_cursos)
          AND (ge.timecreated BETWEEN EXTRACT(EPOCH FROM TIMESTAMP '$fecha_inicio') 
                          AND EXTRACT(EPOCH FROM TIMESTAMP '$fecha_fin') OR ge.timecreated IS NULL)
        GROUP BY ge.userid, g.course, g.id, gg.final_grade, gg.grade_date, gg.retroalimentada
        
        UNION ALL
        
        SELECT 
            st.userid,
            s.course AS courseid,
            s.id AS activity_id,
            'SCORM' AS activity_type,
            COUNT(DISTINCT st.id) AS visitas,
            COUNT(DISTINCT st.id) AS envios,
            ROUND(COALESCE(sg.final_grade, 0), 2) AS nota,
            TO_TIMESTAMP(sg.grade_date) AS fecha_calificacion,
            COALESCE(sg.retroalimentada, 'NO') AS retroalimentada,
            TO_TIMESTAMP(MAX(st.timemodified)) AS fecha_presentacion
        FROM mdl_scorm s
        LEFT JOIN mdl_scorm_scoes_track st ON st.scormid = s.id
        LEFT JOIN scorm_grades sg ON sg.userid = st.userid AND sg.scorm_id = s.id
        WHERE s.course IN ($lista_cursos)
          AND (st.timemodified BETWEEN EXTRACT(EPOCH FROM TIMESTAMP '$fecha_inicio') 
                          AND EXTRACT(EPOCH FROM TIMESTAMP '$fecha_fin') OR st.timemodified IS NULL)
        GROUP BY st.userid, s.course, s.id, sg.final_grade, sg.grade_date, sg.retroalimentada
        
        UNION ALL
        
        SELECT 
            sub.authorid AS userid,
            w.course AS courseid,
            w.id AS activity_id,
            'Taller' AS activity_type,
            COUNT(DISTINCT sub.id) AS visitas,
            COUNT(DISTINCT sub.id) AS envios,
            ROUND(COALESCE(wg.final_grade, 0), 2) AS nota,
            TO_TIMESTAMP(wg.grade_date) AS fecha_calificacion,
            COALESCE(wg.retroalimentada, 'NO') AS retroalimentada,
            TO_TIMESTAMP(MAX(sub.timemodified)) AS fecha_presentacion
        FROM mdl_workshop w
        LEFT JOIN mdl_workshop_submissions sub ON sub.workshopid = w.id
        LEFT JOIN workshop_grades wg ON wg.userid = sub.authorid AND wg.workshop_id = w.id
        WHERE w.course IN ($lista_cursos)
          AND (sub.timemodified BETWEEN EXTRACT(EPOCH FROM TIMESTAMP '$fecha_inicio') 
                          AND EXTRACT(EPOCH FROM TIMESTAMP '$fecha_fin') OR sub.timemodified IS NULL)
        GROUP BY sub.authorid, w.course, w.id, wg.final_grade, wg.grade_date, wg.retroalimentada
    )
    SELECT 
        u.username AS codigo,
        u.firstname AS nombre,
        u.lastname AS apellidos,
        r.name AS rol,
        u.email AS correo,
        c.id AS id_curso,
        c.fullname AS curso,
        ua.activity_type AS tipo_actividad,
        ua.activity_name AS nombre_actividad,
        TO_CHAR(ua.fecha_entrega, 'YYYY-MM-DD HH24:MI:SS') AS fecha_entrega,
        TO_CHAR(ua.fecha_limite, 'YYYY-MM-DD HH24:MI:SS') AS fecha_limite,
        TO_CHAR(i.fecha_presentacion, 'YYYY-MM-DD HH24:MI:SS') AS fecha_presentacion,
        COALESCE(
            (SELECT s.semana_nombre 
             FROM semanas_academicas s 
             WHERE i.fecha_presentacion BETWEEN s.inicio_semana AND s.fin_semana),
            'NO PRESENTADO'
        ) AS semana_presentacion,
        COALESCE(i.visitas, 0) AS visitas,
        COALESCE(i.envios, 0) AS envios,
        COALESCE(i.nota, 0) AS nota,
        TO_CHAR(i.fecha_calificacion, 'YYYY-MM-DD HH24:MI:SS') AS fecha_calificacion,
        COALESCE(
            (SELECT s.semana_nombre 
             FROM semanas_academicas s 
             WHERE i.fecha_calificacion BETWEEN s.inicio_semana AND s.fin_semana),
            'SIN CALIFICAR'
        ) AS semana_calificacion,
        COALESCE(i.retroalimentada, 'NO') AS retroalimentada,
        CASE 
            WHEN ua.activity_type = 'Foro' THEN COALESCE(fr.tiene_rubrica, 'NO')
            WHEN ua.activity_type = 'Tarea' THEN COALESCE(ar.tiene_rubrica, 'NO')
            WHEN ua.activity_type = 'Taller' THEN COALESCE(wr.tiene_rubrica, 'NO')
            ELSE 'NO'
        END AS retroalimentacion_rubrica
    FROM user_activities ua
    JOIN mdl_user u ON ua.userid = u.id
    JOIN mdl_course c ON ua.courseid = c.id
    JOIN mdl_role_assignments ra ON ra.userid = u.id AND ra.contextid = (SELECT contextid FROM all_users au WHERE au.userid = u.id AND au.courseid = c.id)
    JOIN mdl_role r ON ra.roleid = r.id
    LEFT JOIN interacciones i ON ua.userid = i.userid 
        AND ua.courseid = i.courseid 
        AND ua.activity_id = i.activity_id 
        AND ua.activity_type = i.activity_type
    LEFT JOIN forum_rubrics fr ON ua.activity_id = fr.activity_id AND ua.activity_type = 'Foro'
    LEFT JOIN assign_rubrics ar ON ua.activity_id = ar.activity_id AND ua.activity_type = 'Tarea'
    LEFT JOIN workshop_rubrics wr ON ua.activity_id = wr.activity_id AND ua.activity_type = 'Taller'
    ORDER BY c.id, u.lastname, u.firstname, ua.activity_type, ua.activity_name;";

    return $sql;
}

/**
 * Genera el archivo Excel con los resultados de la consulta
 */
function generarExcel($resultados, $semanas) {
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    
    // Configurar encabezados
    $headers = [
        'Código', 'Nombre', 'Apellidos', 'Rol', 'Correo', 
        'ID Curso', 'Curso', 'Tipo Actividad', 'Nombre Actividad',
        'Fecha Entrega', 'Fecha Límite', 'Fecha Presentación',
        'Semana Presentación', 'Visitas', 'Envíos', 'Nota',
        'Fecha Calificación', 'Semana Calificación', 'Retroalimentada',
        'Retroalimentación Rúbrica'
    ];
    
    $sheet->fromArray($headers, NULL, 'A1');
    
    // Estilo para los encabezados
    $headerStyle = [
        'font' => ['bold' => true, 'color' => ['rgb' => 'FFFFFF']],
        'alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER],
        'fill' => ['fillType' => Fill::FILL_SOLID, 'startColor' => ['rgb' => '4472C4']]
    ];
    $sheet->getStyle('A1:T1')->applyFromArray($headerStyle);
    
    // Llenar datos
    $row = 2;
    foreach ($resultados as $fila) {
        $sheet->fromArray($fila, NULL, "A{$row}");
        $row++;
    }
    
    // Autoajustar columnas
    foreach (range('A', 'T') as $col) {
        $sheet->getColumnDimension($col)->setAutoSize(true);
    }
    
    // Guardar archivo Excel
    $fecha_actual = date('Y-m-d');
    $nombre_archivo = "actividades_asignaturas_pregrado_{$fecha_actual}.xlsx";
    $writer = new Xlsx($spreadsheet);
    $writer->save($nombre_archivo);
    
    return $nombre_archivo;
}

/**
 * Envía los correos electrónicos con el archivo adjunto usando Postfix
 */
function enviarCorreos($archivo_excel, $correos_destino, $correo_notificacion, $modo_prueba = false) {
    $fecha_para_nombre = date('d-m-Y');
    $asunto = 'Reporte de Actividades de asignaturas de Pregrado - ' . $fecha_para_nombre;
    
    if ($modo_prueba) {
        $asunto = '[PRUEBA] ' . $asunto;
    }
    
    try {
        // Configurar y enviar correo con PHPMailer usando Postfix
        $mail = new PHPMailer(true);
        $mail->isSendmail(); // Usar sendmail/postfix
        
        $mail->setFrom('noreply-univirtual@utp.edu.co', 'Reporte Moodle');
        foreach ($correos_destino as $correo) {
            $mail->addAddress($correo);
        }
        
        $mail->Subject = $asunto;
        $mail->Body = 'Cordial Saludo, ' . ($modo_prueba ? 'ESTE ES UN CORREO DE PRUEBA. ' : '') . 
                      'Adjunto el Reporte de Actividades de asignaturas de Pregrado.';
        
        if (file_exists($archivo_excel)) {
            $mail->addAttachment($archivo_excel, "actividades_asignaturas_pregrado_{$fecha_para_nombre}.xlsx");
        }
        
        $mail->send();

        // Enviar notificación de éxito
        $mensaje_exito = 'El reporte de actividades de pregrado fue enviado correctamente.' . 
                         ($modo_prueba ? ' (MODO PRUEBA)' : '');
        mail($correo_notificacion, 'Estado Reporte Actividades Pregrado', $mensaje_exito);
    } catch (Exception $e) {
        // Enviar notificación de error
        $mensaje_error = 'Error: ' . $e->getMessage() . ($modo_prueba ? ' (MODO PRUEBA)' : '');
        mail($correo_notificacion, 'Error Reporte Actividades Pregrado', $mensaje_error);
        throw $e;
    }
}
