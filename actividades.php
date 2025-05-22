<?php
// Configuración inicial
date_default_timezone_set('America/Bogota');
require 'vendor/autoload.php';

use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\Exception;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// Configuración de la base de datos
$db_host = 'localhost';
$db_port = '5432';
$db_name = 'moodle';
$db_user = 'moodle';
$db_pass = 'M00dl3';

// Configuración de correo
$correo_destino = ['daniel.pardo@utp.edu.co'];
$correo_notificacion = 'daniel.pardo@utp.edu.co';
$remitente = 'noreply-univirtual@utp.edu.co';

// Parámetros de ejecución
$semanas_totales = 16; // 16 semanas académicas
$semana_inicial = [
    'numero' => 1,
    'inicio' => '2025-02-03 00:00:00',
    'fin' => '2025-02-10 23:59:59'
];

// Función para calcular las fechas de las semanas
function calcularSemanas($semana_inicial, $semanas_totales) {
    $semanas = [];
    $inicio = new DateTime($semana_inicial['inicio']);
    $fin = new DateTime($semana_inicial['fin']);
    
    for ($i = 1; $i <= $semanas_totales; $i++) {
        $semanas[$i] = [
            'numero' => $i,
            'nombre' => 'SEMANA ' . $i,
            'inicio' => $inicio->format('Y-m-d H:i:s'),
            'fin' => $fin->format('Y-m-d H:i:s')
        ];
        
        // Avanzar al siguiente período (6 días para semana 1, 7 días para las demás)
        $inicio->add(new DateInterval($i == 1 ? 'P7D' : 'P6D'));
        $fin->add(new DateInterval('P6D'));
    }
    
    return $semanas;
}

// Función para ejecutar la consulta SQL
function ejecutarConsulta($db_conn, $semana_actual) {
    $sql = "
    WITH 
        all_users AS (
            SELECT DISTINCT 
                u.id AS userid, 
                c.id AS courseid
            FROM mdl_user u
            JOIN mdl_role_assignments ra ON ra.userid = u.id
            JOIN mdl_context mc ON ra.contextid = mc.id AND mc.contextlevel = 50
            JOIN mdl_course c ON mc.instanceid = c.id 
            WHERE c.id IN ('494','415','507','481','508','482','509','485','526','510','511','486','490','416','503','504','527','417','496','497','418','498','419','475','421','420','422','423','512','513','515','488','489','424','516','517','491','518','492','519','493','520','425','476','426','505','506','479','521','428','430','522','495','499','431','453','500','523','434','524','435','436','437','438','440','502','439','452','525','442')
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
            WHERE f.course IN ('494','415','507','481','508','482','509','485','526','510','511','486','490','416','503','504','527','417','496','497','418','498','419','475','421','420','422','423','512','513','515','488','489','424','516','517','491','518','492','519','493','520','425','476','426','505','506','479','521','428','430','522','495','499','431','453','500','523','434','524','435','436','437','438','440','502','439','452','525','442')
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
            WHERE q.course IN ('494','415','507','481','508','482','509','485','526','510','511','486','490','416','503','504','527','417','496','497','418','498','419','475','421','420','422','423','512','513','515','488','489','424','516','517','491','518','492','519','493','520','425','476','426','505','506','479','521','428','430','522','495','499','431','453','500','523','434','524','435','436','437','438','440','502','439','452','525','442')
            
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
            WHERE a.course IN ('494','415','507','481','508','482','509','485','526','510','511','486','490','416','503','504','527','417','496','497','418','498','419','475','421','420','422','423','512','513','515','488','489','424','516','517','491','518','492','519','493','520','425','476','426','505','506','479','521','428','430','522','495','499','431','453','500','523','434','524','435','436','437','438','440','502','439','452','525','442')
            
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
            WHERE gi.courseid IN ('494','415','507','481','508','482','509','485','526','510','511','486','490','416','503','504','527','417','496','497','418','498','419','475','421','420','422','423','512','513','515','488','489','424','516','517','491','518','492','519','493','520','425','476','426','505','506','479','521','428','430','522','495','499','431','453','500','523','434','524','435','436','437','438','440','502','439','452','525','442')
              AND gi.itemmodule IS NULL
              AND gi.itemtype = 'manual'
              AND gi.itemname IS NOT NULL
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
            WHERE f.course IN ('494','415','507','481','508','482','509','485','526','510','511','486','490','416','503','504','527','417','496','497','418','498','419','475','421','420','422','423','512','513','515','488','489','424','516','517','491','518','492','519','493','520','425','476','426','505','506','479','521','428','430','522','495','499','431','453','500','523','434','524','435','436','437','438','440','502','439','452','525','442')
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
            WHERE a.course IN ('494','415','507','481','508','482','509','485','526','510','511','486','490','416','503','504','527','417','496','497','418','498','419','475','421','420','422','423','512','513','515','488','489','424','516','517','491','518','492','519','493','520','425','476','426','505','506','479','521','428','430','522','495','499','431','453','500','523','434','524','435','436','437','438','440','502','439','452','525','442')
            GROUP BY a.id
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
              AND gi.courseid IN ('494','415','507','481','508','482','509','485','526','510','511','486','490','416','503','504','527','417','496','497','418','498','419','475','421','420','422','423','512','513','515','488','489','424','516','517','491','518','492','519','493','520','425','476','426','505','506','479','521','428','430','522','495','499','431','453','500','523','434','524','435','436','437','438','440','502','439','452','525','442')
        ),
        interacciones AS (
            SELECT 
                u.id AS userid,
                c.id AS courseid,
                f.id AS activity_id,
                'Foro' AS activity_type,
                COUNT(DISTINCT fp.id) AS visitas,
                COUNT(DISTINCT fp.id) AS envios,
                ROUND(COALESCE(gg.finalgrade, 0), 2) AS nota,
                TO_TIMESTAMP(gg.timemodified) AS fecha_calificacion,
                CASE WHEN gg.feedback IS NOT NULL AND gg.feedback != '' THEN 'SÍ' ELSE 'NO' END AS retroalimentada,
                TO_TIMESTAMP(MAX(fp.created)) AS fecha_presentacion
            FROM mdl_user u
            JOIN mdl_forum_posts fp ON fp.userid = u.id
            JOIN mdl_forum_discussions fd ON fd.id = fp.discussion
            JOIN mdl_forum f ON f.id = fd.forum
            JOIN mdl_course c ON c.id = f.course
            JOIN mdl_context ctx ON ctx.instanceid = c.id AND ctx.contextlevel = 50
            JOIN mdl_role_assignments ra ON ra.userid = u.id AND ra.contextid = ctx.id
            LEFT JOIN mdl_grade_items gi ON gi.courseid = c.id AND gi.itemmodule = 'forum' AND gi.iteminstance = f.id
            LEFT JOIN mdl_grade_grades gg ON gg.itemid = gi.id AND gg.userid = u.id
            WHERE c.id IN ('494','415','507','481','508','482','509','485','526','510','511','486','490','416','503','504','527','417','496','497','418','498','419','475','421','420','422','423','512','513','515','488','489','424','516','517','491','518','492','519','493','520','425','476','426','505','506','479','521','428','430','522','495','499','431','453','500','523','434','524','435','436','437','438','440','502','439','452','525','442')
              AND ra.roleid IN ('5','9','16','17')
              AND f.name != 'Avisos'
              AND fp.created BETWEEN EXTRACT(EPOCH FROM TIMESTAMP '".$semana_actual['inicio']."') 
                                  AND EXTRACT(EPOCH FROM TIMESTAMP '".$semana_actual['fin']."')
            GROUP BY u.id, c.id, f.id, gg.finalgrade, gg.feedback, gg.timemodified
            
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
            LEFT JOIN mdl_quiz_attempts qa ON qa.quiz = q.id
            LEFT JOIN quiz_grades qg ON qg.quiz_id = q.id AND qg.userid = qa.userid
            WHERE q.course IN ('494','415','507','481','508','482','509','485','526','510','511','486','490','416','503','504','527','417','496','497','418','498','419','475','421','420','422','423','512','513','515','488','489','424','516','517','491','518','492','519','493','520','425','476','426','505','506','479','521','428','430','522','495','499','431','453','500','523','434','524','435','436','437','438','440','502','439','452','525','442')
              AND (qa.timefinish BETWEEN EXTRACT(EPOCH FROM TIMESTAMP '".$semana_actual['inicio']."') 
                                     AND EXTRACT(EPOCH FROM TIMESTAMP '".$semana_actual['fin']."') OR qa.timefinish IS NULL)
            GROUP BY qa.userid, q.course, q.id, qg.final_grade, qg.grade_date, qg.retroalimentada
            
            UNION ALL
            
            SELECT 
                sub.userid,
                a.course AS courseid,
                a.id AS activity_id,
                'Tarea' AS activity_type,
                COUNT(DISTINCT sub.id) AS visitas,
                COUNT(DISTINCT sub.id) AS envios,
                ROUND(COALESCE(gg.finalgrade, 0), 2) AS nota,
                TO_TIMESTAMP(gg.timemodified) AS fecha_calificacion,
                CASE WHEN gg.feedback IS NOT NULL AND gg.feedback != '' THEN 'SÍ' ELSE 'NO' END AS retroalimentada,
                TO_TIMESTAMP(MAX(sub.timemodified)) AS fecha_presentacion
            FROM mdl_assign a
            JOIN mdl_assign_submission sub ON sub.assignment = a.id
            LEFT JOIN mdl_grade_items gi ON gi.itemmodule = 'assign' AND gi.iteminstance = a.id
            LEFT JOIN mdl_grade_grades gg ON gg.userid = sub.userid AND gg.itemid = gi.id
            WHERE a.course IN ('494','415','507','481','508','482','509','485','526','510','511','486','490','416','503','504','527','417','496','497','418','498','419','475','421','420','422','423','512','513','515','488','489','424','516','517','491','518','492','519','493','520','425','476','426','505','506','479','521','428','430','522','495','499','431','453','500','523','434','524','435','436','437','438','440','502','439','452','525','442')
              AND (sub.timecreated BETWEEN EXTRACT(EPOCH FROM TIMESTAMP '".$semana_actual['inicio']."') 
                                        AND EXTRACT(EPOCH FROM TIMESTAMP '".$semana_actual['fin']."') OR sub.timecreated IS NULL)
              AND (sub.status IN ('submitted', 'draft') OR sub.status IS NULL)
            GROUP BY sub.userid, a.course, a.id, gg.finalgrade, gg.feedback, gg.timemodified
            
            UNION ALL
            
            SELECT 
                gg.userid,
                gi.courseid AS courseid,
                gi.id AS activity_id,
                'Calificación Manual' AS activity_type,
                0 AS visitas,
                0 AS envios,
                ROUND(COALESCE(gg.finalgrade, 0), 2) AS nota,
                TO_TIMESTAMP(gg.timemodified) AS fecha_calificacion,
                CASE WHEN gg.feedback IS NOT NULL AND gg.feedback != '' THEN 'SÍ' ELSE 'NO' END AS retroalimentada,
                TO_TIMESTAMP(gg.timemodified) AS fecha_presentacion
            FROM mdl_grade_items gi
            LEFT JOIN mdl_grade_grades gg ON gg.itemid = gi.id
            WHERE gi.courseid IN ('494','415','507','481','508','482','509','485','526','510','511','486','490','416','503','504','527','417','496','497','418','498','419','475','421','420','422','423','512','513','515','488','489','424','516','517','491','518','492','519','493','520','425','476','426','505','506','479','521','428','430','522','495','499','431','453','500','523','434','524','435','436','437','438','440','502','439','452','525','442')
              AND gi.itemmodule IS NULL
              AND gi.itemtype = 'manual'
              AND gi.itemname IS NOT NULL
              AND (gg.timemodified BETWEEN EXTRACT(EPOCH FROM TIMESTAMP '".$semana_actual['inicio']."') 
                                        AND EXTRACT(EPOCH FROM TIMESTAMP '".$semana_actual['fin']."') OR gg.timemodified IS NULL)
            GROUP BY gg.userid, gi.courseid, gi.id, gg.finalgrade, gg.feedback, gg.timemodified
        ),
        semanas_academicas AS (
            SELECT 
                generate_series(1, 16) AS semana_numero,
                'SEMANA ' || generate_series(1, 16) AS semana_nombre,
                (TIMESTAMP '2025-02-05 00:00:00' + 
                    (CASE 
                        WHEN generate_series(1, 16) = 1 THEN INTERVAL '0 days'
                        ELSE INTERVAL '7 days' + INTERVAL '6 days' * (generate_series(1, 16) - 2)
                    END)) AS inicio_semana,
                (TIMESTAMP '2025-02-05 00:00:00' + 
                    (CASE 
                        WHEN generate_series(1, 16) = 1 THEN INTERVAL '5 days 23 hours 59 minutes 59 seconds'
                        ELSE INTERVAL '7 days' + INTERVAL '6 days' * (generate_series(1, 16) - 2) + INTERVAL '6 days 23 hours 59 minutes 59 seconds'
                    END)) AS fin_semana
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
                'NO CALIFICADO'
            ) AS semana_calificacion,
            COALESCE(i.retroalimentada, 'NO') AS retroalimentada,
            CASE 
                WHEN ua.activity_type = 'Foro' THEN COALESCE(fr.tiene_rubrica, 'NO')
                WHEN ua.activity_type = 'Tarea' THEN COALESCE(ar.tiene_rubrica, 'NO')
                ELSE 'NO'
            END AS retroalimentacion_rubrica
        FROM user_activities ua
        JOIN mdl_user u ON ua.userid = u.id
        JOIN mdl_course c ON ua.courseid = c.id
        JOIN mdl_role_assignments ra ON ra.userid = u.id 
            AND ra.contextid = (SELECT id FROM mdl_context WHERE contextlevel = 50 AND instanceid = c.id)
        JOIN mdl_role r ON ra.roleid = r.id
        LEFT JOIN interacciones i ON ua.userid = i.userid 
            AND ua.courseid = i.courseid 
            AND ua.activity_id = i.activity_id 
            AND ua.activity_type = i.activity_type
        LEFT JOIN forum_rubrics fr ON ua.activity_id = fr.activity_id AND ua.activity_type = 'Foro'
        LEFT JOIN assign_rubrics ar ON ua.activity_id = ar.activity_id AND ua.activity_type = 'Tarea'
        ORDER BY c.id, u.lastname, u.firstname, ua.activity_type, ua.activity_name;
    ";
    
    return pg_query($db_conn, $sql);
}

// Función para generar el archivo Excel
function generarExcel($resultados, $semana_actual) {
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    
    // Encabezados
    $sheet->setCellValue('A1', 'Código');
    $sheet->setCellValue('B1', 'Nombre');
    $sheet->setCellValue('C1', 'Apellidos');
    $sheet->setCellValue('D1', 'Rol');
    $sheet->setCellValue('E1', 'Correo');
    $sheet->setCellValue('F1', 'ID Curso');
    $sheet->setCellValue('G1', 'Curso');
    $sheet->setCellValue('H1', 'Tipo Actividad');
    $sheet->setCellValue('I1', 'Nombre Actividad');
    $sheet->setCellValue('J1', 'Fecha Entrega');
    $sheet->setCellValue('K1', 'Fecha Límite');
    $sheet->setCellValue('L1', 'Fecha Presentación');
    $sheet->setCellValue('M1', 'Semana Presentación');
    $sheet->setCellValue('N1', 'Visitas');
    $sheet->setCellValue('O1', 'Envíos');
    $sheet->setCellValue('P1', 'Nota');
    $sheet->setCellValue('Q1', 'Fecha Calificación');
    $sheet->setCellValue('R1', 'Semana Calificación');
    $sheet->setCellValue('S1', 'Retroalimentada');
    $sheet->setCellValue('T1', 'Retroalimentación Rúbrica');
    
    // Datos
    $row = 2;
    while ($fila = pg_fetch_assoc($resultados)) {
        $sheet->setCellValue('A'.$row, $fila['codigo']);
        $sheet->setCellValue('B'.$row, $fila['nombre']);
        $sheet->setCellValue('C'.$row, $fila['apellidos']);
        $sheet->setCellValue('D'.$row, $fila['rol']);
        $sheet->setCellValue('E'.$row, $fila['correo']);
        $sheet->setCellValue('F'.$row, $fila['id_curso']);
        $sheet->setCellValue('G'.$row, $fila['curso']);
        $sheet->setCellValue('H'.$row, $fila['tipo_actividad']);
        $sheet->setCellValue('I'.$row, $fila['nombre_actividad']);
        $sheet->setCellValue('J'.$row, $fila['fecha_entrega']);
        $sheet->setCellValue('K'.$row, $fila['fecha_limite']);
        $sheet->setCellValue('L'.$row, $fila['fecha_presentacion']);
        $sheet->setCellValue('M'.$row, $fila['semana_presentacion']);
        $sheet->setCellValue('N'.$row, $fila['visitas']);
        $sheet->setCellValue('O'.$row, $fila['envios']);
        $sheet->setCellValue('P'.$row, $fila['nota']);
        $sheet->setCellValue('Q'.$row, $fila['fecha_calificacion']);
        $sheet->setCellValue('R'.$row, $fila['semana_calificacion']);
        $sheet->setCellValue('S'.$row, $fila['retroalimentada']);
        $sheet->setCellValue('T'.$row, $fila['retroalimentacion_rubrica']);
        $row++;
    }
    
    // Autoajustar columnas
    foreach(range('A','T') as $columnID) {
        $sheet->getColumnDimension($columnID)->setAutoSize(true);
    }
    
    // Guardar archivo
    $filename = 'reporte_actividades_'.str_replace(' ', '_', $semana_actual['nombre']).'_'.date('Y-m-d').'.xlsx';
    $writer = new Xlsx($spreadsheet);
    $writer->save($filename);
    
    return $filename;
}

// Función para enviar correo
function enviarCorreo($filename, $semana_actual, $resultados, $exito = true, $mensaje = '') {
    global $correo_destino, $correo_notificacion, $remitente;
    
    $mail = new PHPMailer(true);
    
    try {
        // Configuración del servidor
        $mail->isSMTP();
        $mail->Host = 'localhost';
        $mail->SMTPAuth = false;
        $mail->Port = 25;
        
        // Remitente y destinatario
        $mail->setFrom($remitente, 'Reporte Moodle');
        foreach ($correo_destino as $correo) {
            $mail->addAddress($correo);
        }
        
        if ($exito) {
            // Formatear fechas
            $fecha_inicio = (new DateTime($semana_actual['inicio']))->format('d/m/Y');
            $fecha_fin = (new DateTime($semana_actual['fin']))->format('d/m/Y');
            
            // Contenido del correo con archivo adjunto
            $mail->isHTML(true);
            $mail->Subject = 'Reporte de Actividades Semanal - ' . $semana_actual['nombre'];
            $mail->Body = "Cordial Saludo,<br><br>
                         Adjunto el Reporte de Actividades Semanal de asignaturas de pregrado.<br><br>
                         <strong>Período del reporte:</strong> $fecha_inicio a $fecha_fin<br>
                         <strong>Total registros:</strong> " . pg_num_rows($resultados) . "<br>";
            
            // Adjuntar archivo
            $mail->addAttachment($filename, "reporte_actividades_{$semana_actual['nombre']}.xlsx");
            
            $mail->send();
            
            // Enviar notificación de éxito
            mail($correo_notificacion, 'Reporte de Actividades Exitoso', 
                "El reporte de actividades fue generado correctamente.\n" .
                "Período: $fecha_inicio a $fecha_fin\n" .
                "Total registros: " . pg_num_rows($resultados) . "\n" .
                "Semana: " . $semana_actual['nombre']);
        } else {
            // Correo de error
            $mail->Subject = 'ERROR: Reporte de Actividades - ' . $semana_actual['nombre'];
            $mail->Body = "Ocurrió un error al generar el reporte para la {$semana_actual['nombre']}:<br><br>" . $mensaje;
            $mail->send();
            
            // Enviar notificación de error
            mail($correo_notificacion, 'Error en Reporte de Actividades', 
                "Error: " . $mensaje . "\n" .
                "Semana: " . $semana_actual['nombre'] . "\n" .
                "Hora: " . date('Y-m-d H:i:s'));
        }
        
        return true;
    } catch (Exception $e) {
        error_log("Error al enviar correo: ".$e->getMessage());
        return false;
    }
}

// Función principal
function main() {
    global $db_host, $db_port, $db_name, $db_user, $db_pass, $semanas_totales, $semana_inicial;
    
    try {
        // Calcular semana actual
        $semanas = calcularSemanas($semana_inicial, $semanas_totales);
        $hoy = new DateTime();
        $semana_actual = null;
        
        foreach ($semanas as $semana) {
            $inicio = new DateTime($semana['inicio']);
            $fin = new DateTime($semana['fin']);
            
            if ($hoy >= $inicio && $hoy <= $fin) {
                $semana_actual = $semana;
                break;
            }
        }
        
        if (!$semana_actual) {
            throw new Exception("No se encontró una semana académica activa para la fecha actual.");
        }
        
        // Conectar a la base de datos
        $db_conn = pg_connect("host=$db_host port=$db_port dbname=$db_name user=$db_user password=$db_pass");
        if (!$db_conn) {
            throw new Exception("No se pudo conectar a la base de datos.");
        }
        
        // Ejecutar consulta
        $resultados = ejecutarConsulta($db_conn, $semana_actual);
        if (!$resultados) {
            throw new Exception("Error al ejecutar la consulta SQL: " . pg_last_error($db_conn));
        }
        
        // Generar Excel
        $filename = generarExcel($resultados, $semana_actual);
        
        // Enviar correo con archivo
        if (!enviarCorreo($filename, $semana_actual, $resultados)) {
            throw new Exception("Error al enviar el correo con el archivo adjunto.");
        }
        
        // Limpiar
        pg_close($db_conn);
        unlink($filename);
        
        echo "Proceso completado con éxito para la ".$semana_actual['nombre'];
    } catch (Exception $e) {
        // Enviar correo de error
        if (isset($semana_actual)) {
            enviarCorreo(null, $semana_actual, null, false, $e->getMessage());
        } else {
            // Si no se pudo determinar la semana, enviar correo genérico
            mail($correo_notificacion, 'ERROR: Reporte de Actividades', 
                "Ocurrió un error grave en el script de generación de reportes:\n\n".$e->getMessage());
        }
        
        error_log("Error en el script: ".$e->getMessage());
        echo "Error: ".$e->getMessage();
        
        if (isset($db_conn)) {
            pg_close($db_conn);
        }
        
        if (isset($filename) && file_exists($filename)) {
            unlink($filename);
        }
    }
}

// Ejecutar script
main();
?>
