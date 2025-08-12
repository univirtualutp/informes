<?php
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\Exception;

// Establecer la zona horaria
date_default_timezone_set('America/Bogota');

// Configuración de la base de datos
$host = 'localhost';
$dbname = 'moodle';
$user = 'moodle';
$pass = 'M00dl3';

// Configuración de correos 
$correo_destino = ['soporteunivirtual@utp.edu.co','pedagogiaunivirtual@utp.edu.co','univirtual-utp@utp.edu.co'];
$correo_notificacion = 'soporteunivirtual@utp.edu.co';

// Definición de cursos en variable 
$cursos_seleccionados = [
    '786','583','804','596','820','790','821','819','789','580','799','616','617','797','798',
    '842','844','621','805','584','581','802','619','627','800','801','618','588','815','589',
    '590','594','595','787','788','605','607','793','573','791','606','792','608','609','795',
    '611','794','610','586','814','623','622','624','733','574','604','796','576','612','577',
    '614','830','615','783','579','784','785','591','587','810','625','582','803','620','592',
    '816','817','593'
];

// Rangos de fechas
$hoy = new DateTime();
$lunes = clone $hoy;
$lunes->modify('last monday');
$fecha_inicio = new DateTime(date('Y') . '-08-04 00:00:00');
$fecha_fin = clone $lunes;
$fecha_fin->setTime(23, 59, 59);

try {
    $pdo = new PDO("pgsql:host=$host;dbname=$dbname", $user, $pass);
    $pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);

    // Lista de grupos de cursos usando la variable definida arriba
    $grupos_cursos = [
        'actividades_pregrado' => $cursos_seleccionados
    ];

    // Parámetros para las consultas
    $params = [
        ':fecha_inicio' => $fecha_inicio->format('Y-m-d H:i:s'),
        ':fecha_fin' => $fecha_fin->format('Y-m-d H:i:s')
    ];

    // Directorio temporal
    $temp_dir = sys_get_temp_dir() . '/reportes_moodle_' . uniqid();
    if (!mkdir($temp_dir)) {
        throw new Exception("No se pudo crear el directorio temporal: $temp_dir");
    }

    $archivos_excel = [];

    // Consulta SQL modificada (usando la variable de cursos)
    $sql = "
    WITH 
        all_users AS (
            SELECT DISTINCT 
                u.id AS userid, 
                c.id AS courseid
            FROM mdl_user u
            JOIN mdl_role_assignments ra ON ra.userid = u.id
            JOIN mdl_context mc ON ra.contextid = mc.id AND mc.contextlevel = 50
            JOIN mdl_course c ON mc.instanceid = c.id AND c.id IN (" . implode(',', $cursos_seleccionados) . ")
            WHERE u.username <> '12345678'
              AND ra.roleid IN ('3','5', '9', '16','17')
        ),
        all_activities AS (
            SELECT 
                f.course AS courseid,
                f.id AS activity_id,
                'Foro' AS activity_type,
                f.name AS activity_name,
                TO_TIMESTAMP(f.duedate) AS fecha_apertura,
                TO_TIMESTAMP(f.cutoffdate) AS fecha_cierre
            FROM mdl_forum f
            WHERE f.course IN (" . implode(',', $cursos_seleccionados) . ") AND f.name != 'Avisos'
            UNION ALL
            SELECT 
                l.course AS courseid,
                l.id AS activity_id,
                'Lección' AS activity_type,
                l.name AS activity_name,
                TO_TIMESTAMP(l.available) AS fecha_apertura,
                TO_TIMESTAMP(l.deadline) AS fecha_cierre
            FROM mdl_lesson l
            WHERE l.course IN (" . implode(',', $cursos_seleccionados) . ")
            UNION ALL
            SELECT 
                a.course AS courseid,
                a.id AS activity_id,
                'Tarea' AS activity_type,
                a.name AS activity_name,
                TO_TIMESTAMP(a.allowsubmissionsfromdate) AS fecha_apertura,
                TO_TIMESTAMP(a.duedate) AS fecha_cierre
            FROM mdl_assign a
            WHERE a.course IN (" . implode(',', $cursos_seleccionados) . ")
            UNION ALL
            SELECT 
                q.course AS courseid,
                q.id AS activity_id,
                'Quiz' AS activity_type,
                q.name AS activity_name,
                TO_TIMESTAMP(q.timeopen) AS fecha_apertura,
                TO_TIMESTAMP(q.timeclose) AS fecha_cierre
            FROM mdl_quiz q
            WHERE q.course IN (" . implode(',', $cursos_seleccionados) . ")
            UNION ALL
            SELECT 
                g.course AS courseid,
                g.id AS activity_id,
                'Glosario' AS activity_type,
                g.name AS activity_name,
                TO_TIMESTAMP(g.timecreated) AS fecha_apertura,
                TO_TIMESTAMP(g.timemodified) AS fecha_cierre
            FROM mdl_glossary g
            WHERE g.course IN (" . implode(',', $cursos_seleccionados) . ")
            UNION ALL
            SELECT 
                s.course AS courseid,
                s.id AS activity_id,
                'SCORM' AS activity_type,
                s.name AS activity_name,
                TO_TIMESTAMP(s.timeopen) AS fecha_apertura,
                TO_TIMESTAMP(s.timeclose) AS fecha_cierre
            FROM mdl_scorm s
            WHERE s.course IN (" . implode(',', $cursos_seleccionados) . ")
        ),
        user_activities AS (
            SELECT 
                au.userid,
                au.courseid,
                aa.activity_id,
                aa.activity_type,
                aa.activity_name,
                aa.fecha_apertura,
                aa.fecha_cierre
            FROM all_users au
            JOIN all_activities aa ON au.courseid = aa.courseid
        ),
        interacciones AS (
            SELECT 
                fp.userid,
                f.course AS courseid,
                f.id AS activity_id,
                'Foro' AS activity_type,
                COALESCE(gg.finalgrade, 0) AS nota,
                TO_TIMESTAMP(gg.timemodified) AS fecha_calificacion,
                CASE WHEN gg.feedback IS NOT NULL AND gg.feedback != '' THEN 'SÍ' ELSE 'NO' END AS retroalimentada,
                CASE WHEN EXISTS (
                    SELECT 1 FROM mdl_grading_definitions gd
                    JOIN mdl_grading_areas ga ON gd.areaid = ga.id
                    JOIN mdl_context ctx ON ga.contextid = ctx.id AND ctx.contextlevel = 70
                    JOIN mdl_course_modules cm ON ctx.instanceid = cm.id
                    JOIN mdl_modules m ON cm.module = m.id AND m.name = 'forum'
                    WHERE cm.course = f.course AND cm.instance = f.id
                    AND gd.method = 'rubric'
                ) THEN 'SÍ' ELSE 'NO' END AS retroalimentacion_rubrica
            FROM mdl_forum f
            LEFT JOIN mdl_forum_discussions fd ON f.id = fd.forum AND fd.name LIKE '%Momento%'
            LEFT JOIN mdl_forum_posts fp ON fd.id = fp.discussion
            LEFT JOIN mdl_grade_items gi ON gi.itemmodule = 'forum' AND gi.iteminstance = f.id
            LEFT JOIN mdl_grade_grades gg ON gg.userid = fp.userid AND gg.itemid = gi.id
            WHERE f.course IN (" . implode(',', $cursos_seleccionados) . ")
              AND f.name != 'Avisos'
              AND (fp.created BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) 
                                  AND EXTRACT(EPOCH FROM :fecha_fin::timestamp) OR fp.created IS NULL)
            GROUP BY fp.userid, f.course, f.id, gg.finalgrade, gg.timemodified, gg.feedback
            UNION ALL
            SELECT 
                lg.userid,
                l.course AS courseid,
                l.id AS activity_id,
                'Lección' AS activity_type,
                COALESCE(gg.finalgrade, 0) AS nota,
                TO_TIMESTAMP(gg.timemodified) AS fecha_calificacion,
                CASE WHEN gg.feedback IS NOT NULL AND gg.feedback != '' THEN 'SÍ' ELSE 'NO' END AS retroalimentada,
                'NO' AS retroalimentacion_rubrica
            FROM mdl_lesson l
            LEFT JOIN mdl_lesson_grades lg ON lg.lessonid = l.id
            LEFT JOIN mdl_grade_items gi ON gi.itemmodule = 'lesson' AND gi.iteminstance = l.id
            LEFT JOIN mdl_grade_grades gg ON gg.userid = lg.userid AND gg.itemid = gi.id
            WHERE l.course IN (" . implode(',', $cursos_seleccionados) . ")
              AND (lg.completed BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) 
                                    AND EXTRACT(EPOCH FROM :fecha_fin::timestamp) OR lg.completed IS NULL)
            GROUP BY lg.userid, l.course, l.id, gg.finalgrade, gg.timemodified, gg.feedback
            UNION ALL
            SELECT 
                sub.userid,
                a.course AS courseid,
                a.id AS activity_id,
                'Tarea' AS activity_type,
                COALESCE(gg.finalgrade, 0) AS nota,
                TO_TIMESTAMP(gg.timemodified) AS fecha_calificacion,
                CASE WHEN gg.feedback IS NOT NULL AND gg.feedback != '' THEN 'SÍ' ELSE 'NO' END AS retroalimentada,
                CASE WHEN EXISTS (
                    SELECT 1 FROM mdl_grading_definitions gd
                    JOIN mdl_grading_areas ga ON gd.areaid = ga.id
                    JOIN mdl_context ctx ON ga.contextid = ctx.id AND ctx.contextlevel = 70
                    JOIN mdl_course_modules cm ON ctx.instanceid = cm.id
                    JOIN mdl_modules m ON cm.module = m.id AND m.name = 'assign'
                    WHERE cm.course = a.course AND cm.instance = a.id
                    AND gd.method = 'rubric'
                ) THEN 'SÍ' ELSE 'NO' END AS retroalimentacion_rubrica
            FROM mdl_assign a
            LEFT JOIN mdl_assign_submission sub ON sub.assignment = a.id
            LEFT JOIN mdl_grade_items gi ON gi.itemmodule = 'assign' AND gi.iteminstance = a.id
            LEFT JOIN mdl_grade_grades gg ON gg.userid = sub.userid AND gg.itemid = gi.id
            WHERE a.course IN (" . implode(',', $cursos_seleccionados) . ")
              AND (sub.timecreated BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) 
                                       AND EXTRACT(EPOCH FROM :fecha_fin::timestamp) OR sub.timecreated IS NULL)
            GROUP BY sub.userid, a.course, a.id, gg.finalgrade, gg.timemodified, gg.feedback
            UNION ALL
            SELECT 
                qa.userid,
                q.course AS courseid,
                q.id AS activity_id,
                'Quiz' AS activity_type,
                COALESCE(gg.finalgrade, 0) AS nota,
                TO_TIMESTAMP(gg.timemodified) AS fecha_calificacion,
                CASE WHEN gg.feedback IS NOT NULL AND gg.feedback != '' THEN 'SÍ' ELSE 'NO' END AS retroalimentada,
                'NO' AS retroalimentacion_rubrica
            FROM mdl_quiz q
            LEFT JOIN mdl_quiz_attempts qa ON qa.quiz = q.id
            LEFT JOIN mdl_grade_items gi ON gi.itemmodule = 'quiz' AND gi.iteminstance = q.id
            LEFT JOIN mdl_grade_grades gg ON gg.userid = qa.userid AND gg.itemid = gi.id
            WHERE q.course IN (" . implode(',', $cursos_seleccionados) . ")
              AND (qa.timefinish BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) 
                                     AND EXTRACT(EPOCH FROM :fecha_fin::timestamp) OR qa.timefinish IS NULL)
            GROUP BY qa.userid, q.course, q.id, gg.finalgrade, gg.timemodified, gg.feedback
            UNION ALL
            SELECT 
                ge.userid,
                g.course AS courseid,
                g.id AS activity_id,
                'Glosario' AS activity_type,
                COALESCE(gg.finalgrade, 0) AS nota,
                TO_TIMESTAMP(gg.timemodified) AS fecha_calificacion,
                CASE WHEN gg.feedback IS NOT NULL AND gg.feedback != '' THEN 'SÍ' ELSE 'NO' END AS retroalimentada,
                'NO' AS retroalimentacion_rubrica
            FROM mdl_glossary g
            LEFT JOIN mdl_glossary_entries ge ON ge.glossaryid = g.id
            LEFT JOIN mdl_grade_items gi ON gi.itemmodule = 'glossary' AND gi.iteminstance = g.id
            LEFT JOIN mdl_grade_grades gg ON gg.userid = ge.userid AND gg.itemid = gi.id
            WHERE g.course IN (" . implode(',', $cursos_seleccionados) . ")
              AND (ge.timecreated BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) 
                                      AND EXTRACT(EPOCH FROM :fecha_fin::timestamp) OR ge.timecreated IS NULL)
            GROUP BY ge.userid, g.course, g.id, gg.finalgrade, gg.timemodified, gg.feedback
            UNION ALL
            SELECT 
                st.userid,
                s.course AS courseid,
                s.id AS activity_id,
                'SCORM' AS activity_type,
                COALESCE(gg.finalgrade, 0) AS nota,
                TO_TIMESTAMP(gg.timemodified) AS fecha_calificacion,
                CASE WHEN gg.feedback IS NOT NULL AND gg.feedback != '' THEN 'SÍ' ELSE 'NO' END AS retroalimentada,
                'NO' AS retroalimentacion_rubrica
            FROM mdl_scorm s
            LEFT JOIN mdl_scorm_scoes_track st ON st.scormid = s.id
            LEFT JOIN mdl_grade_items gi ON gi.itemmodule = 'scorm' AND gi.iteminstance = s.id
            LEFT JOIN mdl_grade_grades gg ON gg.userid = st.userid AND gg.itemid = gi.id
            WHERE s.course IN (" . implode(',', $cursos_seleccionados) . ")
              AND (st.timemodified BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) 
                                       AND EXTRACT(EPOCH FROM :fecha_fin::timestamp) OR st.timemodified IS NULL)
            GROUP BY st.userid, s.course, s.id, gg.finalgrade, gg.timemodified, gg.feedback
        )
        SELECT 
            u.username AS codigo,
            u.firstname AS nombre,
            u.lastname AS apellidos,
            c.id AS id_curso,
            c.fullname AS curso,
            ua.activity_type AS tipo_actividad,
            ua.activity_name AS nombre_actividad,
            TO_CHAR(ua.fecha_apertura, 'YYYY-MM-DD HH24:MI:SS') AS fecha_apertura,
            TO_CHAR(ua.fecha_cierre, 'YYYY-MM-DD HH24:MI:SS') AS fecha_cierre,
            COALESCE(i.nota, 0) AS nota,
            TO_CHAR(i.fecha_calificacion, 'YYYY-MM-DD HH24:MI:SS') AS fecha_calificacion,
            COALESCE(i.retroalimentada, 'NO') AS retroalimentada,
            COALESCE(i.retroalimentacion_rubrica, 'NO') AS retroalimentacion_rubrica
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
        ORDER BY c.id, u.lastname, u.firstname, ua.activity_type, ua.activity_name";

    // Ejecutar la consulta
    $stmt = $pdo->prepare($sql);
    $stmt->execute($params);
    $resultados = $stmt->fetchAll(PDO::FETCH_ASSOC);

    // Crear el archivo Excel
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();

    // Encabezados
    $headers = [
        'Código', 'Nombre', 'Apellidos', 'ID Curso', 'Curso', 
        'Tipo Actividad', 'Nombre Actividad', 'Fecha Apertura', 'Fecha Cierre',
        'Nota', 'Fecha Calificación', 'Retroalimentada', 'Retroalimentación Rúbrica'
    ];
    $sheet->fromArray($headers, null, 'A1');

    // Llenar datos
    $row = 2;
    foreach ($resultados as $data) {
        $sheet->fromArray($data, null, 'A' . $row);
        $row++;
    }

    // Guardar el archivo Excel
    $nombre_archivo = "reporte_actividades_" . date('Y-m-d') . ".xlsx";
    $ruta_archivo = "$temp_dir/$nombre_archivo";
    $writer = new Xlsx($spreadsheet);
    $writer->save($ruta_archivo);
    $archivos_excel[] = $ruta_archivo;

    // Crear archivo ZIP (manteniendo la estructura original)
    $zip = new ZipArchive();
    $zip_file = "$temp_dir/actividades_pregrado.zip";
    if ($zip->open($zip_file, ZipArchive::CREATE | ZipArchive::OVERWRITE) === true) {
        foreach ($archivos_excel as $archivo) {
            $zip->addFile($archivo, basename($archivo));
        }
        $zip->close();
    } else {
        throw new Exception("No se pudo crear el archivo ZIP");
    }

    // Enviar correo con Postfix 
    $mail = new PHPMailer();
    $mail->isSendmail(); // Usar sendmail/postfix
    
    $mail->setFrom('noreply@utp.edu.co', 'Reportes Moodle');
    foreach ($correo_destino as $correo) {
        $mail->addAddress($correo);
    }
    $mail->Subject = "Reporte de actividades asignaturas de pregrado - " . $hoy->format('Y-m-d');
    $mail->Body = "Cordial Saludo,\n\nAdjunto el reporte de actividades asignaturas de pregrado en un archivo ZIP.";
    $mail->addAttachment($zip_file, 'actividades_pregrado.zip');
    
    if (!$mail->send()) {
        throw new Exception("Error al enviar el correo: " . $mail->ErrorInfo);
    }

    // Notificar éxito
    mail($correo_notificacion, "Estado Reporte", "El reporte de actividades asignaturas de pregrado fue enviado correctamente.");

    // Limpiar archivos temporales
    foreach ($archivos_excel as $archivo) {
        if (file_exists($archivo)) {
            unlink($archivo);
        }
    }
    if (file_exists($zip_file)) {
        unlink($zip_file);
    }
    if (is_dir($temp_dir)) {
        rmdir($temp_dir);
    }

} catch (Exception $e) {
    // Manejo de errores
    if (isset($archivos_excel)) {
        foreach ($archivos_excel as $archivo) {
            if (file_exists($archivo)) {
                unlink($archivo);
            }
        }
    }
    if (isset($zip_file) && file_exists($zip_file)) {
        unlink($zip_file);
    }
    if (isset($temp_dir) && is_dir($temp_dir)) {
        rmdir($temp_dir);
    }
    
    mail($correo_notificacion, 'Error Reporte', 'Error: ' . $e->getMessage());
    die('Error: ' . $e->getMessage());
}