<?php
// Incluir la librería PhpSpreadsheet
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// Establecer la zona horaria a Bogotá, Colombia
date_default_timezone_set('America/Bogota');

// Iniciar el contador de tiempo
$startTime = microtime(true);

// Datos de conexión a PostgreSQL
$host = 'localhost';       // Servidor de la base de datos
$dbname = 'moodle';        // Nombre de la base de datos
$user = 'moodle';          // Usuario de la base de datos
$pass = 'M00dl3';          // Contraseña del usuario
$port = '5432';            // Puerto de PostgreSQL (por defecto es 5432)

try {
    // Conexión a PostgreSQL con PDO
    $dsn = "pgsql:host=$host;port=$port;dbname=$dbname;user=$user;password=$pass";
    $pdo = new PDO($dsn);
    $pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);

    // Consulta SQL adaptada para PostgreSQL
    $sql = "WITH 
    all_users AS (
        SELECT DISTINCT 
            u.id AS userid, 
            c.id AS courseid
        FROM mdl_user u
        JOIN mdl_role_assignments ra ON ra.userid = u.id
        JOIN mdl_context mc ON ra.contextid = mc.id AND mc.contextlevel = 50
        JOIN mdl_course c ON mc.instanceid = c.id AND c.id IN ('494','415','507','481','508','482','509','485','526','510','511','486','490','416','503','504','527','417','496','497','418','498','419','475','421','420','422','423','512','513','515','488','489','424','516','517','491','518','492','519','493','520','425','476','426','505','506','479','521','428','430','522','495','499','431','453','500','523','434','524','435','436','437','438','440','502','439','452','525','442')
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
        WHERE f.course IN ('494','415','507','481','508','482','509','485','526','510','511','486','490','416','503','504','527','417','496','497','418','498','419','475','421','420','422','423','512','513','515','488','489','424','516','517','491','518','492','519','493','520','425','476','426','505','506','479','521','428','430','522','495','499','431','453','500','523','434','524','435','436','437','438','440','502','439','452','525','442') AND f.name != 'Avisos'
        UNION ALL
        SELECT 
            l.course AS courseid,
            l.id AS activity_id,
            'Lección' AS activity_type,
            l.name AS activity_name,
            TO_TIMESTAMP(l.available) AS fecha_apertura,
            TO_TIMESTAMP(l.deadline) AS fecha_cierre
        FROM mdl_lesson l
        WHERE l.course IN ('494','415','507','481','508','482','509','485','526','510','511','486','490','416','503','504','527','417','496','497','418','498','419','475','421','420','422','423','512','513','515','488','489','424','516','517','491','518','492','519','493','520','425','476','426','505','506','479','521','428','430','522','495','499','431','453','500','523','434','524','435','436','437','438','440','502','439','452','525','442')
        UNION ALL
        SELECT 
            a.course AS courseid,
            a.id AS activity_id,
            'Tarea' AS activity_type,
            a.name AS activity_name,
            TO_TIMESTAMP(a.allowsubmissionsfromdate) AS fecha_apertura,
            TO_TIMESTAMP(a.duedate) AS fecha_cierre
        FROM mdl_assign a
        WHERE a.course IN ('494','415','507','481','508','482','509','485','526','510','511','486','490','416','503','504','527','417','496','497','418','498','419','475','421','420','422','423','512','513','515','488','489','424','516','517','491','518','492','519','493','520','425','476','426','505','506','479','521','428','430','522','495','499','431','453','500','523','434','524','435','436','437','438','440','502','439','452','525','442')
        UNION ALL
        SELECT 
            q.course AS courseid,
            q.id AS activity_id,
            'Quiz' AS activity_type,
            q.name AS activity_name,
            TO_TIMESTAMP(q.timeopen) AS fecha_apertura,
            TO_TIMESTAMP(q.timeclose) AS fecha_cierre
        FROM mdl_quiz q
        WHERE q.course IN ('494','415','507','481','508','482','509','485','526','510','511','486','490','416','503','504','527','417','496','497','418','498','419','475','421','420','422','423','512','513','515','488','489','424','516','517','491','518','492','519','493','520','425','476','426','505','506','479','521','428','430','522','495','499','431','453','500','523','434','524','435','436','437','438','440','502','439','452','525','442')
        UNION ALL
        SELECT 
            g.course AS courseid,
            g.id AS activity_id,
            'Glosario' AS activity_type,
            g.name AS activity_name,
            TO_TIMESTAMP(g.timecreated) AS fecha_apertura,
            TO_TIMESTAMP(g.timemodified) AS fecha_cierre
        FROM mdl_glossary g
        WHERE g.course IN ('494','415','507','481','508','482','509','485','526','510','511','486','490','416','503','504','527','417','496','497','418','498','419','475','421','420','422','423','512','513','515','488','489','424','516','517','491','518','492','519','493','520','425','476','426','505','506','479','521','428','430','522','495','499','431','453','500','523','434','524','435','436','437','438','440','502','439','452','525','442')
        UNION ALL
        SELECT 
            s.course AS courseid,
            s.id AS activity_id,
            'SCORM' AS activity_type,
            s.name AS activity_name,
            TO_TIMESTAMP(s.timeopen) AS fecha_apertura,
            TO_TIMESTAMP(s.timeclose) AS fecha_cierre
        FROM mdl_scorm s
        WHERE s.course IN ('494','415','507','481','508','482','509','485','526','510','511','486','490','416','503','504','527','417','496','497','418','498','419','475','421','420','422','423','512','513','515','488','489','424','516','517','491','518','492','519','493','520','425','476','426','505','506','479','521','428','430','522','495','499','431','453','500','523','434','524','435','436','437','438','440','502','439','452','525','442')
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
            COUNT(DISTINCT fp.id) AS visitas,
            COUNT(DISTINCT fp.id) AS envios,
            COALESCE(gg.finalgrade, 0) AS nota,
            CASE WHEN gg.feedback IS NOT NULL AND gg.feedback != '' THEN 'SÍ' ELSE 'NO' END AS retroalimentada
        FROM mdl_forum f
        LEFT JOIN mdl_forum_discussions fd ON f.id = fd.forum AND fd.name LIKE '%Momento%'
        LEFT JOIN mdl_forum_posts fp ON fd.id = fp.discussion
        LEFT JOIN mdl_grade_items gi ON gi.itemmodule = 'forum' AND gi.iteminstance = f.id
        LEFT JOIN mdl_grade_grades gg ON gg.userid = fp.userid AND gg.itemid = gi.id
        WHERE f.course IN ('494','415','507','481','508','482','509','485','526','510','511','486','490','416','503','504','527','417','496','497','418','498','419','475','421','420','422','423','512','513','515','488','489','424','516','517','491','518','492','519','493','520','425','476','426','505','506','479','521','428','430','522','495','499','431','453','500','523','434','524','435','436','437','438','440','502','439','452','525','442')
          AND f.name != 'Avisos'
          AND (fp.created BETWEEN EXTRACT(EPOCH FROM TIMESTAMP '2025-02-03 00:00:00') 
                              AND EXTRACT(EPOCH FROM CURRENT_DATE + INTERVAL '23 hours 59 minutes') OR fp.created IS NULL)
        GROUP BY fp.userid, f.course, f.id, gg.finalgrade, gg.feedback
        UNION ALL
        SELECT 
            lg.userid,
            l.course AS courseid,
            l.id AS activity_id,
            'Lección' AS activity_type,
            COUNT(DISTINCT lg.id) AS visitas,
            COUNT(DISTINCT lg.id) AS envios,
            COALESCE(gg.finalgrade, 0) AS nota,
            CASE WHEN gg.feedback IS NOT NULL AND gg.feedback != '' THEN 'SÍ' ELSE 'NO' END AS retroalimentada
        FROM mdl_lesson l
        LEFT JOIN mdl_lesson_grades lg ON lg.lessonid = l.id
        LEFT JOIN mdl_grade_items gi ON gi.itemmodule = 'lesson' AND gi.iteminstance = l.id
        LEFT JOIN mdl_grade_grades gg ON gg.userid = lg.userid AND gg.itemid = gi.id
        WHERE l.course IN ('494','415','507','481','508','482','509','485','526','510','511','486','490','416','503','504','527','417','496','497','418','498','419','475','421','420','422','423','512','513','515','488','489','424','516','517','491','518','492','519','493','520','425','476','426','505','506','479','521','428','430','522','495','499','431','453','500','523','434','524','435','436','437','438','440','502','439','452','525','442')
          AND (lg.completed BETWEEN EXTRACT(EPOCH FROM TIMESTAMP '2025-02-03 00:00:00') 
                                AND EXTRACT(EPOCH FROM CURRENT_DATE + INTERVAL '23 hours 59 minutes') OR lg.completed IS NULL)
        GROUP BY lg.userid, l.course, l.id, gg.finalgrade, gg.feedback
        UNION ALL
        SELECT 
            sub.userid,
            a.course AS courseid,
            a.id AS activity_id,
            'Tarea' AS activity_type,
            COUNT(DISTINCT sub.id) AS visitas,
            COUNT(DISTINCT sub.id) AS envios,
            COALESCE(gg.finalgrade, 0) AS nota,
            CASE WHEN gg.feedback IS NOT NULL AND gg.feedback != '' THEN 'SÍ' ELSE 'NO' END AS retroalimentada
        FROM mdl_assign a
        LEFT JOIN mdl_assign_submission sub ON sub.assignment = a.id
        LEFT JOIN mdl_grade_items gi ON gi.itemmodule = 'assign' AND gi.iteminstance = a.id
        LEFT JOIN mdl_grade_grades gg ON gg.userid = sub.userid AND gg.itemid = gi.id
        WHERE a.course IN ('494','415','507','481','508','482','509','485','526','510','511','486','490','416','503','504','527','417','496','497','418','498','419','475','421','420','422','423','512','513','515','488','489','424','516','517','491','518','492','519','493','520','425','476','426','505','506','479','521','428','430','522','495','499','431','453','500','523','434','524','435','436','437','438','440','502','439','452','525','442')
          AND (sub.timecreated BETWEEN EXTRACT(EPOCH FROM TIMESTAMP '2025-02-03 00:00:00') 
                                   AND EXTRACT(EPOCH FROM CURRENT_DATE + INTERVAL '23 hours 59 minutes') OR sub.timecreated IS NULL)
        GROUP BY sub.userid, a.course, a.id, gg.finalgrade, gg.feedback
        UNION ALL
        SELECT 
            qa.userid,
            q.course AS courseid,
            q.id AS activity_id,
            'Quiz' AS activity_type,
            COUNT(DISTINCT qa.id) AS visitas,
            COUNT(DISTINCT qa.id) AS envios,
            COALESCE(gg.finalgrade, 0) AS nota,
            CASE WHEN gg.feedback IS NOT NULL AND gg.feedback != '' THEN 'SÍ' ELSE 'NO' END AS retroalimentada
        FROM mdl_quiz q
        LEFT JOIN mdl_quiz_attempts qa ON qa.quiz = q.id
        LEFT JOIN mdl_grade_items gi ON gi.itemmodule = 'quiz' AND gi.iteminstance = q.id
        LEFT JOIN mdl_grade_grades gg ON gg.userid = qa.userid AND gg.itemid = gi.id
        WHERE q.course IN ('494','415','507','481','508','482','509','485','526','510','511','486','490','416','503','504','527','417','496','497','418','498','419','475','421','420','422','423','512','513','515','488','489','424','516','517','491','518','492','519','493','520','425','476','426','505','506','479','521','428','430','522','495','499','431','453','500','523','434','524','435','436','437','438','440','502','439','452','525','442')
          AND (qa.timefinish BETWEEN EXTRACT(EPOCH FROM TIMESTAMP '2025-02-03 00:00:00') 
                                 AND EXTRACT(EPOCH FROM CURRENT_DATE + INTERVAL '23 hours 59 minutes') OR qa.timefinish IS NULL)
        GROUP BY qa.userid, q.course, q.id, gg.finalgrade, gg.feedback
        UNION ALL
        SELECT 
            ge.userid,
            g.course AS courseid,
            g.id AS activity_id,
            'Glosario' AS activity_type,
            COUNT(DISTINCT ge.id) AS visitas,
            COUNT(DISTINCT ge.id) AS envios,
            COALESCE(gg.finalgrade, 0) AS nota,
            CASE WHEN gg.feedback IS NOT NULL AND gg.feedback != '' THEN 'SÍ' ELSE 'NO' END AS retroalimentada
        FROM mdl_glossary g
        LEFT JOIN mdl_glossary_entries ge ON ge.glossaryid = g.id
        LEFT JOIN mdl_grade_items gi ON gi.itemmodule = 'glossary' AND gi.iteminstance = g.id
        LEFT JOIN mdl_grade_grades gg ON gg.userid = ge.userid AND gg.itemid = gi.id
        WHERE g.course IN ('494','415','507','481','508','482','509','485','526','510','511','486','490','416','503','504','527','417','496','497','418','498','419','475','421','420','422','423','512','513','515','488','489','424','516','517','491','518','492','519','493','520','425','476','426','505','506','479','521','428','430','522','495','499','431','453','500','523','434','524','435','436','437','438','440','502','439','452','525','442')
          AND (ge.timecreated BETWEEN EXTRACT(EPOCH FROM TIMESTAMP '2025-02-03 00:00:00') 
                                  AND EXTRACT(EPOCH FROM CURRENT_DATE + INTERVAL '23 hours 59 minutes') OR ge.timecreated IS NULL)
        GROUP BY ge.userid, g.course, g.id, gg.finalgrade, gg.feedback
        UNION ALL
        SELECT 
            st.userid,
            s.course AS courseid,
            s.id AS activity_id,
            'SCORM' AS activity_type,
            COUNT(DISTINCT st.id) AS visitas,
            COUNT(DISTINCT st.id) AS envios,
            COALESCE(gg.finalgrade, 0) AS nota,
            CASE WHEN gg.feedback IS NOT NULL AND gg.feedback != '' THEN 'SÍ' ELSE 'NO' END AS retroalimentada
        FROM mdl_scorm s
        LEFT JOIN mdl_scorm_scoes_track st ON st.scormid = s.id
        LEFT JOIN mdl_grade_items gi ON gi.itemmodule = 'scorm' AND gi.iteminstance = s.id
        LEFT JOIN mdl_grade_grades gg ON gg.userid = st.userid AND gg.itemid = gi.id
        WHERE s.course IN ('494','415','507','481','508','482','509','485','526','510','511','486','490','416','503','504','527','417','496','497','418','498','419','475','421','420','422','423','512','513','515','488','489','424','516','517','491','518','492','519','493','520','425','476','426','505','506','479','521','428','430','522','495','499','431','453','500','523','434','524','435','436','437','438','440','502','439','452','525','442')
          AND (st.timemodified BETWEEN EXTRACT(EPOCH FROM TIMESTAMP '2025-02-03 00:00:00') 
                                   AND EXTRACT(EPOCH FROM CURRENT_DATE + INTERVAL '23 hours 59 minutes') OR st.timemodified IS NULL)
        GROUP BY st.userid, s.course, s.id, gg.finalgrade, gg.feedback
    )
    SELECT 
        u.username AS codigo,
        u.firstname AS nombre,
        u.lastname AS apellidos,
        r.name AS rol,
        u.email AS correo,
        c.id AS id_curso,
        c.fullname AS curso,
        u.institution,
        u.department,
        ua.activity_type AS tipo_actividad,
        ua.activity_name AS nombre_actividad,
        TO_CHAR(ua.fecha_apertura, 'YYYY-MM-DD HH24:MI:SS') AS fecha_apertura,
        TO_CHAR(ua.fecha_cierre, 'YYYY-MM-DD HH24:MI:SS') AS fecha_cierre,
        COALESCE(i.visitas, 0) AS visitas,
        COALESCE(i.envios, 0) AS envios,
        COALESCE(i.nota, 0) AS nota,
        COALESCE(i.retroalimentada, 'NO') AS retroalimentada
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
    ORDER BY c.id, u.lastname, u.firstname, ua.activity_type, ua.activity_name;";

    // Ejecutar la consulta
    $stmt = $pdo->query($sql);
    $results = $stmt->fetchAll(PDO::FETCH_ASSOC);

    // Detener el contador de tiempo y calcular la duración
    $endTime = microtime(true);
    $executionTime = ($endTime - $startTime) / 60; // Duración en minutos

    // Crear el archivo Excel con PhpSpreadsheet
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();

    // Encabezados de las columnas
    $headers = [
        'Código', 'Nombre', 'Apellidos', 'Rol', 'Correo', 'ID Curso', 'Curso', 
        'Institución', 'Departamento', 'Tipo Actividad', 'Nombre Actividad', 
        'Fecha Apertura', 'Fecha Cierre', 'Visitas', 'Envíos', 'Nota', 'Retroalimentada'
    ];
    $sheet->fromArray($headers, null, 'A1');

    // Insertar los datos desde la fila 2
    $row = 2;
    foreach ($results as $result) {
        $sheet->fromArray(array_values($result), null, "A$row");
        $row++;
    }

    // Generar el archivo Excel en memoria
    $writer = new Xlsx($spreadsheet);
    ob_start();
    $writer->save('php://output');
    $excelContent = ob_get_clean();

    // Nombre del archivo con la fecha actual
    $fileName = 'reporte_actividades_' . date('Y-m-d') . '.xlsx';

    // Preparar el correo con el archivo adjunto
    $to = 'daniel.pardo.utp.edu.co';
    $subject = 'Reporte de Actividades Moodle Pregrado - ' . date('Y-m-d');
    $message = "Adjunto encontrarás el reporte de actividades generado el " . date('Y-m-d') . ".\n\n";
    $message .= "Duración de la ejecución de la consulta: " . round($executionTime, 2) . " minutos.\n\n";
    $message .= "Saludos,\nSistema Automático";
    $boundary = md5(time());
    $headers = "MIME-Version: 1.0\r\n";
    $headers .= "From: no-reply@tu-dominio.com\r\n";
    $headers .= "Content-Type: multipart/mixed; boundary=\"$boundary\"\r\n";

    // Cuerpo del correo con el adjunto
    $body = "--$boundary\r\n";
    $body .= "Content-Type: text/plain; charset=UTF-8\r\n";
    $body .= "Content-Transfer-Encoding: 8bit\r\n\r\n";
    $body .= $message . "\r\n";
    $body .= "--$boundary\r\n";
    $body .= "Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet; name=\"$fileName\"\r\n";
    $body .= "Content-Disposition: attachment; filename=\"$fileName\"\r\n";
    $body .= "Content-Transfer-Encoding: base64\r\n\r\n";
    $body .= chunk_split(base64_encode($excelContent)) . "\r\n";
    $body .= "--$boundary--";

    // Enviar el correo con el reporte
    $mailSent = mail($to, $subject, $body, $headers);

    // Enviar correo de confirmación
    $confirmSubject = $mailSent ? 'Confirmación: Reporte Enviado Exitosamente' : 'Error: Fallo al Enviar el Reporte';
    $confirmMessage = $mailSent 
        ? "El reporte de actividades del " . date('Y-m-d') . " fue enviado exitosamente a $to.\n"
        : "Hubo un error al enviar el reporte de actividades del " . date('Y-m-d') . " a $to.\n";
    $confirmMessage .= "Duración de la ejecución de la consulta: " . round($executionTime, 2) . " minutos.";
    $confirmHeaders = "From: no-reply@tu-dominio.com\r\n";
    $confirmHeaders .= "Content-Type: text/plain; charset=UTF-8\r\n";
    mail($to, $confirmSubject, $confirmMessage, $confirmHeaders);

} catch (PDOException $e) {
    // Manejo de errores de conexión o consulta
    $errorMessage = "Error en la base de datos: " . $e->getMessage();
    $headers = "From: no-reply@tu-dominio.com\r\nContent-Type: text/plain; charset=UTF-8\r\n";
    mail('daniel.pardo.utp.edu.co', 'Error en el Script de Reporte', $errorMessage, $headers);
} catch (Exception $e) {
    // Manejo de otros errores (por ejemplo, PhpSpreadsheet)
    $errorMessage = "Error general: " . $e->getMessage();
    $headers = "From: no-reply@tu-dominio.com\r\nContent-Type: text/plain; charset=UTF-8\r\n";
    mail('daniel.pardo.utp.edu.co', 'Error en el Script de Reporte', $errorMessage, $headers);
}
?>
