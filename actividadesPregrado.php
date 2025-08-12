<?php
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\Exception;

// Establecer la zona horaria explícitamente
date_default_timezone_set('America/Bogota');

// Configuración de la base de datos
$host = 'localhost';
$dbname = 'moodle';
$user = 'moodle';
$pass = 'M00dl3';

// Configuración de correos
$modo_prueba = true; // Cambiar a false para modo producción
$correo_prueba = 'daniel.pardo@utp.edu.co';
$correos_produccion = ['correo1@utp.edu.co', 'correo2@utp.edu.co', 'correo3@utp.edu.co'];
$correo_notificacion = 'daniel.pardo@utp.edu.co';

// Lista de cursos como variable
$cursos = [
    '786', '583', '804', '596', '820', '790', '821', '819', '789', '580', '799', '616', '617', '797',
    '798', '842', '844', '621', '805', '584', '581', '802', '619', '627', '800', '801', '618', '588',
    '815', '589', '590', '594', '595', '787', '788', '605', '607', '793', '573', '791', '606', '792',
    '608', '609', '795', '611', '794', '610', '586', '814', '623', '622', '624', '733', '574', '604',
    '796', '576', '612', '577', '614', '830', '615', '783', '579', '784', '785', '591', '587', '810',
    '625', '582', '803', '620', '592', '816', '817', '593'
];

// Fechas para el filtro
$hoy = new DateTime();
$fecha_inicio = new DateTime('2025-08-04 00:00:00');
$fecha_fin = clone $hoy;
$fecha_fin->setTime(23, 59, 59);

try {
    // Conexión a la base de datos
    $pdo = new PDO("pgsql:host=$host;dbname=$dbname", $user, $pass);
    $pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);

    // Directorio temporal para los archivos
    $temp_dir = sys_get_temp_dir() . '/reportes_moodle_' . uniqid();
    if (!mkdir($temp_dir)) {
        throw new Exception("No se pudo crear el directorio temporal: $temp_dir");
    }

    // Inicializar spreadsheet
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();

    // Definir encabezados
    $headers = [
        'Código', 'Nombre', 'Apellidos', 'ID Curso', 'Curso', 'Tipo Actividad', 
        'Nombre Actividad', 'Fecha Entrega', 'Fecha Límite', 'Nota', 
        'Fecha Calificación', 'Retroalimentada', 'Retroalimentación Rúbrica'
    ];
    $sheet->fromArray($headers, null, 'A1');

    // Consulta SQL (usando solo placeholders posicionales)
    $sql = "
        WITH 
        all_users AS (
            SELECT DISTINCT 
                u.id AS userid, 
                c.id AS courseid
            FROM mdl_user u
            JOIN mdl_role_assignments ra ON ra.userid = u.id
            JOIN mdl_context mc ON ra.contextid = mc.id AND mc.contextlevel = 50
            JOIN mdl_course c ON mc.instanceid = c.id AND c.id IN (" . implode(',', array_fill(0, count($cursos), '?')) . ")
            WHERE u.username <> '12345678'
              AND ra.roleid IN (3, 5, 9, 16, 17)
        ),
        all_activities AS (
            SELECT 
                f.course AS courseid,
                f.id AS activity_id,
                'Foro' AS activity_type,
                f.name AS activity_name,
                TO_TIMESTAMP(f.duedate) AS fecha_entrega,
                TO_TIMESTAMP(f.cutoffdate) AS fecha_limite
            FROM mdl_forum f
            WHERE f.course IN (" . implode(',', array_fill(0, count($cursos), '?')) . ") AND f.name != 'Avisos'
            UNION ALL
            SELECT 
                l.course AS courseid,
                l.id AS activity_id,
                'Lección' AS activity_type,
                l.name AS activity_name,
                TO_TIMESTAMP(l.available) AS fecha_entrega,
                TO_TIMESTAMP(l.deadline) AS fecha_limite
            FROM mdl_lesson l
            WHERE l.course IN (" . implode(',', array_fill(0, count($cursos), '?')) . ")
            UNION ALL
            SELECT 
                a.course AS courseid,
                a.id AS activity_id,
                'Tarea' AS activity_type,
                a.name AS activity_name,
                TO_TIMESTAMP(a.allowsubmissionsfromdate) AS fecha_entrega,
                TO_TIMESTAMP(a.duedate) AS fecha_limite
            FROM mdl_assign a
            WHERE a.course IN (" . implode(',', array_fill(0, count($cursos), '?')) . ")
            UNION ALL
            SELECT 
                q.course AS courseid,
                q.id AS activity_id,
                'Quiz' AS activity_type,
                q.name AS activity_name,
                TO_TIMESTAMP(q.timeopen) AS fecha_entrega,
                TO_TIMESTAMP(q.timeclose) AS fecha_limite
            FROM mdl_quiz q
            WHERE q.course IN (" . implode(',', array_fill(0, count($cursos), '?')) . ")
            UNION ALL
            SELECT 
                g.course AS courseid,
                g.id AS activity_id,
                'Glosario' AS activity_type,
                g.name AS activity_name,
                TO_TIMESTAMP(g.timecreated) AS fecha_entrega,
                TO_TIMESTAMP(g.timemodified) AS fecha_limite
            FROM mdl_glossary g
            WHERE g.course IN (" . implode(',', array_fill(0, count($cursos), '?')) . ")
            UNION ALL
            SELECT 
                s.course AS courseid,
                s.id AS activity_id,
                'SCORM' AS activity_type,
                s.name AS activity_name,
                TO_TIMESTAMP(s.timeopen) AS fecha_entrega,
                TO_TIMESTAMP(s.timeclose) AS fecha_limite
            FROM mdl_scorm s
            WHERE s.course IN (" . implode(',', array_fill(0, count($cursos), '?')) . ")
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
                    SELECT 1 FROM mdl_gradingform_rubric_fillings grf
                    JOIN mdl_grading_instances gi ON grf.instanceid = gi.id
                    JOIN mdl_grading_definitions gd ON gi.definitionid = gd.id
                    WHERE gd.areaid = (
                        SELECT id FROM mdl_grading_areas WHERE contextid = (
                            SELECT id FROM mdl_context WHERE contextlevel = 70 AND instanceid = (
                                SELECT cm.id FROM mdl_course_modules cm 
                                WHERE cm.course = f.course AND cm.instance = f.id AND cm.module = (
                                    SELECT id FROM mdl_modules WHERE name = 'forum'
                                )
                            )
                        )
                    ) AND gi.itemid = gg.itemid AND gi.raterid = gg.usermodified
                ) THEN 'SÍ' ELSE 'NO' END AS retroalimentacion_rubrica
            FROM mdl_forum f
            LEFT JOIN mdl_forum_discussions fd ON f.id = fd.forum AND fd.name LIKE '%Momento%'
            LEFT JOIN mdl_forum_posts fp ON fd.id = fp.discussion
            LEFT JOIN mdl_grade_items gi ON gi.itemmodule = 'forum' AND gi.iteminstance = f.id
            LEFT JOIN mdl_grade_grades gg ON gg.userid = fp.userid AND gg.itemid = gi.id
            WHERE f.course IN (" . implode(',', array_fill(0, count($cursos), '?')) . ")
              AND f.name != 'Avisos'
              AND (fp.created BETWEEN EXTRACT(EPOCH FROM ?::timestamp) 
                                  AND EXTRACT(EPOCH FROM ?::timestamp) OR fp.created IS NULL)
            GROUP BY fp.userid, f.course, f.id, gg.finalgrade, gg.timemodified, gg.feedback, gg.itemid, gg.usermodified
            UNION ALL
            SELECT 
                lg.userid,
                l.course AS courseid,
                l.id AS activity_id,
                'Lección' AS activity_type,
                COALESCE(gg.finalgrade, 0) AS nota,
                TO_TIMESTAMP(gg.timemodified) AS fecha_calificacion,
                CASE WHEN gg.feedback IS NOT NULL AND gg.feedback != '' THEN 'SÍ' ELSE 'NO' END AS retroalimentada,
                CASE WHEN EXISTS (
                    SELECT 1 FROM mdl_gradingform_rubric_fillings grf
                    JOIN mdl_grading_instances gi ON grf.instanceid = gi.id
                    JOIN mdl_grading_definitions gd ON gi.definitionid = gd.id
                    WHERE gd.areaid = (
                        SELECT id FROM mdl_grading_areas WHERE contextid = (
                            SELECT id FROM mdl_context WHERE contextlevel = 70 AND instanceid = (
                                SELECT cm.id FROM mdl_course_modules cm 
                                WHERE cm.course = l.course AND cm.instance = l.id AND cm.module = (
                                    SELECT id FROM mdl_modules WHERE name = 'lesson'
                                )
                            )
                        )
                    ) AND gi.itemid = gg.itemid AND gi.raterid = gg.usermodified
                ) THEN 'SÍ' ELSE 'NO' END AS retroalimentacion_rubrica
            FROM mdl_lesson l
            LEFT JOIN mdl_lesson_grades lg ON lg.lessonid = l.id
            LEFT JOIN mdl_grade_items gi ON gi.itemmodule = 'lesson' AND gi.iteminstance = l.id
            LEFT JOIN mdl_grade_grades gg ON gg.userid = lg.userid AND gg.itemid = gi.id
            WHERE l.course IN (" . implode(',', array_fill(0, count($cursos), '?')) . ")
              AND (lg.completed BETWEEN EXTRACT(EPOCH FROM ?::timestamp) 
                                    AND EXTRACT(EPOCH FROM ?::timestamp) OR lg.completed IS NULL)
            GROUP BY lg.userid, l.course, l.id, gg.finalgrade, gg.timemodified, gg.feedback, gg.itemid, gg.usermodified
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
                    SELECT 1 FROM mdl_gradingform_rubric_fillings grf
                    JOIN mdl_grading_instances gi ON grf.instanceid = gi.id
                    JOIN mdl_grading_definitions gd ON gi.definitionid = gd.id
                    WHERE gd.areaid = (
                        SELECT id FROM mdl_grading_areas WHERE contextid = (
                            SELECT id FROM mdl_context WHERE contextlevel = 70 AND instanceid = (
                                SELECT cm.id FROM mdl_course_modules cm 
                                WHERE cm.course = a.course AND cm.instance = a.id AND cm.module = (
                                    SELECT id FROM mdl_modules WHERE name = 'assign'
                                )
                            )
                        )
                    ) AND gi.itemid = gg.itemid AND gi.raterid = gg.usermodified
                ) THEN 'SÍ' ELSE 'NO' END AS retroalimentacion_rubrica
            FROM mdl_assign a
            LEFT JOIN mdl_assign_submission sub ON sub.assignment = a.id
            LEFT JOIN mdl_grade_items gi ON gi.itemmodule = 'assign' AND gi.iteminstance = a.id
            LEFT JOIN mdl_grade_grades gg ON gg.userid = sub.userid AND gg.itemid = gi.id
            WHERE a.course IN (" . implode(',', array_fill(0, count($cursos), '?')) . ")
              AND (sub.timecreated BETWEEN EXTRACT(EPOCH FROM ?::timestamp) 
                                      AND EXTRACT(EPOCH FROM ?::timestamp) OR sub.timecreated IS NULL)
            GROUP BY sub.userid, a.course, a.id, gg.finalgrade, gg.timemodified, gg.feedback, gg.itemid, gg.usermodified
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
            WHERE q.course IN (" . implode(',', array_fill(0, count($cursos), '?')) . ")
              AND (qa.timefinish BETWEEN EXTRACT(EPOCH FROM ?::timestamp) 
                                    AND EXTRACT(EPOCH FROM ?::timestamp) OR qa.timefinish IS NULL)
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
            WHERE g.course IN (" . implode(',', array_fill(0, count($cursos), '?')) . ")
              AND (ge.timecreated BETWEEN EXTRACT(EPOCH FROM ?::timestamp) 
                                     AND EXTRACT(EPOCH FROM ?::timestamp) OR ge.timecreated IS NULL)
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
            WHERE s.course IN (" . implode(',', array_fill(0, count($cursos), '?')) . ")
              AND (st.timemodified BETWEEN EXTRACT(EPOCH FROM ?::timestamp) 
                                      AND EXTRACT(EPOCH FROM ?::timestamp) OR st.timemodified IS NULL)
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
            TO_CHAR(ua.fecha_entrega, 'YYYY-MM-DD HH24:MI:SS') AS fecha_entrega,
            TO_CHAR(ua.fecha_limite, 'YYYY-MM-DD HH24:MI:SS') AS fecha_limite,
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
        ORDER BY c.id, u.lastname, u.firstname, ua.activity_type, ua.activity_name;
    ";

    // Preparar parámetros para la consulta
    $params = [];
    // Añadir los cursos para cada subconsulta (8 subconsultas: 1 para all_users, 6 para all_activities, 6 para interacciones)
    for ($i = 0; $i < 13; $i++) {
        $params = array_merge($params, $cursos);
    }
    // Añadir las fechas para cada tipo de actividad en interacciones (6 tipos * 2 fechas)
    for ($i = 0; $i < 6; $i++) {
        $params[] = $fecha_inicio->format('Y-m-d H:i:s');
        $params[] = $fecha_fin->format('Y-m-d H:i:s');
    }

    // Ejecutar consulta
    $stmt = $pdo->prepare($sql);
    $stmt->execute($params);
    $resultados = $stmt->fetchAll(PDO::FETCH_ASSOC);

    // Añadir datos al spreadsheet
    $row_number = 2;
    foreach ($resultados as $row) {
        $row_data = [
            $row['codigo'],
            $row['nombre'],
            $row['apellidos'],
            $row['id_curso'],
            $row['curso'],
            $row['tipo_actividad'],
            $row['nombre_actividad'],
            $row['fecha_entrega'],
            $row['fecha_limite'],
            $row['nota'],
            $row['fecha_calificacion'],
            $row['retroalimentada'],
            $row['retroalimentacion_rubrica']
        ];
        $sheet->fromArray($row_data, null, 'A' . $row_number);
        $row_number++;
    }

    // Guardar el archivo Excel
    $nombre_archivo = "reporte_actividades_pregrado.xlsx";
    $ruta_archivo = "$temp_dir/$nombre_archivo";
    $writer = new Xlsx($spreadsheet);
    $writer->save($ruta_archivo);

    // Crear el archivo ZIP
    $zip = new ZipArchive();
    $zip_file = "$temp_dir/actividades_pregrado.zip";
    if ($zip->open($zip_file, ZipArchive::CREATE | ZipArchive::OVERWRITE) !== true) {
        throw new Exception("No se pudo crear el archivo ZIP: $zip_file");
    }
    $zip->addFile($ruta_archivo, basename($ruta_archivo));
    $zip->close();

    // Configurar correos según el modo
    $correos_destino = $modo_prueba ? [$correo_prueba] : $correos_produccion;

    // Enviar el correo con el ZIP adjunto
    $mail = new PHPMailer(true);
    $mail->setFrom('noreply@example.com', 'Reporte Moodle');
    foreach ($correos_destino as $correo) {
        $mail->addAddress($correo);
    }
    $mail->Subject = "Reporte de actividades pregrado - " . $hoy->format('Y-m-d');
    $mail->Body = "Cordial Saludo, Adjunto el reporte de actividades pregrado en un archivo ZIP que contiene el reporte consolidado.";
    $mail->addAttachment($zip_file, 'actividades_pregrado.zip');
    $mail->send();

    // Notificar éxito
    mail($correo_notificacion, "Estado Reporte", "El reporte de actividades pregrado fue enviado correctamente como actividades_pregrado.zip.");

    // Limpiar archivos temporales
    unlink($ruta_archivo);
    unlink($zip_file);
    rmdir($temp_dir);

} catch (Exception $e) {
    // Limpiar archivos temporales en caso de error
    if (file_exists($ruta_archivo)) {
        unlink($ruta_archivo);
    }
    if (file_exists($zip_file)) {
        unlink($zip_file);
    }
    if (is_dir($temp_dir)) {
        rmdir($temp_dir);
    }
    mail($correo_notificacion, 'Error Reporte', 'Error: ' . $e->getMessage());
}
?>