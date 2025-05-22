<?php
require 'vendor/autoload.php'; 
use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\Exception;

$host = 'localhost';
$dbname = 'moodle';
$user = 'moodle';
$pass = 'M00dl3';
$correo_destino = ['soporteunivirtual@utp.edu.co', 'univirtual-utp@utp.edu.co'];
$correo_notificacion = 'soporteunivirtual@utp.edu.co';

$fecha_inicio_fija = new DateTime('2025-03-19'); 
$hoy = new DateTime();
$lunes = clone $hoy;
$lunes->modify('last monday');

$fecha_inicio = $fecha_inicio_fija->format('Y-m-d 00:00:00');
$fecha_fin = $lunes->format('Y-m-d 23:59:59');

try {
    $pdo = new PDO("pgsql:host=$host;dbname=$dbname", $user, $pass);
    $pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);

    $sql = "SELECT DISTINCT
        u.username AS codigo,
        u.firstname AS nombre,
        u.lastname AS apellidos,
        COALESCE(
            (SELECT r.name 
             FROM mdl_role_assignments ra
             JOIN mdl_role r ON ra.roleid = r.id
             WHERE ra.userid = u.id AND ra.roleid IN (3, 5, 9, 16, 17)
             LIMIT 1),
            'Estudiante'
        ) AS rol,
        r.id AS rol_id,
        u.email AS correo,
        c.id AS id_curso,
        c.fullname AS curso,
        -- CONTEO DE INGRESOS
        (SELECT COUNT(DISTINCT DATE(to_timestamp(l.timecreated)))
         FROM mdl_logstore_standard_log l
         WHERE l.userid = u.id 
         AND l.courseid = c.id
         AND l.timecreated BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) 
                             AND EXTRACT(EPOCH FROM :fecha_fin::timestamp)) AS ingresos,
        -- INTERACCIONES EN EL AULA
        (SELECT COUNT(*)
         FROM mdl_logstore_standard_log l
         WHERE l.userid = u.id 
         AND l.courseid = c.id 
         AND l.action = 'viewed' 
         AND l.target IN ('course', 'course_module')
         AND l.timecreated BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) 
                             AND EXTRACT(EPOCH FROM :fecha_fin::timestamp)) AS interacciones_en_aula,
        -- ÚLTIMO ACCESO
        COALESCE(
            (SELECT TO_CHAR(to_timestamp(MAX(l.timecreated)), 'YYYY-MM-DD HH24:MI:SS')
             FROM mdl_logstore_standard_log l
             WHERE l.userid = u.id 
             AND l.courseid = c.id
             AND l.timecreated BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) 
                                 AND EXTRACT(EPOCH FROM :fecha_fin::timestamp)),
            'Nunca ha accedido'
        ) AS ultimo_acceso,
        -- MENSAJES ENVIADOS
        COALESCE((
            SELECT COUNT(DISTINCT m.id)
            FROM mdl_messages m
            JOIN mdl_message_conversations mc ON m.conversationid = mc.id
            JOIN mdl_message_conversation_members mcm ON mc.id = mcm.conversationid
            JOIN mdl_user u2 ON mcm.userid = u2.id AND u2.id != u.id
            JOIN mdl_role_assignments ra2 ON u2.id = ra2.userid
            JOIN mdl_context ctx2 ON ra2.contextid = ctx2.id AND ctx2.contextlevel = 50
            WHERE m.useridfrom = u.id
            AND u.username <> '12345678'
            AND ctx2.instanceid = c.id
            AND ra2.roleid IN (5, 11, 16, 17)
            AND m.timecreated BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) 
                                AND EXTRACT(EPOCH FROM :fecha_fin::timestamp)
        ), 0) AS mensajes_enviados,
        -- MENSAJES RECIBIDOS
        COALESCE((
            SELECT COUNT(DISTINCT m.id)
            FROM mdl_messages m
            JOIN mdl_message_conversations mc ON m.conversationid = mc.id
            JOIN mdl_message_conversation_members mcm ON mc.id = mcm.conversationid
            JOIN mdl_user u2 ON m.useridfrom = u2.id AND u2.id != u.id
            JOIN mdl_role_assignments ra2 ON u2.id = ra2.userid
            JOIN mdl_context ctx2 ON ra2.contextid = ctx2.id AND ctx2.contextlevel = 50
            WHERE mcm.userid = u.id
            AND ctx2.instanceid = c.id
            AND ra2.roleid IN (5, 11, 16, 17)
            AND m.timecreated BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) 
                                AND EXTRACT(EPOCH FROM :fecha_fin::timestamp)
        ), 0) AS mensajes_recibidos,
        -- MENSAJES SIN LEER
        COALESCE((
            SELECT COUNT(DISTINCT m.id)
            FROM mdl_messages m
            JOIN mdl_message_conversations mc ON m.conversationid = mc.id
            JOIN mdl_message_conversation_members mcm ON mc.id = mcm.conversationid
            JOIN mdl_user u2 ON m.useridfrom = u2.id AND u2.id != u.id
            JOIN mdl_role_assignments ra2 ON u2.id = ra2.userid
            JOIN mdl_context ctx2 ON ra2.contextid = ctx2.id AND ctx2.contextlevel = 50
            LEFT JOIN mdl_message_user_actions mua ON m.id = mua.messageid AND mua.userid = u.id AND mua.action = 1
            WHERE mcm.userid = u.id
            AND ctx2.instanceid = c.id
            AND ra2.roleid IN (5, 11, 16, 17)
            AND mua.id IS NULL
            AND m.timecreated BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) 
                                AND EXTRACT(EPOCH FROM :fecha_fin::timestamp)
        ), 0) AS mensajes_sin_leer,
        -- ÚLTIMO MENSAJE ENVIADO
        COALESCE((
            SELECT TO_CHAR(TO_TIMESTAMP(MAX(m.timecreated)), 'YYYY-MM-DD HH24:MI:SS')
            FROM mdl_messages m
            JOIN mdl_message_conversations mc ON m.conversationid = mc.id
            JOIN mdl_message_conversation_members mcm ON mc.id = mcm.conversationid
            JOIN mdl_user u2 ON mcm.userid = u2.id AND u2.id != u.id
            JOIN mdl_role_assignments ra2 ON u2.id = ra2.userid
            JOIN mdl_context ctx2 ON ra2.contextid = ctx2.id AND ctx2.contextlevel = 50
            WHERE m.useridfrom = u.id
            AND u.username <> '12345678'
            AND ctx2.instanceid = c.id
            AND ra2.roleid IN (5, 11, 16, 17)
            AND m.timecreated BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) 
                                AND EXTRACT(EPOCH FROM :fecha_fin::timestamp)
        ), 'No ha enviado Mensaje') AS ultimo_mensaje_enviado,
        -- ÚLTIMO MENSAJE RECIBIDO
        COALESCE((
            SELECT TO_CHAR(TO_TIMESTAMP(MAX(m.timecreated)), 'YYYY-MM-DD HH24:MI:SS')
            FROM mdl_messages m
            JOIN mdl_message_conversations mc ON m.conversationid = mc.id
            JOIN mdl_message_conversation_members mcm ON mc.id = mcm.conversationid
            JOIN mdl_user u2 ON m.useridfrom = u2.id AND u2.id != u.id
            JOIN mdl_role_assignments ra2 ON u2.id = ra2.userid
            JOIN mdl_context ctx2 ON ra2.contextid = ctx2.id AND ctx2.contextlevel = 50
            WHERE mcm.userid = u.id
            AND ctx2.instanceid = c.id
            AND ra2.roleid IN (5, 11, 16, 17)
            AND m.timecreated BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) 
                                AND EXTRACT(EPOCH FROM :fecha_fin::timestamp)
        ), 'No ha recibido Mensaje') AS ultimo_mensaje_recibido,
        -- ÚLTIMO MENSAJE LEÍDO
        COALESCE((
            SELECT TO_CHAR(TO_TIMESTAMP(MAX(m.timecreated)), 'YYYY-MM-DD HH24:MI:SS')
            FROM mdl_messages m
            JOIN mdl_message_conversations mc ON m.conversationid = mc.id
            JOIN mdl_message_conversation_members mcm ON mc.id = mcm.conversationid
            JOIN mdl_user u2 ON m.useridfrom = u2.id AND u2.id != u.id
            JOIN mdl_role_assignments ra2 ON u2.id = ra2.userid
            JOIN mdl_context ctx2 ON ra2.contextid = ctx2.id AND ctx2.contextlevel = 50
            JOIN mdl_message_user_actions mua ON m.id = mua.messageid AND mua.userid = u.id AND mua.action = 1
            WHERE mcm.userid = u.id
            AND ctx2.instanceid = c.id
            AND ra2.roleid IN (5, 11, 16, 17)
            AND m.timecreated BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) 
                                AND EXTRACT(EPOCH FROM :fecha_fin::timestamp)
        ), 'No ha leído Mensaje') AS ultimo_mensaje_leido,
        -- ÚLTIMO MENSAJE ENVIADO A
        COALESCE((
            SELECT CONCAT(u2.firstname, ' ', u2.lastname)
            FROM mdl_messages m
            JOIN mdl_message_conversations mc ON m.conversationid = mc.id
            JOIN mdl_message_conversation_members mcm ON mc.id = mcm.conversationid
            JOIN mdl_user u2 ON mcm.userid = u2.id AND u2.id != u.id
            JOIN mdl_role_assignments ra2 ON u2.id = ra2.userid
            JOIN mdl_context ctx2 ON ra2.contextid = ctx2.id AND ctx2.contextlevel = 50
            WHERE m.useridfrom = u.id
            AND u.username <> '12345678'
            AND ctx2.instanceid = c.id
            AND ra2.roleid IN (5, 11, 16, 17)
            AND m.timecreated BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) 
                                AND EXTRACT(EPOCH FROM :fecha_fin::timestamp)
            AND m.timecreated = (
                SELECT MAX(m2.timecreated)
                FROM mdl_messages m2
                JOIN mdl_message_conversations mc2 ON m2.conversationid = mc2.id
                JOIN mdl_message_conversation_members mcm2 ON mc2.id = mcm2.conversationid
                JOIN mdl_user u3 ON mcm2.userid = u3.id AND u3.id != u.id
                JOIN mdl_role_assignments ra3 ON u3.id = ra3.userid
                JOIN mdl_context ctx3 ON ra3.contextid = ctx3.id AND ctx3.contextlevel = 50
                WHERE m2.useridfrom = u.id
                AND ctx3.instanceid = c.id
                AND ra3.roleid IN (5, 11, 16, 17)
                AND m2.timecreated BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) 
                                    AND EXTRACT(EPOCH FROM :fecha_fin::timestamp)
            )
            LIMIT 1
        ), 'N/A') AS ultimo_mensaje_enviado_a,
        -- ÚLTIMO MENSAJE LEÍDO DE
        COALESCE((
            SELECT CONCAT(u2.firstname, ' ', u2.lastname)
            FROM mdl_messages m
            JOIN mdl_message_conversations mc ON m.conversationid = mc.id
            JOIN mdl_message_conversation_members mcm ON mc.id = mcm.conversationid
            JOIN mdl_user u2 ON m.useridfrom = u2.id AND u2.id != u.id
            JOIN mdl_role_assignments ra2 ON u2.id = ra2.userid
            JOIN mdl_context ctx2 ON ra2.contextid = ctx2.id AND ctx2.contextlevel = 50
            JOIN mdl_message_user_actions mua ON m.id = mua.messageid AND mua.userid = u.id AND mua.action = 1
            WHERE mcm.userid = u.id
            AND ctx2.instanceid = c.id
            AND ra2.roleid IN (5, 11, 16, 17)
            AND m.timecreated BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) 
                                AND EXTRACT(EPOCH FROM :fecha_fin::timestamp)
            AND m.timecreated = (
                SELECT MAX(m2.timecreated)
                FROM mdl_messages m2
                JOIN mdl_message_conversations mc2 ON m2.conversationid = mc2.id
                JOIN mdl_message_conversation_members mcm2 ON mc2.id = mcm2.conversationid
                JOIN mdl_user u3 ON m2.useridfrom = u3.id AND u3.id != u.id
                JOIN mdl_role_assignments ra3 ON u3.id = ra3.userid
                JOIN mdl_context ctx3 ON ra3.contextid = ctx3.id AND ctx3.contextlevel = 50
                JOIN mdl_message_user_actions mua2 ON m2.id = mua2.messageid AND mua2.userid = u.id AND mua2.action = 1
                WHERE mcm2.userid = u.id
                AND ctx3.instanceid = c.id
                AND ra3.roleid IN (5, 11, 16, 17)
                AND m2.timecreated BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) 
                                    AND EXTRACT(EPOCH FROM :fecha_fin::timestamp)
            )
            LIMIT 1
        ), 'N/A') AS ultimo_mensaje_leido_de
    FROM mdl_user u
    JOIN mdl_role_assignments ra ON u.id = ra.userid
    JOIN mdl_role r ON ra.roleid = r.id
    JOIN mdl_context ctx ON ra.contextid = ctx.id AND ctx.contextlevel = 50
    JOIN mdl_course c ON ctx.instanceid = c.id
    WHERE u.username <> '12345678'
      AND c.id IN (556,557,533,550,532,552,534,554,566,547,563,537,565,536)
      AND r.id IN (3, 5, 9, 16, 17)
    ORDER BY c.fullname, u.lastname, u.firstname;";

    $stmt = $pdo->prepare($sql);
    $params = [
        ':fecha_inicio' => $fecha_inicio,
        ':fecha_fin' => $fecha_fin
    ];
    $stmt->execute($params);
    $resultados = $stmt->fetchAll(PDO::FETCH_ASSOC);

    function replaceEmptyWithZero($array) {
        return array_map(function($value) {
            return ($value === null || $value === '') ? 0 : $value;
        }, $array);
    }

    $resultados = array_map('replaceEmptyWithZero', $resultados);

    $estudiantes = array_values(array_filter($resultados, function($fila) {
        return in_array($fila['rol_id'], [5, 9, 16, 17]);
    }));
    $profesores = array_values(array_filter($resultados, function($fila) {
        return $fila['rol_id'] == 3;
    }));

    $fecha_para_nombre = date('Ymd', strtotime($fecha_fin));
    $nombre_estudiantes = "estudiantes_regencia_{$fecha_para_nombre}.csv";
    $nombre_profesores = "profesores_regencia_{$fecha_para_nombre}.csv";

    $temp_dir = sys_get_temp_dir();
    $temp_estudiantes = $temp_dir . '/' . uniqid('estudiantes_', true) . '.csv';
    $temp_profesores = $temp_dir . '/' . uniqid('profesores_', true) . '.csv';

    if (!empty($estudiantes)) {
        $file = fopen($temp_estudiantes, 'w');
        fwrite($file, "\xEF\xBB\xBF");
        fputcsv($file, array_keys($estudiantes[0]));
        foreach ($estudiantes as $fila) {
            fputcsv($file, $fila);
        }
        fclose($file);
    }

    if (!empty($profesores)) {
        $file = fopen($temp_profesores, 'w');
        fwrite($file, "\xEF\xBB\xBF");
        fputcsv($file, array_keys($profesores[0]));
        foreach ($profesores as $fila) {
            fputcsv($file, $fila);
        }
        fclose($file);
    }

    $zip = new ZipArchive();
    $zip_file = $temp_dir . '/' . uniqid('reporte_', true) . '.zip';
    if ($zip->open($zip_file, ZipArchive::CREATE) === TRUE) {
        if (file_exists($temp_estudiantes)) {
            $zip->addFile($temp_estudiantes, $nombre_estudiantes);
        }
        if (file_exists($temp_profesores)) {
            $zip->addFile($temp_profesores, $nombre_profesores);
        }
        $zip->close();
    }

    $mail = new PHPMailer(true);
    $mail->setFrom('noreply-univirtual@utp.edu.co', 'Reporte Moodle Regencia');
    foreach ($correo_destino as $correo) {
        $mail->addAddress($correo);
    }
    $mail->Subject = 'Reporte de Ingresos y Mensajes Acumulado de las asignaturas de regencia - ' . $fecha_para_nombre;
    $mail->Body = 'Cordial Saludo, Adjunto el Reporte de Ingresos y Mensajes Acumulado de las asignaturas de regencia, que contiene los archivos de estudiantes y profesores.';
    if (file_exists($zip_file)) {
        $mail->addAttachment($zip_file, "reporte_asignaturas_regencia_{$fecha_para_nombre}.zip");
    }
    $mail->send();

    if (file_exists($temp_estudiantes)) {
        unlink($temp_estudiantes);
    }
    if (file_exists($temp_profesores)) {
        unlink($temp_profesores);
    }
    if (file_exists($zip_file)) {
        unlink($zip_file);
    }

    mail($correo_notificacion, 'Estado Reporte', 'El reporte Ingresos y Mensajes Acumulado de las asignaturas de regencia fue enviado correctamente.');
} catch (Exception $e) {
    mail($correo_notificacion, 'Error Reporte regencia', 'Error: ' . $e->getMessage());
    echo "Error: " . $e->getMessage();
}
?>
