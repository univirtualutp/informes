<?php
// Cargar dependencias (solo PHPMailer para el envío de correos)
require 'vendor/autoload.php'; 
// Incluir archivo de configuración de la base de datos
require_once __DIR__ . '/db_moodle_config.php';
use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\Exception;

// Configuración de la base de datos y correos
$host = 'localhost';
$dbname = 'moodle';
$user = 'moodle';
$pass = 'M00dl3';
$correo_destino = ['soporteunivirtual@utp.edu.co', 'univirtual-utp@utp.edu.co'];
$correo_notificacion = 'soporteunivirtual@utp.edu.co';

// Calcular fechas (martes a lunes)
$hoy = new DateTime();
$lunes = clone $hoy;
$lunes->modify('last monday');
$martes = clone $lunes;
$martes->modify('-6 days');

$fecha_inicio = $martes->format('Y-m-d 00:00:00');
$fecha_fin = $lunes->format('Y-m-d 23:59:59');

try {
    // Conexión a la base de datos PostgreSQL
    $pdo = new PDO(
        "pgsql:host=".DB_HOST.";dbname=".DB_NAME.";port=".DB_PORT, 
        DB_USER, 
        DB_PASS
    );
    $pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);

    // Consulta SQL
    $sql = "WITH AllDays AS (
        SELECT generate_series(
            :fecha_inicio::timestamp,
            :fecha_fin::timestamp,
            interval '1 day'
        )::DATE AS fecha
    ),
    CourseLog AS (
        SELECT 
            u.id AS userid,
            c.id AS courseid,
            COUNT(DISTINCT DATE_TRUNC('day', to_timestamp(l.timecreated))) AS dias_activos,
            COUNT(DISTINCT l.id) AS ingresos_totales
        FROM mdl_logstore_standard_log AS l
        JOIN mdl_user AS u ON l.userid = u.id
        JOIN mdl_context AS ctx ON l.contextid = ctx.id
        JOIN mdl_course AS c ON ctx.instanceid = c.id
        WHERE 
            l.action = 'viewed' 
            AND l.target IN ('course', 'course_module')
            AND l.timecreated BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) 
                                  AND EXTRACT(EPOCH FROM :fecha_fin::timestamp)
        GROUP BY u.id, c.id
    )
    SELECT DISTINCT
        u.username AS codigo,
        u.firstname AS nombre,
        u.lastname AS apellidos,
        r.id AS rol_id,
        r.name AS rol,
        u.email AS correo,
        c.id AS id_curso,
        c.fullname AS curso,
        COALESCE(cl.dias_activos, 0) AS ingresos,
        COALESCE(cl.ingresos_totales, 0) AS interacciones_en_aula,
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
        COALESCE((
            SELECT TO_TIMESTAMP(MAX(m.timecreated))
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
        ), NULL) AS ultimo_mensaje_enviado,
        COALESCE((
            SELECT TO_TIMESTAMP(MAX(m.timecreated))
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
        ), NULL) AS ultimo_mensaje_recibido,
        COALESCE((
            SELECT TO_TIMESTAMP(MAX(m.timecreated))
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
        ), NULL) AS ultimo_mensaje_leido,
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
        ), NULL) AS ultimo_mensaje_enviado_a,
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
        ), NULL) AS ultimo_mensaje_leido_de
    FROM mdl_user u
    JOIN mdl_role_assignments ra ON ra.userid = u.id
    JOIN mdl_role r ON ra.roleid = r.id 
    JOIN mdl_context mc ON ra.contextid = mc.id
    JOIN mdl_course c ON mc.instanceid = c.id
    LEFT JOIN CourseLog cl ON cl.userid = u.id AND cl.courseid = c.id
    WHERE mc.contextlevel = 50
      AND u.username <> '12345678'
      AND c.id IN (556,557,533,550,532,552,534,554,566,547,563,537,565,536)
      AND r.id IN (3,5,9,16,17)
    ORDER BY c.fullname, u.lastname, u.firstname;";

    // Preparar y ejecutar la consulta
    $stmt = $pdo->prepare($sql);

    // Verificar que los parámetros estén correctamente definidos
    $params = [
        ':fecha_inicio' => $fecha_inicio,
        ':fecha_fin' => $fecha_fin
    ];

    $stmt->execute($params); // Pasar los parámetros aquí
    $resultados = $stmt->fetchAll(PDO::FETCH_ASSOC);

    // Función para reemplazar valores nulos o vacíos con 0
    function replaceEmptyWithZero($array) {
        return array_map(function($value) {
            return ($value === null || $value === '') ? 0 : $value;
        }, $array);
    }

    // Aplicar la función a cada fila de resultados
    $resultados = array_map('replaceEmptyWithZero', $resultados);

    // Separar resultados en estudiantes (roles 5, 9, 16) y profesores (rol 3)
    $estudiantes = array_values(array_filter($resultados, function($fila) {
        return in_array($fila['rol_id'], [5, 9, 16, 17]);
    }));
    $profesores = array_values(array_filter($resultados, function($fila) {
        return $fila['rol_id'] == 3;
    }));

    // Definir nombres de archivos con fecha
    $fecha_para_nombre = date('Ymd', strtotime($fecha_fin));
    $nombre_estudiantes = "estudiantes_regencia_{$fecha_para_nombre}.csv";
    $nombre_profesores = "profesores_regencia_{$fecha_para_nombre}.csv";

    // Crear archivos temporales
    $temp_dir = sys_get_temp_dir();
    $temp_estudiantes = $temp_dir . '/' . uniqid('estudiantes_', true) . '.csv';
    $temp_profesores = $temp_dir . '/' . uniqid('profesores_', true) . '.csv';

    // Generar archivo CSV para estudiantes si hay datos
    if (!empty($estudiantes)) {
        $file = fopen($temp_estudiantes, 'w');
        // Agregar BOM (Byte Order Mark) para UTF-8
        fwrite($file, "\xEF\xBB\xBF");
        // Escribir la cabecera (nombres de las columnas)
        fputcsv($file, array_keys($estudiantes[0]));
        // Escribir los datos
        foreach ($estudiantes as $fila) {
            fputcsv($file, $fila);
        }
        fclose($file);
    }

    // Generar archivo CSV para profesores si hay datos
    if (!empty($profesores)) {
        $file = fopen($temp_profesores, 'w');
        // Agregar BOM (Byte Order Mark) para UTF-8
        fwrite($file, "\xEF\xBB\xBF");
        // Escribir la cabecera (nombres de las columnas)
        fputcsv($file, array_keys($profesores[0]));
        // Escribir los datos
        foreach ($profesores as $fila) {
            fputcsv($file, $fila);
        }
        fclose($file);
    }

    // Crear archivo ZIP
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

    // Configurar y enviar correo con PHPMailer
    $mail = new PHPMailer(true);
    $mail->setFrom('noreply-univirtual@utp.edu.co', 'Reporte Moodle Regencia');
    foreach ($correo_destino as $correo) {
        $mail->addAddress($correo);
    }
    $mail->Subject = 'Reporte de Ingresos y Mensajes Semanal del programa en Regencia en Farmacia -' . $fecha_para_nombre;
    $mail->Body = 'Cordial Saludo, Adjunto el Reporte de Ingresos y Mensajes Semanal del programa en Regencia en Farmacia, que contiene los archivos de estudiantes y profesores.';
    if (file_exists($zip_file)) {
        $mail->addAttachment($zip_file, "reporte_asignaturas_regencia_{$fecha_para_nombre}.zip");
    }
    $mail->send();

    // Eliminar archivos temporales
    if (file_exists($temp_estudiantes)) {
        unlink($temp_estudiantes);
    }
    if (file_exists($temp_profesores)) {
        unlink($temp_profesores);
    }
    if (file_exists($zip_file)) {
        unlink($zip_file);
    }

    // Enviar notificación de éxito
    mail($correo_notificacion, 'Estado Reporte', 'El reporte  ingresos y mensajes del programa en Regencia en Farmacia fue enviado correctamente.');
} catch (Exception $e) {
    // Enviar notificación de error
    mail($correo_notificacion, 'Error Reporte regencia', 'Error: ' . $e->getMessage());
    echo "Error: " . $e->getMessage(); // Mostrar el error en la consola
}
?>
