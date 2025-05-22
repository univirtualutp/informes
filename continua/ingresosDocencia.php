<?php
// Cargar dependencias (solo PHPMailer para el envío de correos)
require '/root/scripts/informes/vendor/autoload.php';
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
    $pdo = new PDO("pgsql:host=$host;dbname=$dbname", $user, $pass);
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
            AND l.target = 'course'
            AND l.timecreated BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) 
                                  AND EXTRACT(EPOCH FROM :fecha_fin::timestamp)
        GROUP BY u.id, c.id
    ),
    UltimoMensajeEnviado AS (
        SELECT 
            mm.useridfrom AS userid,
            MAX(mm.timecreated) AS ultimo_mensaje_enviado
        FROM mdl_messages mm
        WHERE mm.timecreated BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) 
                                 AND EXTRACT(EPOCH FROM :fecha_fin::timestamp)
        GROUP BY mm.useridfrom
    )
    SELECT 
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
        (SELECT COALESCE(count(*), 0)
         FROM mdl_message_user_actions mmua 
         WHERE mmua.timecreated BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) 
           AND EXTRACT(EPOCH FROM :fecha_fin::timestamp)
           AND mmua.userid = u.id
           AND EXISTS (
               SELECT 1
               FROM mdl_user um
               JOIN mdl_role_assignments am ON am.userid = um.id
               JOIN mdl_role rm ON am.roleid = rm.id
               JOIN mdl_context tm ON tm.id = am.contextid
               JOIN mdl_course cm ON cm.id = tm.instanceid
               WHERE tm.contextlevel = 50
                 AND um.username NOT IN ('12345678')
                 AND cm.id = c.id
                 AND rm.id IN ('3','5')
           )
        ) AS mensajes_recibidos_leidos,
        (SELECT COALESCE(count(*), 0)
         FROM mdl_messages mm 
         WHERE mm.timecreated BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) 
           AND EXTRACT(EPOCH FROM :fecha_fin::timestamp)
           AND mm.useridfrom = u.id
           AND mm.conversationid IN (
               SELECT um.id
               FROM mdl_user um
               JOIN mdl_role_assignments am ON am.userid = um.id
               JOIN mdl_role rm ON am.roleid = rm.id
               JOIN mdl_context tm ON tm.id = am.contextid
               JOIN mdl_course cm ON cm.id = tm.instanceid
               WHERE tm.contextlevel = 50
                 AND um.username NOT IN ('12345678')
                 AND cm.id = c.id
                 AND rm.id IN ('3','5')
           )
        ) AS mensajes_enviados_sin_leer,
        CASE
            WHEN ume.ultimo_mensaje_enviado IS NULL OR ume.ultimo_mensaje_enviado = 0 THEN 'No ha enviado Mensaje'
            ELSE to_char(to_timestamp(ume.ultimo_mensaje_enviado), 'YYYY-MM-DD HH24:MI:SS')
        END AS ultimo_mensaje_enviado
    FROM mdl_user u
    JOIN mdl_role_assignments ra ON ra.userid = u.id
    JOIN mdl_role r ON ra.roleid = r.id 
    JOIN mdl_context mc ON ra.contextid = mc.id
    JOIN mdl_course c ON mc.instanceid = c.id
    LEFT JOIN CourseLog cl ON cl.userid = u.id AND cl.courseid = c.id
    LEFT JOIN UltimoMensajeEnviado ume ON ume.userid = u.id
    WHERE mc.contextlevel = 50
      AND u.username <> '12345678'
      AND c.id IN ('540','541','542','543','544','545')
      AND r.id IN ('3','5')
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
        return in_array($fila['rol_id'], [5, 9, 16]);
    }));
    $profesores = array_values(array_filter($resultados, function($fila) {
        return $fila['rol_id'] == 3;
    }));

    // Definir nombres de archivos con fecha
    $fecha_para_nombre = date('Ymd', strtotime($fecha_fin));
    $nombre_estudiantes = "estudiantes_docencia_{$fecha_para_nombre}.csv";
    $nombre_profesores = "profesores_docencia_{$fecha_para_nombre}.csv";

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
    $zip_file = $temp_dir . '/' . uniqid('reporteDDD_', true) . '.zip';
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
    $mail->setFrom('noreply-univirtual@utp.edu.co', 'Reporte Moodle Docencia Digital');
    foreach ($correo_destino as $correo) {
        $mail->addAddress($correo);
    }
    $mail->Subject = 'Reporte de Ingresos y Mensajes Semanal del Diplomado en Docencia Digital  - ' . $fecha_para_nombre;
    $mail->Body = 'Cordial Saludo, Adjunto el Reporte de Ingresos y Mensajes Semanal del diplomado en Docencia Digital, que contiene los archivos de estudiantes y profesores.';
    if (file_exists($zip_file)) {
        $mail->addAttachment($zip_file, "reporte_semanal_dgp__{$fecha_para_nombre}.zip");
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
    mail($correo_notificacion, 'Estado Reporte', 'El reporte del diplomado en docencia digital fue enviado correctamente.');
} catch (Exception $e) {
    // Enviar notificación de error
    mail($correo_notificacion, 'Error Reporte', 'Error: ' . $e->getMessage());
    echo "Error: " . $e->getMessage(); // Mostrar el error en la consola
}
?>
