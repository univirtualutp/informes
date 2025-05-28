<?php
// Configuración inicial
date_default_timezone_set('America/Bogota');
// Cargar dependencias (solo PHPMailer para el envío de correos)
require 'vendor/autoload.php'; 
// Incluir archivo de configuración de la base de datos
require_once __DIR__ . '/db_moodle_config.php';

use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\Exception;

// Configuración de correos
$correo_destino = ['soporteunivirtual@utp.edu.co', 'univirtual-utp@utp.edu.co'];
$correo_notificacion = 'soporteunivirtual@utp.edu.co';
// Correo para pruebas
$correo_pruebas = 'daniel.pardo@utp.edu.co';

// Modo prueba (cambiar a false para producción)
$modo_prueba = true;

// Calcular fechas (martes a lunes)
$hoy = new DateTime();
$lunes = clone $hoy;
$lunes->modify('last monday');
$martes = clone $lunes;
$martes->modify('-6 days');
$fecha_inicio = $martes->format('Y-m-d 00:00:00');
$fecha_fin = $lunes->format('Y-m-d 23:59:59');

$cursos = [494, 415, 507, 481, 508, 482, 509, 485, 526, 510, 511, 486, 490, 416, 503, 504, 527, 417, 496, 497, 418, 498, 419, 475, 421, 420, 422, 423, 512, 513, 515, 488, 489, 424, 516, 517, 491, 518, 492, 519, 493, 520, 425, 476, 426, 505, 506, 479, 521, 428, 430, 522, 495, 499, 431, 453, 500, 523, 434, 524, 435, 436, 437, 438, 440, 502, 439, 452, 525, 442];

try {
    // Conexión a la base de datos PostgreSQL usando constantes definidas
    $pdo = new PDO(
        "pgsql:host=".DB_HOST.";dbname=".DB_NAME.";port=".DB_PORT, 
        DB_USER, 
        DB_PASS
    );
    $pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);

    // Crear parámetros nombrados para los cursos
    $cursos_params = [];
    foreach ($cursos as $i => $curso_id) {
        $cursos_params[":curso_$i"] = $curso_id;
    }
    $cursos_in = implode(',', array_keys($cursos_params));

    $sql = "WITH AllDays AS (
      SELECT generate_series(
            :fecha_inicio::timestamp,
            :fecha_fin::timestamp,
            interval '1 day'
        )::DATE AS fecha
    ),
    UserCourseDays AS (
      SELECT 
        u.id AS userid,
        c.id AS courseid,
        mr.name AS rol,
        mr.id AS roleid,
        COUNT(DISTINCT CASE WHEN mlsl.id IS NOT NULL THEN ad.fecha END) AS total_ingresos
      FROM mdl_user u
      CROSS JOIN AllDays ad
      LEFT JOIN mdl_role_assignments mra ON mra.userid = u.id
      LEFT JOIN mdl_role mr ON mra.roleid = mr.id
      LEFT JOIN mdl_context mc ON mc.id = mra.contextid
      LEFT JOIN mdl_course c ON c.id = mc.instanceid 
      LEFT JOIN mdl_logstore_standard_log mlsl ON mlsl.courseid = c.id
        AND mlsl.userid = u.id
        AND mlsl.action = 'viewed'
        AND mlsl.target IN ('course', 'course_module')
        AND CAST(to_timestamp(mlsl.timecreated) AS DATE) = ad.fecha
      WHERE mc.contextlevel = 50
        AND u.username NOT IN ('12345678')
        AND c.id IN ($cursos_in)
      GROUP BY u.id, c.id, mr.id
    )
    SELECT 
      u.username AS codigo,
      u.firstname AS nombre,
      u.lastname AS apellidos,
      u.email AS correo,
      c.fullname AS curso,
      mr.name AS rol,
      mr.id AS rol_id,
      COALESCE(ucd.total_ingresos, 0) AS total_ingresos
    FROM mdl_user u
    LEFT JOIN mdl_role_assignments mra ON mra.userid = u.id
    LEFT JOIN mdl_role mr ON mra.roleid = mr.id
    LEFT JOIN mdl_context mc ON mc.id = mra.contextid
    LEFT JOIN mdl_course c ON c.id = mc.instanceid 
    LEFT JOIN UserCourseDays ucd ON ucd.userid = u.id AND ucd.courseid = c.id AND ucd.roleid = mr.id
    WHERE mc.contextlevel = 50
      AND u.username NOT IN ('12345678')
      AND c.id IN ($cursos_in)
    GROUP BY u.id, u.username, u.firstname, u.lastname, u.email, c.id, c.fullname, mr.name, mr.id, ucd.total_ingresos
    ORDER BY c.fullname, u.lastname, u.firstname";

    // Preparar y ejecutar la consulta
    $stmt = $pdo->prepare($sql);

    // Combinar todos los parámetros
    $params = [
        ':fecha_inicio' => $fecha_inicio, 
        ':fecha_fin' => $fecha_fin
    ] + $cursos_params + $cursos_params; // Duplicamos para los dos IN clauses

    $stmt->execute($params);
    $resultados = $stmt->fetchAll(PDO::FETCH_ASSOC);

    // Función para reemplazar valores nulos o vacíos con 0
    function replaceEmptyWithZero($array) {
        return array_map(function($value) {
            return ($value === null || $value === '') ? 0 : $value;
        }, $array);
    }

    // Aplicar la función a cada fila de resultados
    $resultados = array_map('replaceEmptyWithZero', $resultados);

    // Separar resultados en estudiantes (roles 5, 9, 16, 17) y profesores (rol 3)
    $estudiantes = array_values(array_filter($resultados, function($fila) {
        return in_array($fila['rol_id'], [5, 9, 16, 17]);
    }));
    $profesores = array_values(array_filter($resultados, function($fila) {
        return $fila['rol_id'] == 3;
    }));

    // Definir nombres de archivos con fecha
    $fecha_para_nombre = date('Ymd', strtotime($fecha_fin));
    $nombre_estudiantes = "estudiantes_pregrado_{$fecha_para_nombre}.csv";
    $nombre_profesores = "profesores_pregrado_{$fecha_para_nombre}.csv";

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
    $mail->setFrom('noreply-univirtual@utp.edu.co', 'Reporte Moodle');
    
    if ($modo_prueba) {
        // En modo prueba solo se envía al correo de pruebas
        $mail->addAddress($correo_pruebas);
        $mail->Subject = '[PRUEBA] Reporte de Ingresos y Mensajes Semanal de asignaturas de pregrado - ' . $fecha_para_nombre;
    } else {
        // En modo producción se envía a los destinatarios reales
        foreach ($correo_destino as $correo) {
            $mail->addAddress($correo);
        }
        $mail->Subject = 'Reporte de Ingresos y Mensajes Semanal de asignaturas de pregrado - ' . $fecha_para_nombre;
    }
    
    $mail->Body = "Cordial Saludo,\n\nAdjunto el Reporte de Ingresos y Mensajes Semanal de asignaturas de pregrado correspondiente al período del {$fecha_inicio} al {$fecha_fin}. El reporte contiene los archivos de estudiantes y profesores.\n\nSaludos,\nReporte Moodle";
    $mail->isHTML(false); // Asegura que el correo sea texto plano
    
    if (file_exists($zip_file)) {
        $mail->addAttachment($zip_file, "reporte_asignaturas_pregrado_{$fecha_para_nombre}.zip");
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
    $mensaje_notificacion = $modo_prueba ? 
        "[PRUEBA] El reporte ingresos y mensajes pregrado fue generado correctamente en modo prueba." :
        "El reporte ingresos y mensajes pregrado fue enviado correctamente a los destinatarios.";
    
    mail($correo_notificacion, 'Estado Reporte', $mensaje_notificacion);
    
    if ($modo_prueba) {
        echo "Modo prueba: El reporte se generó correctamente y se envió solo a $correo_pruebas";
    }
} catch (PDOException $e) {
    mail($correo_notificacion, 'Error Reporte - Base de Datos', 'Error en la base de datos: ' . $e->getMessage());
    echo "Error en la base de datos: " . $e->getMessage();
} catch (PHPMailer\PHPMailer\Exception $e) {
    mail($correo_notificacion, 'Error Reporte - Correo', 'Error al enviar el correo: ' . $e->getMessage());
    echo "Error al enviar el correo: " . $e->getMessage();
} catch (Exception $e) {
    mail($correo_notificacion, 'Error Reporte - General', 'Error: ' . $e->getMessage());
    echo "Error: " . $e->getMessage();
}