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

// Calcular fechas (martes a lunes) - ajustado para 04 de agosto de 2025
$hoy = new DateTime('2025-08-04'); // Fecha específica solicitada
$lunes = clone $hoy;
$lunes->modify('next monday'); // Encuentra el próximo lunes
$martes = clone $lunes;
$martes->modify('-6 days'); // Retrocede 6 días para llegar al martes anterior
$fecha_inicio = $martes->format('Y-m-d 00:00:00');
$fecha_fin = $lunes->format('Y-m-d 23:59:59');

$cursos = ['786','583','804','596','820','790','821','819','789','580','799','616','617','797','798','842','844','621','805','584','581','802','619','627','800','801','618','588','815','589','590','594','595','787','788','605','607','793','573','791','606','792','608','609','795','611','794','610','586','814','623','622','624','733','574','604','796','576','612','577','614','830','615','783','579','784','785','591','587','810','625','582','803','620','592','816','817','593'];

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
    UserInfo AS (
      SELECT 
        uid.userid,
        MAX(CASE WHEN ufield.shortname = 'programa' THEN uid.data END) AS idprograma,
        MAX(CASE WHEN ufield.shortname = 'facultad' THEN uid.data END) AS idfacultad,
        MAX(CASE WHEN ufield.shortname = 'edad' THEN uid.data END) AS edad,
        MAX(CASE WHEN ufield.shortname = 'genero' THEN uid.data END) AS genero,
        MAX(CASE WHEN ufield.shortname = 'celular' THEN uid.data END) AS celular,
        MAX(CASE WHEN ufield.shortname = 'estrato' THEN uid.data END) AS estrato
      FROM mdl_user_info_data uid
      JOIN mdl_user_info_field ufield ON ufield.id = uid.fieldid
      WHERE ufield.shortname IN ('programa', 'facultad', 'edad', 'genero', 'celular', 'estrato')
      GROUP BY uid.userid
    ),
    CourseInfo AS (
      SELECT 
        cfidata.instanceid,
        MAX(CASE WHEN cfield.shortname = 'codigo_curso' THEN cfidata.value END) AS idcodigo,
        MAX(CASE WHEN cfield.shortname = 'grupo_curso' THEN cfidata.value END) AS grupo,
        MAX(CASE WHEN cfield.shortname = 'periodo' THEN cfidata.value END) AS periodo,
        MAX(CASE WHEN cfield.shortname = 'nivel_educativo' THEN cfidata.value END) AS nivel
      FROM mdl_customfield_data cfidata
      JOIN mdl_customfield_field cfield ON cfield.id = cfidata.fieldid
      WHERE cfield.shortname IN ('codigo_curso', 'grupo_curso', 'periodo', 'nivel_educativo')
      GROUP BY cfidata.instanceid
    )
    SELECT 
      u.username AS codigo,
      u.firstname AS nombre,
      u.lastname AS apellidos,
      u.email AS correo,
      c.fullname AS curso,
      mr.name as rol,
      mr.id as rol_id, 
      ui.idprograma,
      ui.idfacultad,
      ci.idcodigo,
      ci.grupo,
      ci.periodo,
      ci.nivel,
      ui.edad,
      ui.genero,
      ui.celular,
      ui.estrato,
      ad.fecha,
      CASE WHEN EXISTS (
        SELECT 1
        FROM mdl_logstore_standard_log mlsl
        WHERE mlsl.courseid = c.id
          AND mlsl.userid = u.id
          AND mlsl.action = 'viewed'
          AND mlsl.target IN ('course', 'course_module')
          AND CAST(to_timestamp(mlsl.timecreated) AS DATE) = ad.fecha
      ) THEN 1 ELSE 0 END AS ingreso_dia,
      -- Subconsulta para obtener el número de documento del docente
      (SELECT CONCAT(u2.idnumber) AS Teacher
       FROM mdl_role_assignments AS ra
       JOIN mdl_context AS ctx ON ra.contextid = ctx.id
       JOIN mdl_user AS u2 ON u2.id = ra.userid
       WHERE ra.roleid = 3
         AND ctx.instanceid = c.id
       LIMIT 1) AS nrodoc
    FROM mdl_user u
    CROSS JOIN AllDays ad
    LEFT JOIN mdl_role_assignments mra ON mra.userid = u.id
    LEFT JOIN mdl_role mr ON mra.roleid = mr.id
    LEFT JOIN mdl_context mc ON mc.id = mra.contextid
    LEFT JOIN mdl_course c ON c.id = mc.instanceid 
    LEFT JOIN UserInfo ui ON ui.userid = u.id
    LEFT JOIN CourseInfo ci ON ci.instanceid = c.id
    WHERE mc.contextlevel = 50
      AND u.username NOT IN ('12345678')
      AND c.id IN ($cursos_in)
      AND mr.id IN ('9')
    GROUP BY u.id, c.id, u.email, mr.id, mr.name, ad.fecha, ui.idprograma, ui.idfacultad, ci.idcodigo, ci.grupo, ci.periodo, ci.nivel, ui.edad, ui.genero, ui.celular, ui.estrato
    ORDER BY c.fullname, ad.fecha, u.lastname, u.firstname";

    // Preparar y ejecutar la consulta
    $stmt = $pdo->prepare($sql);

    // Combinar todos los parámetros
    $params = [
        ':fecha_inicio' => $fecha_inicio, 
        ':fecha_fin' => $fecha_fin
    ] + $cursos_params;

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
        $mail->Subject = '[PRUEBA] Reporte de Ingresos Semanal de asignaturas de pregrado - ' . $fecha_para_nombre;
    } else {
        // En modo producción se envía a los destinatarios reales
        foreach ($correo_destino as $correo) {
            $mail->addAddress($correo);
        }
        $mail->Subject = 'Reporte de Ingresos Semanal de asignaturas de pregrado - ' . $fecha_para_nombre;
    }
    
    $mail->Body = "Cordial Saludo,\n\nAdjunto el Reporte de Ingresos Semanal de asignaturas de pregrado correspondiente al período del {$fecha_inicio} al {$fecha_fin}. El reporte contiene los archivos de estudiantes y profesores.\n\nSaludos,\nReporte Moodle";
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