<?php
// Configuración inicial
date_default_timezone_set('America/Bogota');
require 'vendor/autoload.php'; 
require_once __DIR__ . '/db_moodle_config.php';

use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\Exception;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Cell\DataType;

// Configuración de correos
$correo_destino = ['soporteunivirtual@utp.edu.co', 'univirtual-utp@utp.edu.co'];
$correo_notificacion = 'soporteunivirtual@utp.edu.co';
$correo_pruebas = 'daniel.pardo@utp.edu.co';

// Modo prueba (cambiar a false para producción)
$modo_prueba = true;

// Configuración de fechas
$fecha_inicio = '2025-08-13 00:00:00'; // Fecha fija de inicio
$fecha_inicio_simple = '2025-08-13';

// Calcular fecha final (último lunes a las 23:59:59)
$hoy = new DateTime();
$lunes_pasado = clone $hoy;
$lunes_pasado->modify('last monday');
$fecha_fin = $lunes_pasado->format('Y-m-d 23:59:59');
$fecha_fin_simple = $lunes_pasado->format('Y-m-d');

$cursos = ['734','841','736','839','735','838','737','840'];

// Función para reemplazar valores nulos o vacíos con 0
function sanitizeData($data) {
    foreach ($data as &$row) {
        foreach ($row as &$value) {
            if ($value === null || $value === '') {
                $value = 0;
            }
        }
    }
    return $data;
}

try {
    // Conexión a la base de datos
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

    // CONSULTA 1 - Datos detallados (Hoja "Power BI")
    $sql_detalle = "WITH AllDays AS (
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
        AND mr.id IN ('5','9')
      GROUP BY u.id, c.id, mr.id
    ),
    UserCustomFields AS (
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
      mr.name AS rol,
      mr.id AS rol_id,
      COALESCE(ucf.idprograma, '0') AS idprograma,
      COALESCE(ucf.idfacultad, '0') AS idfacultad,
      COALESCE(ci.idcodigo, '0') AS idcodigo,
      COALESCE(ci.grupo, '0') AS grupo,
      COALESCE(ci.periodo, '0') AS periodo,
      COALESCE(ci.nivel, '0') AS nivel,
      COALESCE(ucf.edad, '0') AS edad,
      COALESCE(ucf.genero, '0') AS genero,
      COALESCE(ucf.celular, '0') AS celular,
      COALESCE(ucf.estrato, '0') AS estrato,
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
      COALESCE(ucd.total_ingresos, 0) AS total_ingresos,
      COALESCE((SELECT CONCAT(u2.idnumber) 
       FROM mdl_role_assignments AS ra
       JOIN mdl_context AS ctx ON ra.contextid = ctx.id
       JOIN mdl_user AS u2 ON u2.id = ra.userid
       WHERE ra.roleid = 3
         AND ctx.instanceid = c.id
       LIMIT 1), '0') AS nrodoc
    FROM mdl_user u
    CROSS JOIN AllDays ad
    LEFT JOIN mdl_role_assignments mra ON mra.userid = u.id
    LEFT JOIN mdl_role mr ON mra.roleid = mr.id
    LEFT JOIN mdl_context mc ON mc.id = mra.contextid
    LEFT JOIN mdl_course c ON c.id = mc.instanceid 
    LEFT JOIN UserCustomFields ucf ON ucf.userid = u.id
    LEFT JOIN CourseInfo ci ON ci.instanceid = c.id
    LEFT JOIN UserCourseDays ucd ON ucd.userid = u.id AND ucd.courseid = c.id AND ucd.roleid = mr.id
    WHERE mc.contextlevel = 50
      AND u.username NOT IN ('12345678')
      AND c.id IN ($cursos_in)
      AND mr.id IN ('5','9','3')
    GROUP BY u.id, u.username, u.firstname, u.lastname, u.email, c.id, c.fullname, mr.id, mr.name, ad.fecha, 
             ucf.idprograma, ucf.idfacultad, ci.idcodigo, ci.grupo, ci.periodo, ci.nivel, 
             ucf.edad, ucf.genero, ucf.celular, ucf.estrato, ucd.total_ingresos
    ORDER BY c.fullname, ad.fecha, u.lastname, u.firstname";

    // CONSULTA 2 - Datos resumidos (Hoja "Resumen")
    $sql_resumen = "WITH AllDays AS (
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
        AND mr.id IN ('5','9')
      GROUP BY u.id, c.id, mr.id
    ),
    UserCustomFields AS (
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
    )
    SELECT 
      u.username AS codigo,
      u.firstname AS nombre,
      u.lastname AS apellidos,
      u.email AS correo,
      c.fullname AS curso,
      mr.name AS rol,
      COALESCE(ucf.idprograma, '0') AS idprograma,
      COALESCE(u.department, '0') as programa,
      COALESCE(ucf.idfacultad, '0') AS idfacultad,
      COALESCE(u.institution, '0') as facultad,
      COALESCE(ucf.edad, '0') AS edad,
      COALESCE(ucf.genero, '0') AS genero,
      COALESCE(ucf.celular, '0') AS celular,
      COALESCE(ucf.estrato, '0') AS estrato,
      COALESCE(ucd.total_ingresos, 0) AS total_ingresos
    FROM mdl_user u
    LEFT JOIN mdl_role_assignments mra ON mra.userid = u.id
    LEFT JOIN mdl_role mr ON mra.roleid = mr.id
    LEFT JOIN mdl_context mc ON mc.id = mra.contextid
    LEFT JOIN mdl_course c ON c.id = mc.instanceid 
    LEFT JOIN UserCourseDays ucd ON ucd.userid = u.id AND ucd.courseid = c.id AND ucd.roleid = mr.id
    LEFT JOIN UserCustomFields ucf ON ucf.userid = u.id
    WHERE mc.contextlevel = 50
      AND u.username NOT IN ('12345678')
      AND c.id IN ($cursos_in)
      AND mr.id IN ('5','9','3')
    GROUP BY u.id, u.username, u.firstname, u.lastname, u.email, c.id, c.fullname, mr.name, ucd.total_ingresos, 
             ucf.idprograma, ucf.idfacultad, ucf.edad, ucf.genero, ucf.celular, ucf.estrato, u.department, u.institution
    ORDER BY c.fullname, u.lastname, u.firstname";

    // Combinar todos los parámetros
    $params = [
        ':fecha_inicio' => $fecha_inicio, 
        ':fecha_fin' => $fecha_fin
    ] + $cursos_params;

    // Función para crear una hoja en el Excel con manejo de valores nulos
    function createSheet($spreadsheet, $data, $title, $sheetIndex = 0) {
        if ($sheetIndex > 0) {
            $sheet = $spreadsheet->createSheet();
        } else {
            $sheet = $spreadsheet->getActiveSheet();
        }
        
        $sheet->setTitle(substr($title, 0, 31));
        
        if (!empty($data)) {
            // Escribir encabezados
            $headers = array_keys($data[0]);
            $sheet->fromArray($headers, null, 'A1');
            
            // Escribir datos fila por fila
            $rowNumber = 2;
            foreach ($data as $row) {
                $colLetter = 'A';
                foreach ($row as $value) {
                    // Convertir valores nulos o vacíos a 0
                    if ($value === null || $value === '') {
                        $value = 0;
                    }
                    
                    // Establecer el valor en la celda
                    $cell = $sheet->getCell($colLetter . $rowNumber);
                    if (is_numeric($value)) {
                        $cell->setValueExplicit($value, DataType::TYPE_NUMERIC);
                    } else {
                        $cell->setValueExplicit($value, DataType::TYPE_STRING);
                    }
                    
                    $colLetter++;
                }
                $rowNumber++;
            }
        }
        
        return $sheet;
    }

    // Procesar consulta detalle
    $stmt_detalle = $pdo->prepare($sql_detalle);
    $stmt_detalle->execute($params);
    $resultados_detalle = $stmt_detalle->fetchAll(PDO::FETCH_ASSOC);
    $resultados_detalle = sanitizeData($resultados_detalle);

    // Procesar consulta resumen
    $stmt_resumen = $pdo->prepare($sql_resumen);
    $stmt_resumen->execute($params);
    $resultados_resumen = $stmt_resumen->fetchAll(PDO::FETCH_ASSOC);
    $resultados_resumen = sanitizeData($resultados_resumen);

    // Filtrar resultados
    $estudiantes_detalle = array_values(array_filter($resultados_detalle, function($fila) {
        return in_array($fila['rol_id'], [5, 9]);
    }));
    
    $profesores_detalle = array_values(array_filter($resultados_detalle, function($fila) {
        return $fila['rol_id'] == 3;
    }));

    $estudiantes_resumen = array_values(array_filter($resultados_resumen, function($fila) {
        return in_array($fila['rol'], ['Estudiante', 'Student']);
    }));
    
    $profesores_resumen = array_values(array_filter($resultados_resumen, function($fila) {
        return $fila['rol'] == 'Profesor';
    }));

    // Generar archivos Excel con dos hojas
    $fecha_para_nombre = $lunes_pasado->format('Ymd');
    $nombre_estudiantes = "estudiantes_pregrado_regencia_c2_s1_acumulado_{$fecha_para_nombre}.xlsx";
    $nombre_profesores = "profesores_pregrado_regencia_c2_s1_acumulado_{$fecha_para_nombre}.xlsx";

    $temp_dir = sys_get_temp_dir();
    
    // Archivo estudiantes
    if (!empty($estudiantes_detalle) || !empty($estudiantes_resumen)) {
        $spreadsheet = new Spreadsheet();
        
        // Hoja 1: Power BI (detalle)
        createSheet($spreadsheet, $estudiantes_detalle, 'Power BI', 0);
        
        // Hoja 2: Resumen
        createSheet($spreadsheet, $estudiantes_resumen, 'Resumen', 1);
        
        $writer = new Xlsx($spreadsheet);
        $temp_estudiantes = $temp_dir . '/' . uniqid('estudiantes_', true) . '.xlsx';
        $writer->save($temp_estudiantes);
    }

    // Archivo profesores
    if (!empty($profesores_detalle) || !empty($profesores_resumen)) {
        $spreadsheet = new Spreadsheet();
        
        // Hoja 1: Power BI (detalle)
        createSheet($spreadsheet, $profesores_detalle, 'Power BI', 0);
        
        // Hoja 2: Resumen
        createSheet($spreadsheet, $profesores_resumen, 'Resumen', 1);
        
        $writer = new Xlsx($spreadsheet);
        $temp_profesores = $temp_dir . '/' . uniqid('profesores_', true) . '.xlsx';
        $writer->save($temp_profesores);
    }

    // Crear ZIP
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

    // Enviar correo
    $mail = new PHPMailer(true);
    $mail->setFrom('noreply-univirtual@utp.edu.co', 'Reporte Moodle');
    
    if ($modo_prueba) {
        $mail->addAddress($correo_pruebas);
        $mail->Subject = '[PRUEBA] Reporte de Ingresos Acumulado (13 Ago - '.$fecha_fin_simple.')';
    } else {
        foreach ($correo_destino as $correo) {
            $mail->addAddress($correo);
        }
        $mail->Subject = 'Reporte de Ingresos Acumulado (13 Ago - '.$fecha_fin_simple.')';
    }
    
    $mail->Body = "Reporte correspondiente al período del {$fecha_inicio_simple} al {$fecha_fin_simple}";
    $mail->isHTML(false);
    
    if (file_exists($zip_file)) {
        $mail->addAttachment($zip_file, "reporte_regencia_c2_s1_acumulado".$fecha_fin_simple.".zip");
    }
    
    $mail->send();

    // Limpieza
    if (file_exists($temp_estudiantes)) unlink($temp_estudiantes);
    if (file_exists($temp_profesores)) unlink($temp_profesores);
    if (file_exists($zip_file)) unlink($zip_file);

    // Notificación
    $mensaje_notificacion = $modo_prueba ? 
        "[PRUEBA] Reporte generado correctamente acumulado (13 Ago - $fecha_fin_simple)" :
        "Reporte enviado correctamente acumulado (13 Ago - $fecha_fin_simple)";
    
    mail($correo_notificacion, 'Estado Reporte', $mensaje_notificacion);
    
    if ($modo_prueba) {
        echo "Modo prueba: Reporte generado y enviado a $correo_pruebas<br>";
        echo "Periodo del reporte: {$fecha_inicio_simple} al {$fecha_fin_simple}<br>";
        echo "Archivos generados:<br>";
        echo "- $nombre_estudiantes (con hojas 'Power BI' y 'Resumen')<br>";
        echo "- $nombre_profesores (con hojas 'Power BI' y 'Resumen')<br>";
    }
} catch (PDOException $e) {
    mail($correo_notificacion, 'Error Reporte - BD', 'Error: ' . $e->getMessage());
    echo "Error BD: " . $e->getMessage();
} catch (PHPMailer\PHPMailer\Exception $e) {
    mail($correo_notificacion, 'Error Reporte - Correo', 'Error: ' . $e->getMessage());
    echo "Error correo: " . $e->getMessage();
} catch (Exception $e) {
    mail($correo_notificacion, 'Error Reporte', 'Error: ' . $e->getMessage());
    echo "Error: " . $e->getMessage();
}