<?php
// Configuración inicial
date_default_timezone_set('America/Bogota');

// Autoload de Composer
require __DIR__ . '/vendor/autoload.php';

// Incluir archivo de configuración de la base de datos
require_once __DIR__ . '/db_moodle_config.php';

use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\Exception;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// =============================================
// CONFIGURACIÓN
// =============================================
$correo_destino = ['soporteunivirtual@utp.edu.co','univirtual-utp@utp.edu.co'];
$correo_notificacion = 'soporteunivirtual@utp.edu.co';
$remitente = 'noreply-univirtual@utp.edu.co';

$cursos = [494, 415, 507, 481, 508, 482, 509, 485, 526, 510, 511, 486, 490, 416, 503, 504, 527, 417, 496, 497, 418, 498, 419, 475, 421, 420, 422, 423, 512, 513, 515, 488, 489, 424, 516, 517, 491, 518, 492, 519, 493, 520, 425, 476, 426, 505, 506, 479, 521, 428, 430, 522, 495, 499, 431, 453, 500, 523, 434, 524, 435, 436, 437, 438, 440, 502, 439, 452, 525, 442];
$cursos_str = implode("','", $cursos);

// =============================================
// CÁLCULO DE FECHAS (inicio fijo, fin dinámico)
// =============================================
$hoy = new DateTime('now', new DateTimeZone('America/Bogota'));

// Fecha de inicio FIJA (2025-02-03 00:00:00)
$fecha_inicio = '2025-02-03 00:00:00';
$fecha_inicio_simple = '2025-02-03';

// Fecha final DINÁMICA (último lunes a las 23:59:59)
$lunes_pasado = clone $hoy;
$lunes_pasado->modify('last monday');
$fecha_fin = $lunes_pasado->format('Y-m-d 23:59:59');
$fecha_fin_simple = $lunes_pasado->format('Y-m-d');

// Crear objeto DateTime para la fecha de inicio para formatearla
$fecha_inicio_obj = new DateTime($fecha_inicio, new DateTimeZone('America/Bogota'));

// Fecha para el nombre del archivo (opcional, según necesidad)
$fecha_para_nombre = $hoy->format('Y-m-d');

// Nombres de archivos
$temp_profesores = 'profesores_' . $fecha_para_nombre . '.xlsx';
$temp_estudiantes = 'estudiantes_' . $fecha_para_nombre . '.xlsx';
$temp_asesores = 'asesores_' . $fecha_para_nombre . '.xlsx';
$temp_general = 'general_' . $fecha_para_nombre . '.xlsx';
$zip_file = 'reporte_asignaturas_pregrado_' . $fecha_para_nombre . '.zip';

try {
    // Conexión a la base de datos usando constantes del archivo de configuración
    $dsn = "pgsql:host=".DB_HOST.";port=".DB_PORT.";dbname=".DB_NAME;
    $db = new PDO($dsn, DB_USER, DB_PASS);
    $db->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
    $db->exec("SET TIME ZONE 'America/Bogota'");

    // CONSULTA SQL IDÉNTICA A LA ORIGINAL
    $sql = "
WITH AllDays AS (
  SELECT generate_series(
        timestamp '$fecha_inicio',
        timestamp '$fecha_fin',
    interval '1 day'
  )::DATE AS fecha
),
UserCourseAccess AS (
  SELECT 
    u.id AS userid,
    c.id AS courseid,
    COUNT(DISTINCT CAST(to_timestamp(mlsl.timecreated) AS DATE)) AS access_days
  FROM mdl_user u
  JOIN mdl_role_assignments mra ON mra.userid = u.id
  JOIN mdl_role mr ON mra.roleid = mr.id
  JOIN mdl_context mc ON mc.id = mra.contextid
  JOIN mdl_course c ON c.id = mc.instanceid
  LEFT JOIN mdl_logstore_standard_log mlsl ON mlsl.courseid = c.id
    AND mlsl.userid = u.id
    AND mlsl.action = 'viewed'
    AND mlsl.target IN ('course', 'course_module')
    AND CAST(to_timestamp(mlsl.timecreated) AS DATE) BETWEEN '$fecha_inicio_simple ' AND '$fecha_fin_simple'
  WHERE mc.contextlevel = 50
    AND u.username NOT IN ('12345678')
    AND c.id IN ('$cursos_str')
    AND mr.id IN ('3','5','9','11','16','17')
  GROUP BY u.id, c.id
)
SELECT 
  u.username AS codigo,
  mr.name AS rol,
  u.firstname AS nombre,
  u.lastname AS apellidos,
  u.email AS correo,
  c.fullname AS curso,
  COALESCE(uca.access_days, 0) AS total_ingresos
FROM mdl_user u
JOIN mdl_role_assignments mra ON mra.userid = u.id
JOIN mdl_role mr ON mra.roleid = mr.id
JOIN mdl_context mc ON mc.id = mra.contextid
JOIN mdl_course c ON c.id = mc.instanceid
LEFT JOIN UserCourseAccess uca ON uca.userid = u.id AND uca.courseid = c.id
WHERE mc.contextlevel = 50
  AND u.username NOT IN ('12345678')
  AND c.id IN ('$cursos_str')
  AND mr.id IN ('3','5','9','11','16','17')
GROUP BY u.id, u.username, u.firstname, u.lastname, u.email, c.id, c.fullname, mr.name, uca.access_days
ORDER BY c.fullname, u.lastname, u.firstname";

    // Ejecutar consulta
    $stmt = $db->query($sql);
    $resultados = $stmt->fetchAll(PDO::FETCH_ASSOC);
    
    if (count($resultados) === 0) {
        throw new Exception("No se encontraron resultados para el período especificado ($fecha_inicio_simple a $fecha_fin_simple)");
    }

    // Separar resultados por rol (usando el nombre del rol ahora)
    $profesores = array_filter($resultados, function($fila) {
        return strpos($fila['rol'], 'Profesor') !== false;
    });
    
    $estudiantes = array_filter($resultados, function($fila) {
        return strpos($fila['rol'], 'Estudiante') !== false;
    });
    
    $asesores = array_filter($resultados, function($fila) {
        return strpos($fila['rol'], 'Asesor') !== false;
    });

    $general = $resultados; // Todos los registros

    function generarArchivoProduccion($datos, $nombreArchivo, $tituloHoja) {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setTitle($tituloHoja);
        
        // Encabezados
        $headers = ['Código', 'Rol', 'Nombre', 'Apellidos', 'Correo', 'Curso', 'Total Ingresos'];
        $sheet->fromArray($headers, null, 'A1');
        
        // Datos
        $row = 2;
        foreach ($datos as $item) {
            // Asegurar que 'total_ingresos' sea un número (incluso si viene como NULL o vacío)
            $totalIngresos = isset($item['total_ingresos']) ? (int)$item['total_ingresos'] : 0;
            
            // Escribir los datos normalmente
            $sheet->setCellValue('A' . $row, $item['codigo']);
            $sheet->setCellValue('B' . $row, $item['rol']);
            $sheet->setCellValue('C' . $row, $item['nombre']);
            $sheet->setCellValue('D' . $row, $item['apellidos']);
            $sheet->setCellValue('E' . $row, $item['correo']);
            $sheet->setCellValue('F' . $row, $item['curso']);
            
            // Forzar el valor numérico en la celda de Total Ingresos
            $sheet->setCellValueExplicit(
                'G' . $row,
                $totalIngresos,
                \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC
            );
            
            $row++;
        }
        
        // Opcional: Formatear columna "Total Ingresos" como número
        $sheet->getStyle('G2:G' . ($row - 1))
              ->getNumberFormat()
              ->setFormatCode('0'); // Formato numérico sin decimales
        
        (new Xlsx($spreadsheet))->save($nombreArchivo);
    }

    // Generar archivos
    generarArchivoProduccion($profesores, $temp_profesores, 'Profesores');
    generarArchivoProduccion($estudiantes, $temp_estudiantes, 'Estudiantes');
    generarArchivoProduccion($asesores, $temp_asesores, 'Asesores');
    generarArchivoProduccion($general, $temp_general, 'General');

    // Crear ZIP
    $zip = new ZipArchive();
    if ($zip->open($zip_file, ZipArchive::CREATE) !== TRUE) {
        throw new Exception("No se pudo crear el archivo ZIP");
    }
    $zip->addFile($temp_profesores, basename($temp_profesores));
    $zip->addFile($temp_estudiantes, basename($temp_estudiantes));
    $zip->addFile($temp_asesores, basename($temp_asesores));
    $zip->addFile($temp_general, basename($temp_general));
    if (!$zip->close()) {
        throw new Exception("Error al cerrar el archivo ZIP");
    }

    // Configurar y enviar correo con PHPMailer
    $mail = new PHPMailer(true);
    $mail->setFrom($remitente, 'Reporte Moodle');
    foreach ($correo_destino as $correo) {
        $mail->addAddress($correo);
    }
    
    $fecha_inicio_formatted = $fecha_inicio_obj->format('d/m/Y');
    $fecha_fin_formatted = $lunes_pasado->format('d/m/Y');
    
    $mail->Subject = 'Reporte de Ingresos Acumulado - ' . $fecha_para_nombre;
    $mail->Body = "Cordial Saludo,<br><br>
                 Adjunto el Reporte de Ingresos Acumulado de las asignaturas de pregrado.<br><br>
                 <strong>Período del reporte:</strong> $fecha_inicio_formatted a $fecha_fin_formatted<br>
                 <strong>Total registros:</strong> " . count($resultados) . "<br>
                 <strong>Desglose:</strong><br>
                 - Profesores: " . count($profesores) . "<br>
                 - Estudiantes: " . count($estudiantes) . "<br>
                 - General: " . count($general) . "<br>
                 - Asesores: " . count($asesores)
                 ;
    $mail->isHTML(true);
    
    if (file_exists($zip_file)) {
        $mail->addAttachment($zip_file, "reporte_ingresos_acumulado_asignaturas_pregrado_{$fecha_para_nombre}.zip");
    }
    
    $mail->send();

    // Eliminar archivos temporales
    @unlink($temp_profesores);
    @unlink($temp_estudiantes);
    @unlink($temp_asesores);
    @unlink($temp_general);
    @unlink($zip_file);

    // Enviar notificación de éxito
    mail($correo_notificacion, 'Reporte Exitoso', 
        "El reporte fue generado correctamente.\n" .
        "Período: $fecha_inicio_formatted a $fecha_fin_formatted\n" .
        "Total registros: " . count($resultados) . "\n" .
        "Profesores: " . count($profesores) . "\n" .
        "Estudiantes: " . count($estudiantes) . "\n" .
        "General: " . count($general) . "\n" .
        "Asesores: " . count($asesores));

} catch (Exception $e) {
    // Enviar notificación de error
    mail($correo_notificacion, 'Error en Reporte', 
        "Error: " . $e->getMessage() . "\n" .
        "Período: $fecha_inicio_simple a $fecha_fin_simple\n" .
        "Hora: " . date('Y-m-d H:i:s'));
    exit("Error: " . $e->getMessage());
}