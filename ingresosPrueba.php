<?php
// Configuración inicial
date_default_timezone_set('America/Bogota');

// Autoload de Composer
require __DIR__ . '/vendor/autoload.php';

use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\Exception;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// =============================================
// CONFIGURACIÓN
// =============================================
$db_config = [
    'host' => 'localhost',
    'name' => 'moodle',
    'user' => 'moodle',
    'pass' => 'M00dl3',
    'port' => '5432'
];

$correo_destino = ['daniel.pardo@utp.edu.co'];
$correo_notificacion = 'daniel.pardo@utp.edu.co';
$remitente = 'noreply-univirtual@utp.edu.co';

// Configuración de diagnóstico
$usuario_problematico = '1007462692'; // Reemplazar con el username real del usuario con discrepancia
$curso_problematico = '502';      // Reemplazar con el ID del curso problemático

// FECHAS ESTÁTICAS (comentar para volver a dinámicas)
$fecha_inicio = '2025-05-06 00:00:00';
$fecha_fin = '2025-05-12 23:59:59';
$fecha_inicio_simple = '2025-05-06';
$fecha_fin_simple = '2025-05-12';
$fecha_para_nombre = date('Y-m-d'); // Fecha actual para el nombre del archivo

/*
// FECHAS DINÁMICAS (descomentar para usar)
$hoy = new DateTime('now', new DateTimeZone('America/Bogota'));
$lunes_pasado = clone $hoy;
$lunes_pasado->modify('last monday');
$martes_anterior = clone $lunes_pasado;
$martes_anterior->modify('-6 days');

$fecha_inicio = $martes_anterior->format('Y-m-d') . ' 00:00:00';
$fecha_fin = $lunes_pasado->format('Y-m-d') . ' 23:59:59';
$fecha_inicio_simple = $martes_anterior->format('Y-m-d');
$fecha_fin_simple = $lunes_pasado->format('Y-m-d');
$fecha_para_nombre = $hoy->format('Y-m-d');
*/
// =============================================

// Nombre de archivo de diagnóstico
$diagnostic_file = 'diagnostico_usuario_' . $usuario_problematico . '_' . $fecha_para_nombre . '.xlsx';

try {
    // Conexión a la base de datos
    $dsn = "pgsql:host={$db_config['host']};port={$db_config['port']};dbname={$db_config['name']}";
    $db = new PDO($dsn, $db_config['user'], $db_config['pass']);
    $db->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
    $db->exec("SET TIME ZONE 'America/Bogota'");

    // CONSULTA PRINCIPAL DE DIAGNÓSTICO
    $sql = "
    SELECT 
    u.username,
    u.firstname || ' ' || u.lastname AS nombre_completo,
    c.id AS courseid,
    c.fullname AS course_name,
    mr.name AS role_name,
  -- Lista completa de fechas de acceso
  array_agg(DISTINCT to_char(to_timestamp(mlsl.timecreated) AT TIME ZONE 'America/Bogota', 'YYYY-MM-DD')) AS access_dates,
  -- Conteo verificable
  COUNT(DISTINCT DATE(to_timestamp(mlsl.timecreated) AT TIME ZONE 'America/Bogota')) AS access_days,
  -- Detalle completo para diagnóstico
  string_agg(
    to_char(to_timestamp(mlsl.timecreated) AT TIME ZONE 'America/Bogota', 'YYYY-MM-DD HH24:MI:SS') || ' | ' || 
    mlsl.action || ' | ' || mlsl.target, 
    '\n'
    ) AS access_details
  FROM mdl_user u
  JOIN mdl_role_assignments mra ON mra.userid = u.id
  JOIN mdl_role mr ON mra.roleid = mr.id
  JOIN mdl_context mc ON mc.id = mra.contextid
  JOIN mdl_course c ON c.id = mc.instanceid
  LEFT JOIN mdl_logstore_standard_log mlsl ON (
      mlsl.courseid = c.id AND
      mlsl.userid = u.id AND
      mlsl.action = 'viewed' AND
      mlsl.target IN ('course', 'course_module') AND
      to_timestamp(mlsl.timecreated) >= to_timestamp(extract(epoch FROM timestamp '$fecha_inicio'::timestamp)) AND
      to_timestamp(mlsl.timecreated) <= to_timestamp(extract(epoch FROM timestamp '$fecha_fin'::timestamp))
  )
  WHERE mc.contextlevel = 50
  AND u.username = '$usuario_problematico'
  AND c.id = '$curso_problematico'
  AND mr.id IN ('3','5','9','11','16','17')
  GROUP BY u.username, u.firstname, u.lastname, c.id, c.fullname, mr.name";

    // CONSULTA DE VERIFICACIÓN (LOGS CRUDOS)
  $sql_verificacion = "
  SELECT 
  to_char(to_timestamp(timecreated) AT TIME ZONE 'America/Bogota', 'YYYY-MM-DD HH24:MI:SS') AS fecha_hora,
  action,
  target,
  objecttable,
  objectid
  FROM mdl_logstore_standard_log
  WHERE userid = (SELECT id FROM mdl_user WHERE username = '$usuario_problematico')
  AND courseid = $curso_problematico
  AND action = 'viewed'
  AND target IN ('course', 'course_module')
  AND timecreated >= extract(epoch FROM timestamp '$fecha_inicio'::timestamp)
  AND timecreated <= extract(epoch FROM timestamp '$fecha_fin'::timestamp)
  ORDER BY timecreated";

    // Ejecutar consultas
  $stmt = $db->query($sql);
  $resultados = $stmt->fetchAll(PDO::FETCH_ASSOC);

  $stmt_verificacion = $db->query($sql_verificacion);
  $logs_crudos = $stmt_verificacion->fetchAll(PDO::FETCH_ASSOC);

  if (count($resultados) === 0) {
    throw new Exception("No se encontraron resultados para el usuario $usuario_problematico en el curso $curso_problematico");
}

    // Crear archivo Excel con diagnóstico completo
$spreadsheet = new Spreadsheet();

    // Hoja 1: Resumen de diagnóstico
$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle('Resumen Diagnóstico');

    // Encabezados
$headers = [
'Username', 'Nombre', 'Curso ID', 'Nombre Curso', 'Rol',
'Días de Acceso', 'Fechas de Acceso (Únicas)', 'Detalles Completos'
];
$sheet->fromArray($headers, null, 'A1');

    // Datos
$row = 2;
foreach ($resultados as $registro) {
    $sheet->setCellValue('A'.$row, $registro['username']);
    $sheet->setCellValue('B'.$row, $registro['nombre_completo']);
    $sheet->setCellValue('C'.$row, $registro['courseid']);
    $sheet->setCellValue('D'.$row, $registro['course_name']);
    $sheet->setCellValue('E'.$row, $registro['role_name']);
    $sheet->setCellValue('F'.$row, $registro['access_days']);
    $sheet->setCellValue('G'.$row, str_replace(['{','}','"'], '', $registro['access_dates']));
    $sheet->setCellValue('H'.$row, $registro['access_details']);

    $sheet->getRowDimension($row)->setRowHeight(60);
    $row++;
}

    // Hoja 2: Logs crudos
$spreadsheet->createSheet();
$sheet2 = $spreadsheet->getSheet(1);
$sheet2->setTitle('Logs Crudos');

$headers_logs = ['Fecha/Hora', 'Action', 'Target', 'Object Table', 'Object ID'];
$sheet2->fromArray($headers_logs, null, 'A1');

$row_log = 2;
foreach ($logs_crudos as $log) {
    $sheet2->fromArray($log, null, "A{$row_log}");
    $row_log++;
}

    // Ajustes de formato
foreach(range('A','H') as $col) {
    $sheet->getColumnDimension($col)->setAutoSize(true);
}
$sheet->getStyle('G:H')->getAlignment()->setWrapText(true);

    // Guardar archivo
(new Xlsx($spreadsheet))->save($diagnostic_file);

    // Configurar y enviar correo con PHPMailer
$mail = new PHPMailer(true);
$mail->setFrom($remitente, 'Reporte Moodle - Diagnóstico');
foreach ($correo_destino as $correo) {
    $mail->addAddress($correo);
}

    // Calcular días únicos de logs crudos
$fechas_unicas_logs = [];
foreach ($logs_crudos as $log) {
    $fecha = substr($log['fecha_hora'], 0, 10);
    if (!in_array($fecha, $fechas_unicas_logs)) {
        $fechas_unicas_logs[] = $fecha;
    }
}
$total_dias_logs = count($fechas_unicas_logs);

$mail->Subject = 'Diagnóstico de Usuario - ' . $usuario_problematico . ' - ' . $fecha_para_nombre;
$mail->Body = "Cordial Saludo,<br><br>
Adjunto el reporte de diagnóstico detallado para el usuario <strong>$usuario_problematico</strong> en el curso <strong>$curso_problematico</strong>.<br><br>
<strong>Resumen de resultados:</strong><br>
- Días de acceso contados por el sistema: <strong>{$resultados[0]['access_days']}</strong><br>
- Días con accesos según logs crudos: <strong>$total_dias_logs</strong><br>
- Total de eventos registrados: <strong>" . count($logs_crudos) . "</strong><br><br>";

    // Análisis de discrepancia
if ($resultados[0]['access_days'] != $total_dias_logs) {
    $mail->Body .= "<span style='color:red;font-weight:bold'>¡SE DETECTÓ DISCREPANCIA EN LOS CONTEO!</span><br><br>";
    $mail->Body .= "<strong>Posibles causas:</strong><br>
    1. Eventos cerca del cambio de día (verificar horas)<br>
    2. Diferencia en el manejo de huso horario<br>
    3. Eventos que no cumplen todos los criterios de filtrado<br><br>";
}

$mail->Body .= "<strong>Instrucciones:</strong><br>
1. Revise la hoja 'Resumen Diagnóstico' para ver el conteo y fechas<br>
2. Verifique la hoja 'Logs Crudos' para ver todos los eventos registrados<br>
3. Compare las fechas/horas con su consulta manual<br><br>
Período analizado: $fecha_inicio_simple a $fecha_fin_simple";

$mail->isHTML(true);
$mail->addAttachment($diagnostic_file);
$mail->send();

    // Eliminar archivo temporal
if (file_exists($diagnostic_file)) {
    unlink($diagnostic_file);
}

    // Enviar notificación de éxito
mail($correo_notificacion, 'Diagnóstico Completado', 
"El diagnóstico para el usuario $usuario_problematico fue generado y enviado.\n" .
"Días contados: {$resultados[0]['access_days']}\n" .
"Eventos encontrados: " . count($logs_crudos));

} catch (Exception $e) {
    // Enviar notificación de error
    mail($correo_notificacion, 'Error en Diagnóstico', 'Error: ' . $e->getMessage());
    exit("Error: " . $e->getMessage());
}
