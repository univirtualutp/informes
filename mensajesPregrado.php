<?php
// Configuración inicial del entorno
date_default_timezone_set('America/Bogota'); // Establece la zona horaria a Bogotá

// Carga del autoload de Composer para usar PHPMailer y PhpSpreadsheet
require __DIR__ . '/vendor/autoload.php';

// Inclusión del archivo de configuración de la base de datos
require_once __DIR__ . '/db_moodle_config.php';

// Importación de clases necesarias
use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\Exception;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// =============================================
// CONFIGURACIÓN DE CORREOS Y CURSOS
// =============================================

// Correos de destino para el reporte
$correo_destino = ['soporteunivirtual@utp.edu.co'];
// Correo para notificaciones de éxito o error
$correo_notificacion = 'soporteunivirtual@utp.edu.co';
// Correo remitente para el envío del reporte
$remitente = 'noreply-univirtual@utp.edu.co';

// Lista de IDs de cursos a consultar
$cursos = [
    786, 583, 804, 596, 820, 790, 821, 819, 789, 580, 799, 616, 617, 797, 798,
    842, 844, 621, 805, 584, 581, 802, 619, 627, 800, 801, 618, 588, 815, 589,
    590, 594, 595, 787, 788, 605, 607, 793, 573, 791, 606, 792, 608, 609, 795,
    611, 794, 610, 586, 814, 623, 622, 624, 733, 574, 604, 796, 576, 612, 577,
    614, 830, 615, 783, 579, 784, 785, 591, 587, 810, 625, 582, 803, 620, 592,
    816, 817, 593, 740, 822, 823, 824, 825, 741, 826, 827, 828
];
$cursos_str = implode(",", $cursos); // Convierte la lista de cursos en una cadena para la consulta SQL

// =============================================
// CÁLCULO DE FECHAS DINÁMICAS
// =============================================

// Obtiene la fecha actual en la zona horaria de Bogotá
$hoy = new DateTime('now', new DateTimeZone('America/Bogota'));
// Calcula el lunes pasado
$lunes_pasado = clone $hoy;
$lunes_pasado->modify('last monday');
// Calcula el martes de la semana anterior
$martes_anterior = clone $lunes_pasado;
$martes_anterior->modify('-6 days');

// Formatea las fechas para la consulta y el nombre de los archivos
$fecha_inicio = $martes_anterior->format('Y-m-d 00:00:00');
$fecha_fin = $lunes_pasado->format('Y-m-d 23:59:59');
$fecha_inicio_simple = $martes_anterior->format('Y-m-d');
$fecha_fin_simple = $lunes_pasado->format('Y-m-d');
$fecha_para_nombre = $hoy->format('Y-m-d');

// Nombres de los archivos temporales y del archivo ZIP
$temp_profesores = 'profesores_' . $fecha_para_nombre . '.xlsx';
$temp_estudiantes = 'estudiantes_' . $fecha_para_nombre . '.xlsx';
$temp_asesores = 'asesores_' . $fecha_para_nombre . '.xlsx';
$temp_general = 'general_' . $fecha_para_nombre . '.xlsx';
$zip_file = 'reporte_asignaturas_pregrado_' . $fecha_para_nombre . '.zip';

try {
    // =============================================
    // CONEXIÓN A LA BASE DE DATOS
    // =============================================
    
    // Establece la conexión a PostgreSQL usando las constantes definidas en db_moodle_config.php
    $dsn = "pgsql:host=" . DB_HOST . ";port=" . DB_PORT . ";dbname=" . DB_NAME;
    $db = new PDO($dsn, DB_USER, DB_PASS);
    $db->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
    $db->exec("SET TIME ZONE 'America/Bogota'");

    // =============================================
    // CONSULTA SQL PARA OBTENER DATOS DE MENSAJERÍA
    // =============================================
    
    $sql = "
    SELECT 
        u.username,
        u.firstname AS nombres,
        u.lastname AS apellidos,
        u.email AS correo,
        r.name AS rol,
        c.id AS id_curso,
        c.fullname AS nombre_curso,
        -- Estadísticas de mensajes (solo enviados)
        COUNT(DISTINCT CASE WHEN mu.role = 1 THEN m.id END) AS total_mensajes_enviados,
        SUM(CASE WHEN mu.unread = 1 THEN 1 ELSE 0 END) AS mensajes_sin_leer,
        SUM(CASE WHEN mu.unread = 0 THEN 1 ELSE 0 END) AS mensajes_leidos,
        -- Información del último mensaje (con fecha formateada)
        to_char(to_timestamp(MAX(m.time)), 'YYYY-MM-DD HH24:MI:SS') AS fecha_ultimo_mensaje,
        -- Asunto del último mensaje recibido
        (SELECT m2.subject 
         FROM mdl_local_mail_messages m2
         JOIN mdl_local_mail_message_users mu2 ON m2.id = mu2.messageid
         WHERE mu2.userid = u.id AND m2.time = (SELECT MAX(m3.time) 
                                              FROM mdl_local_mail_messages m3
                                              JOIN mdl_local_mail_message_users mu3 ON m3.id = mu3.messageid
                                              WHERE mu3.userid = u.id)
         LIMIT 1) AS asunto_ultimo_mensaje,
        -- Remitente del último mensaje recibido
        (SELECT CONCAT(u2.firstname, ' ', u2.lastname) 
         FROM mdl_local_mail_messages m2
         JOIN mdl_local_mail_message_refs mr ON m2.id = mr.messageid
         JOIN mdl_local_mail_messages m_ref ON mr.reference = m_ref.id
         JOIN mdl_local_mail_message_users mu_ref ON m_ref.id = mu_ref.messageid
         JOIN mdl_user u2 ON mu_ref.userid = u2.id
         WHERE mu_ref.userid != u.id AND m2.time = (SELECT MAX(m3.time) 
                                                 FROM mdl_local_mail_messages m3
                                                 JOIN mdl_local_mail_message_users mu3 ON m3.id = mu3.messageid
                                                 WHERE mu3.userid = u.id)
         LIMIT 1) AS ultimo_remitente,
        -- Destinatario del último mensaje enviado (con fecha formateada)
        (SELECT CONCAT(
                u2.firstname, ' ', u2.lastname, ' - ', 
                to_char(to_timestamp(m2.time), 'YYYY-MM-DD HH24:MI:SS')
            ) 
         FROM mdl_local_mail_messages m2
         JOIN mdl_local_mail_message_users mu bolsinger ON m2.id = mu_sender.messageid
         JOIN mdl_local_mail_message_users mu_recipient ON m2.id = mu_recipient.messageid
         JOIN mdl_user u2 ON mu_recipient.userid = u2.id
         WHERE mu_sender.userid = u.id 
         AND mu_sender.role = 1
         AND mu_recipient.userid != u.id
         AND m2.time = (SELECT MAX(m3.time) 
                       FROM mdl_local_mail_messages m3
                       JOIN mdl_local_mail_message_users mu3 ON m3.id = mu3.messageid
                       WHERE mu3.userid = u.id AND mu3.role = 1)
         LIMIT 1) AS ultimo_destinatario_con_fecha
    FROM mdl_user u
    JOIN mdl_role_assignments ra ON u.id = ra.userid
    JOIN mdl_role r ON ra.roleid = r.id
    LEFT JOIN mdl_context ctx ON ra.contextid = ctx.id
    LEFT JOIN mdl_course c ON ctx.instanceid = c.id AND ctx.contextlevel = 50
    LEFT JOIN mdl_local_mail_message_users mu ON u.id = mu.userid
    LEFT JOIN mdl_local_mail_messages m ON mu.messageid = m.id
    WHERE r.id IN (3, 5, 9, 16, 17)
    AND c.id IN ($cursos_str)
    AND u.username <> '12345678'
    GROUP BY u.id, u.username, u.firstname, u.lastname, u.email, r.name, c.id, c.fullname
    ORDER BY c.id, u.lastname, u.firstname";

    // Ejecutar la consulta y obtener los resultados
    $stmt = $db->query($sql);
    $resultados = $stmt->fetchAll(PDO::FETCH_ASSOC);

    // Verificar si se obtuvieron resultados
    if (count($resultados) === 0) {
        throw new Exception("No se encontraron resultados para el período especificado ($fecha_inicio_simple a $fecha_fin_simple)");
    }

    // =============================================
    // SEPARACIÓN DE RESULTADOS POR ROL
    // =============================================
    
    // Filtrar resultados para profesores
    $profesores = array_filter($resultados, function($fila) {
        return strpos($fila['rol'], 'Profesor') !== false;
    });
    
    // Filtrar resultados para estudiantes
    $estudiantes = array_filter($resultados, function($fila) {
        return strpos($fila['rol'], 'Estudiante') !== false;
    });
    
    // Filtrar resultados para asesores
    $asesores = array_filter($resultados, function($fila) {
        return strpos($fila['rol'], 'Asesor') !== false;
    });

    // Todos los registros (general)
    $general = $resultados;

    // =============================================
    // FUNCIÓN PARA GENERAR ARCHIVOS EXCEL
    // =============================================
    
    function generarArchivoProduccion($datos, $nombreArchivo, $tituloHoja) {
        // Crear una nueva hoja de cálculo
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setTitle($tituloHoja);

        // Definir los encabezados según los campos de la nueva consulta
        $headers = [
            'Código',
            'Nombres',
            'Apellidos',
            'Correo',
            'Rol',
            'ID Curso',
            'Nombre Curso',
            'Total Mensajes Enviados',
            'Mensajes Sin Leer',
            'Mensajes Leídos',
            'Fecha Último Mensaje',
            'Asunto Último Mensaje',
            'Último Remitente',
            'Último Destinatario con Fecha'
        ];

        // Establecer los encabezados en la primera fila
        $sheet->fromArray($headers, null, 'A1');

        // Llenar los datos
        $row = 2;
        foreach ($datos as $item) {
            $sheet->setCellValue('A' . $row, $item['username']);
            $sheet->setCellValue('B' . $row, $item['nombres']);
            $sheet->setCellValue('C' . $row, $item['apellidos']);
            $sheet->setCellValue('D' . $row, $item['correo']);
            $sheet->setCellValue('E' . $row, $item['rol']);
            $sheet->setCellValue('F' . $row, $item['id_curso']);
            $sheet->setCellValue('G' . $row, $item['nombre_curso']);
            // Forzar valores numéricos para los contadores
            $sheet->setCellValueExplicit('H' . $row, $item['total_mensajes_enviados'], \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC);
            $sheet->setCellValueExplicit('I' . $row, $item['mensajes_sin_leer'], \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC);
            $sheet->setCellValueExplicit('J' . $row, $item['mensajes_leidos'], \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC);
            $sheet->setCellValue('K' . $row, $item['fecha_ultimo_mensaje']);
            $sheet->setCellValue('L' . $row, $item['asunto_ultimo_mensaje']);
            $sheet->setCellValue('M' . $row, $item['ultimo_remitente']);
            $sheet->setCellValue('N' . $row, $item['ultimo_destinatario_con_fecha']);
            $row++;
        }

        // Formatear columnas numéricas
        $sheet->getStyle('H2:J' . ($row - 1))
              ->getNumberFormat()
              ->setFormatCode('0'); // Formato numérico sin decimales

        // Autoajustar el ancho de las columnas
        foreach (range('A', 'N') as $col) {
            $sheet->getColumnDimension($col)->setAutoSize(true);
        }

        // Guardar el archivo Excel
        (new Xlsx($spreadsheet))->save($nombreArchivo);
    }

    // Generar los archivos Excel para cada grupo
    generarArchivoProduccion($profesores, $temp_profesores, 'Profesores');
    generarArchivoProduccion($estudiantes, $temp_estudiantes, 'Estudiantes');
    generarArchivoProduccion($asesores, $temp_asesores, 'Asesores');
    generarArchivoProduccion($general, $temp_general, 'General');

    // =============================================
    // CREACIÓN DEL ARCHIVO ZIP
    // =============================================
    
    // Crear un archivo ZIP con los archivos generados
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

    // =============================================
    // ENVÍO DEL CORREO CON EL REPORTE
    // =============================================
    
    // Configurar PHPMailer para enviar el correo
    $mail = new PHPMailer(true);
    $mail->setFrom($remitente, 'Reporte Moodle mensajes pregrado');
    foreach ($correo_destino as $correo) {
        $mail->addAddress($correo);
    }

    // Formatear fechas para el cuerpo del correo
    $fecha_inicio_formatted = $martes_anterior->format('d/m/Y');
    $fecha_fin_formatted = $lunes_pasado->format('d/m/Y');

    // Configurar el asunto y cuerpo del correo
    $mail->Subject = 'Reporte de Mensajes Semanal de asignaturas de pregrado - ' . $fecha_para_nombre;
    $mail->Body = "Cordial Saludo,<br><br>
                 Adjunto el Reporte de Mensajería Semanal de asignaturas de pregrado.<br><br>
                 <strong>Período del reporte:</strong> $fecha_inicio_formatted a $fecha_fin_formatted<br>
                 <strong>Total registros:</strong> " . count($resultados) . "<br>
                 <strong>Desglose:</strong><br>
                 - Profesores: " . count($profesores) . "<br>
                 - Estudiantes: " . count($estudiantes) . "<br>
                 - General: " . count($general) . "<br>
                 - Asesores: " . count($asesores);
    $mail->isHTML(true);

    // Adjuntar el archivo ZIP al correo
    if (file_exists($zip_file)) {
        $mail->addAttachment($zip_file, "reporte_mensajeria_pregrado_{$fecha_para_nombre}.zip");
    }

    // Enviar el correo
    $mail->send();

    // =============================================
    // LIMPIEZA DE ARCHIVOS TEMPORALES
    // =============================================
    
    // Eliminar archivos temporales
    @unlink($temp_profesores);
    @unlink($temp_estudiantes);
    @unlink($temp_asesores);
    @unlink($temp_general);
    @unlink($zip_file);

    // =============================================
    // NOTIFICACIÓN DE ÉXITO
    // =============================================
    
    // Enviar correo de notificación de éxito
    mail($correo_notificacion, 'Reporte Exitoso de mensajes pregrado',
        "El reporte de mensajería fue generado correctamente.\n" .
        "Período: $fecha_inicio_formatted a $fecha_fin_formatted\n" .
        "Total registros: " . count($resultados) . "\n" .
        "Profesores: " . count($profesores) . "\n" .
        "Estudiantes: " . count($estudiantes) . "\n" .
        "General: " . count($general) . "\n" .
        "Asesores: " . count($asesores));

} catch (Exception $e) {
    // =============================================
    // MANEJO DE ERRORES
    // =============================================
    
    // Enviar correo de notificación de error
    mail($correo_notificacion, 'Error en Reporte',
        "Error: " . $e->getMessage() . "\n" .
        "Período: $fecha_inicio_simple a $fecha_fin_simple\n" .
        "Hora: " . date('Y-m-d H:i:s'));
    exit("Error: " . $e->getMessage());
}
?>