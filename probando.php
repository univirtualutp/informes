<?php
// Configuración inicial
date_default_timezone_set('America/Bogota');

// Autoload de Composer
require __DIR__ . '/vendor/autoload.php';
require_once ' db_moodle_config.php';

use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\Exception;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// =============================================
// CONFIGURACIÓN
// =============================================
/*$db_config = [
    'host' => 'localhost',
    'name' => 'moodle',
    'user' => 'moodle',
    'pass' => 'M00dl3',
    'port' => '5432'
];*/

$correo_destino = ['soporteunivirtual@utp.edu.co','univirtual-utp@utp.edu.co'];
$correo_notificacion = 'soporteunivirtual@utp.edu.co';
$remitente = 'noreply-univirtual@utp.edu.co';

$cursos = [494, 415, 507, 481, 508, 482, 509, 485, 526, 510, 511, 486, 490, 416, 503, 504, 527, 417, 496, 497, 418, 498, 419, 475, 421, 420, 422, 423, 512, 513, 515, 488, 489, 424, 516, 517, 491, 518, 492, 519, 493, 520, 425, 476, 426, 505, 506, 479, 521, 428, 430, 522, 495, 499, 431, 453, 500, 523, 434, 524, 435, 436, 437, 438, 440, 502, 439, 452, 525, 442];
$cursos_str = implode(",", $cursos);

// =============================================
// CÁLCULO DE FECHAS DINÁMICAS (fecha fija: 03 de febrero de 2025)
// =============================================
$hoy = new DateTime('now', new DateTimeZone('America/Bogota'));
$fecha_inicio = new DateTime('2025-02-03 00:00:00', new DateTimeZone('America/Bogota'));
$fecha_fin = clone $hoy;

$fecha_inicio_str = $fecha_inicio->format('Y-m-d 00:00:00');
$fecha_fin_str = $fecha_fin->format('Y-m-d 23:59:59');
$fecha_inicio_simple = $fecha_inicio->format('Y-m-d');
$fecha_fin_simple = $fecha_fin->format('Y-m-d');
$fecha_para_nombre = $hoy->format('Y-m-d');

// Nombres de archivos
$temp_profesores = 'profesores_' . $fecha_para_nombre . '.xlsx';
$temp_estudiantes = 'estudiantes_' . $fecha_para_nombre . '.xlsx';
$temp_asesores = 'asesores_' . $fecha_para_nombre . '.xlsx';
$temp_general = 'general_' . $fecha_para_nombre . '.xlsx';
$zip_file = 'reporte_asignaturas_pregrado_' . $fecha_para_nombre . '.zip';

try {
    // Conexión a la base de datos
    $dsn = "pgsql:host={$db_config['host']};port={$db_config['port']};dbname={$db_config['name']}";
    $db = new PDO($dsn, $db_config['user'], $db_config['pass']);
    $db->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
    $db->exec("SET TIME ZONE 'America/Bogota'");

    // NUEVA CONSULTA SQL PARA MENSAJERÍA CON ROL INCLUIDO
    $sql = "
SELECT DISTINCT
    u.username as codigo,
    u.firstname as nombre,
    u.lastname as apellido,
    mr.name as rol,
    u.email as correo,
    c.id AS id_curso,
    c.fullname AS curso,
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
        AND ra2.roleid IN (3,5,9,11,16,17)
        AND m.timecreated >= EXTRACT(EPOCH FROM TO_TIMESTAMP('$fecha_inicio_str', 'YYYY-MM-DD HH24:MI:SS'))
        AND m.timecreated <= EXTRACT(EPOCH FROM TO_TIMESTAMP('$fecha_fin_str', 'YYYY-MM-DD HH24:MI:SS'))
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
        AND u.username <> '12345678'
        AND ra2.roleid IN (3,5,9,11,16,17)
        AND m.timecreated >= EXTRACT(EPOCH FROM TO_TIMESTAMP('$fecha_inicio_str', 'YYYY-MM-DD HH24:MI:SS'))
        AND m.timecreated <= EXTRACT(EPOCH FROM TO_TIMESTAMP('$fecha_fin_str', 'YYYY-MM-DD HH24:MI:SS'))
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
        AND u.username <> '12345678'
        AND ra2.roleid IN (3,5,9,11,16,17)
        AND mua.id IS NULL
        AND m.timecreated >= EXTRACT(EPOCH FROM TO_TIMESTAMP('$fecha_inicio_str', 'YYYY-MM-DD HH24:MI:SS'))
        AND m.timecreated <= EXTRACT(EPOCH FROM TO_TIMESTAMP('$fecha_fin_str', 'YYYY-MM-DD HH24:MI:SS'))
    ), 0) AS mensajes_sin_leer,
    
    CASE 
        WHEN EXISTS (
            SELECT 1
            FROM mdl_messages m
            JOIN mdl_message_conversations mc ON m.conversationid = mc.id
            JOIN mdl_message_conversation_members mcm ON mc.id = mcm.conversationid
            JOIN mdl_user u2 ON mcm.userid = u2.id AND u2.id != u.id
            JOIN mdl_role_assignments ra2 ON u2.id = ra2.userid
            JOIN mdl_context ctx2 ON ra2.contextid = ctx2.id AND ctx2.contextlevel = 50
            WHERE m.useridfrom = u.id
            AND u.username <> '12345678'
            AND ctx2.instanceid = c.id
            AND ra2.roleid IN (3,5,9,11,16,17)
            AND m.timecreated >= EXTRACT(EPOCH FROM TO_TIMESTAMP('$fecha_inicio_str', 'YYYY-MM-DD HH24:MI:SS'))
            AND m.timecreated <= EXTRACT(EPOCH FROM TO_TIMESTAMP('$fecha_fin_str', 'YYYY-MM-DD HH24:MI:SS'))
        ) THEN TO_CHAR(TO_TIMESTAMP((
            SELECT MAX(m.timecreated)
            FROM mdl_messages m
            JOIN mdl_message_conversations mc ON m.conversationid = mc.id
            JOIN mdl_message_conversation_members mcm ON mc.id = mcm.conversationid
            JOIN mdl_user u2 ON mcm.userid = u2.id AND u2.id != u.id AND u2.username <> '12345678'
            JOIN mdl_role_assignments ra2 ON u2.id = ra2.userid
            JOIN mdl_context ctx2 ON ra2.contextid = ctx2.id AND ctx2.contextlevel = 50
            WHERE m.useridfrom = u.id
            AND u.username <> '12345678'
            AND ctx2.instanceid = c.id
            AND ra2.roleid IN (3,5,9,11,16,17)
            AND m.timecreated >= EXTRACT(EPOCH FROM TO_TIMESTAMP('$fecha_inicio_str', 'YYYY-MM-DD HH24:MI:SS'))
            AND m.timecreated <= EXTRACT(EPOCH FROM TO_TIMESTAMP('$fecha_fin_str', 'YYYY-MM-DD HH24:MI:SS'))
        )), 'YYYY-MM-DD HH24:MI:SS')
        ELSE 'No ha enviado mensajes'
    END AS ultimo_mensaje_enviado,
    
    CASE 
        WHEN EXISTS (
            SELECT 1
            FROM mdl_messages m
            JOIN mdl_message_conversations mc ON m.conversationid = mc.id
            JOIN mdl_message_conversation_members mcm ON mc.id = mcm.conversationid
            JOIN mdl_user u2 ON m.useridfrom = u2.id AND u2.id != u.id
            JOIN mdl_role_assignments ra2 ON u2.id = ra2.userid
            JOIN mdl_context ctx2 ON ra2.contextid = ctx2.id AND ctx2.contextlevel = 50
            WHERE mcm.userid = u.id
            AND ctx2.instanceid = c.id
            AND u.username <> '12345678'
            AND ra2.roleid IN (3,5,9,11,16,17)
            AND m.timecreated >= EXTRACT(EPOCH FROM TO_TIMESTAMP('$fecha_inicio_str', 'YYYY-MM-DD HH24:MI:SS'))
            AND m.timecreated <= EXTRACT(EPOCH FROM TO_TIMESTAMP('$fecha_fin_str', 'YYYY-MM-DD HH24:MI:SS'))
        ) THEN TO_CHAR(TO_TIMESTAMP((
            SELECT MAX(m.timecreated)
            FROM mdl_messages m
            JOIN mdl_message_conversations mc ON m.conversationid = mc.id
            JOIN mdl_message_conversation_members mcm ON mc.id = mcm.conversationid
            JOIN mdl_user u2 ON m.useridfrom = u2.id AND u2.id != u.id
            JOIN mdl_role_assignments ra2 ON u2.id = ra2.userid
            JOIN mdl_context ctx2 ON ra2.contextid = ctx2.id AND ctx2.contextlevel = 50
            WHERE mcm.userid = u.id
            AND ctx2.instanceid = c.id
            AND u.username <> '12345678'
            AND ra2.roleid IN (3,5,9,11,16,17)
            AND m.timecreated >= EXTRACT(EPOCH FROM TO_TIMESTAMP('$fecha_inicio_str', 'YYYY-MM-DD HH24:MI:SS'))
            AND m.timecreated <= EXTRACT(EPOCH FROM TO_TIMESTAMP('$fecha_fin_str', 'YYYY-MM-DD HH24:MI:SS'))
        )), 'YYYY-MM-DD HH24:MI:SS')
        ELSE 'No ha recibido mensajes'
    END AS ultimo_mensaje_recibido,
    
    CASE 
        WHEN EXISTS (
            SELECT 1
            FROM mdl_messages m
            JOIN mdl_message_conversations mc ON m.conversationid = mc.id
            JOIN mdl_message_conversation_members mcm ON mc.id = mcm.conversationid
            JOIN mdl_user u2 ON m.useridfrom = u2.id AND u2.id != u.id
            JOIN mdl_role_assignments ra2 ON u2.id = ra2.userid
            JOIN mdl_context ctx2 ON ra2.contextid = ctx2.id AND ctx2.contextlevel = 50
            JOIN mdl_message_user_actions mua ON m.id = mua.messageid AND mua.userid = u.id AND mua.action = 1
            WHERE mcm.userid = u.id
            AND ctx2.instanceid = c.id
            AND u.username <> '12345678'
            AND ra2.roleid IN (3,5,9,11,16,17)
            AND m.timecreated >= EXTRACT(EPOCH FROM TO_TIMESTAMP('$fecha_inicio_str', 'YYYY-MM-DD HH24:MI:SS'))
            AND m.timecreated <= EXTRACT(EPOCH FROM TO_TIMESTAMP('$fecha_fin_str', 'YYYY-MM-DD HH24:MI:SS'))
        ) THEN TO_CHAR(TO_TIMESTAMP((
            SELECT MAX(m.timecreated)
            FROM mdl_messages m
            JOIN mdl_message_conversations mc ON m.conversationid = mc.id
            JOIN mdl_message_conversation_members mcm ON mc.id = mcm.conversationid
            JOIN mdl_user u2 ON m.useridfrom = u2.id AND u2.id != u.id
            JOIN mdl_role_assignments ra2 ON u2.id = ra2.userid
            JOIN mdl_context ctx2 ON ra2.contextid = ctx2.id AND ctx2.contextlevel = 50
            JOIN mdl_message_user_actions mua ON m.id = mua.messageid AND mua.userid = u.id AND mua.action = 1
            WHERE mcm.userid = u.id
            AND u.username <> '12345678'
            AND ctx2.instanceid = c.id
            AND ra2.roleid IN (3,5,9,11,16,17)
            AND m.timecreated >= EXTRACT(EPOCH FROM TO_TIMESTAMP('$fecha_inicio_str', 'YYYY-MM-DD HH24:MI:SS'))
            AND m.timecreated <= EXTRACT(EPOCH FROM TO_TIMESTAMP('$fecha_fin_str', 'YYYY-MM-DD HH24:MI:SS'))
        )), 'YYYY-MM-DD HH24:MI:SS')
        ELSE 'No ha leído mensajes'
    END AS ultimo_mensaje_leido,
    
    CASE 
        WHEN EXISTS (
            SELECT 1
            FROM mdl_messages m
            JOIN mdl_message_conversations mc ON m.conversationid = mc.id
            JOIN mdl_message_conversation_members mcm ON mc.id = mcm.conversationid
            JOIN mdl_user u2 ON mcm.userid = u2.id AND u2.id != u.id
            JOIN mdl_role_assignments ra2 ON u2.id = ra2.userid
            JOIN mdl_context ctx2 ON ra2.contextid = ctx2.id AND ctx2.contextlevel = 50
            WHERE m.useridfrom = u.id
            AND u.username <> '12345678'
            AND u2.username <> '12345678'
            AND ctx2.instanceid = c.id
            AND ra2.roleid IN (3,5,9,11,16,17)
            AND m.timecreated >= EXTRACT(EPOCH FROM TO_TIMESTAMP('$fecha_inicio_str', 'YYYY-MM-DD HH24:MI:SS'))
            AND m.timecreated <= EXTRACT(EPOCH FROM TO_TIMESTAMP('$fecha_fin_str', 'YYYY-MM-DD HH24:MI:SS'))
        ) THEN (
            SELECT CONCAT(u2.firstname, ' ', u2.lastname)
            FROM mdl_messages m
            JOIN mdl_message_conversations mc ON m.conversationid = mc.id
            JOIN mdl_message_conversation_members mcm ON mc.id = mcm.conversationid
            JOIN mdl_user u2 ON mcm.userid = u2.id AND u2.id != u.id AND u2.username <> '12345678'
            JOIN mdl_role_assignments ra2 ON u2.id = ra2.userid
            JOIN mdl_context ctx2 ON ra2.contextid = ctx2.id AND ctx2.contextlevel = 50
            WHERE m.useridfrom = u.id
            AND u.username <> '12345678'
            AND u2.username <> '12345678'
            AND ctx2.instanceid = c.id
            AND ra2.roleid IN (3,5,9,11,16,17)
            AND m.timecreated >= EXTRACT(EPOCH FROM TO_TIMESTAMP('$fecha_inicio_str', 'YYYY-MM-DD HH24:MI:SS'))
            AND m.timecreated <= EXTRACT(EPOCH FROM TO_TIMESTAMP('$fecha_fin_str', 'YYYY-MM-DD HH24:MI:SS'))
            AND m.timecreated = (
                SELECT MAX(m2.timecreated)
                FROM mdl_messages m2
                JOIN mdl_message_conversations mc2 ON m2.conversationid = mc2.id
                JOIN mdl_message_conversation_members mcm2 ON mc2.id = mcm2.conversationid
                JOIN mdl_user u3 ON mcm2.userid = u3.id AND u3.id != u.id AND u3.username <> '12345678'
                JOIN mdl_role_assignments ra3 ON u3.id = ra3.userid
                JOIN mdl_context ctx3 ON ra3.contextid = ctx3.id AND ctx3.contextlevel = 50
                WHERE m2.useridfrom = u.id
                AND u.username <> '12345678'
                AND u3.username <> '12345678'
                AND ctx3.instanceid = c.id
                AND ra3.roleid IN (3,5,9,11,16,17)
                AND m2.timecreated >= EXTRACT(EPOCH FROM TO_TIMESTAMP('$fecha_inicio_str', 'YYYY-MM-DD HH24:MI:SS'))
                AND m2.timecreated <= EXTRACT(EPOCH FROM TO_TIMESTAMP('$fecha_fin_str', 'YYYY-MM-DD HH24:MI:SS'))
            )
            LIMIT 1
        )
        ELSE 'No ha enviado mensajes'
    END AS ultimo_mensaje_enviado_a,
    
    CASE 
        WHEN EXISTS (
            SELECT 1
            FROM mdl_messages m
            JOIN mdl_message_conversations mc ON m.conversationid = mc.id
            JOIN mdl_message_conversation_members mcm ON mc.id = mcm.conversationid
            JOIN mdl_user u2 ON m.useridfrom = u2.id AND u2.id != u.id
            JOIN mdl_role_assignments ra2 ON u2.id = ra2.userid
            JOIN mdl_context ctx2 ON ra2.contextid = ctx2.id AND ctx2.contextlevel = 50
            JOIN mdl_message_user_actions mua ON m.id = mua.messageid AND mua.userid = u.id AND mua.action = 1
            WHERE mcm.userid = u.id
            AND u.username <> '12345678'
            AND ctx2.instanceid = c.id
            AND ra2.roleid IN (3,5,9,11,16,17)
            AND m.timecreated >= EXTRACT(EPOCH FROM TO_TIMESTAMP('$fecha_inicio_str', 'YYYY-MM-DD HH24:MI:SS'))
            AND m.timecreated <= EXTRACT(EPOCH FROM TO_TIMESTAMP('$fecha_fin_str', 'YYYY-MM-DD HH24:MI:SS'))
        ) THEN (
            SELECT CONCAT(u2.firstname, ' ', u2.lastname)
            FROM mdl_messages m
            JOIN mdl_message_conversations mc ON m.conversationid = mc.id
            JOIN mdl_message_conversation_members mcm ON mc.id = mcm.conversationid
            JOIN mdl_user u2 ON m.useridfrom = u2.id AND u2.id != u.id
            JOIN mdl_role_assignments ra2 ON u2.id = ra2.userid
            JOIN mdl_context ctx2 ON ra2.contextid = ctx2.id AND ctx2.contextlevel = 50
            JOIN mdl_message_user_actions mua ON m.id = mua.messageid AND mua.userid = u.id AND mua.action = 1
            WHERE mcm.userid = u.id
            AND u.username <> '12345678'
            AND ctx2.instanceid = c.id
            AND ra2.roleid IN (3,5,9,11,16,17)
            AND m.timecreated >= EXTRACT(EPOCH FROM TO_TIMESTAMP('$fecha_inicio_str', 'YYYY-MM-DD HH24:MI:SS'))
            AND m.timecreated <= EXTRACT(EPOCH FROM TO_TIMESTAMP('$fecha_fin_str', 'YYYY-MM-DD HH24:MI:SS'))
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
                AND u.username <> '12345678'
                AND ctx3.instanceid = c.id
                AND ra3.roleid IN (3,5,9,11,16,17)
                AND m2.timecreated >= EXTRACT(EPOCH FROM TO_TIMESTAMP('$fecha_inicio_str', 'YYYY-MM-DD HH24:MI:SS'))
                AND m2.timecreated <= EXTRACT(EPOCH FROM TO_TIMESTAMP('$fecha_fin_str', 'YYYY-MM-DD HH24:MI:SS'))
            )
            LIMIT 1
        )
        ELSE 'No ha leído mensajes'
    END AS ultimo_mensaje_leido_de
FROM mdl_user u
JOIN mdl_role_assignments ra ON u.id = ra.userid
JOIN mdl_role mr ON ra.roleid = mr.id
JOIN mdl_context ctx ON ra.contextid = ctx.id AND ctx.contextlevel = 50
JOIN mdl_course c ON ctx.instanceid = c.id
WHERE ctx.instanceid in ($cursos_str)
AND ra.roleid in (3,5,9,11,16,17)
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
    
    // ENCABEZADOS ACTUALIZADOS CON ROL EN LA POSICIÓN CORRECTA
    $headers = [
        'Código', 
        'Nombre', 
        'Apellido', 
        'Rol',
        'Correo', 
        'ID Curso', 
        'Curso', 
        'Mensajes Enviados', 
        'Mensajes Recibidos', 
        'Mensajes Sin Leer',
        'Último Mensaje Enviado',
        'Último Mensaje Recibido',
        'Último Mensaje Leído',
        'Último Mensaje Enviado A',
        'Último Mensaje Leído De'
    ];
    
    $sheet->fromArray($headers, null, 'A1');
    
    // Datos
    $row = 2;
    foreach ($datos as $item) {
        $sheet->setCellValue('A' . $row, $item['codigo']);
        $sheet->setCellValue('B' . $row, $item['nombre']);
        $sheet->setCellValue('C' . $row, $item['apellido']);
        $sheet->setCellValue('D' . $row, $item['rol']);
        $sheet->setCellValue('E' . $row, $item['correo']);
        $sheet->setCellValue('F' . $row, $item['id_curso']);
        $sheet->setCellValue('G' . $row, $item['curso']);
        
        // Forzar valores numéricos para los contadores
        $sheet->setCellValueExplicit('H' . $row, $item['mensajes_enviados'], \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC);
        $sheet->setCellValueExplicit('I' . $row, $item['mensajes_recibidos'], \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC);
        $sheet->setCellValueExplicit('J' . $row, $item['mensajes_sin_leer'], \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC);
        
        // Resto de campos
        $sheet->setCellValue('K' . $row, $item['ultimo_mensaje_enviado']);
        $sheet->setCellValue('L' . $row, $item['ultimo_mensaje_recibido']);
        $sheet->setCellValue('M' . $row, $item['ultimo_mensaje_leido']);
        $sheet->setCellValue('N' . $row, $item['ultimo_mensaje_enviado_a']);
        $sheet->setCellValue('O' . $row, $item['ultimo_mensaje_leido_de']);
        
        $row++;
    }
    
    // Formatear columnas numéricas
    $sheet->getStyle('H2:J' . ($row - 1))
          ->getNumberFormat()
          ->setFormatCode('0'); // Formato numérico sin decimales
    
    // Autoajustar el ancho de las columnas
    foreach (range('A', 'O') as $col) {
        $sheet->getColumnDimension($col)->setAutoSize(true);
    }
    
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
    $mail->setFrom($remitente, 'Reporte Moodle mensajes acumulado');
    foreach ($correo_destino as $correo) {
        $mail->addAddress($correo);
    }
    
    $fecha_inicio_formatted = $fecha_inicio->format('d/m/Y');
    $fecha_fin_formatted = $fecha_fin->format('d/m/Y');
    
    $mail->Subject = 'Reporte de Mensajes - Desde 03/02/2025 hasta ' . $fecha_fin_formatted;
    $mail->Body = "Cordial Saludo,<br><br>
                 Adjunto el Reporte de Mensajes acumulado de las asignaturas de pregrado.<br><br>
                 <strong>Período del reporte:</strong> 03/02/2025 a $fecha_fin_formatted<br>
                 <strong>Total registros:</strong> " . count($resultados) . "<br>
                 <strong>Desglose:</strong><br>
                 - Profesores: " . count($profesores) . "<br>
                 - Estudiantes: " . count($estudiantes) . "<br>
                 - General: " . count($general) . "<br>
                 - Asesores: " . count($asesores)
                 ;
    $mail->isHTML(true);
    
    if (file_exists($zip_file)) {
        $mail->addAttachment($zip_file, "reporte_mensajes_asignaturas_pregrado_{$fecha_para_nombre}.zip");
    }
    
    $mail->send();

    // Eliminar archivos temporales
    @unlink($temp_profesores);
    @unlink($temp_estudiantes);
    @unlink($temp_asesores);
    @unlink($temp_general);
    @unlink($zip_file);

    // Enviar notificación de éxito
    mail($correo_notificacion, 'Reporte Exitoso Mensajes asignaturas pregrado', 
        "El reporte de mensajería fue generado correctamente.\n" .
        "Período: 03/02/2025 a $fecha_fin_formatted\n" .
        "Total registros: " . count($resultados) . "\n" .
        "Profesores: " . count($profesores) . "\n" .
        "Estudiantes: " . count($estudiantes) . "\n" .
        "General: " . count($general) . "\n" .
        "Asesores: " . count($asesores));

} catch (Exception $e) {
    // Enviar notificación de error
    mail($correo_notificacion, 'Error en Reporte', 
        "Error: " . $e->getMessage() . "\n" .
        "Período: 03/02/2025 a $fecha_fin_simple\n" .
        "Hora: " . date('Y-m-d H:i:s'));
    exit("Error: " . $e->getMessage());
}
