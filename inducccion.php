<?php
// Cargar dependencias (PHPMailer y PhpSpreadsheet para Excel)
require 'vendor/autoload.php'; 
use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\Exception;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// Configurar zona horaria para Bogotá, Colombia
date_default_timezone_set('America/Bogota');

// Configuración de la base de datos y correos
$host = 'localhost';
$dbname = 'moodle';
$user = 'moodle';
$pass = 'M00dl3';
$correo_destino = ['soporteunivirtual@utp.edu.co','m.monroy@utp.edu.co','beatriz.gutierrez@utp.edu.co'];
$correo_notificacion = 'soporteunivirtual@utp.edu.co';

try {
    // Conexión a la base de datos PostgreSQL
    $pdo = new PDO("pgsql:host=$host;dbname=$dbname", $user, $pass);
    $pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);

    // Consulta SQL con las modificaciones solicitadas
    $sql = "SELECT 
      u.username AS Documento,
      mr.name AS rol,
      u.firstname AS Nombres,
      u.lastname AS Apellidos,
      u.email AS correo,
      c.fullname AS curso,

      CASE 
        WHEN to_char(to_timestamp(MAX(mue.timecreated)), 'YYYY-MM-DD') IN ('1969-12-31', '1970-01-01') THEN 'NUNCA'
        ELSE to_char(to_timestamp(MAX(mue.timecreated)), 'YYYY-MM-DD')
      END AS fecha_matricula,

      CASE 
        WHEN to_char(to_timestamp(MAX(u.firstaccess)), 'YYYY-MM-DD') IN ('1969-12-31', '1970-01-01') THEN 'NUNCA'
        ELSE to_char(to_timestamp(MAX(u.firstaccess)), 'YYYY-MM-DD')
      END AS fecha_primer_acceso,

      CASE 
        WHEN to_char(to_timestamp(MAX(u.lastaccess)), 'YYYY-MM-DD') IN ('1969-12-31', '1970-01-01') THEN 'NUNCA'
        ELSE to_char(to_timestamp(MAX(u.lastaccess)), 'YYYY-MM-DD')
      END AS fecha_ultimo_acceso,

      CASE 
        WHEN ROUND(AVG(a.Nota), 2) >= 4.0 THEN MAX(a.fecha_realizacion)
        ELSE NULL
      END AS fecha_culminación,

      CASE 
        WHEN ROUND(AVG(a.Nota), 2) >= 4.0 THEN 'APROBÓ'
        ELSE 'No Aprobado'
      END AS promedio_nota,

      ROUND(AVG(a.Nota), 2) AS nota_numerica

    FROM mdl_user u
    JOIN mdl_user_enrolments mue ON mue.userid = u.id
    JOIN mdl_role_assignments mra ON mra.userid = u.id
    JOIN mdl_role mr ON mr.id = mra.roleid
    JOIN mdl_context mc ON mc.id = mra.contextid
    JOIN mdl_course c ON c.id = mc.instanceid
    LEFT JOIN 
    (
        SELECT 
          mgm.userid,
          mg.name as Grupo
        FROM mdl_groups mg 
        JOIN mdl_groups_members mgm ON mgm.groupid = mg.id
        WHERE mg.courseid = '199'
        AND EXISTS (
            SELECT 1 FROM mdl_role_assignments mra 
            JOIN mdl_role mr ON mr.id = mra.roleid
            WHERE mra.userid = mgm.userid AND mr.id = '5'
        )
    ) g ON g.userid = u.id
    LEFT JOIN 
    (
        SELECT 
          mqg.userid AS userid,
          c.id AS course_id,
          to_char(to_timestamp(MAX(mqg.timemodified)), 'YYYY-MM-DD') AS fecha_realizacion,
          COALESCE(AVG(gr.finalgrade), 0) AS Nota
        FROM mdl_quiz_grades mqg
        JOIN mdl_quiz q ON q.id = mqg.quiz AND q.id IN (166, 177)
        JOIN mdl_course c ON c.id = q.course
        JOIN mdl_grade_items i ON i.iteminstance = q.id AND i.itemmodule = 'quiz'
        JOIN mdl_grade_grades gr ON gr.itemid = i.id AND gr.userid = mqg.userid
        GROUP BY mqg.userid, c.id
    ) a ON a.userid = u.id AND a.course_id = c.id
    WHERE mc.contextlevel = 50
    AND c.id = '199'
    AND mr.id = '5'
    GROUP BY u.username, mr.name, u.firstname, u.lastname, u.email, c.fullname
    ORDER BY u.lastname, u.firstname;";

    // Preparar y ejecutar la consulta
    $stmt = $pdo->prepare($sql);
    $stmt->execute();
    $resultados = $stmt->fetchAll(PDO::FETCH_ASSOC);

    // Función para limpiar y formatear los datos
    function cleanData($value, $key) {
        if ($value === null || $value === '') {
            return '';
        }
        // Manejar fechas que indican "nunca"
        if (in_array($value, ['1969-12-31', '1970-01-01'])) {
            return 'NUNCA';
        }
        // Para la columna fecha_culminación, mantener vacío si no aprobó
        if ($key === 'fecha_culminación' && $value === NULL) {
            return '';
        }
        return $value;
    }

    // Aplicar la función de limpieza a cada fila de resultados
    $resultados = array_map(function($row) {
        $cleanedRow = [];
        foreach ($row as $key => $value) {
            $cleanedRow[$key] = cleanData($value, $key);
        }
        return $cleanedRow;
    }, $resultados);

    // Crear un nuevo archivo Excel
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    $sheet->setTitle('Reporte Estudiantes');

    // Escribir encabezados
    if (!empty($resultados)) {
        // Definir los encabezados que queremos mostrar (excluyendo nota_numerica)
        $headersToShow = array_filter(array_keys($resultados[0]), function($key) {
            return $key !== 'nota_numerica';
        });
        
        $sheet->fromArray($headersToShow, NULL, 'A1');
        
        // Escribir datos (excluyendo la columna nota_numerica)
        $row = 2;
        foreach ($resultados as $data) {
            $filteredData = array_filter($data, function($key) {
                return $key !== 'nota_numerica';
            }, ARRAY_FILTER_USE_KEY);
            $sheet->fromArray($filteredData, NULL, 'A'.$row);
            $row++;
        }
        
        // Autoajustar el ancho de las columnas
        foreach (range('A', $sheet->getHighestColumn()) as $col) {
            $sheet->getColumnDimension($col)->setAutoSize(true);
        }
        
        // Congelar la primera fila (encabezados)
        $sheet->freezePane('A2');
        
        // Formato condicional para resaltar aprobados/no aprobados
        $lastRow = count($resultados) + 1;
        $conditionalStyles = [
            'APROBÓ' => [
                'fill' => [
                    'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                    'color' => ['argb' => 'FFC6EFCE']
                ]
            ],
            'No Aprobado' => [
                'fill' => [
                    'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                    'color' => ['argb' => 'FFFFC7CE']
                ]
            ]
        ];
        
        $columnLetter = 'F'; // Columna del promedio_nota (ajustar según posición real)
        foreach ($conditionalStyles as $text => $style) {
            $conditional = new \PhpOffice\PhpSpreadsheet\Style\Conditional();
            $conditional->setConditionType(\PhpOffice\PhpSpreadsheet\Style\Conditional::CONDITION_CELLIS)
                       ->setOperatorType(\PhpOffice\PhpSpreadsheet\Style\Conditional::OPERATOR_EQUAL)
                       ->addCondition('"'.$text.'"');
            $conditional->getStyle()->applyFromArray($style);
            
            $conditionalStylesArray[] = $conditional;
        }
        
        $sheet->getStyle($columnLetter.'2:'.$columnLetter.$lastRow)
              ->setConditionalStyles($conditionalStylesArray);
    }

    // Guardar el archivo Excel temporalmente
    $fecha_para_nombre = date('Ymd');
    $nombre_archivo = "reporte_estudiantes_curso_induccion_{$fecha_para_nombre}.xlsx";
    $temp_dir = sys_get_temp_dir();
    $temp_file = $temp_dir . '/' . uniqid('reporte_', true) . '.xlsx';
    
    $writer = new Xlsx($spreadsheet);
    $writer->save($temp_file);

    // Configurar y enviar correo con PHPMailer
    $mail = new PHPMailer(true);
    $mail->setFrom('noreply-univirtual@utp.edu.co', 'Reporte Induccion');
    foreach ($correo_destino as $correo) {
        $mail->addAddress($correo);
    }
    $mail->Subject = 'Reporte de Estudiantes - Curso de indccion administrativa - ' . date('d/m/Y');
    $mail->Body = 'Cordial Saludo, Adjunto el Reporte de estudiantes que estan matriculados en el curso de indccion administrativa.' . "\n\n" . 'En caso de mayor información por favor comunicarse a <strong>soporteunivirtual@utp.edu.co</strong>.';
    
    // Adjuntar el archivo Excel
    if (file_exists($temp_file)) {
        $mail->addAttachment($temp_file, $nombre_archivo);
    }
    
    $mail->send();

    // Eliminar archivo temporal
    if (file_exists($temp_file)) {
        unlink($temp_file);
    }

    // Enviar notificación de éxito
    mail($correo_notificacion, 'Estado Reporte', 'El reporte de estudiantes del curso de indccion administrativa fue enviado correctamente.');
} catch (Exception $e) {
    // Enviar notificación de error
    mail($correo_notificacion, 'Error Reporte del curso indccion administrativa', 'Error: ' . $e->getMessage());
    echo "Error: " . $e->getMessage(); // Mostrar el error en la consola
}
?>
