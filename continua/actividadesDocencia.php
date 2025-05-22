<?php
require '/root/scripts/informes/vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\Exception;

// Establecer la zona horaria explícitamente (ajusta según tu ubicación)
date_default_timezone_set('America/Bogota');

$host = 'localhost';
$dbname = 'moodle';
$user = 'moodle';
$pass = 'M00dl3';
$correo_destino = ['soporteunivirtual@utp.edu.co', 'univirtual-utp@utp.edu.co'];
$correo_notificacion = 'soporteunivirtual@utp.edu.co';

$hoy = new DateTime();
$lunes = clone $hoy;
$lunes->modify('last monday');
$fecha_inicio = new DateTime(date('Y') . '-03-25 00:00:00');
$fecha_fin = clone $lunes;
$fecha_fin->setTime(23, 59, 59);

try {
    $pdo = new PDO("pgsql:host=$host;dbname=$dbname", $user, $pass);
    $pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);

    // Lista de grupos de cursos con nombres asociados
    $grupos_cursos = [
    	'M1' => ['540'],
        'M2' => ['541'],
        'M3' => ['542'],
        'M4' => ['543'],
        'M5' => ['544'],
        'M6' => ['545']
    ];

    // Parámetros para las consultas de tablas temporales
    $params = [
        ':fecha_inicio' => $fecha_inicio->format('Y-m-d H:i:s'),
        ':fecha_fin' => $fecha_fin->format('Y-m-d H:i:s')
    ];

    // Directorio temporal para almacenar los archivos Excel
    $temp_dir = sys_get_temp_dir() . '/reportes_moodle_' . uniqid();
    if (!mkdir($temp_dir)) {
        throw new Exception("No se pudo crear el directorio temporal: $temp_dir");
    }

    // Array para almacenar las rutas de los archivos generados
    $archivos_excel = [];

    // Iterar sobre cada grupo de cursos
    foreach ($grupos_cursos as $nombre_grupo => $cursos) {
        // Obtener nombres y fechas de actividades del curso de referencia (primer curso del grupo)
        $curso_referencia = $cursos[0];

        // Foros
        $forum_sql = "SELECT DISTINCT id AS forum_number, name AS forumname, duedate AS fecha_apertura, cutoffdate AS fecha_cierre 
                      FROM mdl_forum WHERE course = :curso AND name != 'Avisos' ORDER BY name";
        $stmt = $pdo->prepare($forum_sql);
        $stmt->execute([':curso' => $curso_referencia]);
        $forum_results = $stmt->fetchAll(PDO::FETCH_ASSOC);
        $forum_names = [];
        $forum_dates = [];
        $max_forums = count($forum_results);
        foreach ($forum_results as $i => $row) {
            $forum_names[$i + 1] = $row['forumname'];
            $forum_dates[$i + 1] = [
                'apertura' => $row['fecha_apertura'] ? date('Y-m-d H:i:s', (int)$row['fecha_apertura']) : '',
                'cierre' => $row['fecha_cierre'] ? date('Y-m-d H:i:s', (int)$row['fecha_cierre']) : ''
            ];
        }

        // Lecciones
        $lesson_sql = "SELECT DISTINCT id AS lesson_number, name AS lessonname, available AS fecha_apertura, deadline AS fecha_cierre 
                       FROM mdl_lesson WHERE course = :curso ORDER BY name";
        $stmt = $pdo->prepare($lesson_sql);
        $stmt->execute([':curso' => $curso_referencia]);
        $lesson_results = $stmt->fetchAll(PDO::FETCH_ASSOC);
        $lesson_names = [];
        $lesson_dates = [];
        $max_lessons = count($lesson_results);
        foreach ($lesson_results as $i => $row) {
            $lesson_names[$i + 1] = $row['lessonname'];
            $lesson_dates[$i + 1] = [
                'apertura' => $row['fecha_apertura'] ? date('Y-m-d H:i:s', (int)$row['fecha_apertura']) : '',
                'cierre' => $row['fecha_cierre'] ? date('Y-m-d H:i:s', (int)$row['fecha_cierre']) : ''
            ];
        }

        // Tareas
        $assign_sql = "SELECT DISTINCT id AS assign_number, name AS assignname, allowsubmissionsfromdate AS fecha_apertura, duedate AS fecha_cierre 
                       FROM mdl_assign WHERE course = :curso ORDER BY name";
        $stmt = $pdo->prepare($assign_sql);
        $stmt->execute([':curso' => $curso_referencia]);
        $assign_results = $stmt->fetchAll(PDO::FETCH_ASSOC);
        $assign_names = [];
        $assign_dates = [];
        $max_assigns = count($assign_results);
        foreach ($assign_results as $i => $row) {
            $assign_names[$i + 1] = $row['assignname'];
            $assign_dates[$i + 1] = [
                'apertura' => $row['fecha_apertura'] ? date('Y-m-d H:i:s', (int)$row['fecha_apertura']) : '',
                'cierre' => $row['fecha_cierre'] ? date('Y-m-d H:i:s', (int)$row['fecha_cierre']) : ''
            ];
        }

        // Quizzes
        $quiz_sql = "SELECT DISTINCT id AS quiz_number, name AS quizname, timeopen AS fecha_apertura, timeclose AS fecha_cierre 
                     FROM mdl_quiz WHERE course = :curso ORDER BY name";
        $stmt = $pdo->prepare($quiz_sql);
        $stmt->execute([':curso' => $curso_referencia]);
        $quiz_results = $stmt->fetchAll(PDO::FETCH_ASSOC);
        $quiz_names = [];
        $quiz_dates = [];
        $max_quizzes = count($quiz_results);
        foreach ($quiz_results as $i => $row) {
            $quiz_names[$i + 1] = $row['quizname'];
            $quiz_dates[$i + 1] = [
                'apertura' => $row['fecha_apertura'] ? date('Y-m-d H:i:s', (int)$row['fecha_apertura']) : '',
                'cierre' => $row['fecha_cierre'] ? date('Y-m-d H:i:s', (int)$row['fecha_cierre']) : ''
            ];
        }

        // Glosarios
        $glossary_sql = "SELECT DISTINCT id AS glossary_number, name AS glossaryname, timecreated AS fecha_apertura, timemodified AS fecha_cierre 
                         FROM mdl_glossary WHERE course = :curso ORDER BY name";
        $stmt = $pdo->prepare($glossary_sql);
        $stmt->execute([':curso' => $curso_referencia]);
        $glossary_results = $stmt->fetchAll(PDO::FETCH_ASSOC);
        $glossary_names = [];
        $glossary_dates = [];
        $max_glossaries = count($glossary_results);
        foreach ($glossary_results as $i => $row) {
            $glossary_names[$i + 1] = $row['glossaryname'];
            $glossary_dates[$i + 1] = [
                'apertura' => $row['fecha_apertura'] ? date('Y-m-d H:i:s', (int)$row['fecha_apertura']) : '',
                'cierre' => $row['fecha_cierre'] ? date('Y-m-d H:i:s', (int)$row['fecha_cierre']) : ''
            ];
        }

        // SCORM
        $scorm_sql = "SELECT DISTINCT id AS scorm_number, name AS scormname, timeopen AS fecha_apertura, timeclose AS fecha_cierre 
                      FROM mdl_scorm WHERE course = :curso ORDER BY name";
        $stmt = $pdo->prepare($scorm_sql);
        $stmt->execute([':curso' => $curso_referencia]);
        $scorm_results = $stmt->fetchAll(PDO::FETCH_ASSOC);
        $scorm_names = [];
        $scorm_dates = [];
        $max_scorms = count($scorm_results);
        foreach ($scorm_results as $i => $row) {
            $scorm_names[$i + 1] = $row['scormname'];
            $scorm_dates[$i + 1] = [
                'apertura' => $row['fecha_apertura'] ? date('Y-m-d H:i:s', (int)$row['fecha_apertura']) : '',
                'cierre' => $row['fecha_cierre'] ? date('Y-m-d H:i:s', (int)$row['fecha_cierre']) : ''
            ];
        }

        // Generar encabezados
        $headers = array_fill(0, 9, '');
        $sub_headers = [
            'codigo', 'nombre', 'apellidos', 'rol', 'correo', 'id_curso', 'curso',
            'facultad', 'programa'
        ];

        for ($i = 1; $i <= $max_forums; $i++) {
            $forum_name = $forum_names[$i] ?? "Foro $i";
            $headers[] = $forum_name;
            $headers[] = ''; $headers[] = ''; $headers[] = ''; $headers[] = ''; $headers[] = '';
            $sub_headers[] = 'fecha_apertura';
            $sub_headers[] = 'fecha_cierre';
            $sub_headers[] = 'visitas';
            $sub_headers[] = 'numero_envios';
            $sub_headers[] = 'nota_final';
            $sub_headers[] = 'retroalimentada';
        }

        for ($i = 1; $i <= $max_lessons; $i++) {
            $lesson_name = $lesson_names[$i] ?? "Lección $i";
            $headers[] = $lesson_name;
            $headers[] = ''; $headers[] = ''; $headers[] = ''; $headers[] = ''; $headers[] = '';
            $sub_headers[] = 'fecha_apertura';
            $sub_headers[] = 'fecha_cierre';
            $sub_headers[] = 'visitas';
            $sub_headers[] = 'numero_intentos';
            $sub_headers[] = 'nota_final';
            $sub_headers[] = 'retroalimentada';
        }

        for ($i = 1; $i <= $max_assigns; $i++) {
            $assign_name = $assign_names[$i] ?? "Tarea $i";
            $headers[] = $assign_name;
            $headers[] = ''; $headers[] = ''; $headers[] = ''; $headers[] = ''; $headers[] = '';
            $sub_headers[] = 'fecha_apertura';
            $sub_headers[] = 'fecha_cierre';
            $sub_headers[] = 'visitas';
            $sub_headers[] = 'numero_envios';
            $sub_headers[] = 'nota_final';
            $sub_headers[] = 'retroalimentada';
        }

        for ($i = 1; $i <= $max_quizzes; $i++) {
            $quiz_name = $quiz_names[$i] ?? "Quiz $i";
            $headers[] = $quiz_name;
            $headers[] = ''; $headers[] = ''; $headers[] = ''; $headers[] = ''; $headers[] = '';
            $sub_headers[] = 'fecha_apertura';
            $sub_headers[] = 'fecha_cierre';
            $sub_headers[] = 'visitas';
            $sub_headers[] = 'numero_intentos';
            $sub_headers[] = 'nota_final';
            $sub_headers[] = 'retroalimentada';
        }

        for ($i = 1; $i <= $max_glossaries; $i++) {
            $glossary_name = $glossary_names[$i] ?? "Glosario $i";
            $headers[] = $glossary_name;
            $headers[] = ''; $headers[] = ''; $headers[] = ''; $headers[] = ''; $headers[] = '';
            $sub_headers[] = 'fecha_apertura';
            $sub_headers[] = 'fecha_cierre';
            $sub_headers[] = 'visitas';
            $sub_headers[] = 'numero_envios';
            $sub_headers[] = 'nota_final';
            $sub_headers[] = 'retroalimentada';
        }

        for ($i = 1; $i <= $max_scorms; $i++) {
            $scorm_name = $scorm_names[$i] ?? "SCORM $i";
            $headers[] = $scorm_name;
            $headers[] = ''; $headers[] = ''; $headers[] = ''; $headers[] = ''; $headers[] = '';
            $sub_headers[] = 'fecha_apertura';
            $sub_headers[] = 'fecha_cierre';
            $sub_headers[] = 'visitas';
            $sub_headers[] = 'numero_envios';
            $sub_headers[] = 'nota_final';
            $sub_headers[] = 'retroalimentada';
        }

        // Inicializar spreadsheet para este grupo
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->fromArray($headers, null, 'A1');
        $sheet->fromArray($sub_headers, null, 'A2');

        // Fila inicial para los datos
        $row_number = 3;

        // Procesar cada curso del grupo
        foreach ($cursos as $curso_id) {
            $pdo->exec("DROP TABLE IF EXISTS temp_forum_agg");
            $forum_agg_sql = "
                CREATE TEMP TABLE temp_forum_agg AS
                SELECT 
                    fp.userid,
                    f.course AS courseid,
                    f.name AS forumname,
                    COUNT(DISTINCT fp.id) AS num_posts,
                    COALESCE(gg.finalgrade, 0) AS nota_final,
                    CASE WHEN gg.feedback IS NOT NULL AND gg.feedback != '' THEN 'SÍ' ELSE 'NO' END AS retroalimentada,
                    ROW_NUMBER() OVER (PARTITION BY fp.userid, f.course ORDER BY f.name) AS forum_number
                FROM mdl_forum_posts fp
                JOIN mdl_forum_discussions fd ON fp.discussion = fd.id
                JOIN mdl_forum f ON fd.forum = f.id
                LEFT JOIN mdl_grade_grades gg ON gg.userid = fp.userid 
                    AND gg.itemid IN (SELECT id FROM mdl_grade_items WHERE itemmodule = 'forum' AND iteminstance = f.id)
                WHERE fp.created BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) 
                                    AND EXTRACT(EPOCH FROM :fecha_fin::timestamp)
                  AND f.name != 'Avisos'
                  AND f.course = :curso
                GROUP BY fp.userid, f.course, f.name, gg.finalgrade, gg.feedback";
            $stmt = $pdo->prepare($forum_agg_sql);
            $stmt->execute(array_merge($params, [':curso' => $curso_id]));
            $pdo->exec("CREATE INDEX ON temp_forum_agg (userid, courseid)");

            $pdo->exec("DROP TABLE IF EXISTS temp_lesson_agg");
            $lesson_agg_sql = "
                CREATE TEMP TABLE temp_lesson_agg AS
                SELECT 
                    lg.userid,
                    l.course AS courseid,
                    l.name AS lessonname,
                    COUNT(DISTINCT lg.id) AS num_attempts,
                    COALESCE(gg.finalgrade, 0) AS nota_final,
                    CASE WHEN gg.feedback IS NOT NULL AND gg.feedback != '' THEN 'SÍ' ELSE 'NO' END AS retroalimentada,
                    ROW_NUMBER() OVER (PARTITION BY lg.userid, l.course ORDER BY l.name) AS lesson_number
                FROM mdl_lesson_grades lg
                JOIN mdl_lesson l ON lg.lessonid = l.id
                LEFT JOIN mdl_grade_grades gg ON gg.userid = lg.userid 
                    AND gg.itemid IN (SELECT id FROM mdl_grade_items WHERE itemmodule = 'lesson' AND iteminstance = l.id)
                WHERE lg.completed BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) 
                                      AND EXTRACT(EPOCH FROM :fecha_fin::timestamp)
                  AND l.course = :curso
                GROUP BY lg.userid, l.course, l.name, gg.finalgrade, gg.feedback";
            $stmt = $pdo->prepare($lesson_agg_sql);
            $stmt->execute(array_merge($params, [':curso' => $curso_id]));
            $pdo->exec("CREATE INDEX ON temp_lesson_agg (userid, courseid)");

            $pdo->exec("DROP TABLE IF EXISTS temp_assign_agg");
            $assign_agg_sql = "
                CREATE TEMP TABLE temp_assign_agg AS
                SELECT 
                    sub.userid,
                    a.course AS courseid,
                    a.name AS assignname,
                    COUNT(DISTINCT sub.id) AS num_submissions,
                    COALESCE(gg.finalgrade, 0) AS nota_final,
                    CASE WHEN gg.feedback IS NOT NULL AND gg.feedback != '' THEN 'SÍ' ELSE 'NO' END AS retroalimentada,
                    ROW_NUMBER() OVER (PARTITION BY sub.userid, a.course ORDER BY a.name) AS assign_number
                FROM mdl_assign_submission sub
                JOIN mdl_assign a ON sub.assignment = a.id
                LEFT JOIN mdl_grade_grades gg ON gg.userid = sub.userid 
                    AND gg.itemid IN (SELECT id FROM mdl_grade_items WHERE itemmodule = 'assign' AND iteminstance = a.id)
                WHERE sub.timecreated BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) 
                                         AND EXTRACT(EPOCH FROM :fecha_fin::timestamp)
                  AND a.course = :curso
                GROUP BY sub.userid, a.course, a.name, gg.finalgrade, gg.feedback";
            $stmt = $pdo->prepare($assign_agg_sql);
            $stmt->execute(array_merge($params, [':curso' => $curso_id]));
            $pdo->exec("CREATE INDEX ON temp_assign_agg (userid, courseid)");

            $pdo->exec("DROP TABLE IF EXISTS temp_quiz_agg");
            $quiz_agg_sql = "
                CREATE TEMP TABLE temp_quiz_agg AS
                SELECT 
                    qa.userid,
                    q.course AS courseid,
                    q.name AS quizname,
                    COUNT(DISTINCT qa.id) AS num_attempts,
                    COALESCE(gg.finalgrade, 0) AS nota_final,
                    CASE WHEN gg.feedback IS NOT NULL AND gg.feedback != '' THEN 'SÍ' ELSE 'NO' END AS retroalimentada,
                    ROW_NUMBER() OVER (PARTITION BY qa.userid, q.course ORDER BY q.name) AS quiz_number
                FROM mdl_quiz_attempts qa
                JOIN mdl_quiz q ON qa.quiz = q.id
                LEFT JOIN mdl_grade_grades gg ON gg.userid = qa.userid 
                    AND gg.itemid IN (SELECT id FROM mdl_grade_items WHERE itemmodule = 'quiz' AND iteminstance = q.id)
                WHERE qa.timefinish BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) 
                                       AND EXTRACT(EPOCH FROM :fecha_fin::timestamp)
                  AND q.course = :curso
                GROUP BY qa.userid, q.course, q.name, gg.finalgrade, gg.feedback";
            $stmt = $pdo->prepare($quiz_agg_sql);
            $stmt->execute(array_merge($params, [':curso' => $curso_id]));
            $pdo->exec("CREATE INDEX ON temp_quiz_agg (userid, courseid)");

            $pdo->exec("DROP TABLE IF EXISTS temp_glossary_agg");
            $glossary_agg_sql = "
                CREATE TEMP TABLE temp_glossary_agg AS
                SELECT 
                    ge.userid,
                    g.course AS courseid,
                    g.name AS glossaryname,
                    COUNT(DISTINCT ge.id) AS num_entries,
                    COALESCE(gg.finalgrade, 0) AS nota_final,
                    CASE WHEN gg.feedback IS NOT NULL AND gg.feedback != '' THEN 'SÍ' ELSE 'NO' END AS retroalimentada,
                    ROW_NUMBER() OVER (PARTITION BY ge.userid, g.course ORDER BY g.name) AS glossary_number
                FROM mdl_glossary_entries ge
                JOIN mdl_glossary g ON ge.glossaryid = g.id
                LEFT JOIN mdl_grade_grades gg ON gg.userid = ge.userid 
                    AND gg.itemid IN (SELECT id FROM mdl_grade_items WHERE itemmodule = 'glossary' AND iteminstance = g.id)
                WHERE ge.timecreated BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) 
                                        AND EXTRACT(EPOCH FROM :fecha_fin::timestamp)
                  AND g.course = :curso
                GROUP BY ge.userid, g.course, g.name, gg.finalgrade, gg.feedback";
            $stmt = $pdo->prepare($glossary_agg_sql);
            $stmt->execute(array_merge($params, [':curso' => $curso_id]));
            $pdo->exec("CREATE INDEX ON temp_glossary_agg (userid, courseid)");

            $pdo->exec("DROP TABLE IF EXISTS temp_scorm_agg");
            $scorm_agg_sql = "
                CREATE TEMP TABLE temp_scorm_agg AS
                SELECT 
                    st.userid,
                    s.course AS courseid,
                    s.name AS scormname,
                    COUNT(DISTINCT st.id) AS num_attempts,
                    COALESCE(gg.finalgrade, 0) AS nota_final,
                    CASE WHEN gg.feedback IS NOT NULL AND gg.feedback != '' THEN 'SÍ' ELSE 'NO' END AS retroalimentada,
                    ROW_NUMBER() OVER (PARTITION BY st.userid, s.course ORDER BY s.name) AS scorm_number
                FROM mdl_scorm_scoes_track st
                JOIN mdl_scorm s ON st.scormid = s.id
                LEFT JOIN mdl_grade_grades gg ON gg.userid = st.userid 
                    AND gg.itemid IN (SELECT id FROM mdl_grade_items WHERE itemmodule = 'scorm' AND iteminstance = s.id)
                WHERE st.timemodified BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) 
                                         AND EXTRACT(EPOCH FROM :fecha_fin::timestamp)
                  AND s.course = :curso
                GROUP BY st.userid, s.course, s.name, gg.finalgrade, gg.feedback";
            $stmt = $pdo->prepare($scorm_agg_sql);
            $stmt->execute(array_merge($params, [':curso' => $curso_id]));
            $pdo->exec("CREATE INDEX ON temp_scorm_agg (userid, courseid)");

            $pdo->exec("DROP TABLE IF EXISTS temp_log_agg");
            $log_agg_sql = "
                CREATE TEMP TABLE temp_log_agg AS
                SELECT 
                    log.userid,
                    log.courseid,
                    log.component,
                    log.target,
                    log.objectid,
                    COUNT(*) AS num_views
                FROM mdl_logstore_standard_log log
                WHERE log.action = 'viewed'
                  AND log.timecreated BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) 
                                         AND EXTRACT(EPOCH FROM :fecha_fin::timestamp)
                  AND log.courseid = :curso
                  AND log.component IN ('mod_forum', 'mod_lesson', 'mod_assign', 'mod_quiz', 'mod_glossary', 'mod_scorm')
                GROUP BY log.userid, log.courseid, log.component, log.target, log.objectid";
            $stmt = $pdo->prepare($log_agg_sql);
            $stmt->execute(array_merge($params, [':curso' => $curso_id]));
            $pdo->exec("CREATE INDEX ON temp_log_agg (userid, courseid, component, target, objectid)");

            // Consulta principal para el curso actual
            $sql = "SELECT 
                u.username AS codigo,
                u.firstname AS nombre,
                u.lastname AS apellidos,
                r.name AS rol,
                u.email AS correo,
                c.id AS id_curso,
                c.fullname AS curso,
                u.institution AS facultad, 
                u.department AS programa,";

            for ($i = 1; $i <= $max_forums; $i++) {
                $sql .= "
                MAX(CASE WHEN fi.forum_number = $i THEN '" . ($forum_dates[$i]['apertura'] ?? '') . "' END) AS fecha_apertura_forum_$i,
                MAX(CASE WHEN fi.forum_number = $i THEN '" . ($forum_dates[$i]['cierre'] ?? '') . "' END) AS fecha_cierre_forum_$i,
                MAX(CASE WHEN fi.forum_number = $i THEN GREATEST(COALESCE(flog.num_views, 0), COALESCE(fi.num_posts, 0)) END) AS visitas_forum_$i,
                MAX(CASE WHEN fi.forum_number = $i THEN fi.num_posts END) AS numero_envios_$i,
                MAX(CASE WHEN fi.forum_number = $i THEN fi.nota_final END) AS nota_final_forum_$i,
                MAX(CASE WHEN fi.forum_number = $i THEN fi.retroalimentada END) AS retroalimentada_forum_$i,";
            }

            for ($i = 1; $i <= $max_lessons; $i++) {
                $sql .= "
                MAX(CASE WHEN li.lesson_number = $i THEN '" . ($lesson_dates[$i]['apertura'] ?? '') . "' END) AS fecha_apertura_lesson_$i,
                MAX(CASE WHEN li.lesson_number = $i THEN '" . ($lesson_dates[$i]['cierre'] ?? '') . "' END) AS fecha_cierre_lesson_$i,
                MAX(CASE WHEN li.lesson_number = $i THEN GREATEST(COALESCE(llog.num_views, 0), COALESCE(li.num_attempts, 0)) END) AS visitas_lesson_$i,
                MAX(CASE WHEN li.lesson_number = $i THEN li.num_attempts END) AS numero_intentos_$i,
                MAX(CASE WHEN li.lesson_number = $i THEN li.nota_final END) AS nota_final_lesson_$i,
                MAX(CASE WHEN li.lesson_number = $i THEN li.retroalimentada END) AS retroalimentada_lesson_$i,";
            }

            for ($i = 1; $i <= $max_assigns; $i++) {
                $sql .= "
                MAX(CASE WHEN ai.assign_number = $i THEN '" . ($assign_dates[$i]['apertura'] ?? '') . "' END) AS fecha_apertura_assign_$i,
                MAX(CASE WHEN ai.assign_number = $i THEN '" . ($assign_dates[$i]['cierre'] ?? '') . "' END) AS fecha_cierre_assign_$i,
                MAX(CASE WHEN ai.assign_number = $i THEN GREATEST(COALESCE(alog.num_views, 0), COALESCE(ai.num_submissions, 0)) END) AS visitas_assign_$i,
                MAX(CASE WHEN ai.assign_number = $i THEN ai.num_submissions END) AS numero_envios_assign_$i,
                MAX(CASE WHEN ai.assign_number = $i THEN ai.nota_final END) AS nota_final_assign_$i,
                MAX(CASE WHEN ai.assign_number = $i THEN ai.retroalimentada END) AS retroalimentada_assign_$i,";
            }

            for ($i = 1; $i <= $max_quizzes; $i++) {
                $sql .= "
                MAX(CASE WHEN qi.quiz_number = $i THEN '" . ($quiz_dates[$i]['apertura'] ?? '') . "' END) AS fecha_apertura_quiz_$i,
                MAX(CASE WHEN qi.quiz_number = $i THEN '" . ($quiz_dates[$i]['cierre'] ?? '') . "' END) AS fecha_cierre_quiz_$i,
                MAX(CASE WHEN qi.quiz_number = $i THEN GREATEST(COALESCE(qlog.num_views, 0), COALESCE(qi.num_attempts, 0)) END) AS visitas_quiz_$i,
                MAX(CASE WHEN qi.quiz_number = $i THEN qi.num_attempts END) AS numero_intentos_quiz_$i,
                MAX(CASE WHEN qi.quiz_number = $i THEN qi.nota_final END) AS nota_final_quiz_$i,
                MAX(CASE WHEN qi.quiz_number = $i THEN qi.retroalimentada END) AS retroalimentada_quiz_$i,";
            }

            for ($i = 1; $i <= $max_glossaries; $i++) {
                $sql .= "
                MAX(CASE WHEN gi.glossary_number = $i THEN '" . ($glossary_dates[$i]['apertura'] ?? '') . "' END) AS fecha_apertura_glossary_$i,
                MAX(CASE WHEN gi.glossary_number = $i THEN '" . ($glossary_dates[$i]['cierre'] ?? '') . "' END) AS fecha_cierre_glossary_$i,
                MAX(CASE WHEN gi.glossary_number = $i THEN GREATEST(COALESCE(glog.num_views, 0), COALESCE(gi.num_entries, 0)) END) AS visitas_glossary_$i,
                MAX(CASE WHEN gi.glossary_number = $i THEN gi.num_entries END) AS numero_envios_glossary_$i,
                MAX(CASE WHEN gi.glossary_number = $i THEN gi.nota_final END) AS nota_final_glossary_$i,
                MAX(CASE WHEN gi.glossary_number = $i THEN gi.retroalimentada END) AS retroalimentada_glossary_$i,";
            }

            for ($i = 1; $i <= $max_scorms; $i++) {
                $sql .= "
                MAX(CASE WHEN si.scorm_number = $i THEN '" . ($scorm_dates[$i]['apertura'] ?? '') . "' END) AS fecha_apertura_scorm_$i,
                MAX(CASE WHEN si.scorm_number = $i THEN '" . ($scorm_dates[$i]['cierre'] ?? '') . "' END) AS fecha_cierre_scorm_$i,
                MAX(CASE WHEN si.scorm_number = $i THEN GREATEST(COALESCE(slog.num_views, 0), COALESCE(si.num_attempts, 0)) END) AS visitas_scorm_$i,
                MAX(CASE WHEN si.scorm_number = $i THEN si.num_attempts END) AS numero_envios_scorm_$i,
                MAX(CASE WHEN si.scorm_number = $i THEN si.nota_final END) AS nota_final_scorm_$i,
                MAX(CASE WHEN si.scorm_number = $i THEN si.retroalimentada END) AS retroalimentada_scorm_$i,";
            }

            $sql = rtrim($sql, ',');

            $sql .= "
            FROM mdl_user u
            JOIN mdl_role_assignments ra ON ra.userid = u.id
            JOIN mdl_role r ON ra.roleid = r.id 
            JOIN mdl_context mc ON ra.contextid = mc.id
            JOIN mdl_course c ON mc.instanceid = c.id
            LEFT JOIN temp_forum_agg fi ON fi.userid = u.id AND fi.courseid = c.id
            LEFT JOIN temp_lesson_agg li ON li.userid = u.id AND li.courseid = c.id
            LEFT JOIN temp_assign_agg ai ON ai.userid = u.id AND ai.courseid = c.id
            LEFT JOIN temp_quiz_agg qi ON qi.userid = u.id AND qi.courseid = c.id
            LEFT JOIN temp_glossary_agg gi ON gi.userid = u.id AND gi.courseid = c.id
            LEFT JOIN temp_scorm_agg si ON si.userid = u.id AND si.courseid = c.id
            LEFT JOIN temp_log_agg flog ON flog.userid = u.id AND flog.courseid = c.id AND flog.component = 'mod_forum' AND flog.target IN ('discussion', 'forum')
            LEFT JOIN temp_log_agg llog ON llog.userid = u.id AND llog.courseid = c.id AND llog.component = 'mod_lesson' AND llog.target = 'lesson'
            LEFT JOIN temp_log_agg alog ON alog.userid = u.id AND alog.courseid = c.id AND alog.component = 'mod_assign' AND alog.target = 'assign'
            LEFT JOIN temp_log_agg qlog ON qlog.userid = u.id AND qlog.courseid = c.id AND qlog.component = 'mod_quiz' AND qlog.target = 'quiz'
            LEFT JOIN temp_log_agg glog ON glog.userid = u.id AND glog.courseid = c.id AND glog.component = 'mod_glossary' AND glog.target = 'glossary'
            LEFT JOIN temp_log_agg slog ON slog.userid = u.id AND slog.courseid = c.id AND slog.component = 'mod_scorm' AND slog.target = 'scorm'
            WHERE mc.contextlevel = 50
              AND u.username <> '12345678'
              AND c.id = :curso
              AND r.id IN ('5', '9', '16')
            GROUP BY u.id, u.username, u.firstname, u.lastname, r.name, u.email, c.id, c.fullname
            ORDER BY u.lastname, u.firstname;";

            $stmt = $pdo->prepare($sql);
            $stmt->execute([':curso' => $curso_id]);
            $resultados = $stmt->fetchAll(PDO::FETCH_ASSOC);

            // Añadir datos al spreadsheet
            foreach ($resultados as $row) {
                $row_data = [
                    $row['codigo'],
                    $row['nombre'],
                    $row['apellidos'],
                    $row['rol'],
                    $row['correo'],
                    $row['id_curso'],
                    $row['curso'],
                    $row['facultad'],
                    $row['programa']
                ];

                for ($i = 1; $i <= $max_forums; $i++) {
                    $row_data[] = $row["fecha_apertura_forum_$i"] ?? $forum_dates[$i]['apertura'];
                    $row_data[] = $row["fecha_cierre_forum_$i"] ?? $forum_dates[$i]['cierre'];
                    $row_data[] = $row["visitas_forum_$i"] ?? '';
                    $row_data[] = $row["numero_envios_$i"] ?? '';
                    $row_data[] = $row["nota_final_forum_$i"] ?? '';
                    $row_data[] = $row["retroalimentada_forum_$i"] ?? 'NO';
                }

                for ($i = 1; $i <= $max_lessons; $i++) {
                    $row_data[] = $row["fecha_apertura_lesson_$i"] ?? $lesson_dates[$i]['apertura'];
                    $row_data[] = $row["fecha_cierre_lesson_$i"] ?? $lesson_dates[$i]['cierre'];
                    $row_data[] = $row["visitas_lesson_$i"] ?? '';
                    $row_data[] = $row["numero_intentos_$i"] ?? '';
                    $row_data[] = $row["nota_final_lesson_$i"] ?? '';
                    $row_data[] = $row["retroalimentada_lesson_$i"] ?? 'NO';
                }

                for ($i = 1; $i <= $max_assigns; $i++) {
                    $row_data[] = $row["fecha_apertura_assign_$i"] ?? $assign_dates[$i]['apertura'];
                    $row_data[] = $row["fecha_cierre_assign_$i"] ?? $assign_dates[$i]['cierre'];
                    $row_data[] = $row["visitas_assign_$i"] ?? '';
                    $row_data[] = $row["numero_envios_assign_$i"] ?? '';
                    $row_data[] = $row["nota_final_assign_$i"] ?? '';
                    $row_data[] = $row["retroalimentada_assign_$i"] ?? 'NO';
                }

                for ($i = 1; $i <= $max_quizzes; $i++) {
                    $row_data[] = $row["fecha_apertura_quiz_$i"] ?? $quiz_dates[$i]['apertura'];
                    $row_data[] = $row["fecha_cierre_quiz_$i"] ?? $quiz_dates[$i]['cierre'];
                    $row_data[] = $row["visitas_quiz_$i"] ?? '';
                    $row_data[] = $row["numero_intentos_quiz_$i"] ?? '';
                    $row_data[] = $row["nota_final_quiz_$i"] ?? '';
                    $row_data[] = $row["retroalimentada_quiz_$i"] ?? 'NO';
                }

                for ($i = 1; $i <= $max_glossaries; $i++) {
                    $row_data[] = $row["fecha_apertura_glossary_$i"] ?? $glossary_dates[$i]['apertura'];
                    $row_data[] = $row["fecha_cierre_glossary_$i"] ?? $glossary_dates[$i]['cierre'];
                    $row_data[] = $row["visitas_glossary_$i"] ?? '';
                    $row_data[] = $row["numero_envios_glossary_$i"] ?? '';
                    $row_data[] = $row["nota_final_glossary_$i"] ?? '';
                    $row_data[] = $row["retroalimentada_glossary_$i"] ?? 'NO';
                }

                for ($i = 1; $i <= $max_scorms; $i++) {
                    $row_data[] = $row["fecha_apertura_scorm_$i"] ?? $scorm_dates[$i]['apertura'];
                    $row_data[] = $row["fecha_cierre_scorm_$i"] ?? $scorm_dates[$i]['cierre'];
                    $row_data[] = $row["visitas_scorm_$i"] ?? '';
                    $row_data[] = $row["numero_envios_scorm_$i"] ?? '';
                    $row_data[] = $row["nota_final_scorm_$i"] ?? '';
                    $row_data[] = $row["retroalimentada_scorm_$i"] ?? 'NO';
                }

                $sheet->fromArray($row_data, null, 'A' . $row_number);
                $row_number++;
            }
        }

        // Guardar el archivo Excel en el directorio temporal
        $nombre_archivo = "reporte_actividades_{$nombre_grupo}.xlsx";
        $ruta_archivo = "$temp_dir/$nombre_archivo";
        $writer = new Xlsx($spreadsheet);
        $writer->save($ruta_archivo);
        $archivos_excel[] = $ruta_archivo;
    }

    // Crear el archivo ZIP
    $zip = new ZipArchive();
    $zip_file = "$temp_dir/actividades_ddd.zip";
    if ($zip->open($zip_file, ZipArchive::CREATE | ZipArchive::OVERWRITE) !== true) {
        throw new Exception("No se pudo crear el archivo ZIP: $zip_file");
    }

    // Añadir todos los archivos Excel al ZIP
    foreach ($archivos_excel as $archivo) {
        $zip->addFile($archivo, basename($archivo));
    }
    $zip->close();

    // Enviar el correo con el ZIP adjunto
    $mail = new PHPMailer(true);
    $mail->setFrom('noreply@example.com', 'Reporte Moodle DDD');
    foreach ($correo_destino as $correo) {
        $mail->addAddress($correo);
    }
    $mail->Subject = "Reporte de actividades Diplomado Docencia Digital - " . $hoy->format('Y-m-d');
    $mail->Body = "Cordial Saludo, Adjunto el reporte de actividades Diplomado Docencia Digital en un archivo ZIP que contiene los reportes de todos los grupos de cursos.";
    $mail->addAttachment($zip_file, 'actividades_ddd.zip');
    $mail->send();

    // Notificar éxito
    mail($correo_notificacion, "Estado Reporte", "El reporte de actividades Diplomado Docencia Digital  fue enviado correctamente como actividades_regencia.zip.");

    // Limpiar archivos temporales
    foreach ($archivos_excel as $archivo) {
        unlink($archivo);
    }
    unlink($zip_file);
    rmdir($temp_dir);

} catch (Exception $e) {
    // Intentar limpiar archivos temporales en caso de error
    if (isset($archivos_excel)) {
        foreach ($archivos_excel as $archivo) {
            if (file_exists($archivo)) {
                unlink($archivo);
            }
        }
    }
    if (isset($zip_file) && file_exists($zip_file)) {
        unlink($zip_file);
    }
    if (isset($temp_dir) && is_dir($temp_dir)) {
        rmdir($temp_dir);
    }
    mail($correo_notificacion, 'Error Reporte', 'Error: ' . $e->getMessage());
}
?>
