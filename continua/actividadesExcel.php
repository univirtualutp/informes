<?php
require '/root/scripts/informes/vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

date_default_timezone_set('America/Bogota');

$host = 'localhost';
$dbname = 'moodle';
$user = 'moodle';
$pass = 'M00dl3';
$correo_destino = ['juapabgonzalez@utp.edu.co','univirtual-utp@utp.edu.co'];
$correo_notificacion = 'juapabgonzalez@utp.edu.co';

$hoy = new DateTime();
$lunes = clone $hoy;
$lunes->modify('last monday');
$fecha_inicio = new DateTime(date('Y') . '-04-17 00:00:00');
$fecha_fin = clone $lunes;
$fecha_fin->setTime(23, 59, 59);

try {
    $pdo = new PDO("pgsql:host=$host;dbname=$dbname", $user, $pass);
    $pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);

    $grupos_cursos = [
        'Excel_basico' => ['558'],
        // 'Fundamentos_Quimica_biologia' => ['533','550'],
        // 'Introduccion_regencia' => ['532','552'],
        // 'Uso_herramientas_ofimatica' => ['534', '554']
    ];

    $temp_dir = sys_get_temp_dir() . '/reportes_moodle_' . uniqid();
    if (!mkdir($temp_dir)) {
        throw new Exception("No se pudo crear el directorio temporal: $temp_dir");
    }

    $archivos_excel = [];

    foreach ($grupos_cursos as $nombre_grupo => $cursos) {
        $curso_referencia = $cursos[0];

        // Consulta para obtener actividades del curso de referencia
        $sql_actividades = "SELECT 
                cm.id AS cmid,
                CASE 
                    WHEN m.name = 'forum' THEN 'Foro'
                    WHEN m.name = 'quiz' THEN 'Quiz'
                    WHEN m.name = 'assign' THEN 'Tarea'
                    WHEN m.name = 'lesson' THEN 'Lección'
                    WHEN m.name = 'glossary' THEN 'Glosario'
                    WHEN m.name = 'scorm' THEN 'SCORM'
                END AS activity_type,
                CASE 
                    WHEN m.name = 'forum' THEN f.name
                    WHEN m.name = 'quiz' THEN q.name
                    WHEN m.name = 'assign' THEN a.name
                    WHEN m.name = 'lesson' THEN l.name
                    WHEN m.name = 'glossary' THEN g.name
                    WHEN m.name = 'scorm' THEN s.name
                END AS activity_name,
                CASE 
                    WHEN m.name = 'forum' THEN 
                        CASE WHEN f.duedate > 0 THEN TO_TIMESTAMP(f.duedate) ELSE TO_TIMESTAMP(c.startdate) END
                    WHEN m.name = 'quiz' THEN 
                        CASE WHEN q.timeopen > 0 THEN TO_TIMESTAMP(q.timeopen) ELSE TO_TIMESTAMP(c.startdate) END
                    WHEN m.name = 'assign' THEN 
                        CASE WHEN a.allowsubmissionsfromdate > 0 THEN TO_TIMESTAMP(a.allowsubmissionsfromdate) ELSE TO_TIMESTAMP(c.startdate) END
                    WHEN m.name = 'lesson' THEN 
                        CASE WHEN l.available > 0 THEN TO_TIMESTAMP(l.available) ELSE TO_TIMESTAMP(c.startdate) END
                    WHEN m.name = 'glossary' THEN 
                        CASE WHEN g.timecreated > 0 THEN TO_TIMESTAMP(g.timecreated) ELSE TO_TIMESTAMP(c.startdate) END
                    WHEN m.name = 'scorm' THEN 
                        CASE WHEN s.timeopen > 0 THEN TO_TIMESTAMP(s.timeopen) ELSE TO_TIMESTAMP(c.startdate) END
                END AS fecha_apertura,
                CASE 
                    WHEN m.name = 'forum' THEN 
                        CASE WHEN f.cutoffdate > 0 THEN TO_TIMESTAMP(f.cutoffdate) ELSE TO_TIMESTAMP(c.enddate) END
                    WHEN m.name = 'quiz' THEN 
                        CASE WHEN q.timeclose > 0 THEN TO_TIMESTAMP(q.timeclose) ELSE TO_TIMESTAMP(c.enddate) END
                    WHEN m.name = 'assign' THEN 
                        CASE WHEN a.duedate > 0 THEN TO_TIMESTAMP(a.duedate) ELSE TO_TIMESTAMP(c.enddate) END
                    WHEN m.name = 'lesson' THEN 
                        CASE WHEN l.deadline > 0 THEN TO_TIMESTAMP(l.deadline) ELSE TO_TIMESTAMP(c.enddate) END
                    WHEN m.name = 'glossary' THEN 
                        CASE WHEN g.timemodified > 0 AND g.timemodified != g.timecreated THEN TO_TIMESTAMP(g.timemodified) ELSE TO_TIMESTAMP(c.enddate) END
                    WHEN m.name = 'scorm' THEN 
                        CASE WHEN s.timeclose > 0 THEN TO_TIMESTAMP(s.timeclose) ELSE TO_TIMESTAMP(c.enddate) END
                END AS fecha_cierre,
                CASE 
                    WHEN m.name = 'forum' THEN f.id
                    WHEN m.name = 'quiz' THEN q.id
                    WHEN m.name = 'assign' THEN a.id
                    WHEN m.name = 'lesson' THEN l.id
                    WHEN m.name = 'glossary' THEN g.id
                    WHEN m.name = 'scorm' THEN s.id
                END AS activity_id
            FROM mdl_course_modules cm
            JOIN mdl_modules m ON cm.module = m.id
            JOIN mdl_course c ON cm.course = c.id
            LEFT JOIN mdl_forum f ON m.name = 'forum' AND cm.instance = f.id
            LEFT JOIN mdl_quiz q ON m.name = 'quiz' AND cm.instance = q.id
            LEFT JOIN mdl_assign a ON m.name = 'assign' AND cm.instance = a.id
            LEFT JOIN mdl_lesson l ON m.name = 'lesson' AND cm.instance = l.id
            LEFT JOIN mdl_glossary g ON m.name = 'glossary' AND cm.instance = g.id
            LEFT JOIN mdl_scorm s ON m.name = 'scorm' AND cm.instance = s.id
            WHERE c.id = :curso_id
              AND m.name IN ('forum', 'quiz', 'assign', 'lesson', 'glossary', 'scorm')
              AND (m.name != 'forum' OR f.name != 'Avisos')
            ORDER BY 
                CASE 
                    WHEN m.name = 'forum' THEN 1
                    WHEN m.name = 'lesson' THEN 2
                    WHEN m.name = 'assign' THEN 3
                    WHEN m.name = 'quiz' THEN 4
                    WHEN m.name = 'glossary' THEN 5
                    WHEN m.name = 'scorm' THEN 6
                END,
                cm.id";

        $stmt = $pdo->prepare($sql_actividades);
        $stmt->execute([':curso_id' => $curso_referencia]);
        $actividades = $stmt->fetchAll(PDO::FETCH_ASSOC);

        // Organizar actividades por tipo
        $foros = array_filter($actividades, function($a) { return $a['activity_type'] == 'Foro'; });
        $lecciones = array_filter($actividades, function($a) { return $a['activity_type'] == 'Lección'; });
        $tareas = array_filter($actividades, function($a) { return $a['activity_type'] == 'Tarea'; });
        $quizzes = array_filter($actividades, function($a) { return $a['activity_type'] == 'Quiz'; });
        $glosarios = array_filter($actividades, function($a) { return $a['activity_type'] == 'Glosario'; });
        $scorms = array_filter($actividades, function($a) { return $a['activity_type'] == 'SCORM'; });

        // Configurar encabezados del Excel (7 columnas iniciales)
        $headers = array_fill(0, 7, '');
        $sub_headers = [
            'codigo', 'nombre', 'apellidos', 'rol', 'correo', 'id_curso', 'curso'
        ];

        foreach ([$foros, $lecciones, $tareas, $quizzes, $glosarios, $scorms] as $tipo_actividades) {
            foreach ($tipo_actividades as $actividad) {
                $headers[] = $actividad['activity_name'];
                $headers = array_merge($headers, array_fill(0, 6, ''));
                $sub_headers = array_merge($sub_headers, [
                    'fecha_apertura',
                    'fecha_cierre',
                    'ultima_fecha_envio',
                    'visitas',
                    'numero_envios',
                    'nota_final',
                    'retroalimentada'
                ]);
            }
        }

        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->fromArray($headers, null, 'A1');
        $sheet->fromArray($sub_headers, null, 'A2');
        $row_number = 3;

        foreach ($cursos as $curso_id) {
            // CONSULTA PRINCIPAL
            $sql = "WITH 
                user_info AS (
                    SELECT 
                        u.id AS userid,
                        u.username,
                        u.firstname,
                        u.lastname,
                        r.name AS rol,
                        u.email,
                        c.id AS courseid,
                        c.fullname AS curso
                    FROM mdl_user u
                    JOIN mdl_role_assignments ra ON ra.userid = u.id
                    JOIN mdl_context mc ON ra.contextid = mc.id AND mc.contextlevel = 50
                    JOIN mdl_course c ON mc.instanceid = c.id
                    JOIN mdl_role r ON ra.roleid = r.id
                    WHERE c.id = :curso_id
                      AND u.username <> '12345678'
                      AND ra.roleid IN ('5','9')
                ),
                actividades AS (
                    SELECT 
                        cm.id AS cmid,
                        c.id AS courseid,
                        CASE 
                            WHEN m.name = 'forum' THEN 'Foro'
                            WHEN m.name = 'quiz' THEN 'Quiz'
                            WHEN m.name = 'assign' THEN 'Tarea'
                            WHEN m.name = 'lesson' THEN 'Lección'
                            WHEN m.name = 'glossary' THEN 'Glosario'
                            WHEN m.name = 'scorm' THEN 'SCORM'
                        END AS activity_type,
                        CASE 
                            WHEN m.name = 'forum' THEN f.name
                            WHEN m.name = 'quiz' THEN q.name
                            WHEN m.name = 'assign' THEN a.name
                            WHEN m.name = 'lesson' THEN l.name
                            WHEN m.name = 'glossary' THEN g.name
                            WHEN m.name = 'scorm' THEN s.name
                        END AS activity_name,
                        CASE 
                            WHEN m.name = 'forum' THEN f.id
                            WHEN m.name = 'quiz' THEN q.id
                            WHEN m.name = 'assign' THEN a.id
                            WHEN m.name = 'lesson' THEN l.id
                            WHEN m.name = 'glossary' THEN g.id
                            WHEN m.name = 'scorm' THEN s.id
                        END AS activity_id,
                        CASE 
                            WHEN m.name = 'forum' THEN 
                                CASE WHEN f.duedate > 0 THEN TO_TIMESTAMP(f.duedate) ELSE TO_TIMESTAMP(c.startdate) END
                            WHEN m.name = 'quiz' THEN 
                                CASE WHEN q.timeopen > 0 THEN TO_TIMESTAMP(q.timeopen) ELSE TO_TIMESTAMP(c.startdate) END
                            WHEN m.name = 'assign' THEN 
                                CASE WHEN a.allowsubmissionsfromdate > 0 THEN TO_TIMESTAMP(a.allowsubmissionsfromdate) ELSE TO_TIMESTAMP(c.startdate) END
                            WHEN m.name = 'lesson' THEN 
                                CASE WHEN l.available > 0 THEN TO_TIMESTAMP(l.available) ELSE TO_TIMESTAMP(c.startdate) END
                            WHEN m.name = 'glossary' THEN 
                                CASE WHEN g.timecreated > 0 THEN TO_TIMESTAMP(g.timecreated) ELSE TO_TIMESTAMP(c.startdate) END
                            WHEN m.name = 'scorm' THEN 
                                CASE WHEN s.timeopen > 0 THEN TO_TIMESTAMP(s.timeopen) ELSE TO_TIMESTAMP(c.startdate) END
                        END AS fecha_apertura,
                        CASE 
                            WHEN m.name = 'forum' THEN 
                                CASE WHEN f.cutoffdate > 0 THEN TO_TIMESTAMP(f.cutoffdate) ELSE TO_TIMESTAMP(c.enddate) END
                            WHEN m.name = 'quiz' THEN 
                                CASE WHEN q.timeclose > 0 THEN TO_TIMESTAMP(q.timeclose) ELSE TO_TIMESTAMP(c.enddate) END
                            WHEN m.name = 'assign' THEN 
                                CASE WHEN a.duedate > 0 THEN TO_TIMESTAMP(a.duedate) ELSE TO_TIMESTAMP(c.enddate) END
                            WHEN m.name = 'lesson' THEN 
                                CASE WHEN l.deadline > 0 THEN TO_TIMESTAMP(l.deadline) ELSE TO_TIMESTAMP(c.enddate) END
                            WHEN m.name = 'glossary' THEN 
                                CASE WHEN g.timemodified > 0 AND g.timemodified != g.timecreated THEN TO_TIMESTAMP(g.timemodified) ELSE TO_TIMESTAMP(c.enddate) END
                            WHEN m.name = 'scorm' THEN 
                                CASE WHEN s.timeclose > 0 THEN TO_TIMESTAMP(s.timeclose) ELSE TO_TIMESTAMP(c.enddate) END
                        END AS fecha_cierre
                    FROM mdl_course_modules cm
                    JOIN mdl_modules m ON cm.module = m.id
                    JOIN mdl_course c ON cm.course = c.id
                    LEFT JOIN mdl_forum f ON m.name = 'forum' AND cm.instance = f.id
                    LEFT JOIN mdl_quiz q ON m.name = 'quiz' AND cm.instance = q.id
                    LEFT JOIN mdl_assign a ON m.name = 'assign' AND cm.instance = a.id
                    LEFT JOIN mdl_lesson l ON m.name = 'lesson' AND cm.instance = l.id
                    LEFT JOIN mdl_glossary g ON m.name = 'glossary' AND cm.instance = g.id
                    LEFT JOIN mdl_scorm s ON m.name = 'scorm' AND cm.instance = s.id
                    WHERE c.id = :curso_id
                      AND m.name IN ('forum', 'quiz', 'assign', 'lesson', 'glossary', 'scorm')
                      AND (m.name != 'forum' OR f.name != 'Avisos')
                ),
                interacciones AS (
                    SELECT 
                        fp.userid,
                        f.course AS courseid,
                        f.id AS activity_id,
                        'Foro' AS activity_type,
                        COUNT(DISTINCT fp.id) AS num_envios,
                        MAX(fp.created) AS ultima_fecha_envio,
                        (SELECT gg.finalgrade 
                         FROM mdl_grade_grades gg
                         JOIN mdl_grade_items gi ON gg.itemid = gi.id
                         WHERE gi.courseid = f.course 
                           AND gi.itemmodule = 'forum' 
                           AND gi.iteminstance = f.id 
                           AND gg.userid = fp.userid) AS nota_final,
                        (SELECT CASE WHEN gg.feedback IS NOT NULL AND gg.feedback != '' THEN 'SÍ' ELSE 'NO' END
                         FROM mdl_grade_grades gg
                         JOIN mdl_grade_items gi ON gg.itemid = gi.id
                         WHERE gi.courseid = f.course 
                           AND gi.itemmodule = 'forum' 
                           AND gi.iteminstance = f.id 
                           AND gg.userid = fp.userid) AS retroalimentada
                    FROM mdl_forum_posts fp
                    JOIN mdl_forum_discussions fd ON fp.discussion = fd.id
                    JOIN mdl_forum f ON fd.forum = f.id
                    WHERE f.course = :curso_id
                      AND (fp.created BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) 
                                      AND EXTRACT(EPOCH FROM :fecha_fin::timestamp))
                    GROUP BY fp.userid, f.course, f.id
                    
                    UNION ALL
                    
                    SELECT 
                        sub.userid,
                        a.course AS courseid,
                        a.id AS activity_id,
                        'Tarea' AS activity_type,
                        COUNT(DISTINCT sub.id) AS num_envios,
                        MAX(sub.timemodified) AS ultima_fecha_envio,
                        (SELECT gg.finalgrade 
                         FROM mdl_grade_grades gg
                         JOIN mdl_grade_items gi ON gg.itemid = gi.id
                         WHERE gi.courseid = a.course 
                           AND gi.itemmodule = 'assign' 
                           AND gi.iteminstance = a.id 
                           AND gg.userid = sub.userid) AS nota_final,
                        (SELECT CASE WHEN gg.feedback IS NOT NULL AND gg.feedback != '' THEN 'SÍ' ELSE 'NO' END
                         FROM mdl_grade_grades gg
                         JOIN mdl_grade_items gi ON gg.itemid = gi.id
                         WHERE gi.courseid = a.course 
                           AND gi.itemmodule = 'assign' 
                           AND gi.iteminstance = a.id 
                           AND gg.userid = sub.userid) AS retroalimentada
                    FROM mdl_assign_submission sub
                    JOIN mdl_assign a ON sub.assignment = a.id
                    WHERE a.course = :curso_id
                      AND (sub.timemodified BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) 
                                       AND EXTRACT(EPOCH FROM :fecha_fin::timestamp))
                    GROUP BY sub.userid, a.course, a.id
                    
                    UNION ALL
                    
                    SELECT 
                        qg.userid,
                        q.course AS courseid,
                        q.id AS activity_id,
                        'Quiz' AS activity_type,
                        (SELECT COUNT(DISTINCT qa2.id)
                         FROM mdl_quiz_attempts qa2
                         WHERE qa2.quiz = q.id
                           AND qa2.userid = qg.userid
                           AND qa2.state = 'finished'
                           AND qa2.timefinish BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) 
                                             AND EXTRACT(EPOCH FROM :fecha_fin::timestamp)) AS num_envios,
                        (SELECT MAX(qa3.timefinish)
                         FROM mdl_quiz_attempts qa3
                         WHERE qa3.quiz = q.id
                           AND qa3.userid = qg.userid
                           AND qa3.state = 'finished') AS ultima_fecha_envio,
                        qg.grade AS nota_final,
                        (SELECT CASE WHEN gg.feedback IS NOT NULL AND gg.feedback != '' THEN 'SÍ' ELSE 'NO' END
                         FROM mdl_grade_grades gg
                         JOIN mdl_grade_items gi ON gg.itemid = gi.id
                         WHERE gi.courseid = q.course 
                           AND gi.itemmodule = 'quiz' 
                           AND gi.iteminstance = q.id 
                           AND gg.userid = qg.userid) AS retroalimentada
                    FROM mdl_quiz_grades qg
                    JOIN mdl_quiz q ON qg.quiz = q.id
                    WHERE q.course = :curso_id
                    GROUP BY qg.userid, q.course, q.id, qg.grade
                    
                    UNION ALL
                    
                    SELECT 
                        lg.userid,
                        l.course AS courseid,
                        l.id AS activity_id,
                        'Lección' AS activity_type,
                        COUNT(DISTINCT lg.id) AS num_envios,
                        MAX(lg.completed) AS ultima_fecha_envio,
                        (SELECT gg.finalgrade 
                         FROM mdl_grade_grades gg
                         JOIN mdl_grade_items gi ON gg.itemid = gi.id
                         WHERE gi.courseid = l.course 
                           AND gi.itemmodule = 'lesson' 
                           AND gi.iteminstance = l.id 
                           AND gg.userid = lg.userid) AS nota_final,
                        (SELECT CASE WHEN gg.feedback IS NOT NULL AND gg.feedback != '' THEN 'SÍ' ELSE 'NO' END
                         FROM mdl_grade_grades gg
                         JOIN mdl_grade_items gi ON gg.itemid = gi.id
                         WHERE gi.courseid = l.course 
                           AND gi.itemmodule = 'lesson' 
                           AND gi.iteminstance = l.id 
                           AND gg.userid = lg.userid) AS retroalimentada
                    FROM mdl_lesson_grades lg
                    JOIN mdl_lesson l ON lg.lessonid = l.id
                    WHERE l.course = :curso_id
                      AND (lg.completed BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) 
                                    AND EXTRACT(EPOCH FROM :fecha_fin::timestamp))
                    GROUP BY lg.userid, l.course, l.id
                    
                    UNION ALL
                    
                    SELECT 
                        ge.userid,
                        g.course AS courseid,
                        g.id AS activity_id,
                        'Glosario' AS activity_type,
                        COUNT(DISTINCT ge.id) AS num_envios,
                        MAX(ge.timecreated) AS ultima_fecha_envio,
                        (SELECT gg.finalgrade 
                         FROM mdl_grade_grades gg
                         JOIN mdl_grade_items gi ON gg.itemid = gi.id
                         WHERE gi.courseid = g.course 
                           AND gi.itemmodule = 'glossary' 
                           AND gi.iteminstance = g.id 
                           AND gg.userid = ge.userid) AS nota_final,
                        (SELECT CASE WHEN gg.feedback IS NOT NULL AND gg.feedback != '' THEN 'SÍ' ELSE 'NO' END
                         FROM mdl_grade_grades gg
                         JOIN mdl_grade_items gi ON gg.itemid = gi.id
                         WHERE gi.courseid = g.course 
                           AND gi.itemmodule = 'glossary' 
                           AND gi.iteminstance = g.id 
                           AND gg.userid = ge.userid) AS retroalimentada
                    FROM mdl_glossary_entries ge
                    JOIN mdl_glossary g ON ge.glossaryid = g.id
                    WHERE g.course = :curso_id
                      AND (ge.timecreated BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) 
                                      AND EXTRACT(EPOCH FROM :fecha_fin::timestamp))
                    GROUP BY ge.userid, g.course, g.id
                    
                    UNION ALL
                    
                    SELECT 
                        st.userid,
                        s.course AS courseid,
                        s.id AS activity_id,
                        'SCORM' AS activity_type,
                        COUNT(DISTINCT st.id) AS num_envios,
                        MAX(st.timemodified) AS ultima_fecha_envio,
                        (SELECT gg.finalgrade 
                         FROM mdl_grade_grades gg
                         JOIN mdl_grade_items gi ON gg.itemid = gi.id
                         WHERE gi.courseid = s.course 
                           AND gi.itemmodule = 'scorm' 
                           AND gi.iteminstance = s.id 
                           AND gg.userid = st.userid) AS nota_final,
                        (SELECT CASE WHEN gg.feedback IS NOT NULL AND gg.feedback != '' THEN 'SÍ' ELSE 'NO' END
                         FROM mdl_grade_grades gg
                         JOIN mdl_grade_items gi ON gg.itemid = gi.id
                         WHERE gi.courseid = s.course 
                           AND gi.itemmodule = 'scorm' 
                           AND gi.iteminstance = s.id 
                           AND gg.userid = st.userid) AS retroalimentada
                    FROM mdl_scorm_scoes_track st
                    JOIN mdl_scorm s ON st.scormid = s.id
                    WHERE s.course = :curso_id
                      AND (st.timemodified BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) 
                                       AND EXTRACT(EPOCH FROM :fecha_fin::timestamp))
                    GROUP BY st.userid, s.course, s.id
                ),
                visitas AS (
                    SELECT 
                        log.userid,
                        log.courseid,
                        log.objectid AS activity_id,
                        CASE 
                            WHEN log.component = 'mod_forum' THEN 'Foro'
                            WHEN log.component = 'mod_quiz' THEN 'Quiz'
                            WHEN log.component = 'mod_assign' THEN 'Tarea'
                            WHEN log.component = 'mod_lesson' THEN 'Lección'
                            WHEN log.component = 'mod_glossary' THEN 'Glosario'
                            WHEN log.component = 'mod_scorm' THEN 'SCORM'
                        END AS activity_type,
                        COUNT(*) AS num_visitas
                    FROM mdl_logstore_standard_log log
                    WHERE log.action = 'viewed'
                      AND log.timecreated BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) 
                         AND EXTRACT(EPOCH FROM :fecha_fin::timestamp)
                      AND log.courseid = :curso_id
                      AND log.component IN ('mod_forum', 'mod_lesson', 'mod_assign', 'mod_quiz', 'mod_glossary', 'mod_scorm')
                    GROUP BY log.userid, log.courseid, log.objectid, log.component
                )
                SELECT 
                    ui.username AS codigo,
                    ui.firstname AS nombre,
                    ui.lastname AS apellidos,
                    ui.rol,
                    ui.email AS correo,
                    ui.courseid AS id_curso,
                    ui.curso,
                    a.activity_type,
                    a.activity_name,
                    TO_CHAR(a.fecha_apertura, 'YYYY-MM-DD HH24:MI:SS') AS fecha_apertura,
                    TO_CHAR(a.fecha_cierre, 'YYYY-MM-DD HH24:MI:SS') AS fecha_cierre,
                    TO_CHAR(TO_TIMESTAMP(i.ultima_fecha_envio), 'YYYY-MM-DD HH24:MI:SS') AS ultima_fecha_envio,
                    COALESCE(v.num_visitas, 0) AS visitas,
                    COALESCE(i.num_envios, 0) AS num_envios,
                    COALESCE(i.nota_final, 0) AS nota_final,
                    COALESCE(i.retroalimentada, 'NO') AS retroalimentada
                FROM user_info ui
                CROSS JOIN actividades a
                LEFT JOIN interacciones i ON ui.userid = i.userid 
                    AND ui.courseid = i.courseid 
                    AND a.activity_id = i.activity_id 
                    AND a.activity_type = i.activity_type
                LEFT JOIN visitas v ON ui.userid = v.userid 
                    AND ui.courseid = v.courseid 
                    AND a.activity_id = v.activity_id 
                    AND a.activity_type = v.activity_type
                ORDER BY ui.lastname, ui.firstname, 
                    CASE 
                        WHEN a.activity_type = 'Foro' THEN 1
                        WHEN a.activity_type = 'Lección' THEN 2
                        WHEN a.activity_type = 'Tarea' THEN 3
                        WHEN a.activity_type = 'Quiz' THEN 4
                        WHEN a.activity_type = 'Glosario' THEN 5
                        WHEN a.activity_type = 'SCORM' THEN 6
                    END,
                    a.activity_name";

            $params = [
                ':curso_id' => $curso_id,
                ':fecha_inicio' => $fecha_inicio->format('Y-m-d H:i:s'),
                ':fecha_fin' => $fecha_fin->format('Y-m-d H:i:s')
            ];

            $stmt = $pdo->prepare($sql);
            $stmt->execute($params);
            $resultados = $stmt->fetchAll(PDO::FETCH_ASSOC);

            // Organizar datos para el Excel
            $datos_usuarios = [];
            foreach ($resultados as $row) {
                $codigo = $row['codigo'];
                if (!isset($datos_usuarios[$codigo])) {
                    $datos_usuarios[$codigo] = [
                        'info' => [
                            $row['codigo'],
                            $row['nombre'],
                            $row['apellidos'],
                            $row['rol'],
                            $row['correo'],
                            $row['id_curso'],
                            $row['curso']
                        ],
                        'actividades' => []
                    ];
                }

                $datos_usuarios[$codigo]['actividades'][$row['activity_type']][$row['activity_name']] = [
                    'fecha_apertura' => $row['fecha_apertura'],
                    'fecha_cierre' => $row['fecha_cierre'],
                    'ultima_fecha_envio' => $row['ultima_fecha_envio'],
                    'visitas' => $row['visitas'],
                    'num_envios' => $row['num_envios'],
                    'nota_final' => $row['nota_final'],
                    'retroalimentada' => $row['retroalimentada']
                ];
            }

            // Llenar la hoja de cálculo
            foreach ($datos_usuarios as $usuario) {
                $row_data = $usuario['info'];

                foreach ([$foros, $lecciones, $tareas, $quizzes, $glosarios, $scorms] as $tipo_actividades) {
                    foreach ($tipo_actividades as $actividad) {
                        $actividad_data = $usuario['actividades'][$actividad['activity_type']][$actividad['activity_name']] ?? null;
                        
                        $row_data[] = $actividad['fecha_apertura'];
                        $row_data[] = $actividad['fecha_cierre'];
                        $row_data[] = $actividad_data ? $actividad_data['ultima_fecha_envio'] : '';
                        $row_data[] = $actividad_data ? $actividad_data['visitas'] : '';
                        $row_data[] = $actividad_data ? $actividad_data['num_envios'] : '';
                        $row_data[] = $actividad_data ? $actividad_data['nota_final'] : '';
                        $row_data[] = $actividad_data ? $actividad_data['retroalimentada'] : 'NO';
                    }
                }

                $sheet->fromArray($row_data, null, 'A' . $row_number);
                $row_number++;
            }
        }

        // Guardar archivo Excel
        $nombre_archivo = "reporte_actividades_{$nombre_grupo}.xlsx";
        $ruta_archivo = "$temp_dir/$nombre_archivo";
        $writer = new Xlsx($spreadsheet);
        $writer->save($ruta_archivo);
        $archivos_excel[] = $ruta_archivo;
    }

    // Preparar el correo
    $to = implode(',', $correo_destino);
    $subject = "Reporte de actividades Curso Excel para la generación de informes - " . $hoy->format('Y-m-d');
    $boundary = uniqid('np');
    $headers = "From: noreply-univirtual@utp.edu.co\r\n";
    $headers .= "MIME-Version: 1.0\r\n";
    $headers .= "Content-Type: multipart/mixed; boundary=\"$boundary\"\r\n";

    // Verificar si hay solo un grupo de cursos
    if (count($grupos_cursos) == 1) {
        // Caso: Solo un grupo de cursos, enviar el archivo Excel directamente
        $nombre_grupo = array_key_first($grupos_cursos);
        $body = "Cordial Saludo,\n\nAdjunto el reporte de actividades del curso Excel para la generación de informes correspondiente al grupo $nombre_grupo.\n\nSaludos,\nSistema de Reportes Moodle";
        $file_path = $archivos_excel[0];
        $file_name = "reporte_actividades_{$nombre_grupo}.xlsx";
        $mime_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    } else {
        // Caso: Múltiples grupos de cursos, crear y enviar archivo ZIP
        $zip = new ZipArchive();
        $zip_file = "$temp_dir/actividades_regencia.zip";
        if ($zip->open($zip_file, ZipArchive::CREATE | ZipArchive::OVERWRITE) !== true) {
            throw new Exception("No se pudo crear el archivo ZIP: $zip_file");
        }

        foreach ($archivos_excel as $archivo) {
            $zip->addFile($archivo, basename($archivo));
        }
        $zip->close();

        $body = "Cordial Saludo,\n\nAdjunto el reporte de actividades Curso Excel para la generación de informes en un archivo ZIP que contiene los reportes de todos los grupos de cursos.\n\nSaludos,\nSistema de Reportes Moodle";
        $file_path = $zip_file;
        $file_name = 'actividades_excel_generacion_informes.zip';
        $mime_type = 'application/zip';
    }

    // Leer y codificar el archivo adjunto
    $file_content = file_get_contents($file_path);
    $file_encoded = chunk_split(base64_encode($file_content));

    // Construir el cuerpo del correo con el adjunto
    $message = "--$boundary\r\n";
    $message .= "Content-Type: text/plain; charset=UTF-8\r\n";
    $message .= "Content-Transfer-Encoding: 7bit\r\n\r\n";
    $message .= $body . "\r\n";
    $message .= "--$boundary\r\n";
    $message .= "Content-Type: $mime_type; name=\"$file_name\"\r\n";
    $message .= "Content-Transfer-Encoding: base64\r\n";
    $message .= "Content-Disposition: attachment; filename=\"$file_name\"\r\n\r\n";
    $message .= $file_encoded . "\r\n";
    $message .= "--$boundary--";

    // Enviar el correo
    if (!mail($to, $subject, $message, $headers)) {
        throw new Exception("Error al enviar el correo");
    }

    // Notificar éxito
    mail($correo_notificacion, "Estado Reporte", "El reporte de actividades Curso Excel para la generación de informes fue enviado correctamente.");

    // Limpiar archivos temporales
    foreach ($archivos_excel as $archivo) {
        if (file_exists($archivo)) {
            unlink($archivo);
        }
    }
    if (isset($zip_file) && file_exists($zip_file)) {
        unlink($zip_file);
    }
    if (is_dir($temp_dir)) {
        rmdir($temp_dir);
    }

} catch (Exception $e) {
    // Manejo de errores
    error_log("Error en actividadesExcel.php: " . $e->getMessage());
    
    // Intentar limpiar archivos temporales
    if (isset($archivos_excel)) {
        foreach ($archivos_excel as $archivo) {
            if (file_exists($archivo)) {
                @unlink($archivo);
            }
        }
    }
    if (isset($zip_file) && file_exists($zip_file)) {
        @unlink($zip_file);
    }
    if (isset($temp_dir) && is_dir($temp_dir)) {
        @rmdir($temp_dir);
    }
    
    // Notificar error
    mail($correo_notificacion, 'Error Reporte', 'Error en el script actividadesExcel.php: ' . $e->getMessage());
    
    // Salir con código de error
    exit(1);
}
?>
