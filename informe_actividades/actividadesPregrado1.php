<?php
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\Exception;

// Configuración de la base de datos y correos
$host = 'localhost';
$dbname = 'moodle';
$user = 'moodle';
$pass = 'M00dl3';
$correo_destino = ['daniel.pardo@utp.edu.co'];
$correo_notificacion = 'daniel.pardo@utp.edu.co';

// Configuración de fechas para el reporte
$fecha_inicio = new DateTime('2025-02-03 00:00:00'); // Desde el 3 de febrero del año actual
$fecha_fin = new DateTime(date('Y-m-d H:i:s')); // Hasta el día actual

try {
    // Conexión a la base de datos PostgreSQL
    $pdo = new PDO("pgsql:host=$host;dbname=$dbname", $user, $pass);
    $pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);

    // Parámetros para las consultas con fechas
    $params = [
        ':fecha_inicio' => $fecha_inicio->format('Y-m-d H:i:s'),
        ':fecha_fin' => $fecha_fin->format('Y-m-d H:i:s')
    ];

    // Obtener nombres y conteos de actividades (foros, lecciones, tareas, etc.) solo en el rango de fechas
    $forum_sql = "
        SELECT DISTINCT f.id AS forum_number, f.name AS forumname 
        FROM mdl_forum f
        LEFT JOIN mdl_forum_discussions fd ON f.id = fd.forum
        LEFT JOIN mdl_forum_posts fp ON fd.id = fp.discussion
        WHERE f.course IN ('428') 
          AND f.name != 'Avisos' 
          AND (fp.created BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) AND EXTRACT(EPOCH FROM :fecha_fin::timestamp) OR fp.created IS NULL)
        ORDER BY f.name";
    $stmt = $pdo->prepare($forum_sql);
    $stmt->execute($params);
    $forum_results = $stmt->fetchAll(PDO::FETCH_ASSOC);
    $forum_names = [];
    $max_forums = count($forum_results);
    foreach ($forum_results as $i => $row) {
        $forum_names[$i + 1] = $row['forumname'];
    }

    $lesson_sql = "
        SELECT DISTINCT l.id AS lesson_number, l.name AS lessonname 
        FROM mdl_lesson l
        LEFT JOIN mdl_lesson_grades lg ON l.id = lg.lessonid
        WHERE l.course IN ('428') 
          AND (lg.completed BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) AND EXTRACT(EPOCH FROM :fecha_fin::timestamp) OR lg.completed IS NULL)
        ORDER BY l.name";
    $stmt = $pdo->prepare($lesson_sql);
    $stmt->execute($params);
    $lesson_results = $stmt->fetchAll(PDO::FETCH_ASSOC);
    $lesson_names = [];
    $max_lessons = count($lesson_results);
    foreach ($lesson_results as $i => $row) {
        $lesson_names[$i + 1] = $row['lessonname'];
    }

    $assign_sql = "
        SELECT DISTINCT a.id AS assign_number, a.name AS assignname 
        FROM mdl_assign a
        LEFT JOIN mdl_assign_submission sub ON a.id = sub.assignment
        WHERE a.course IN ('428') 
          AND (sub.timecreated BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) AND EXTRACT(EPOCH FROM :fecha_fin::timestamp) OR sub.timecreated IS NULL)
        ORDER BY a.name";
    $stmt = $pdo->prepare($assign_sql);
    $stmt->execute($params);
    $assign_results = $stmt->fetchAll(PDO::FETCH_ASSOC);
    $assign_names = [];
    $max_assigns = count($assign_results);
    foreach ($assign_results as $i => $row) {
        $assign_names[$i + 1] = $row['assignname'];
    }

    $quiz_sql = "
        SELECT DISTINCT q.id AS quiz_number, q.name AS quizname 
        FROM mdl_quiz q
        LEFT JOIN mdl_quiz_attempts qa ON q.id = qa.quiz
        WHERE q.course IN ('428') 
          AND (qa.timefinish BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) AND EXTRACT(EPOCH FROM :fecha_fin::timestamp) OR qa.timefinish IS NULL)
        ORDER BY q.name";
    $stmt = $pdo->prepare($quiz_sql);
    $stmt->execute($params);
    $quiz_results = $stmt->fetchAll(PDO::FETCH_ASSOC);
    $quiz_names = [];
    $max_quizzes = count($quiz_results);
    foreach ($quiz_results as $i => $row) {
        $quiz_names[$i + 1] = $row['quizname'];
    }

    $glossary_sql = "
        SELECT DISTINCT g.id AS glossary_number, g.name AS glossaryname 
        FROM mdl_glossary g
        LEFT JOIN mdl_glossary_entries ge ON g.id = ge.glossaryid
        WHERE g.course IN ('428') 
          AND (ge.timecreated BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) AND EXTRACT(EPOCH FROM :fecha_fin::timestamp) OR ge.timecreated IS NULL)
        ORDER BY g.name";
    $stmt = $pdo->prepare($glossary_sql);
    $stmt->execute($params);
    $glossary_results = $stmt->fetchAll(PDO::FETCH_ASSOC);
    $glossary_names = [];
    $max_glossaries = count($glossary_results);
    foreach ($glossary_results as $i => $row) {
        $glossary_names[$i + 1] = $row['glossaryname'];
    }

    $scorm_sql = "
        SELECT DISTINCT s.id AS scorm_number, s.name AS scormname 
        FROM mdl_scorm s
        LEFT JOIN mdl_scorm_scoes_track st ON s.id = st.scormid
        WHERE s.course IN ('428') 
          AND (st.timemodified BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) AND EXTRACT(EPOCH FROM :fecha_fin::timestamp) OR st.timemodified IS NULL)
        ORDER BY s.name";
    $stmt = $pdo->prepare($scorm_sql);
    $stmt->execute($params);
    $scorm_results = $stmt->fetchAll(PDO::FETCH_ASSOC);
    $scorm_names = [];
    $max_scorms = count($scorm_results);
    foreach ($scorm_results as $i => $row) {
        $scorm_names[$i + 1] = $row['scormname'];
    }

    // Crear tablas temporales para pre-agregación de datos
    // Foro: Usar duedate y cutoffdate
    $pdo->exec("DROP TABLE IF EXISTS temp_forum_agg");
    $forum_agg_sql = "
        CREATE TEMP TABLE temp_forum_agg AS
        SELECT 
            fp.userid,
            f.course AS courseid,
            f.name AS forumname,
            COUNT(DISTINCT fp.id) AS num_posts,
            COALESCE(gg.finalgrade, 0) AS nota_final,
            CASE WHEN f.duedate > 0 THEN to_timestamp(f.duedate)::date ELSE NULL END AS fecha_apertura,
            CASE WHEN f.cutoffdate > 0 THEN to_timestamp(f.cutoffdate)::date ELSE NULL END AS fecha_cierre,
            CASE WHEN COUNT(DISTINCT CASE WHEN ra2.roleid IN (3, 4) AND fp2.userid != fp.userid THEN fp2.id END) > 0 THEN 'SÍ' ELSE 'NO' END AS retroalimenta,
            ROW_NUMBER() OVER (PARTITION BY fp.userid, f.course ORDER BY f.name) AS forum_number
        FROM mdl_forum_posts fp
        JOIN mdl_forum_discussions fd ON fp.discussion = fd.id
        JOIN mdl_forum f ON fd.forum = f.id
        LEFT JOIN mdl_forum_posts fp2 ON fp2.discussion = fd.id
        LEFT JOIN mdl_user u2 ON fp2.userid = u2.id
        LEFT JOIN mdl_role_assignments ra2 ON u2.id = ra2.userid AND ra2.contextid IN (
            SELECT contextid FROM mdl_context WHERE instanceid = f.course AND contextlevel = 50
        )
        LEFT JOIN mdl_grade_grades gg ON gg.userid = fp.userid 
            AND gg.itemid IN (SELECT id FROM mdl_grade_items WHERE itemmodule = 'forum' AND iteminstance = f.id)
        WHERE fp.created BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) 
                            AND EXTRACT(EPOCH FROM :fecha_fin::timestamp)
          AND f.name != 'Avisos'
          AND f.course IN ('428')
        GROUP BY fp.userid, f.course, f.name, gg.finalgrade, f.duedate, f.cutoffdate";
    $stmt = $pdo->prepare($forum_agg_sql);
    $stmt->execute($params);
    $pdo->exec("CREATE INDEX ON temp_forum_agg (userid, courseid)");

    // Lección
    $pdo->exec("DROP TABLE IF EXISTS temp_lesson_agg");
    $lesson_agg_sql = "
        CREATE TEMP TABLE temp_lesson_agg AS
        SELECT 
            lg.userid,
            l.course AS courseid,
            l.name AS lessonname,
            COUNT(DISTINCT lg.id) AS num_attempts,
            COALESCE(gg.finalgrade, 0) AS nota_final,
            CASE WHEN l.available > 0 THEN to_timestamp(l.available)::date ELSE NULL END AS fecha_apertura,
            CASE WHEN l.deadline > 0 THEN to_timestamp(l.deadline)::date ELSE NULL END AS fecha_cierre,
            CASE WHEN gg.feedback IS NOT NULL THEN 'SÍ' ELSE 'NO' END AS retroalimenta,
            ROW_NUMBER() OVER (PARTITION BY lg.userid, l.course ORDER BY l.name) AS lesson_number
        FROM mdl_lesson_grades lg
        JOIN mdl_lesson l ON lg.lessonid = l.id
        LEFT JOIN mdl_grade_grades gg ON gg.userid = lg.userid 
            AND gg.itemid IN (SELECT id FROM mdl_grade_items WHERE itemmodule = 'lesson' AND iteminstance = l.id)
        WHERE lg.completed BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) 
                              AND EXTRACT(EPOCH FROM :fecha_fin::timestamp)
          AND l.course IN ('428')
        GROUP BY lg.userid, l.course, l.name, gg.finalgrade, l.available, l.deadline, gg.feedback";
    $stmt = $pdo->prepare($lesson_agg_sql);
    $stmt->execute($params);
    $pdo->exec("CREATE INDEX ON temp_lesson_agg (userid, courseid)");

    // Tarea
    $pdo->exec("DROP TABLE IF EXISTS temp_assign_agg");
    $assign_agg_sql = "
        CREATE TEMP TABLE temp_assign_agg AS
        SELECT 
            sub.userid,
            a.course AS courseid,
            a.name AS assignname,
            COUNT(DISTINCT sub.id) AS num_submissions,
            COALESCE(gg.finalgrade, 0) AS nota_final,
            CASE WHEN a.allowsubmissionsfromdate > 0 THEN to_timestamp(a.allowsubmissionsfromdate)::date ELSE NULL END AS fecha_apertura,
            CASE WHEN a.duedate > 0 THEN to_timestamp(a.duedate)::date ELSE NULL END AS fecha_cierre,
            CASE WHEN COUNT(DISTINCT CASE WHEN afc.commenttext IS NOT NULL THEN afc.id END) > 0 THEN 'SÍ' ELSE 'NO' END AS retroalimenta,
            ROW_NUMBER() OVER (PARTITION BY sub.userid, a.course ORDER BY a.name) AS assign_number
        FROM mdl_assign_submission sub
        JOIN mdl_assign a ON sub.assignment = a.id
        LEFT JOIN mdl_grade_grades gg ON gg.userid = sub.userid 
            AND gg.itemid IN (SELECT id FROM mdl_grade_items WHERE itemmodule = 'assign' AND iteminstance = a.id)
        LEFT JOIN mdl_assign_grades ag ON ag.assignment = a.id AND ag.userid = sub.userid
        LEFT JOIN mdl_assignfeedback_comments afc ON afc.grade = ag.id
        WHERE sub.timecreated BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) 
                                 AND EXTRACT(EPOCH FROM :fecha_fin::timestamp)
          AND a.course IN ('428')
        GROUP BY sub.userid, a.course, a.name, gg.finalgrade, a.allowsubmissionsfromdate, a.duedate";
    $stmt = $pdo->prepare($assign_agg_sql);
    $stmt->execute($params);
    $pdo->exec("CREATE INDEX ON temp_assign_agg (userid, courseid)");

    // Quiz
    $pdo->exec("DROP TABLE IF EXISTS temp_quiz_agg");
    $quiz_agg_sql = "
        CREATE TEMP TABLE temp_quiz_agg AS
        SELECT 
            qa.userid,
            q.course AS courseid,
            q.name AS quizname,
            COUNT(DISTINCT qa.id) AS num_attempts,
            COALESCE(gg.finalgrade, 0) AS nota_final,
            CASE WHEN q.timeopen > 0 THEN to_timestamp(q.timeopen)::date ELSE NULL END AS fecha_apertura,
            CASE WHEN q.timeclose > 0 THEN to_timestamp(q.timeclose)::date ELSE NULL END AS fecha_cierre,
            CASE WHEN gg.feedback IS NOT NULL THEN 'SÍ' ELSE 'NO' END AS retroalimenta,
            ROW_NUMBER() OVER (PARTITION BY qa.userid, q.course ORDER BY q.name) AS quiz_number
        FROM mdl_quiz_attempts qa
        JOIN mdl_quiz q ON qa.quiz = q.id
        LEFT JOIN mdl_grade_grades gg ON gg.userid = qa.userid 
            AND gg.itemid IN (SELECT id FROM mdl_grade_items WHERE itemmodule = 'quiz' AND iteminstance = q.id)
        WHERE qa.timefinish BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) 
                               AND EXTRACT(EPOCH FROM :fecha_fin::timestamp)
          AND q.course IN ('428')
        GROUP BY qa.userid, q.course, q.name, gg.finalgrade, q.timeopen, q.timeclose, gg.feedback";
    $stmt = $pdo->prepare($quiz_agg_sql);
    $stmt->execute($params);
    $pdo->exec("CREATE INDEX ON temp_quiz_agg (userid, courseid)");

    // Glosario
    $pdo->exec("DROP TABLE IF EXISTS temp_glossary_agg");
    $glossary_agg_sql = "
        CREATE TEMP TABLE temp_glossary_agg AS
        SELECT 
            ge.userid,
            g.course AS courseid,
            g.name AS glossaryname,
            COUNT(DISTINCT ge.id) AS num_entries,
            COALESCE(gg.finalgrade, 0) AS nota_final,
            regexp_replace(g.intro, '.*Fecha de apertura:</td>\s*<p>([^<]*)</p>.*', '\\1', 'g') AS fecha_apertura,
            regexp_replace(g.intro, '.*Fecha de cierre:</td>\s*<p>([^<]*)</p>.*', '\\1', 'g') AS fecha_cierre,
            CASE WHEN COUNT(DISTINCT CASE WHEN ra2.roleid IN (3, 4) AND ge2.userid != ge.userid THEN ge2.id END) > 0 THEN 'SÍ' ELSE 'NO' END AS retroalimenta,
            ROW_NUMBER() OVER (PARTITION BY ge.userid, g.course ORDER BY g.name) AS glossary_number
        FROM mdl_glossary_entries ge
        JOIN mdl_glossary g ON ge.glossaryid = g.id
        LEFT JOIN mdl_glossary_entries ge2 ON ge2.glossaryid = g.id
        LEFT JOIN mdl_user u2 ON ge2.userid = u2.id
        LEFT JOIN mdl_role_assignments ra2 ON u2.id = ra2.userid AND ra2.contextid IN (
            SELECT contextid FROM mdl_context WHERE instanceid = g.course AND contextlevel = 50
        )
        LEFT JOIN mdl_grade_grades gg ON gg.userid = ge.userid 
            AND gg.itemid IN (SELECT id FROM mdl_grade_items WHERE itemmodule = 'glossary' AND iteminstance = g.id)
        WHERE ge.timecreated BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) 
                                AND EXTRACT(EPOCH FROM :fecha_fin::timestamp)
          AND g.course IN ('428')
        GROUP BY ge.userid, g.course, g.name, gg.finalgrade, g.intro";
    $stmt = $pdo->prepare($glossary_agg_sql);
    $stmt->execute($params);
    $pdo->exec("CREATE INDEX ON temp_glossary_agg (userid, courseid)");

    // SCORM
    $pdo->exec("DROP TABLE IF EXISTS temp_scorm_agg");
    $scorm_agg_sql = "
        CREATE TEMP TABLE temp_scorm_agg AS
        SELECT 
            st.userid,
            s.course AS courseid,
            s.name AS scormname,
            COUNT(DISTINCT st.id) AS num_attempts,
            COALESCE(gg.finalgrade, 0) AS nota_final,
            CASE WHEN s.timeopen > 0 THEN to_timestamp(s.timeopen)::date ELSE NULL END AS fecha_apertura,
            CASE WHEN s.timeclose > 0 THEN to_timestamp(s.timeclose)::date ELSE NULL END AS fecha_cierre,
            CASE WHEN COUNT(DISTINCT CASE WHEN st2.element LIKE '%comment%' AND ra2.roleid IN (3, 4) THEN st2.id END) > 0 
                 OR gg.feedback IS NOT NULL THEN 'SÍ' ELSE 'NO' END AS retroalimenta,
            ROW_NUMBER() OVER (PARTITION BY st.userid, s.course ORDER BY s.name) AS scorm_number
        FROM mdl_scorm_scoes_track st
        JOIN mdl_scorm s ON st.scormid = s.id
        LEFT JOIN mdl_scorm_scoes_track st2 ON st2.scormid = s.id
        LEFT JOIN mdl_user u2 ON st2.userid = u2.id
        LEFT JOIN mdl_role_assignments ra2 ON u2.id = ra2.userid AND ra2.contextid IN (
            SELECT contextid FROM mdl_context WHERE instanceid = s.course AND contextlevel = 50
        )
        LEFT JOIN mdl_grade_grades gg ON gg.userid = st.userid 
            AND gg.itemid IN (SELECT id FROM mdl_grade_items WHERE itemmodule = 'scorm' AND iteminstance = s.id)
        WHERE st.timemodified BETWEEN EXTRACT(EPOCH FROM :fecha_inicio::timestamp) 
                                 AND EXTRACT(EPOCH FROM :fecha_fin::timestamp)
          AND s.course IN ('428')
        GROUP BY st.userid, s.course, s.name, gg.finalgrade, s.timeopen, s.timeclose, gg.feedback";
    $stmt = $pdo->prepare($scorm_agg_sql);
    $stmt->execute($params);
    $pdo->exec("CREATE INDEX ON temp_scorm_agg (userid, courseid)");

    // Tabla temporal para logs de vistas
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
          AND log.courseid IN ('428')
          AND log.component IN ('mod_forum', 'mod_lesson', 'mod_assign', 'mod_quiz', 'mod_glossary', 'mod_scorm')
        GROUP BY log.userid, log.courseid, log.component, log.target, log.objectid";
    $stmt = $pdo->prepare($log_agg_sql);
    $stmt->execute($params);
    $pdo->exec("CREATE INDEX ON temp_log_agg (userid, courseid, component, target, objectid)");

    // Consulta principal con información de usuarios y cursos
    $sql = "WITH UserInfo AS (
        SELECT 
            uid.userid,
            MAX(CASE WHEN ufield.shortname = 'programa' THEN uid.data END) AS idprograma,
            MAX(CASE WHEN ufield.shortname = 'facultad' THEN uid.data END) AS idfacultad,
            MAX(CASE WHEN ufield.shortname = 'edad' THEN uid.data END) AS edad,
            MAX(CASE WHEN ufield.shortname = 'genero' THEN uid.data END) AS genero,
            MAX(CASE WHEN ufield.shortname = 'celular' THEN uid.data END) AS celular,
            MAX(CASE WHEN ufield.shortname = 'jornada' THEN uid.data END) AS jornada,
            MAX(CASE WHEN ufield.shortname = 'estrato' THEN uid.data END) AS estrato
        FROM mdl_user_info_data uid
        JOIN mdl_user_info_field ufield ON ufield.id = uid.fieldid
        WHERE ufield.shortname IN ('programa', 'facultad', 'edad', 'genero', 'celular', 'jornada', 'estrato')
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
        r.name AS rol,
        u.email AS correo,
        c.id AS id_curso,
        c.fullname AS curso,
        ci.idcodigo AS codigo_curso,
        ui.idprograma,
        u.institution, 
        ui.idfacultad,
        u.department, 
        ui.jornada,
        ci.grupo,
        ci.periodo,
        ci.nivel,
        ui.edad,
        ui.genero,
        ui.celular,
        ui.estrato,";

    // Columnas dinámicas para foros
    for ($i = 1; $i <= $max_forums; $i++) {
        $sql .= "
        MAX(CASE WHEN fi.forum_number = $i THEN fi.fecha_apertura END) AS fecha_apertura_forum_$i,
        MAX(CASE WHEN fi.forum_number = $i THEN fi.fecha_cierre END) AS fecha_cierre_forum_$i,
        MAX(CASE WHEN fi.forum_number = $i THEN GREATEST(COALESCE(flog.num_views, 0), COALESCE(fi.num_posts, 0)) END) AS visitas_forum_$i,
        MAX(CASE WHEN fi.forum_number = $i THEN fi.num_posts END) AS numero_envios_$i,
        MAX(CASE WHEN fi.forum_number = $i THEN fi.nota_final END) AS nota_final_forum_$i,
        MAX(CASE WHEN fi.forum_number = $i THEN fi.retroalimenta END) AS retroalimenta_forum_$i,";
    }

    // Columnas dinámicas para lecciones
    for ($i = 1; $i <= $max_lessons; $i++) {
        $sql .= "
        MAX(CASE WHEN li.lesson_number = $i THEN li.fecha_apertura END) AS fecha_apertura_lesson_$i,
        MAX(CASE WHEN li.lesson_number = $i THEN li.fecha_cierre END) AS fecha_cierre_lesson_$i,
        MAX(CASE WHEN li.lesson_number = $i THEN GREATEST(COALESCE(llog.num_views, 0), COALESCE(li.num_attempts, 0)) END) AS visitas_lesson_$i,
        MAX(CASE WHEN li.lesson_number = $i THEN li.num_attempts END) AS numero_intentos_$i,
        MAX(CASE WHEN li.lesson_number = $i THEN li.nota_final END) AS nota_final_lesson_$i,
        MAX(CASE WHEN li.lesson_number = $i THEN li.retroalimenta END) AS retroalimenta_lesson_$i,";
    }

    // Columnas dinámicas para tareas
    for ($i = 1; $i <= $max_assigns; $i++) {
        $sql .= "
        MAX(CASE WHEN ai.assign_number = $i THEN ai.fecha_apertura END) AS fecha_apertura_assign_$i,
        MAX(CASE WHEN ai.assign_number = $i THEN ai.fecha_cierre END) AS fecha_cierre_assign_$i,
        MAX(CASE WHEN ai.assign_number = $i THEN GREATEST(COALESCE(alog.num_views, 0), COALESCE(ai.num_submissions, 0)) END) AS visitas_assign_$i,
        MAX(CASE WHEN ai.assign_number = $i THEN ai.num_submissions END) AS numero_envios_$i,
        MAX(CASE WHEN ai.assign_number = $i THEN ai.nota_final END) AS nota_final_assign_$i,
        MAX(CASE WHEN ai.assign_number = $i THEN ai.retroalimenta END) AS retroalimenta_assign_$i,";
    }

    // Columnas dinámicas para quizzes
    for ($i = 1; $i <= $max_quizzes; $i++) {
        $sql .= "
        MAX(CASE WHEN qi.quiz_number = $i THEN qi.fecha_apertura END) AS fecha_apertura_quiz_$i,
        MAX(CASE WHEN qi.quiz_number = $i THEN qi.fecha_cierre END) AS fecha_cierre_quiz_$i,
        MAX(CASE WHEN qi.quiz_number = $i THEN GREATEST(COALESCE(qlog.num_views, 0), COALESCE(qi.num_attempts, 0)) END) AS visitas_quiz_$i,
        MAX(CASE WHEN qi.quiz_number = $i THEN qi.num_attempts END) AS numero_intentos_$i,
        MAX(CASE WHEN qi.quiz_number = $i THEN qi.nota_final END) AS nota_final_quiz_$i,
        MAX(CASE WHEN qi.quiz_number = $i THEN qi.retroalimenta END) AS retroalimenta_quiz_$i,";
    }

    // Columnas dinámicas para glosarios
    for ($i = 1; $i <= $max_glossaries; $i++) {
        $sql .= "
        MAX(CASE WHEN gi.glossary_number = $i THEN gi.fecha_apertura END) AS fecha_apertura_glossary_$i,
        MAX(CASE WHEN gi.glossary_number = $i THEN gi.fecha_cierre END) AS fecha_cierre_glossary_$i,
        MAX(CASE WHEN gi.glossary_number = $i THEN GREATEST(COALESCE(glog.num_views, 0), COALESCE(gi.num_entries, 0)) END) AS visitas_glossary_$i,
        MAX(CASE WHEN gi.glossary_number = $i THEN gi.num_entries END) AS numero_entradas_$i,
        MAX(CASE WHEN gi.glossary_number = $i THEN gi.nota_final END) AS nota_final_glossary_$i,
        MAX(CASE WHEN gi.glossary_number = $i THEN gi.retroalimenta END) AS retroalimenta_glossary_$i,";
    }

    // Columnas dinámicas para SCORM
    for ($i = 1; $i <= $max_scorms; $i++) {
        $sql .= "
        MAX(CASE WHEN si.scorm_number = $i THEN si.fecha_apertura END) AS fecha_apertura_scorm_$i,
        MAX(CASE WHEN si.scorm_number = $i THEN si.fecha_cierre END) AS fecha_cierre_scorm_$i,
        MAX(CASE WHEN si.scorm_number = $i THEN GREATEST(COALESCE(slog.num_views, 0), COALESCE(si.num_attempts, 0)) END) AS visitas_scorm_$i,
        MAX(CASE WHEN si.scorm_number = $i THEN si.num_attempts END) AS numero_intentos_$i,
        MAX(CASE WHEN si.scorm_number = $i THEN si.nota_final END) AS nota_final_scorm_$i,
        MAX(CASE WHEN si.scorm_number = $i THEN si.retroalimenta END) AS retroalimenta_scorm_$i,";
    }

    // Completar la consulta principal
    $sql = rtrim($sql, ',') . "
    FROM mdl_user u
    JOIN mdl_role_assignments ra ON u.id = ra.userid
    JOIN mdl_role r ON ra.roleid = r.id
    JOIN mdl_context ctx ON ra.contextid = ctx.id AND ctx.contextlevel = 50
    JOIN mdl_course c ON ctx.instanceid = c.id
    LEFT JOIN UserInfo ui ON u.id = ui.userid
    LEFT JOIN CourseInfo ci ON c.id = ci.instanceid
    LEFT JOIN temp_forum_agg fi ON u.id = fi.userid AND c.id = fi.courseid
    LEFT JOIN temp_lesson_agg li ON u.id = li.userid AND c.id = li.courseid
    LEFT JOIN temp_assign_agg ai ON u.id = ai.userid AND c.id = ai.courseid
    LEFT JOIN temp_quiz_agg qi ON u.id = qi.userid AND c.id = qi.courseid
    LEFT JOIN temp_glossary_agg gi ON u.id = gi.userid AND c.id = gi.courseid
    LEFT JOIN temp_scorm_agg si ON u.id = si.userid AND c.id = si.courseid
    LEFT JOIN temp_log_agg flog ON u.id = flog.userid AND c.id = flog.courseid AND flog.component = 'mod_forum' AND flog.objectid = fi.forum_number
    LEFT JOIN temp_log_agg llog ON u.id = llog.userid AND c.id = llog.courseid AND llog.component = 'mod_lesson' AND llog.objectid = li.lesson_number
    LEFT JOIN temp_log_agg alog ON u.id = alog.userid AND c.id = alog.courseid AND alog.component = 'mod_assign' AND alog.objectid = ai.assign_number
    LEFT JOIN temp_log_agg qlog ON u.id = qlog.userid AND c.id = qlog.courseid AND qlog.component = 'mod_quiz' AND qlog.objectid = qi.quiz_number
    LEFT JOIN temp_log_agg glog ON u.id = glog.userid AND c.id = glog.courseid AND glog.component = 'mod_glossary' AND glog.objectid = gi.glossary_number
    LEFT JOIN temp_log_agg slog ON u.id = slog.userid AND c.id = slog.courseid AND slog.component = 'mod_scorm' AND slog.objectid = si.scorm_number
    WHERE c.id IN ('428')
    GROUP BY u.username, u.firstname, u.lastname, r.name, u.email, c.id, c.fullname, ci.idcodigo, ui.idprograma, u.institution, 
             ui.idfacultad, u.department, ui.jornada, ci.grupo, ci.periodo, ci.nivel, ui.edad, ui.genero, ui.celular, ui.estrato";

    // Ejecutar la consulta
    $stmt = $pdo->prepare($sql);
    $stmt->execute();
    $results = $stmt->fetchAll(PDO::FETCH_ASSOC);

    // Crear el archivo Excel en memoria
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();

    // Generar encabezados personalizados
    $headers = [
        'Código', 'Nombre', 'Apellidos', 'Rol', 'Correo', 'ID Curso', 'Curso', 'Código Curso', 'Programa', 'Institución', 
        'Facultad', 'Departamento', 'Jornada', 'Grupo', 'Periodo', 'Nivel', 'Edad', 'Género', 'Celular', 'Estrato'
    ];

    // Añadir encabezados dinámicos para cada tipo de actividad en el orden deseado
    for ($i = 1; $i <= $max_forums; $i++) {
        $name = $forum_names[$i] ?? "Foro $i";
        $headers[] = "Fecha_apertura: $name";
        $headers[] = "Fecha_cierre: $name";
        $headers[] = "Visitas: $name";
        $headers[] = "Numero_envios: $name";
        $headers[] = "Nota: $name";
        $headers[] = "Retroalimenta: $name";
    }
    for ($i = 1; $i <= $max_lessons; $i++) {
        $name = $lesson_names[$i] ?? "Lección $i";
        $headers[] = "Fecha_apertura: $name";
        $headers[] = "Fecha_cierre: $name";
        $headers[] = "Visitas: $name";
        $headers[] = "Numero_intentos: $name";
        $headers[] = "Nota: $name";
        $headers[] = "Retroalimenta: $name";
    }
    for ($i = 1; $i <= $max_assigns; $i++) {
        $name = $assign_names[$i] ?? "Tarea $i";
        $headers[] = "Fecha_apertura: $name";
        $headers[] = "Fecha_cierre: $name";
        $headers[] = "Visitas: $name";
        $headers[] = "Numero_envios: $name";
        $headers[] = "Nota: $name";
        $headers[] = "Retroalimenta: $name";
    }
    for ($i = 1; $i <= $max_quizzes; $i++) {
        $name = $quiz_names[$i] ?? "Quiz $i";
        $headers[] = "Fecha_apertura: $name";
        $headers[] = "Fecha_cierre: $name";
        $headers[] = "Visitas: $name";
        $headers[] = "Numero_intentos: $name";
        $headers[] = "Nota: $name";
        $headers[] = "Retroalimenta: $name";
    }
    for ($i = 1; $i <= $max_glossaries; $i++) {
        $name = $glossary_names[$i] ?? "Glosario $i";
        $headers[] = "Fecha_apertura: $name";
        $headers[] = "Fecha_cierre: $name";
        $headers[] = "Visitas: $name";
        $headers[] = "Numero_entradas: $name";
        $headers[] = "Nota: $name";
        $headers[] = "Retroalimenta: $name";
    }
    for ($i = 1; $i <= $max_scorms; $i++) {
        $name = $scorm_names[$i] ?? "SCORM $i";
        $headers[] = "Fecha_apertura: $name";
        $headers[] = "Fecha_cierre: $name";
        $headers[] = "Visitas: $name";
        $headers[] = "Numero_intentos: $name";
        $headers[] = "Nota: $name";
        $headers[] = "Retroalimenta: $name";
    }

    // Escribir encabezados en la fila 1
    $sheet->fromArray($headers, NULL, 'A1');

    // Escribir datos a partir de la fila 2
    $sheet->fromArray($results, NULL, 'A2');

    // Generar el archivo Excel en un stream temporal
    $writer = new Xlsx($spreadsheet);
    $tempFilePath = tempnam(sys_get_temp_dir(), 'reporte_moodle_');
    $writer->save($tempFilePath);

    // Enviar el archivo adjunto por correo
    $mail = new PHPMailer(true);
    $mail->setFrom('noreply@example.com', 'Reporte Moodle');
    foreach ($correo_destino as $correo) {
        $mail->addAddress($correo);
    }
    $mail->Subject = 'Reporte de interacciones en foros Moodle';
    $mail->Body = 'Cordial Saludo, Adjunto el Reporte de interacciones en foros de Moodle.';
    $mail->addAttachment($tempFilePath, 'reporte.xlsx');

    // Enviar correo de notificación
    $mail_notificacion = new PHPMailer(true);
    $mail_notificacion->setFrom('noreply@example.com', 'Notificación Reporte Moodle');
    $mail_notificacion->addAddress($correo_notificacion);

    try {
        $mail->send();
        $mail_notificacion->Subject = 'Éxito: Reporte de Moodle enviado';
        $mail_notificacion->Body = 'El reporte de interacciones en foros de Moodle fue enviado exitosamente a ' . implode(', ', $correo_destino) . ' el ' . date('Y-m-d H:i:s') . '.';
        $mail_notificacion->send();
        echo "Reporte generado y enviado correctamente.";
    } catch (Exception $e) {
        $mail_notificacion->Subject = 'Fallo: Error al enviar reporte de Moodle';
        $mail_notificacion->Body = 'Hubo un error al enviar el reporte de interacciones en foros de Moodle: ' . $e->getMessage();
        $mail_notificacion->send();
        echo "Error al enviar el reporte: " . $e->getMessage();
    }

    // Eliminar el archivo temporal
    unlink($tempFilePath);

} catch (Exception $e) {
    echo "Error: " . $e->getMessage();
}
?>
