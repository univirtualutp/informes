<?php
// Configuración de la conexión
$db_user = 'CONSULTA_UNIVIRTUAL';
$db_pass = 'x10v$rta!5key';
$db_conn_string = '(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=clusteroracle.utp.edu.co)(PORT=1452))(CONNECT_DATA=(SERVICE_NAME=PRODUCT)))';

// Conectar a la base de datos
$conn = oci_connect($db_user, $db_pass, $db_conn_string, 'AL32UTF8');

if (!$conn) {
    $e = oci_error();
    die("Error de conexión: " . $e['message']);
}

// Consulta SQL
$query = "SELECT NUMERODOCUMENTO, INITCAP(NOMBRES), INITCAP(APELLIDOS), EMAIL, CODIGOASIGNATURA, NOMBREASIGNATURA, NUMEROGRUPO, IDGRUPO, IDESTRUCFACULTAD, FACULTAD, CODIGOPROGRAMA, PROGRAMA, JORNADA, FECHANACIMIENTO, SEXO, TELEFONO, ESTRATOSOCIAL
          FROM REGISTRO.VI_RYC_UNIVIRTUALESTUDMATRIC
          WHERE regexp_like(PERIODOACADEMICO, '^20252(.*)[Pp][Rr][Ee][Gg][Rr][Aa][Dd][Oo].*')
          AND CODIGOPROGRAMA NOT IN ('TR')
          ORDER BY NOMBREASIGNATURA, NUMEROGRUPO, INITCAP(NOMBRES) ASC";

// Preparar y ejecutar la consulta
$stid = oci_parse($conn, $query);
if (!$stid) {
    $e = oci_error($conn);
    die("Error al parsear la consulta: " . $e['message']);
}

if (!oci_execute($stid)) {
    $e = oci_error($stid);
    die("Error al ejecutar la consulta: " . $e['message']);
}

// Contar los registros
$count = 0;
while (oci_fetch($stid)) {
    $count++;
}

echo "Número de registros encontrados: $count\n";

// Liberar recursos y cerrar conexión
oci_free_statement($stid);
oci_close($conn);
?>
