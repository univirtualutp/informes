<?php
// Configuración exacta de tu otro servidor
$config = [
    'host'         => 'clusteroracle.utp.edu.co',
    'port'         => '1452',
    'service_name' => 'PRODUCT',
    'database'     => 'PRODUCT',
    'username'     => 'CONSULTA_UNIVIRTUAL',
    'password'     => 'x10v$rta!5key',
    'charset'      => 'AL32UTF8',
    'view_name'    => 'REGISTRO.VI_RYC_UNIVIRTUALASIGNCANCELCO' // Vista que necesitas consultar
];

// Formato de conexión profesional (2 intentos)
$connection_methods = [
    "Método SERVICE_NAME" => "//{$config['host']}:{$config['port']}/{$config['service_name']}",
    "Método Descriptor" => "(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST={$config['host']})(PORT={$config['port']}))(CONNECT_DATA=(SERVICE_NAME={$config['service_name']})))"
];

foreach ($connection_methods as $method_name => $conn_string) {
    echo "<h3>Probando: $method_name</h3>";
    echo "<pre>Cadena: " . htmlspecialchars($conn_string) . "</pre>";
    
    $conn = @oci_connect(
        $config['username'],
        $config['password'],
        $conn_string,
        $config['charset']
    );
    
    if ($conn) {
        echo "<div style='color:green; padding:10px; background:#f0fff0;'>✅ ¡Conexión exitosa!</div>";
        
        // Consulta a tu vista específica
        $query = "SELECT COUNT(*) {$config['view_name']} ";
        $stid = oci_parse($conn, $query);
        
        if (oci_execute($stid)) {
            echo "<h4>Primeros 5 registros de {$config['view_name']}:</h4>";
            echo "<table border='1' cellpadding='5'><tr>";
            
            // Encabezados
            for ($i = 1; $i <= oci_num_fields($stid); $i++) {
                echo "<th>".oci_field_name($stid, $i)."</th>";
            }
            echo "</tr>";
            
            // Datos
            while ($row = oci_fetch_array($stid, OCI_ASSOC+OCI_RETURN_NULLS)) {
                echo "<tr>";
                foreach ($row as $item) {
                    echo "<td>".($item !== null ? htmlentities($item) : "NULL")."</td>";
                }
                echo "</tr>";
            }
            echo "</table>";
        } else {
            $error = oci_error($stid);
            echo "<div style='color:orange; background:#fffaf0; padding:10px;'>⚠️ Error en consulta: ORA-".$error['code'].": ".$error['message']."</div>";
        }
        
        oci_free_statement($stid);
        oci_close($conn);
        break; // Si una conexión funciona, detenemos las pruebas
        
    } else {
        $error = oci_error();
        echo "<div style='color:red; background:#fff0f0; padding:10px;'>❌ Falló: ORA-".$error['code'].": ".$error['message']."</div>";
    }
    echo "<hr>";
}

// Si ambos métodos fallan
if (!$conn) {
    echo "<h3 style='color:darkred;'>🚨 Solución requerida:</h3>";
    echo "<ol>
    <li>Verifica con el DBA que el SERVICE_NAME '{$config['service_name']}' sea correcto</li>
    <li>Confirma que el puerto {$config['port']} esté abierto en el firewall</li>
    <li>Pide el archivo <code>tnsnames.ora</code> para usar un alias TNS</li>
    </ol>";
}
?>
