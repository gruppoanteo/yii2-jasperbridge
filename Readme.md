## yii2-jasperbridge

Yii2 extension to generate JasperReports from PHP via PHP/Java Bridge. Provided by Marco Curatitoli at HalService. Mantained by Pietro Bardone at Anteo Impresa Sociale.

Compile/fill `.jrxml`/`.jasper` templates and export to multiple formats (PDF, HTML, XML, XLS/XLSX, DOCX, PPTX, CSV). Use events to customize queries, data sources, and report structure through Jasper APIs.

### Requirements
- **Yii2** (`yiisoft/yii2`)
- Active **PHP/Java Bridge** (native extension or reachable Java servlet)
- **Java** (JRE) on the server
- **JasperReports** JARs and **JDBC driver** JARs for your DB (must be in the Java bridge classpath)

### Installation
```bash
composer require anteo/yii2-jasperbridge
```

You don't need to register a Yii component; instantiate `anteo\jasperbridge\Report` where you need it.

### PHP/Java Bridge configuration
The extension automatically tries these fallbacks to load `Java.inc`:
1) dynamic `java.so`/`php_java.dll`
2) local `Java.inc` file (bundled here)
3) servlet at `http://127.0.0.1:{port}/java/Java.inc`

If using the servlet, ensure the bridge is reachable and set the port if needed:

```php
$report = new \anteo\jasperbridge\Report('@app/reports/sample.jasper');
$report->javaBridgePort = '8080'; // default 8080
```

For JDBC, the driver JARs (e.g., MySQL/PostgreSQL) and JasperReports must be in the JVM classpath where the bridge runs.

### Quick start
```php
use anteo\jasperbridge\Report;

// .jasper (precompiled) or .jrxml (compiled on the fly)
$report = new Report('@app/reports/invoice.jasper', [
    'id' => 123,
]);

// Output format (PDF by default) and optional exporter params
$report->setOutput(Report::OUTPUT_PDF);

// Data source: by default uses your JDBC connection (see JDBC section)
// Generate and send inline to the browser
return $report->render('', Report::DESTINATION_INLINE);
```

### Available output formats
- `Report::OUTPUT_PDF` (pdf)
- `Report::OUTPUT_HTML` (html)
- `Report::OUTPUT_XML` (xml)
- `Report::OUTPUT_XLS` (xls)
- `Report::OUTPUT_XLS_JAPI` (xls via legacy JRXlsExporter)
- `Report::OUTPUT_XLSX` (xlsx)
- `Report::OUTPUT_DOCX` (docx)
- `Report::OUTPUT_PPTX` (pptx)
- `Report::OUTPUT_CSV` (csv)

Output destinations:
- `Report::DESTINATION_INLINE` sends to browser (inline)
- `Report::DESTINATION_DOWNLOAD` forces download
- `Report::DESTINATION_FILE` saves to file (returns the path)

```php
// XLSX download with exporter params
$report->setOutput(Report::OUTPUT_XLSX, [
    // Example: map JR constants as key => value
    // 'IS_ONE_PAGE_PER_SHEET' => true,
]);
return $report->render('', Report::DESTINATION_DOWNLOAD);
```

### JDBC from Yii2 DB
The report can create a JDBC connection from your `Yii::$app->db`.

```php
$report = new Report('@app/reports/sales.jasper');
$report->db = \Yii::$app->db; // supported driverName: mysql, pgsql
return $report->render('', Report::DESTINATION_INLINE);
```

Notes:
- The Yii2 DSN is translated into a JDBC URL and loads the appropriate driver (e.g., `com.mysql.jdbc.Driver`, `org.postgresql.Driver`).
- Ensure driver JARs are in the JVM classpath.

### Events and customization
Hook into various points of the report lifecycle.

- **`Report::EVENT_LOAD`**: access `JasperDesign` (for `.jrxml`) or the report object for structural changes and queries.
- **`Report::EVENT_FILL`**: provide a custom `JR*DataSource` (ResultSet, JSON, etc.).

Examples:

```php
use anteo\jasperbridge\Report;

$report = new Report('@app/reports/credentials.jasper', ['id' => $id]);

// Modify structure (JasperDesign) or query (jrxml only)
$report->on(Report::EVENT_LOAD, function($event) {
    $jasperDesign = $event->reportObject; // JRDesign
    // Example: set a dynamic query
    $sql = 'SELECT * FROM t WHERE id = $P{id}';
    $jrDesignQuery = new \Java('net.sf.jasperreports.engine.design.JRDesignQuery');
    $jrDesignQuery->setText($sql);
    $jasperDesign->setQuery($jrDesignQuery);
});

// JRResultSetDataSource without parameters
$report->on(Report::EVENT_FILL, function($event) {
    $report = $event->sender;
    $conn = $report->jdbcConnection; // java.sql.Connection
    $rs = $conn->createStatement()->executeQuery('SELECT 1 as v');
    $ds = new \Java('net.sf.jasperreports.engine.JRResultSetDataSource', $rs);
    $fill = $report->getFillManager();
    $report->jasperPrint = $fill->fillReport($event->reportObject, $report->reportParamsToHashMap(), $ds);
});

// JsonDataSource
$json = json_encode(['items' => [['name' => 'Foo']]]);
$report->on(Report::EVENT_FILL, function($event) use ($json) {
    $stream = new \Java('java.io.ByteArrayInputStream', $json);
    $ds = new \Java('net.sf.jasperreports.engine.data.JsonDataSource', $stream);
    $fill = $event->sender->getFillManager();
    $event->sender->jasperPrint = $fill->fillReport($event->reportObject, $event->sender->reportParamsToHashMap(), $ds);
});

return $report->render('', Report::DESTINATION_INLINE);
```

### Report parameters
Pass parameters as an associative array in the constructor or by assigning `$report->reportParams`. For dates, consider `convertValue('YYYY-MM-DD', 'java.util.Date')` to type them correctly in Java.

### `.jrxml` compilation
If you pass a `.jrxml`, the library detects it is not compiled and compiles it before filling.

### Troubleshooting
- Error loading `Java.inc`: ensure PHP/Java Bridge is active and `$report->javaBridgePort` matches the servlet.
- JDBC driver not found: add the driver JAR to the JVM classpath of the bridge.
- Headers already sent: use `DESTINATION_FILE` or remove previous output before `render()`.
- Accented characters/encoding: set proper fonts and encoding in the Jasper template and, if needed, exporter parameters.

### License
BSD 3-Clause (see `composer.json`).

### Author
`p3pp01` — pietro.bardone@gruppoanteo.it — https://anteocoop.it


