<?php

namespace anteo\jasperbridge;

use DateTime;
use Java;
use JavaClass;
use Yii;
use yii\base\Component;
use yii\base\Exception;
use yii\base\InvalidParamException;
use yii\base\NotSupportedException;
use yii\db\Connection;
use yii\di\Instance;

/**
 * Documentazione API e parametri per esportazione
 * @property-read-only isReportFileCompiled return if report has to be compiled (jrxml or jasper file)
 * @property-read-only fillManager
 * @property-read-only jdbcConnection
 * @property jasperPrint
 * 
 *  http://jasperreports.sourceforge.net/api/index-all.html
 *
 * protected function actionPrint($id)
 * {
 *     $report = new \anteo\jasperbridge\Report('@app/reports/credentials.jasper', [
 *          'id' => $id,
 *     ]);
 * 
 *     // for date use $report->convertValue('2014-05-01', 'java.util.Date')
 * 
 *     // To modify report structure:
 *     $report->on(Report::EVENT_LOAD, function($event) {
 *         $JRBands = $event->reportObject->getDetailSection()->getBands();
 *         $element = $JRBands[0]->getElementByKey('key');
 *         $element->setX(50);
 *     });
 * 
 *     // To set a full text query YOU NEED to use jrxml report file
 *     $report->on(Report::EVENT_LOAD, function($event) {
 *         $jasperDesign = $event->reportObject;
 *         $query = $event->data;
 *         $jrDesignQuery = new \Java("net.sf.jasperreports.engine.design.JRDesignQuery");
 *         $jrDesignQuery->setText($sql);
 *         $jasperDesign->setQuery($jrDesignQuery);
 *     });
 * 
 *     // To pass a JR*Source* to fill: in this example a JRResultSetDataSource WITHOUT parameters
 *     $report->on(Report::EVENT_FILL, function($event) {
 *          $report = $event->sender;
 *          $reportObject = $event->reportObject;
 *          $query = $event->data;
 *      
 *          $connection = $report->jdbcConnection;
 *          $statement = $connection->createStatement();
 *          $resultSet = $statement->executeQuery($query);
 *       
 *          $fillManager = $report->getFillManager();
 *          $JRResultSetDataSource = new \Java('net.sf.jasperreports.engine.JRResultSetDataSource', $resultSet);
 *          $report->jasperPrint = $fillManager->fillReport($reportObject, $report->reportParamsToHashMap(), $JRResultSetDataSource);
 *     }, $query);
 * 
 *     // To pass a JR*Source* to fill: in this example a JsonDataSource
 *     $report->on(Report::EVENT_FILL, function($event) {
 *          $report = $event->sender;
 *          $json = $event->data;
 *          $reportObject = $event->reportObject;
 *          $stream = new Java('java.io.ByteArrayInputStream', $json);
 *          $JsonDataSource = new Java('net.sf.jasperreports.engine.data.JsonDataSource', $stream);
 *          $fillManager = $report->getFillManager();
 *          $report->jasperPrint = $fillManager->fillReport($reportObject, $report->reportParamsToHashMap(), $JsonDataSource);
 *      }, $json);
 * 
 *     // To pass a JR*Source* to fill: in this example a JRResultSetDataSource WITH parameters
 *     $report->on(Report::EVENT_FILL, function($event) {
 *          $report = $event->sender;
 *          $reportObject = $event->reportObject;
 *          $query = $event->data;
 *      
 *          $connection = $report->jdbcConnection;
 *          // with parameters
 *          $sql = $query->createCommand()->sql;
 *              foreach($query->params as $k => $v) {
 *              $sql = str_replace($k, '?', $sql);
 *          }
 *          $preparedStatement = $connection->prepareStatement($sql);
 *          $i = 0;
 *          foreach($query->params as $k => $v) {
 *              $i++;
 *              switch ($k) {
 *                  case ':date_from':
 *                  case ':date_to':
 *                      $preparedStatement->setDate($i, $report->convertValue($v, 'java.sql.Date'));
 *                      break;
 *                  case ':field':
 *                      $preparedStatement->setInt($i, $v);
 *                      break;
 *                 ........
 *              }
 *          }
 *          $resultSet = $preparedStatement->executeQuery();
 * 
 *          $fillManager = $report->getFillManager();
 *          $JRResultSetDataSource = new \Java('net.sf.jasperreports.engine.JRResultSetDataSource', $resultSet);
 *          $report->jasperPrint = $fillManager->fillReport($reportObject, $report->reportParamsToHashMap(), $JRResultSetDataSource);
 *     }, $query);
 *      
 *     return $report->execute();
 * }
 */
class Report extends Component
{
    const EVENT_LOAD = 'reportLoad';
    const EVENT_COMPILE = 'reportCompile';
    const EVENT_FILL = 'reportFill';
    const EVENT_OUTPUT = 'reportOutput';
    const OUTPUT_PDF = 1;
    const OUTPUT_XLS = 2;
    const OUTPUT_XLSX = 3;
    const OUTPUT_XLS_JAPI = 4;
    const OUTPUT_DOCX = 5;
    const OUTPUT_HTML = 6;
//	const ODF=7;
//	const ODS=8;
    const OUTPUT_PPTX = 9;
//	const TXT=10;
    const OUTPUT_CSV = 11;
    const OUTPUT_XML = 12;
    const DESTINATION_FILE = 'F';
    const DESTINATION_INLINE = 'I';
    const DESTINATION_DOWNLOAD = 'D';
    const DESTINATION_STREAM = 'S';

    private static $_extensions = [
        self::OUTPUT_PDF => 'pdf',
        self::OUTPUT_XLS => 'xls',
        self::OUTPUT_XLS_JAPI => 'xls',
        self::OUTPUT_XLSX => 'xslx',
        self::OUTPUT_DOCX => 'docx',
        self::OUTPUT_HTML => 'html',
//		self::ODF=>'odf',
//		self::ODS=>'ods',
        self::OUTPUT_PPTX => 'pptx',
//		self::TXT=>'txt',
        self::OUTPUT_CSV => 'csv',
        self::OUTPUT_XML => 'xml',
    ];
    /**
     * @var Connection the database connection
     */
    public $db = 'db';
    /**
     * @var string database port for createJdbcDriver ()
     */
    public $dbPort = '5432';
    /**
     * @var string java bridge listening port (default 8080)
     */
    public $javaBridgePort = '8080';
    /**
     * @var string full path to compiled jasper report.jasper file
     */
    public $report;
    /**
     * @var array report parameters
     */
    public $reportParams = [];
    /**
     * @var integer output format 
     */
    public $outputFormat = self::OUTPUT_PDF;
    /**
     * @var array output params 
     */
    public $outputParams = [];
    
    private $_connection;
    private $_fillManager;
    private $_jasperPrint;

    public function __construct($report = null, $reportParams = [], $config = [])
    {
        $this->report = $report;
        $this->reportParams = $reportParams;
        parent::__construct($config);
    }

    /**
     * Inits
     * @throws \Exception
     * @throws InvalidParamException
     */
    public function init()
    {
        parent::init();

        if ($this->db !== null) {
            $this->db = Instance::ensure($this->db, 'yii\db\Connection');
        }
        if (!$this->checkJavaExtension()) {
            throw new Exception("php-java-bridge is not running");
        }
    }

    public function getIsReportFileCompiled()
    {
        return (strtolower(pathinfo($this->report, PATHINFO_EXTENSION)) != 'jrxml');
    }

    public function getFillManager()
    {
        if ($this->_fillManager !== null) {
            return $this->_fillManager;
        } else {
            return new JavaClass("net.sf.jasperreports.engine.JasperFillManager");
        }
    }

    public function setJasperPrint($jasperPrint)
    {
        $this->_jasperPrint = $jasperPrint;
    }

    public function getJasperPrint()
    {
        return $this->_jasperPrint;
    }

    /**
     * Sets the output format and relative params
     * @param integer $output output format. Must be a FORMAT_* consts of this class
     * @param array $params key => value of output parameters http://jasperreports.sourceforge.net/api/index-all.html and search for JRxxxxExportParameters for full list
     */
    public function setOutput($output = self::OUTPUT_PDF, $params = [])
    {
        if (!in_array($output, array_keys(self::$_extensions))) {
            $this->outputFormat = self::OUTPUT_PDF;
        }
        if (!is_array($params)) {
            $this->outputParams = [];
        }
        $this->outputFormat = $output;
        $this->outputParams = $params;
    }

    /**
     * Loads report jrxml or jasper
     * @return net.sf.jasperreports.engine.JRReport | net.sf.jasperreports.engine.design.JasperDesign
     */
    protected function load()
    {
        if ($this->report === null) {
            throw new InvalidParamException(__CLASS__ . '::report must be set.');
        } else {
            $this->report = Yii::getAlias($this->report);
        }
        if (!file_exists($this->report)) {
            throw new InvalidParamException("Invalid report file: $this->report");
        }
        if ($this->isReportFileCompiled) {
            $JRLoader = new JavaClass("net.sf.jasperreports.engine.util.JRLoader");
            $reportObject = $JRLoader->loadObjectFromFile($this->report);
        } else {
            $JRXmlLoader = new JavaClass("net.sf.jasperreports.engine.xml.JRXmlLoader");
            $reportObject = $JRXmlLoader->load($this->report);
        }
        $event = new ReportEvent($reportObject);
        $this->trigger(self::EVENT_LOAD, $event);

        return $reportObject;
    }

    /**
     * Compiles report jrxml or jasper
     * @param net.sf.jasperreports.engine.JRReport $reportObject
     * @return JRReport|false if not a source
     */
    protected function compile($reportObject)
    {
        if ($this->isReportFileCompiled) {
            return $reportObject;
        }
        $event = new ReportEvent($reportObject);
        $this->trigger(self::EVENT_COMPILE, $event);
        $compilerManager = new Java("net.sf.jasperreports.engine.JasperCompileManager");
        return $compilerManager->compileReport($reportObject);
    }

    /**
     * Hashes report parameters
     * @return java.util.HashMap
     */
    public function reportParamsToHashMap()
    {
        if (!isset($this->reportParams['APP_FOLDER'])) {
            $this->reportParams['APP_FOLDER'] = Yii::getAlias('@app');
        }
        return new Java("java.util.HashMap", $this->reportParams);
    }

    /**
     * Fills report
     * @param JRReport $JRReport
     */
    protected function fill($JRReport)
    {
        $event = new ReportEvent($JRReport);
        $this->trigger(self::EVENT_FILL, $event);

        if ($this->jasperPrint === null) {
            $fillManager = $this->getFillManager();
            $this->jasperPrint = $fillManager->fillReport($JRReport, $this->reportParamsToHashMap(), $this->jdbcConnection);
        }
    }

    /**
     * Fills report width data and calls output generation.
     * @param string $file result filename. Leave empty string for automatic
     * @param string $dest destination output type (I=inline, F=file, D=download)
     * @return string
     */
    public function execute($file = '', $dest = self::DESTINATION_INLINE)
    {
//        if (!$this->checkJavaExtension()) {
//            throw new Exception("php-java-bridge is not running");
//        }

        $reportObject = $this->load();

        if (!$this->isReportFileCompiled) {
            $JRReport = $this->compile($reportObject);
        } else {
            $JRReport = $reportObject;
        }

        try {
            $this->fill($JRReport);
            $out = $this->output($file, $dest);
        } catch (\Exception $e) {
            throw($e);
        } finally {
            if ($this->_connection !== null) {
                $this->_connection->close();
            }
        }
        return $out;
    }

    /**
     * Generates output.
     * @param string|null $file result filename. if '' or null a temporary name will be generated
     * @param string|boolean $dest destination output type (I=inline, F=file, D=download)
     * @throws \Exception
     */
    protected function output($file, $dest)
    {
        $jasperPrint = $this->jasperPrint;

        if ($jasperPrint === null) {
            throw new Exception(__CLASS__ . '::jasperPrint is not set');
        }

        $event = new ReportEvent($jasperPrint);
        $this->trigger(self::EVENT_OUTPUT, $event);

        $dest = strtoupper($dest);
        if ($dest{0} != self::DESTINATION_FILE) {
            $file = preg_replace('/[\s]+/', '_', $file);
            $file = preg_replace('/[^a-zA-Z0-9_\.-]/', '', $file);
        }
        if ($file != '') {
            $outputPath = $file;
            $ext = "";
        } else {
            $outputPath = tempnam(sys_get_temp_dir(), 'report-');
            $ext = self::$_extensions[$this->outputFormat];
            $outputPath .= ".$ext";
        }

        $exportManager = new JavaClass("net.sf.jasperreports.engine.JasperExportManager");

        switch ($this->outputFormat) {
            case self::OUTPUT_PDF:
                if ($dest != self::DESTINATION_FILE) {
                    $stream = new Java('java.io.ByteArrayOutputStream');
                    $exportManager->exportReportToPdfStream($jasperPrint, $stream);
                } else {
                    $exportManager->exportReportToPdfFile($jasperPrint, $outputPath);
                }
                break;
            case self::OUTPUT_HTML:
                if ($dest != self::DESTINATION_FILE) {
                    $exporterHTML = new Java('net.sf.jasperreports.engine.export.JRHtmlExporter');
                    $JRHtmlExporterParameter = Java('net.sf.jasperreports.engine.export.JRHtmlExporterParameter');
                    $stream = new Java('java.io.ByteArrayOutputStream');

                    $exporterHTML->setParameter($JRHtmlExporterParameter->JASPER_PRINT, $jasperPrint);

//					$exporterHTML->setParameter($JRHtmlExporterParameter->IS_WHITE_PAGE_BACKGROUND,false);
//					$exporterHTML->setParameter($JRHtmlExporterParameter->IS_REMOVE_EMPTY_SPACE_BETWEEN_ROWS,false);
//					$exporterHTML->setParameter($JRHtmlExporterParameter->IS_DETECT_CELL_TYPE,TRUE);

                    foreach ($this->outputParams as $p => $v) {
                        $exporterHTML->setParameter($JRHtmlExporterParameter->$p, $v);
                    }
                    $exporterHTML->setParameter($JRHtmlExporterParameter->OUTPUT_STREAM, $stream);
                    $exporterHTML->exportReport();
                } else {
                    $exportManager->exportReportToHtmlFile($jasperPrint, $outputPath);
                }
                break;
            case self::OUTPUT_XML:
                if ($dest != self::DESTINATION_FILE) {
                    $stream = new Java('java.io.ByteArrayOutputStream');
                    $exportManager->exportReportToXmlStream($jasperPrint, $stream);
                } else {
                    $exportManager->exportReportToXmlFile($jasperPrint, $outputPath);
                }
                break;
            case self::OUTPUT_XLS:
                $exporterXLS = new Java('net.sf.jasperreports.engine.export.JRXlsExporter');
                $JRXlsExporterParameter = Java('net.sf.jasperreports.engine.export.JRXlsExporterParameter');

                $exporterXLS->setParameter($JRXlsExporterParameter->JASPER_PRINT, $jasperPrint);

                $exporterXLS->setParameter($JRXlsExporterParameter->IS_WHITE_PAGE_BACKGROUND, false);
                $exporterXLS->setParameter($JRXlsExporterParameter->IS_REMOVE_EMPTY_SPACE_BETWEEN_ROWS, true);
                $exporterXLS->setParameter($JRXlsExporterParameter->IS_REMOVE_EMPTY_SPACE_BETWEEN_COLUMNS, true);
                $exporterXLS->setParameter($JRXlsExporterParameter->IS_DETECT_CELL_TYPE, false);
                $exporterXLS->setParameter($JRXlsExporterParameter->IS_COLLAPSE_ROW_SPAN, true);

                foreach ($this->outputParams as $p => $v) {
                    $exporterXLS->setParameter($JRXlsExporterParameter->$p, $v);
                }

                if ($dest == self::DESTINATION_FILE) {
                    $exporterXLS->setParameter($JRXlsExporterParameter->OUTPUT_FILE_NAME, $outputPath);
                } else {
                    $stream = new Java('java.io.ByteArrayOutputStream');
                    $exporterXLS->setParameter($JRXlsExporterParameter->OUTPUT_STREAM, $stream);
                }

                $exporterXLS->exportReport();

                break;

            case self::OUTPUT_XLS_JAPI:
                $exporterXLS = new Java('net.sf.jasperreports.engine.export.JExcelApiExporter');
                $JRXlsExporterParameter = Java('net.sf.jasperreports.engine.export.JExcelApiExporterParameter');

                $exporterXLS->setParameter($JRXlsExporterParameter->JASPER_PRINT, $jasperPrint);

                $exporterXLS->setParameter($JRXlsExporterParameter->IS_WHITE_PAGE_BACKGROUND, false);
                $exporterXLS->setParameter($JRXlsExporterParameter->IS_REMOVE_EMPTY_SPACE_BETWEEN_ROWS, true);
                $exporterXLS->setParameter($JRXlsExporterParameter->IS_REMOVE_EMPTY_SPACE_BETWEEN_COLUMNS, true);
                $exporterXLS->setParameter($JRXlsExporterParameter->IS_DETECT_CELL_TYPE, false);
                $exporterXLS->setParameter($JRXlsExporterParameter->IS_COLLAPSE_ROW_SPAN, true);
//				$exporterXLS->setParameter($JRXlsExporterParameter->IS_FONT_SIZE_FIX_ENABLED,false);
//				$exporterXLS->setParameter($JRXlsExporterParameter->IS_IGNORE_GRAPHICS,true);

                foreach ($this->outputParams as $p => $v) {
                    $exporterXLS->setParameter($JRXlsExporterParameter->$p, $v);
                }

                if ($dest == self::DESTINATION_FILE) {
                    $exporterXLS->setParameter($JRXlsExporterParameter->OUTPUT_FILE_NAME, $outputPath);
                } else {
                    $stream = new Java('java.io.ByteArrayOutputStream');
                    $exporterXLS->setParameter($JRXlsExporterParameter->OUTPUT_STREAM, $stream);
                }
                $exporterXLS->exportReport();
                break;

            case self::OUTPUT_XLSX:
                $exporterXLSX = new Java('net.sf.jasperreports.engine.export.ooxml.JRXlsxExporter');
                $JRXlsExporterParameter = Java('net.sf.jasperreports.engine.export.JRXlsExporterParameter');

                $exporterXLSX->setParameter($JRXlsExporterParameter->JASPER_PRINT, $jasperPrint);

                $exporterXLSX->setParameter($JRXlsExporterParameter->IS_WHITE_PAGE_BACKGROUND, false);
                $exporterXLSX->setParameter($JRXlsExporterParameter->IS_REMOVE_EMPTY_SPACE_BETWEEN_ROWS, true);
                $exporterXLSX->setParameter($JRXlsExporterParameter->IS_REMOVE_EMPTY_SPACE_BETWEEN_COLUMNS, true);
                $exporterXLSX->setParameter($JRXlsExporterParameter->IS_DETECT_CELL_TYPE, false);
                $exporterXLSX->setParameter($JRXlsExporterParameter->IS_COLLAPSE_ROW_SPAN, true);

                if (!empty($this->outputParams)) {
                    foreach ($this->outputParams as $p => $v) {
                        $exporterXLSX->setParameter($JRXlsExporterParameter->$p, $v);
                    }
                }

                if ($dest == self::DESTINATION_FILE) {
                    $exporterXLSX->setParameter($JRXlsExporterParameter->OUTPUT_FILE_NAME, $outputPath);
                } else {
                    $stream = new Java('java.io.ByteArrayOutputStream');
                    $exporterXLSX->setParameter($JRXlsExporterParameter->OUTPUT_STREAM, $stream);
                }

                $exporterXLSX->exportReport();
                break;

            case self::OUTPUT_DOCX:
                $exporterDOCX = new Java('net.sf.jasperreports.engine.export.ooxml.JRDocxExporter');
                $JRDocxExporterParameter = Java('net.sf.jasperreports.engine.export.ooxml.JRDocxExporterParameter');

                $exporterDOCX->setParameter($JRDocxExporterParameter->JASPER_PRINT, $jasperPrint);

                $exporterDOCX->setParameter($JRDocxExporterParameter->FLEXIBLE_ROW_HEIGHT, true);

                foreach ($this->outputParams as $p => $v) {
                    $exporterDOCX->setParameter($JRDocxExporterParameter->$p, $v);
                }

                if ($dest == self::DESTINATION_FILE) {
                    $exporterDOCX->setParameter($JRDocxExporterParameter->OUTPUT_FILE_NAME, $outputPath);
                } else {
                    $stream = new Java('java.io.ByteArrayOutputStream');
                    $exporterDOCX->setParameter($JRDocxExporterParameter->OUTPUT_STREAM, $stream);
                }

                $exporterDOCX->exportReport();

//				if ($dest == self::DEST_FILE)
//					file_put_contents($outputPath,$stream->toByteArray());
                break;

            case self::OUTPUT_PPTX:
                $exporterPPTX = new Java('net.sf.jasperreports.engine.export.ooxml.JRPptxExporter');
                $JRExporterParameter = Java('net.sf.jasperreports.engine.JRExporterParameter');

                $exporterPPTX->setParameter($JRExporterParameter->JASPER_PRINT, $jasperPrint);

                if ($dest == self::DESTINATION_FILE) {
                    $exporterPPTX->setParameter($JRExporterParameter->OUTPUT_FILE_NAME, $outputPath);
                } else {
                    $stream = new Java('java.io.ByteArrayOutputStream');
                    $exporterPPTX->setParameter($JRExporterParameter->OUTPUT_STREAM, $stream);
                }

                $exporterPPTX->exportReport();

                break;

            case self::OUTPUT_CSV:
                $exporterCSV = new Java('net.sf.jasperreports.engine.export.JRCsvExporter');
                $JRCsvExporterParameter = Java('net.sf.jasperreports.engine.export.JRCsvExporterParameter');
                $stream = new Java('java.io.ByteArrayOutputStream');

                $exporterCSV->setParameter($JRCsvExporterParameter->JASPER_PRINT, $jasperPrint);

                $exporterCSV->setParameter($JRCsvExporterParameter->FIELD_DELIMITER, ";");

                foreach ($this->outputParams as $p => $v) {
                    $exporterCSV->setParameter($JRCsvExporterParameter->$p, $v);
                }
//				$exporterCSV->setParameter($JRCsvExporterParameter->RECORD_DELIMITER,"\n");

                if ($dest == self::DESTINATION_FILE) {
                    $exporterCSV->setParameter($JRCsvExporterParameter->OUTPUT_FILE_NAME, $outputPath);
                } else {
                    $exporterCSV->setParameter($JRCsvExporterParameter->OUTPUT_STREAM, $stream);
                }
                $exporterCSV->exportReport();
                break;

//			case self::TXT:
//				$exporterTXT=new Java('net.sf.jasperreports.engine.export.JRTextExporter');
//				$JRTextExporterParameter=Java('net.sf.jasperreports.engine.export.JRTextExporterParameter');
//				$JRExporterParameter=Java('net.sf.jasperreports.engine.JRExporterParameter');
//				$stream=new Java('java.io.ByteArrayOutputStream');
//				
//				$exporterTXT->setParameter($JRExporterParameter->JASPER_PRINT,$jasperPrint);
//				
//				$exporterTXT->setParameter($JRTextExporterParameter->CHARACTER_WIDTH,5);
//				$exporterTXT->setParameter($JRTextExporterParameter->CHARACTER_HEIGHT,20);
////				$exporterTXT->setParameter($JRTextExporterParameter->CHARACTER_ENCODING, "iso-8859-15");
////				$exporterTXT->setParameter($JRTextExporterParameter->PAGE_HEIGHT,$this->convertValue(90,'java.lang.Integer'));
////				$exporterTXT->setParameter($JRTextExporterParameter->PAGE_WIDTH,$this->convertValue(96,'java.lang.Integer'));
//
//				$exporterTXT->setParameter($JRExporterParameter->OUTPUT_STREAM,$stream);
//				
//				print_r(java_values($stream->toByteArray()));exit;
////				$exporterTXT->exportReport();
////				Yii::$app->end();
//				
//				if ($dest == self::DEST_FILE)
//					file_put_contents($outputPath,$stream->toByteArray());
//				break;
//			default:
//				$exportManager->exportReportToPdfFile($jasperPrint,$outputPath);
        }

        switch ($this->outputFormat) {
            case self::OUTPUT_PDF: $ctype = 'application/pdf';
                break;
            case self::OUTPUT_XLS_JAPI:
            case self::OUTPUT_XLS: $ctype = "application/vnd.ms-excel";
                break;
            case self::OUTPUT_XLSX: $ctype = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                break;
            case self::OUTPUT_PPTX: $ctype = "application/vnd.ms-powerpoint";
                break;
            case self::OUTPUT_PPTX: $ctype = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
                break;
//			case self::DOC: $ctype="application/vnd.msword"; break; 
            case self::OUTPUT_DOCX: $ctype = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                break;
//			case self::TXT: 
            case self::OUTPUT_CSV: $ctype = "application/text";
                break;
            case self::OUTPUT_HTML: $ctype = "text/html";
                break;
        }

        switch ($dest) {
            case self::DESTINATION_INLINE:
                // Send to the standard output
                if (ob_get_contents()) {
                    throw new \Exception("Some data has already been output, can\'t send $ext file");
                }
                // send output to a browser
                if (headers_sent()) {
                    throw new \Exception("Some data has already been output to browser, can\'t send $ext file");
                }
                return Yii::$app->getResponse()->sendContentAsFile($stream->toByteArray(), basename($outputPath), ['mimeType' => $ctype, 'inline' => true]);
            case self::DESTINATION_DOWNLOAD:
                // download as file
                if (ob_get_contents()) {
                    throw new \Exception("Some data has already been output, can\'t send $ext file");
                }
                return Yii::$app->getResponse()->sendContentAsFile($stream->toByteArray(), basename($outputPath), ['mimeType' => $ctype]);
            case self::DESTINATION_FILE:
                break;
//            case 'E':
//                // return PDF as base64 mime multi-part email attachment (RFC 2045)
//                $retval = "Content-Type: $ctype;" . "\r\n";
//                $retval .= ' name="' . basename($outputPath) . '"' . "\r\n";
//                $retval .= 'Content-Transfer-Encoding: base64' . "\r\n";
//                $retval .= 'Content-Disposition: attachment;' . "\r\n";
//                $retval .= ' filename="' . basename($outputPath) . '"' . "\r\n\r\n";
//                $retval .= chunk_split(base64_encode($stream->toByteArray()), 76, "\r\n");
////				unlink($outputPath);
//                return $retval;
//                break;
            case self::DESTINATION_STREAM:
                return $stream->toByteArray();
            default:
                throw new \Exception('Incorrect output destination: ' . $dest);
        }
    }

    public function getJdbcConnection()
    {
        if ($this->_connection === null) {
            $this->_connection = $this->createJdbcConnection();
        }
        return $this->_connection;
    }

    /**
     * Creates a JDBC connection based on CDBConnection parameters
     * @return object Java JDBC connection
     */
    public function createJdbcConnection()
    {
        if ($this->db === false) {
            return new Java("net.sf.jasperreports.engine.JREmptyDataSource");
        }

        $arr = explode(":", $this->db->dsn);
        $string = array_shift($arr);
        $arr2 = explode(";", array_pop($arr));
        $parsed = [];

        foreach ($arr2 as $p) {
            $temp = explode("=", $p);
            $parsed[strtolower($temp[0])] = $temp[1];
        }

        switch ($this->db->driverName) {
            case 'pgsql';
                $jdbcDriver = 'org.postgresql.Driver';
                $dbtype = 'postgresql';
                $port = $this->dbPort;
                break;
            case 'mysql';
                $jdbcDriver = 'com.mysql.jdbc.Driver';
                $dbtype = 'mysql';
                $port = $this->dbPort;
                break;
            default:
                throw new NotSupportedException('Database not supported');
        }

        Java('java.lang.Class')->forName($jdbcDriver);

        $jdbcUrl = strtr("jdbc:{type}://{host}:{port}/{dbname}?user={username}&password={password}", [
            '{type}' => $dbtype,
            '{host}' => $parsed['host'],
            '{port}' => (isset($parsed['port'])) ? $parsed['port'] : $port,
            '{dbname}' => $parsed['dbname'],
            '{username}' => $this->db->username,
            '{password}' => $this->db->password,
        ]);

        return Java('java.sql.DriverManager')->getConnection($jdbcUrl);
    }

    /**
     * convert a PHP value into Java object
     * @param mixed $value value to convert to
     * @param string $className Java class name 
     * @return mixed Java object or false
     */
    public function convertValue($value, $className)
    {
        // if we are a string, just use the normal conversion  
        // methods from the java extension...  
        try {
            if ($className == 'java.lang.String') {
                $temp = new Java('java.lang.String', $value);
                return $temp;
            } else if ($className == 'java.lang.Boolean' ||
            $className == 'java.lang.Integer' ||
            $className == 'java.lang.Long' ||
            $className == 'java.lang.Short' ||
            $className == 'java.lang.Double' ||
            $className == 'java.math.BigDecimal') {
                $temp = new Java($className, $value);
                return $temp;
            } else if ($className == 'java.sql.Timestamp' || $className == 'java.sql.Time') {
                $temp = new Java($className);
                $javaObject = $temp->valueOf($value);
                return $javaObject;
            } else if ($className == 'java.util.Locale') {
                $value_arr = explode("_", $value);
                $temp = new Java($className, $value_arr[0], $value_arr[1]);
                return $temp;
            } else if ($className == "java.util.Date") {
                $format = new Java('java.text.SimpleDateFormat', "yyyy-MM-dd");
                return $format->parse($value);
            } else if ($className == "java.sql.Date") {
                $date = new DateTime($value);
                return new Java('java.sql.Date', $date->getTimestamp() * 1000);
            }
        } catch (\Exception $err) {
            echo ( 'unable to convert value, ' . $value .
            ' could not be converted to ' . $className . ' ');
            //' could not be converted to ' . $className . ' ' . $err);  

            return false;
        }

        echo ( 'unable to convert value, class name ' . $className .
        ' not recognised');
        return false;
    }

    /**
     * See if the java extension was loaded. 
     * @return mixed boolean true for success, string for error
     */
    private function checkJavaExtension()
    {
        if (!extension_loaded('java')) {
            $sapiType = php_sapi_name();
            defined('JAVA_HOSTS') or define('JAVA_HOSTS', "127.0.0.1:{$this->javaBridgePort}");
            if ($sapiType == "cgi" || $sapiType == "cgi-fcgi" || $sapiType == "cli") {
                if (!(PHP_SHLIB_SUFFIX == "so" && @dl('java.so')) && !(PHP_SHLIB_SUFFIX == "dll" && @dl('php_java.dll')) && !(@include_once(dirname(__FILE__) . '/Java.inc')) && !(require_once("http://127.0.0.1:{$this->javaBridgePort}/java/Java.inc"))) {
                    return "java extension not installed.";
                }
            } else {
                if (!(@include_once(dirname(__FILE__) . '/Java.inc'))) {
                    require_once("http://127.0.0.1:{$this->javaBridgePort}/java/Java.inc");
                }
            }
        }
        if (!function_exists("java_get_server_name")) {
            return "The loaded java extension is not the PHP/Java Bridge";
        }
        return true;
    }
}
