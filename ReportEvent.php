<?php

namespace anteo\jasperbridge;

use yii\base\Event;

class ReportEvent extends Event
{
    /**
     * @var mixed report object
     */
    public $reportObject;
    
    /**
     * Constructor.
     * @param Report $report the report associated with this action event.
     * @param array $config name-value pairs that will be used to initialize the object properties
     */
    public function __construct($reportObject, $config = [])
    {
        $this->reportObject = $reportObject;
        parent::__construct($config);
    }
}