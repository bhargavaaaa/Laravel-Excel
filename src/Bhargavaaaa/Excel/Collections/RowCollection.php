<?php namespace Bhargavaaaa\Excel\Collections;

/**
 *
 * LaravelExcel RowCollection
 *
 * @category   Laravel Excel
 * @version    1.0.0
 * @package    bhargavaaaa/excel
 * @copyright  Copyright (c) 2013 - 2014 Bhargavaaaa (http://www.bhargavaaaa.nl)
 * @author     Bhargavaaaa <info@bhargavaaaa.nl>
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 */
class RowCollection extends ExcelCollection {

    /**
     * Sheet heading
     * @var array
     */
    protected $heading;

    /**
     * Get the heading
     * @return array
     */
    public function getHeading()
    {
        return $this->heading;
    }

    /**
     * Set the heading
     * @param array $heading
     */
    public function setHeading(array $heading)
    {
        $this->heading = $heading;
    }
}
