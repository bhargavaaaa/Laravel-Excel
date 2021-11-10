<?php namespace Bhargavaaaa\Excel\Collections;

/**
 *
 * LaravelExcel CellCollection
 *
 * @category   Laravel Excel
 * @package    bhargavaaaa/excel
 * @copyright  Copyright (c) 2013 - 2014 Bhargavaaaa (http://www.bhargavaaaa.nl)
 * @author     Bhargavaaaa <info@bhargavaaaa.nl>
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 */
class CellCollection extends ExcelCollection {

    /**
     * Create a new collection.
     * @param  array $items
     * @return \Bhargavaaaa\Excel\Collections\CellCollection
     */
    public function __construct(array $items = [])
    {
        $this->setItems($items);
    }

    /**
     * Set the items
     * @param array $items
     * @return void
     */
    public function setItems($items)
    {
        foreach ($items as $name => $value) {
            $name = trim($name) !== '' && !empty($name) ? $name : null;
            $value = !empty($value) || is_numeric($value) || is_bool($value) ? $value : null;
            $this->put($name, $value);
        }
    }

    /**
     * Dynamically get values
     * @param  string $key
     * @return string|null
     */
    public function __get($key)
    {
        if ($this->has($key)) {
            return $this->get($key);
        }

        return null;
    }

    /**
     * Determine if an attribute exists on the model.
     *
     * @param  string  $key
     * @return bool
     */
    public function __isset($key)
    {
        return $this->has($key);
    }
}
