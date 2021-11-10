<?php namespace Bhargavaaaa\Excel\Facades;

use Illuminate\Support\Facades\Facade;

/**
 *
 * LaravelExcel Facade
 *
 * @category   Laravel Excel
 * @version    1.0.0
 * @package    bhargavaaaa/excel
 * @copyright  Copyright (c) 2013 - 2014 Bhargavaaaa (http://www.bhargavaaaa.nl)
 * @author     Bhargavaaaa <info@bhargavaaaa.nl>
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 */
class Excel extends Facade {

    /**
     * Return facade accessor
     * @return string
     */
    protected static function getFacadeAccessor()
    {
        return 'excel';
    }
}