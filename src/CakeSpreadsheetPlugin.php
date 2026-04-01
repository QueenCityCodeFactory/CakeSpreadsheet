<?php
declare(strict_types=1);

namespace CakeSpreadsheet;

use Cake\Core\BasePlugin;
use Cake\Core\PluginApplicationInterface;
use Cake\Http\ServerRequest;

class CakeSpreadsheetPlugin extends BasePlugin
{
    /**
     * Plugin name.
     *
     * @var string
     */
    protected ?string $name = 'CakeSpreadsheet';

    /**
     * Load routes or not
     *
     * @var bool
     */
    protected bool $routesEnabled = false;

    /**
     * Console middleware
     *
     * @var bool
     */
    protected bool $consoleEnabled = false;

    /**
     * @inheritDoc
     */
    public function bootstrap(PluginApplicationInterface $app): void
    {
        /**
         * Add a request detector named "xlsx" to check whether the request was for a spreadsheet,
         * either through accept header or file extension
         *
         * @link https://book.cakephp.org/5/en/controllers/request-response.html#checking-request-conditions
         */
        ServerRequest::addDetector(
            'xlsx',
            [
                'accept' => ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'],
                'param' => '_ext',
                'value' => 'xlsx',
            ]
        );
    }
}