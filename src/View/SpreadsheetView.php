<?php
declare(strict_types=1);

namespace CakeSpreadsheet\View;

use Cake\Core\Exception\CakeException;
use Cake\Event\EventManager;
use Cake\Http\Response;
use Cake\Http\ServerRequest;
use Cake\Utility\Text;
use Cake\View\View;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

/**
 * Spreadsheet View
 */
class SpreadsheetView extends View
{
    /**
     * Excel layouts are located in the xlsx sub directory of `Layouts/`
     *
     * @var string
     */
    public $layoutPath = 'xlsx';

    /**
     * Excel views are always located in the 'xlsx' sub directory for a
     * controllers views.
     *
     * @var string
     */
    public $subDir = 'xlsx';

    /**
     * Spreadsheet instance
     *
     * @var \PhpOffice\PhpSpreadsheet\Spreadsheet
     */
    public $Spreadsheet = null;

    /**
     * Constructor
     *
     * @param \Cake\Http\ServerRequest|null $request Request instance.
     * @param \Cake\Http\Response|null $response Response instance.
     * @param \Cake\Event\EventManager|null $eventManager Event manager instance.
     * @param array $viewOptions View options. See View::$_passedVars for list of
     *   options which get set as class properties.
     */
    public function __construct(
        ?ServerRequest $request = null,
        ?Response $response = null,
        ?EventManager $eventManager = null,
        array $viewOptions = []
    ) {
        if (!empty($viewOptions['templatePath']) && $viewOptions['templatePath'] == '/xlsx') {
            $this->subDir = null;
        }

        parent::__construct($request, $response, $eventManager, $viewOptions);

        $this->response = $this->response->withType('xlsx');
        if (isset($viewOptions['templatePath']) && $viewOptions['templatePath'] == 'Error') {
            $this->subDir = null;
            $this->layoutPath = null;
            $this->response = $this->response->withType('html');

            return;
        }

        $this->Spreadsheet = new Spreadsheet();
    }

    /**
     * Renders view for given template file and layout.
     *
     * Render triggers helper callbacks, which are fired before and after the template are rendered,
     * as well as before and after the layout. The helper callbacks are called:
     *
     * - `beforeRender`
     * - `afterRender`
     * - `beforeLayout`
     * - `afterLayout`
     *
     * If View::$autoLayout is set to `false`, the template will be returned bare.
     *
     * Template and layout names can point to plugin templates/layouts. Using the `Plugin.template` syntax
     * a plugin template/layout can be used instead of the app ones. If the chosen plugin is not found
     * the template will be located along the regular view path cascade.
     *
     * @param string|null $template Name of template file to use
     * @param string|false|null $layout Layout to use. False to disable.
     * @return string Rendered content.
     * @throws \Cake\Core\Exception\CakeException If there is an error in the view.
     * @triggers View.beforeRender $this, [$templateFileName]
     * @triggers View.afterRender $this, [$templateFileName]
     */
    public function render(?string $template = null, $layout = null): string
    {
        $content = parent::render($template, $layout);
        if ($this->response->getType() === 'text/html') {
            return $content;
        }

        $this->Blocks->set('content', $this->output());
        $this->response = $this->response->withDownload($this->getFilename());

        return $this->Blocks->get('content');
    }

    /**
     * Generates the binary excel data
     *
     * @return string
     * @throws \Cake\Core\Exception\CakeException If the excel writer does not exist
     */
    protected function output(): string
    {
        ob_start();

        $writer = IOFactory::createWriter($this->Spreadsheet, 'Xlsx');

        if (!isset($writer)) {
            throw new CakeException('Excel writer not found');
        }

        $writer->setPreCalculateFormulas(false);
        $writer->setIncludeCharts(true);
        $writer->save('php://output');

        $output = ob_get_clean();

        return $output;
    }

    /**
     * Gets the filename
     *
     * @return string filename
     */
    public function getFilename(): string
    {
        if (isset($this->viewVars['_filename'])) {
            return $this->viewVars['_filename'] . '.xlsx';
        }

        return Text::slug(str_replace('.xlsx', '', $this->request->getPath())) . '.xlsx';
    }

    /**
     * Get instance of Spreadsheet
     *
     * @return \PhpOffice\PhpSpreadsheet\Spreadsheet The Spreadsheet Object
     * @throws \Cake\Core\Exception\CakeException If the excel writer does not exist
     */
    public function getSpreadsheet(): Spreadsheet
    {
        if ($this->Spreadsheet instanceof Spreadsheet) {
            return $this->Spreadsheet;
        }

        throw new CakeException('Spreadsheet not found');
    }
}
