<?php
namespace CakeSpreadsheet\View;

use Cake\Core\Exception\Exception;
use Cake\Event\EventManager;
use Cake\Http\Response;
use Cake\Http\ServerRequest;
use Cake\Utility\Inflector;
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
     * @var Spreadsheet
     */
    public $Spreadsheet = null;

    /**
     * Constructor
     *
     * @param \Cake\Http\ServerRequest $request Request instance.
     * @param \Cake\Http\Response $response Response instance.
     * @param \Cake\Event\EventManager $eventManager EventManager instance.
     * @param array $viewOptions An array of view options
     */
    public function __construct(
        ServerRequest $request = null,
        Response $response = null,
        EventManager $eventManager = null,
        array $viewOptions = []
    ) {
        if (!empty($viewOptions['templatePath']) && $viewOptions['templatePath'] == '/xlsx') {
            $this->setSubDir(null);
        }

        parent::__construct($request, $response, $eventManager, $viewOptions);

        $this->setResponse($this->getResponse()->withType('xlsx'));
        if (isset($viewOptions['templatePath']) && $viewOptions['templatePath'] == 'Error') {
            $this->setSubDir(null);
            $this->setLayoutPath(null);
            $this->setResponse($this->getResponse()->withType('html'));

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
     * If View::$autoRender is false and no `$layout` is provided, the template will be returned bare.
     *
     * Template and layout names can point to plugin templates/layouts. Using the `Plugin.template` syntax
     * a plugin template/layout can be used instead of the app ones. If the chosen plugin is not found
     * the template will be located along the regular view path cascade.
     *
     * @param string|false|null $view Name of view file to use
     * @param string|null $layout Layout to use.
     * @return string|null Rendered content or null if content already rendered and returned earlier.
     * @throws \Cake\Core\Exception\Exception If there is an error in the view.
     * @triggers View.beforeRender $this, [$viewFileName]
     * @triggers View.afterRender $this, [$viewFileName]
     */
    public function render($view = null, $layout = null)
    {
        $content = parent::render($view, $layout);
        if ($this->getResponse()->getType() === 'text/html') {
            return $content;
        }

        $this->Blocks->set('content', $this->output());
        $this->setResponse($this->getResponse()->withDownload($this->getFilename()));

        return $this->Blocks->get('content');
    }

    /**
     * Generates the binary excel data
     *
     * @return string
     * @throws CakeException If the excel writer does not exist
     */
    protected function output()
    {
        ob_start();

        $writer = IOFactory::createWriter($this->Spreadsheet, 'Xlsx');

        if (!isset($writer)) {
            throw new Exception('Excel writer not found');
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
    public function getFilename()
    {
        if (isset($this->viewVars['_filename'])) {
            return $this->viewVars['_filename'] . '.xlsx';
        }

        return Inflector::slug(str_replace('.xlsx', '', $this->getRequest()->url)) . '.xlsx';
    }

    /**
     * Get instance of Spreadsheet
     *
     * @return \PhpOffice\PhpSpreadsheet\Spreadsheet The Spreadsheet Object
     */
    public function getSpreadsheet()
    {
        if ($this->Spreadsheet instanceof Spreadsheet) {
            return $this->Spreadsheet;
        }

        throw new Exception('Spreadsheet not found');
    }
}
