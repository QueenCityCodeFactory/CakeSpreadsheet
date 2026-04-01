<?php
declare(strict_types=1);

/**
 * Test suite bootstrap for CakeSpreadsheet.
 *
 * This function is used to find the correct path to CakePHP to allow
 * the test suite to run without CakePHP being installed as a dependency
 * of the plugin.
 */

use Cake\Core\Configure;

$findRoot = function ($root) {
    do {
        $lastRoot = $root;
        $root = dirname($root);
        if (is_dir($root . '/vendor/cakephp/cakephp')) {
            return $root;
        }
    } while ($root !== $lastRoot);

    throw new \Exception('Cannot find the root of the application, unable to run tests');
};
$root = $findRoot(__FILE__);
unset($findRoot);

chdir($root);

require_once $root . '/vendor/autoload.php';

/**
 * Define fallback values for required constants and configuration.
 */
if (!defined('DS')) {
    define('DS', DIRECTORY_SEPARATOR);
}
define('ROOT', $root);
define('APP_DIR', 'src');
define('APP', ROOT . '/tests/test_app/src/');
define('CONFIG', ROOT . '/tests/test_app/config/');
define('WWW_ROOT', ROOT . '/tests/test_app/webroot/');
define('TESTS', ROOT . '/tests/');
define('TMP', ROOT . '/tmp/');
define('LOGS', TMP . 'logs/');
define('CACHE', TMP . 'cache/');
define('CAKE_CORE_INCLUDE_PATH', ROOT . '/vendor/cakephp/cakephp');
define('CORE_PATH', CAKE_CORE_INCLUDE_PATH . DS);

Configure::write('App', ['namespace' => 'CakeSpreadsheet\Test\App']);
Configure::write('debug', true);
