<?php
declare(strict_types=1);

namespace CakeSpreadsheet\Test\TestCase;

use CakeSpreadsheet\CakeSpreadsheetPlugin;
use Cake\TestSuite\TestCase;

class CakeSpreadsheetPluginTest extends TestCase
{
    public function testPluginName(): void
    {
        $plugin = new CakeSpreadsheetPlugin();
        $this->assertSame('CakeSpreadsheet', $plugin->getName());
    }
}
