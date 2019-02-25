# CakeSpreadsheet
CakeSpreadsheet is a CakePHP 3.x plugin for generating Excel Files in the xlsx format using PHPOffice/PhpSpreadsheet.

## Requirements

* CakePHP 3.x
* PHP 7.2

## Installation

_[Using [Composer](http://getcomposer.org/)]_

```
composer require queencitycodefactory/cakespreadsheet
```

### Enable plugin

Load the plugin in your app's `config/bootstrap.php` file:

```php
Plugin::load('CakeSpreadsheet', ['bootstrap' => true, 'routes' => true]);
```

## Usage

First, you'll want to setup extension parsing for the `xlsx` extension. To do so, you will need to add the following to your `config/routes.php` file:

```php
// Set this before you specify any routes
Router::extensions('xlsx');
```

Next, we'll need to add a viewClassMap entry to your Controller. You can place the following in your AppController:

```php
public $components = [
    'RequestHandler' => [
        'viewClassMap' => [
            'xlsx' => 'CakeSpreadsheet.Spreadsheet',
        ],
    ]
];
```

Each application *must* have an xlsx layout. The following is a simple layout that can be placed in `src/Template/Layout/xlsx/default.ctp`:

```php
<?= $this->fetch('content') ?>
```

Finally, you can link to the current page with the .xlsx extension. This assumes you've created an `xlsx/index.ctp` file in your particular controller's template directory:

```php
$this->Html->link('Excel file', ['_ext' => 'xlsx']);
```

Inside your view file you will have access to the `PhpOffice\PhpSpreadsheet\Spreadsheet` library by using `$spreadsheet = $this->getSpreadsheet()`. Please see the [PhpOffice\PhpSpreadsheet](https://github.com/PHPOffice/PhpSpreadsheet) documentation for a guide on how to use `PhpOffice\PhpSpreadsheet`.
