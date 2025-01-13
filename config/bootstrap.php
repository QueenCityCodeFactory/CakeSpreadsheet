<?php

use Cake\Event\EventInterface;
use Cake\Event\EventManager;
use Cake\Http\ServerRequest;

// Add event listener for setting view class map for xlsx
EventManager::instance()->on('Controller.initialize', function (EventInterface $event): void {
    $controller = $event->getSubject();
    $controller->viewBuilder()->setClassName('xlsx', 'CakeSpreadsheet.Spreadsheet');
});

// Add the request detector for xlsx
ServerRequest::addDetector('xlsx', [
    'accept' => ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'],
    'param' => '_ext',
    'value' => 'xlsx',
]);
