<?php
use Cake\Event\Event;
use Cake\Event\EventManager;
use Cake\Http\ServerRequest;

EventManager::instance()->on('Controller.initialize', function (Event $event) {
    $controller = $event->getSubject();
    if ($controller->components()->has('RequestHandler')) {
        $controller->RequestHandler->setConfig('viewClassMap.xlsx', 'CakeSpreadsheet.Spreadsheet');
    }
});

ServerRequest::addDetector(
    'xlsx',
    [
        'accept' => ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'],
        'param' => '_ext',
        'value' => 'xlsx',
    ]
);
