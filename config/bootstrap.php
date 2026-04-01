<?php
declare(strict_types=1);

/**
 * CakeSpreadsheet plugin bootstrap.
 *
 * The request detector and other setup is handled in CakeSpreadsheetPlugin::bootstrap().
 *
 * To use the SpreadsheetView for content negotiation, add it to your controller's viewClasses():
 *
 * ```
 * use CakeSpreadsheet\View\SpreadsheetView;
 *
 * public function viewClasses(): array
 * {
 *     return [SpreadsheetView::class];
 * }
 * ```
 */
