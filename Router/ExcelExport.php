<?php

use App\Constants\PermissionTitle;
use App\Http\Controllers\Export\ExcelExportController;
use App\Models\Export\ExcelExport;
use Illuminate\Support\Facades\Route;

Route::middleware(['auth:api'])->group(
    function () {
        Route::apiResource('excel-exports', ExcelExportController::class)->only('store');
    }
);
