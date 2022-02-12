<?php

namespace App\Http\Controllers\Export;

use App\Constants\PermissionTitle;
use App\Exports\ModelExport;
use App\Filters\Export\ExcelExportFilter;
use App\Http\Controllers\Controller;
use App\Http\Resources\Export\ExcelExportResource;
use App\Jobs\Export\ExcelExportJob;
use App\Models\Export\ExcelExport;
use Illuminate\Http\Request;
use Illuminate\Http\Resources\Json\AnonymousResourceCollection;
use Illuminate\Support\Facades\Auth;
use Maatwebsite\Excel\Excel;
use Illuminate\Http\Response;
use Illuminate\Support\Facades\Storage;

class ExcelExportController extends Controller
{
 
    /**
     * @param Request $request Request.
     *
     * @return \Illuminate\Http\JsonResponse
     */
    public function store(Request $request)
    {
        try {
            $modelExport = (new ModelExport($request));
            $className = get_class($modelExport->getModel());
            $userId = Auth::id();
            $fileName = 'export-data-' . $userId . '-' . time() . '.xlsx';
            $storagePath = ExcelExport::STORAGE_PATH . $fileName;

            $modelExport->queue($storagePath)->chain([
                new ExcelExportJob($userId, $className, $fileName)
            ]);

            return $this->getResponse(['message' => __('success.export_started')]);
        } catch (\Exception $exception) {
            return $this->getResponse(['message' => $exception->getMessage()], $exception->getCode());
        }
    }
}
