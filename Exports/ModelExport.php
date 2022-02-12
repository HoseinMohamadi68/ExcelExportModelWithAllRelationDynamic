<?php

namespace App\Exports;

use Illuminate\Http\Request;
use Illuminate\Support\Collection;
use Illuminate\Database\Eloquent\Model;
use Illuminate\View\View;
use Maatwebsite\Excel\Concerns\Exportable;
use Maatwebsite\Excel\Concerns\FromView;
use Maatwebsite\Excel\Concerns\RegistersEventListeners;
use Maatwebsite\Excel\Concerns\ShouldAutoSize;
use Maatwebsite\Excel\Concerns\WithEvents;
use Maatwebsite\Excel\Events\AfterSheet;

class ModelExport implements FromView, WithEvents, ShouldAutoSize
{
    use Exportable;
    use RegistersEventListeners;

    /**
     * @var Model|mixed Model.
     */
    private Model $model;

    /**
     * @var array
     */
    private array $headers;

    /**
     * @var array
     */
    private array $relations = [];

    /**
     * @var integer
     */
    private int $countRows = 1;

    /**
     * @var integer|null
     */
    private ?int $typeId;

    /**
     * @var array|mixed
     */
    private array $fields;

    /**
     * @var array
     */
    private array $requestFields;

    /**
     * @var array
     */
    private array $requestRelations;


    /**
     * @param Request      $request Request.
     * @param integer|null $typeId  Id.
     */
    public function __construct(
        Request $request,
        ?int $typeId = null
    ) {
        $data = collect($request->all());
        $modelName = $data->keys()->first();

        $this->model = new $modelName();
        $this->typeId = $typeId;
        $this->fields = $data[$modelName];
    }

    /**
     * @return View
     */
    public function view(): View
    {
        $relations = $this->fields['relations'] ?? [];

        $model = $this->model::when(count($relations), function ($joinRelations) use ($relations) {
            return $joinRelations->with(collect($relations)->pluck('relation')->toArray());
        })
            ->select('*')
            ->when($this->typeId, function ($addId) {
                return $addId->whereId($this->typeId);
            })
            ->get();

        return view('export.general', [
            'model' => $model,
            'headers' => $this->fields,
            'generateHeaders' => $this->generateHeaders(),
            'items' => $this->generateTdValues($model),
            'relationHeaders' => $this->generateRelationHeaders(),
        ]);
    }

    /**
     * @return Collection
     */
    public function generateHeaders(): Collection
    {
        $constants = collect((new \ReflectionClass($this->model))->getConstants())->values();
        $selectHeaders = collect();

        foreach ($this->fields as $column => $header) {
            if (!$constants->contains($column)) {
                continue;
            }

            $selectHeaders->put($column, $header);
        }

        if (!$selectHeaders->contains($this->model::ID)) {
            $selectHeaders->put($this->model::ID, $this->model::ID);
        }

        return $selectHeaders;
    }

    /**
     * @return array
     */
    public function generateRelationHeaders(): array
    {
        return $this->fields['relations'] ?? [];
    }

    /**
     * @param Collection $items Items.
     *
     * @return Collection
     */
    public function generateTdValues(Collection $items): Collection
    {
        return $items->map(function ($item) {
            $item->tdValues = collect();
            $item->tdValues = $this->generateMainColumnValues($item->tdValues, $item);

            $item->tdValues = $this->generateRelationsColumnValues($item->tdValues, $item);

            return $item;
        });
    }

    /**
     * @param Collection $tdValues Td Values.
     * @param Model      $item     Item.
     *
     * @return Collection
     */
    public function generateMainColumnValues(Collection $tdValues, Model $item): Collection
    {
        foreach ($this->generateHeaders() as $headerKey => $headerValue) {
            if (is_array($headerValue) and isset($headerValue['relation'])) {
                if ($headerValue['relation']) {
                    $tdValues->put($headerKey, (
                        optional($item->{$headerValue['relation']})->{$headerValue['field']} ?? null));
                } elseif (isset($headerValue['values'])) {
                    $tdValues->put($headerKey, ($headerValue['values'][$item->{$headerKey}] ?? null));
                }
            } else {
                $tdValues->put($headerKey, $item->$headerKey);
            }
        }

        return $tdValues;
    }

    /**
     * @param Collection $tdValues Td Values.
     * @param Model      $item     Item.
     *
     * @return Collection
     */
    public function generateRelationsColumnValues(Collection $tdValues, Model $item): collection
    {
        $countMaxRelations = 0;
        foreach ($this->generateRelationHeaders() as $headerValue) {
            if (
                $item->{$headerValue['relation']} instanceof Collection
                and is_object($item->{$headerValue['relation']})
            ) {
                if (
                    !is_null($item->{$headerValue['relation']})
                    and $item->{$headerValue['relation']}->count() > $countMaxRelations
                ) {
                    $countMaxRelations = $item->{$headerValue['relation']}->count();
                }
            } else {
                if (!is_null($item->{$headerValue['relation']})) {
                    $countMaxRelations = 1;
                }
            }
        }

        $items = collect();
        for ($i = 0; $i < $countMaxRelations; $i++) {
            $relationItems = collect();
            foreach ($this->generateRelationHeaders() as $headerValue) {
                $fields = collect();

                foreach ($headerValue['fields'] as $fieldKey => $field) {
                    if (
                        $item->{$headerValue['relation']} instanceof Collection
                        and is_object($item->{$headerValue['relation']})
                    ) {
                        if (isset($item->{$headerValue['relation']}[$i])) {
                            $fields->put($fieldKey, $item->{$headerValue['relation']}[$i]->$fieldKey);
                        }
                    } else {
                        $fields->put($fieldKey, $item->{$headerValue['relation']}->$fieldKey);
                    }
                }

                $relationItems->put(
                    $headerValue['relation'],
                    $fields->count() ? $fields : null
                );
            }
            $items->put($i, $relationItems);
        }

        $tdValues->put('relations', $items);
        $tdValues->put('relations_count', $countMaxRelations);

        return $tdValues;
    }


    /**
     * @param AfterSheet $event AfterSheet.
     *
     * @return void
     */
    public static function afterSheet(AfterSheet $event): void
    {
        $sheet = $event->sheet->getDelegate();
        $sheet->getRowDimension(1)->setRowHeight(30);

        $header = $sheet->getStyle('A1:' . $sheet->getHighestDataColumn() . '1');
        $header->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
        $header->getFont()->setBold(true);
        $header->getFill()->setFillType(
            \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID
        )->getStartColor()->setARGB('00000000');
        $header->getFont()->getColor()->setARGB(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_WHITE);

        $other = $sheet->getStyle('A2:' . $sheet->getHighestDataColumn() . $sheet->getHighestRow());
        $other->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP);

        foreach ([$header, $other] as $item) {
            $item->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT);
            $item->getAlignment()->setWrapText(true);
        }
    }

    /**
     * @return Model|mixed
     */
    public function getModel(): mixed
    {
        return $this->model;
    }
}
