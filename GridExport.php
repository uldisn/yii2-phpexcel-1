<?php

namespace alexgx\phpexcel;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\PageSetup;
use Yii;
use yii\base\BaseObject;

/**
 * One sheet workbook written by GridDataWriter
 *
 * ```php
 * (new GridExport([
 *     'models' => $dataProvider->getModels(),
 *     'columns' => $columns,
 *     'totals' => ['label' => 'Total', 'labelColspan' => 4, 'sumColumns' => ['amount']],
 * ]))->send('Invoices_' . date('ymdHi'));
 * ```
 */
class GridExport extends BaseObject
{
    /**
     * @var array models or arrays, one per row
     */
    public $models = [];

    /**
     * @var array column definitions, see GridDataWriter
     */
    public $columns = [];

    /**
     * @var array attribute => header text
     */
    public $headers = [];

    /**
     * @var array|null totals row, see GridDataWriter::$totals
     */
    public $totals;

    /**
     * @var array document properties, name => value, e.g. ['creator' => 'irekini.lv', 'title' => 'Invoices']
     */
    public $properties = [];

    /**
     * @var string|null
     */
    public $sheetTitle;

    /**
     * @var bool
     */
    public $pageOrientationPortrait = true;

    /**
     * @var bool print the sheet on one page
     */
    public $pageFitToPage = false;

    /**
     * @var int|null
     */
    public $headerHeight;

    /**
     * @var array|null header row style, null keeps the GridDataWriter default
     */
    public $headerStyle;

    /**
     * @var array|null data rows style, null keeps the GridDataWriter default
     */
    public $bodyStyle;

    /**
     * @var bool|null
     */
    public $freezeHeader;

    /**
     * @var bool|null
     */
    public $autoFilter;

    /**
     * @var string|null Excel number format of `date` columns
     */
    public $dateFormat;

    /**
     * @var string|null Excel number format of `date-time` columns
     */
    public $dateTimeFormat;

    /**
     * @var Spreadsheet|null
     */
    private $spreadsheet;

    /**
     * Sends the file as download and ends the application
     * @param string $fileName extension is appended when missing
     * @param string|null $format PhpSpreadsheet writer type, default by file name extension or Xlsx
     */
    public function send(string $fileName, ?string $format = null): void
    {
        (new PhpExcel())->responseFile($this->getSpreadsheet(), $fileName, $format, false);
    }

    /**
     * @param string $directory
     * @param string $fileName extension is appended when missing
     * @param string|null $format PhpSpreadsheet writer type, default by file name extension or Xlsx
     * @return string saved file path
     */
    public function save(string $directory, string $fileName, ?string $format = null): string
    {
        return (new PhpExcel())->saveFile(
            $this->getSpreadsheet(),
            rtrim($directory, '/\\') . DIRECTORY_SEPARATOR . $fileName,
            $format
        );
    }

    /**
     * @return Spreadsheet
     */
    public function getSpreadsheet(): Spreadsheet
    {
        if ($this->spreadsheet === null) {
            $this->spreadsheet = $this->createSpreadsheet();
        }
        return $this->spreadsheet;
    }

    /**
     * @return Spreadsheet
     */
    protected function createSpreadsheet(): Spreadsheet
    {
        $spreadsheet = (new PhpExcel())->create();
        foreach ($this->properties as $name => $value) {
            $spreadsheet->getProperties()->{'set' . ucfirst($name)}($value);
        }
        $sheet = $spreadsheet->getActiveSheet();
        if ($this->sheetTitle !== null) {
            $sheet->setTitle($this->sheetTitle);
        }
        $sheet->getPageSetup()->setOrientation(
            $this->pageOrientationPortrait ? PageSetup::ORIENTATION_PORTRAIT : PageSetup::ORIENTATION_LANDSCAPE
        );
        if ($this->pageFitToPage) {
            $sheet->getPageSetup()->setFitToPage(true);
        }

        $writerOptions = array_filter([
            'headerHeight' => $this->headerHeight,
            'headerStyle' => $this->headerStyle,
            'bodyStyle' => $this->bodyStyle,
            'freezeHeader' => $this->freezeHeader,
            'autoFilter' => $this->autoFilter,
            'dateFormat' => $this->dateFormat,
            'dateTimeFormat' => $this->dateTimeFormat,
        ], static function ($value) {
            return $value !== null;
        });
        // created by the DI container, so application wide defaults can be set for GridDataWriter
        /** @var GridDataWriter $writer */
        $writer = Yii::createObject(array_merge(['class' => GridDataWriter::class], $writerOptions, [
            'sheet' => $sheet,
            'data' => $this->models,
            'columns' => $this->columns,
            'headers' => $this->headers,
            'totals' => $this->totals,
        ]));
        $writer->write();

        return $spreadsheet;
    }
}
