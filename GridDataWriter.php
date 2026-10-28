<?php

namespace alexgx\phpexcel;

use DateTime;
use DateTimeInterface;
use Exception;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Shared\Date;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use Yii;
use yii\base\InvalidConfigException;
use yii\base\Model;
use yii\helpers\ArrayHelper;
use yii\i18n\Formatter;

/**
 * Writes models or arrays as a table with GridView like column definitions.
 * Column keys are compatible with moonland\phpexcel\Excel:
 *
 * - string shorthand 'attribute' or 'attribute:format:header'
 * - attribute, value (closure function ($model, $key, $row, $writer) or attribute path), list (value => label)
 * - header, label, headerStyle, visible, excelWidth, autoSize, excelWrap
 * - format: 'date', 'date-time', 'text', ['decimal', 2] or any yii\i18n\Formatter format
 * - type: ExcelDataWriter cell type ('date', 'url' or PhpSpreadsheet DataType)
 */
class GridDataWriter extends ExcelDataWriter
{
    /**
     * default style of the totals label and sum cells
     */
    public const TOTALS_STYLE = [
        'font' => [
            'bold' => true,
        ],
        'alignment' => [
            'horizontal' => Alignment::HORIZONTAL_RIGHT,
        ],
    ];

    /**
     * @var array attribute => header text, used when the column has no `header` or `label`
     */
    public $headers = [];

    /**
     * @var bool write the header row
     */
    public $setFirstTitle = true;

    /**
     * @var array|null header row style for PhpSpreadsheet applyFromArray()
     */
    public $headerStyle = [
        'font' => [
            'bold' => true,
            'color' => ['rgb' => 'FFFDFE'],
        ],
        'alignment' => [
            'horizontal' => Alignment::HORIZONTAL_CENTER,
            'vertical' => Alignment::VERTICAL_JUSTIFY,
            'wrapText' => true,
        ],
        'borders' => [
            'top' => [
                'borderStyle' => Border::BORDER_THIN,
            ],
        ],
        'fill' => [
            'fillType' => Fill::FILL_SOLID,
            'color' => ['rgb' => '7ebf00'],
        ],
    ];

    /**
     * @var array|null data rows style for PhpSpreadsheet applyFromArray()
     */
    public $bodyStyle = [
        'borders' => [
            'allBorders' => [
                'borderStyle' => Border::BORDER_THIN,
            ],
        ],
    ];

    /**
     * @var int|null header row height
     */
    public $headerHeight;

    /**
     * @var string Excel number format of `date` columns
     */
    public $dateFormat = 'dd/mm/yy';

    /**
     * @var string Excel number format of `date-time` columns
     */
    public $dateTimeFormat = 'd/m/yy h:mm';

    /**
     * @var array|null totals row below the data rows:
     * - label: string, written in the first column
     * - labelColspan: int, how many cells the label is merged over, default 1
     * - sumColumns: string[], attributes of the columns to sum
     * - labelStyle, valueStyle: array, PhpSpreadsheet styles, default self::TOTALS_STYLE
     */
    public $totals;

    /**
     * @var Formatter|array|null formatter for Yii formats, default is a copy of the application formatter
     */
    public $formatter;

    /**
     * @var string|null formatter decimal separator
     */
    public $decimalSeparator;

    /**
     * @var string|null formatter thousand separator
     */
    public $thousandSeparator;

    /**
     * @var bool
     */
    protected $freezeHeader = true;

    /**
     * @var bool
     */
    protected $autoFilter = true;

    /**
     * @var int|null
     */
    protected $headerRow;

    /**
     * @var int|null
     */
    protected $firstDataRow;

    /**
     * @var int|null
     */
    protected $lastDataRow;

    public function init()
    {
        parent::init();
        if ($this->formatter === null) {
            $this->formatter = clone Yii::$app->getFormatter();
        } elseif (is_array($this->formatter)) {
            $this->formatter = Yii::createObject($this->formatter);
        }
        if (!$this->formatter instanceof Formatter) {
            throw new InvalidConfigException('The "formatter" property must be either a Formatter object or a configuration array.');
        }
        $this->formatter->nullDisplay = null;
        if ($this->decimalSeparator !== null) {
            $this->formatter->decimalSeparator = $this->decimalSeparator;
        }
        if ($this->thousandSeparator !== null) {
            $this->formatter->thousandSeparator = $this->thousandSeparator;
        }
    }

    /**
     * @param bool $autoFilter
     */
    public function setAutoFilter($autoFilter)
    {
        $this->autoFilter = $autoFilter;
    }

    public function write()
    {
        if (!is_array($this->data) || !is_array($this->columns)) {
            return;
        }
        $this->columns = $this->normalizeColumns($this->columns);
        if (!$this->columns) {
            return;
        }
        if ($this->setFirstTitle) {
            $this->writeHeaderRow();
        }
        $this->writeDataRows();
        $this->writeFooterRow();
        $this->writeTotalsRow();
    }

    /**
     * @return int|null
     */
    public function getHeaderRow(): ?int
    {
        return $this->headerRow;
    }

    /**
     * @return int|null null when there are no data rows
     */
    public function getFirstDataRow(): ?int
    {
        return $this->lastDataRow === null ? null : $this->firstDataRow;
    }

    /**
     * @return int|null null when there are no data rows
     */
    public function getLastDataRow(): ?int
    {
        return $this->lastDataRow;
    }

    /**
     * @return int index of the last column
     */
    public function getLastColumn(): int
    {
        return $this->startColumn + count($this->columns) - 1;
    }

    /**
     * Converts string columns to arrays and removes invisible ones.
     * Without columns all model attributes are written.
     * @param array $columns
     * @return array
     * @throws InvalidConfigException
     */
    protected function normalizeColumns(array $columns): array
    {
        if (!$columns && ($model = reset($this->data)) !== false) {
            $columns = $model instanceof Model ? $model->attributes() : array_keys((array)$model);
        }
        $normalized = [];
        foreach ($columns as $column) {
            if (is_string($column)) {
                $parts = explode(':', $column);
                $column = ['attribute' => $parts[0]];
                if (isset($parts[1]) && $parts[1] !== '') {
                    $column['format'] = $parts[1];
                }
                if (isset($parts[2])) {
                    $column['header'] = $parts[2];
                }
            } elseif (!isset($column['attribute']) && !isset($column['value'])) {
                throw new InvalidConfigException('Column attribute or value must be defined.');
            }
            if (isset($column['visible']) && !$column['visible']) {
                continue;
            }
            $normalized[] = $column;
        }
        return $normalized;
    }

    protected function writeHeaderRow()
    {
        $this->headerRow = $this->j;
        $model = reset($this->data);
        $i = $this->startColumn;
        foreach ($this->columns as $column) {
            $this->sheet->setCellValueExplicit([$i, $this->j], $this->resolveHeader($column, $model), DataType::TYPE_STRING);
            ++$i;
        }
        if ($this->headerStyle) {
            $this->sheet
                ->getStyle([$this->startColumn, $this->j, $this->getLastColumn(), $this->j])
                ->applyFromArray($this->headerStyle);
        }
        $i = $this->startColumn;
        foreach ($this->columns as $column) {
            $columnStyle = $column['headerStyle'] ?? $column['headerStyles'] ?? null;
            if ($columnStyle) {
                $this->sheet->getStyle([$i, $this->j])->applyFromArray($columnStyle);
            }
            ++$i;
        }
        if ($this->headerHeight) {
            $this->sheet->getRowDimension($this->j)->setRowHeight($this->headerHeight);
        }
        if ($this->freezeHeader) {
            $this->sheet->freezePane([1, $this->j + 1]);
        }
        ++$this->j;
    }

    /**
     * @param array $column
     * @param mixed $model first data row, used for attribute labels
     * @return string
     */
    protected function resolveHeader(array $column, $model): string
    {
        if (isset($column['header'])) {
            return (string)$column['header'];
        }
        if (isset($column['label'])) {
            return (string)$column['label'];
        }
        if (!isset($column['attribute'])) {
            return '';
        }
        if (isset($this->headers[$column['attribute']])) {
            return (string)$this->headers[$column['attribute']];
        }
        if ($model instanceof Model) {
            return $model->getAttributeLabel($column['attribute']);
        }
        return $column['attribute'];
    }

    protected function writeDataRows()
    {
        $this->firstDataRow = $this->j;
        foreach ($this->data as $key => $model) {
            $i = $this->startColumn;
            foreach ($this->columns as $column) {
                $this->writeColumnCell($model, $key, $column, $i, $this->j);
                ++$i;
            }
            ++$this->j;
        }
        $this->lastDataRow = $this->j > $this->firstDataRow ? $this->j - 1 : null;
        $this->applyColumnFormats();
    }

    /**
     * @param mixed $model
     * @param mixed $key
     * @param array $column
     * @param int $i column index
     * @param int $row
     */
    protected function writeColumnCell($model, $key, array $column, int $i, int $row): void
    {
        try {
            $value = $this->getColumnValue($model, $key, $column, $row);
            if (isset($column['type'])) {
                $this->writeCell($value, $i, $row, $column);
                return;
            }
            switch ($this->getFormatName($column)) {
                case null:
                    $this->writeValue([$i, $row], $value);
                    break;
                case 'date':
                case 'date-time':
                    $this->writeValue([$i, $row], $this->toExcelDate($value));
                    break;
                case 'text':
                    if ($value === null || $value === '') {
                        $this->sheet->setCellValue([$i, $row], null);
                    } else {
                        $this->sheet->setCellValueExplicit([$i, $row], (string)$value, DataType::TYPE_STRING);
                    }
                    break;
                case 'decimal':
                    $this->writeValue([$i, $row], is_numeric($value) ? (float)$value : $value);
                    break;
                default:
                    $this->writeValue([$i, $row], $this->formatter->format($value, $column['format']));
            }
        } catch (Exception $exception) {
            Yii::error(
                'Excel column "' . ($column['attribute'] ?? $column['header'] ?? $i) . '" row ' . $row . ': '
                . $exception->getMessage() . PHP_EOL . $exception->getTraceAsString(),
                __METHOD__
            );
            $this->sheet->setCellValue([$i, $row], '???');
        }
    }

    /**
     * @param mixed $model
     * @param mixed $key
     * @param array $column
     * @param int $row
     * @return mixed
     * @throws Exception
     */
    protected function getColumnValue($model, $key, array $column, int $row)
    {
        if (isset($column['value'])) {
            if (is_string($column['value'])) {
                return ArrayHelper::getValue($model, $column['value']);
            }
            return call_user_func($column['value'], $model, $key, $row, $this);
        }
        if (!isset($column['attribute'])) {
            return null;
        }
        $value = ArrayHelper::getValue($model, $column['attribute']);
        if (isset($column['list']) && is_array($column['list'])) {
            if ($value === null || $value === '') {
                return '';
            }
            return $column['list'][$value] ?? $value . '!';
        }
        return $value;
    }

    /**
     * Writes a value, strings starting with "=" are written as text, not as formulas
     * @param array $coordinate
     * @param mixed $value
     */
    protected function writeValue(array $coordinate, $value): void
    {
        if (is_string($value) && strpos($value, '=') === 0) {
            $this->sheet->setCellValueExplicit($coordinate, $value, DataType::TYPE_STRING);
            return;
        }
        $this->sheet->setCellValue($coordinate, $value);
    }

    /**
     * @param mixed $value 'Y-m-d', 'Y-m-d H:i:s' or DateTimeInterface
     * @return mixed Excel date, null for empty dates or the original value when it is not a date
     */
    protected function toExcelDate($value)
    {
        if ($value instanceof DateTimeInterface) {
            return Date::dateTimeToExcel($value);
        }
        if ($value === null || $value === '' || strpos((string)$value, '0000-00-00') === 0) {
            return null;
        }
        foreach (['!Y-m-d H:i:s', '!Y-m-d'] as $format) {
            if ($dateTime = DateTime::createFromFormat($format, (string)$value)) {
                return Date::dateTimeToExcel($dateTime);
            }
        }
        return $value;
    }

    /**
     * @param array $column
     * @return string|null
     */
    protected function getFormatName(array $column): ?string
    {
        if (!isset($column['format'])) {
            return null;
        }
        $name = is_array($column['format']) ? ($column['format'][0] ?? null) : $column['format'];
        return $name === null ? null : (string)$name;
    }

    /**
     * @param array $column
     * @return string|null Excel number format of the column cells
     */
    protected function getNumberFormat(array $column): ?string
    {
        if (isset($column['type'])) {
            return null;
        }
        switch ($this->getFormatName($column)) {
            case 'date':
                return $this->dateFormat;
            case 'date-time':
                return $this->dateTimeFormat;
            case 'text':
                return NumberFormat::FORMAT_TEXT;
            case 'decimal':
                $decimals = $this->getDecimals($column);
                return $decimals > 0 ? '0.' . str_repeat('0', $decimals) : NumberFormat::FORMAT_NUMBER;
        }
        return null;
    }

    /**
     * @param array $column
     * @return int decimals of a ['decimal', n] column
     */
    protected function getDecimals(array $column): int
    {
        return is_array($column['format']) ? (int)($column['format'][1] ?? 2) : 2;
    }

    /**
     * Column widths, number formats, wrapping, body style and auto filter
     */
    protected function applyColumnFormats(): void
    {
        $i = $this->startColumn;
        foreach ($this->columns as $column) {
            $width = $column['excelWidth'] ?? $column['width'] ?? null;
            if (is_numeric($width)) {
                $this->sheet->getColumnDimensionByColumn($i)->setWidth((float)$width);
            } elseif (!isset($column['autoSize']) || $column['autoSize']) {
                $this->sheet->getColumnDimensionByColumn($i)->setAutoSize(true);
            }
            if ($this->lastDataRow !== null) {
                $range = [$i, $this->firstDataRow, $i, $this->lastDataRow];
                if (($numberFormat = $this->getNumberFormat($column)) !== null) {
                    $this->sheet->getStyle($range)->getNumberFormat()->setFormatCode($numberFormat);
                }
                if (isset($column['excelWrap'])) {
                    $this->sheet->getStyle($range)->getAlignment()->setWrapText((bool)$column['excelWrap']);
                }
            }
            ++$i;
        }
        if ($this->bodyStyle && $this->lastDataRow !== null) {
            $this->sheet
                ->getStyle([$this->startColumn, $this->firstDataRow, $this->getLastColumn(), $this->lastDataRow])
                ->applyFromArray($this->bodyStyle);
        }
        if ($this->autoFilter && $this->headerRow !== null) {
            $this->sheet->setAutoFilter([
                $this->startColumn,
                $this->headerRow,
                $this->getLastColumn(),
                $this->lastDataRow ?? $this->headerRow,
            ]);
        }
    }

    /**
     * ExcelDataWriter footer row, written only when a column has footer settings
     */
    protected function writeFooterRow()
    {
        foreach ($this->columns as $column) {
            if (isset($column['footer']) || isset($column['footerOptions']) || isset($column['footerLabel'])) {
                parent::writeFooterRow();
                return;
            }
        }
    }

    /**
     * @throws InvalidConfigException
     */
    protected function writeTotalsRow(): void
    {
        if (!$this->totals || $this->lastDataRow === null) {
            return;
        }
        if (isset($this->totals['label'])) {
            $labelColspan = max(1, (int)($this->totals['labelColspan'] ?? 1));
            $labelRange = [$this->startColumn, $this->j, $this->startColumn + $labelColspan - 1, $this->j];
            $this->sheet->setCellValueExplicit([$this->startColumn, $this->j], (string)$this->totals['label'], DataType::TYPE_STRING);
            if ($labelColspan > 1) {
                $this->sheet->mergeCells($labelRange);
            }
            $this->sheet->getStyle($labelRange)->applyFromArray($this->totals['labelStyle'] ?? self::TOTALS_STYLE);
        }
        foreach ($this->totals['sumColumns'] ?? [] as $attribute) {
            $i = $this->getColumnIndex($attribute);
            $letter = Coordinate::stringFromColumnIndex($i);
            $this->sheet->setCellValue([$i, $this->j], '=SUM(' . $letter . $this->firstDataRow . ':' . $letter . $this->lastDataRow . ')');
            $style = $this->sheet->getStyle([$i, $this->j]);
            $column = $this->columns[$i - $this->startColumn];
            if ($this->getFormatName($column) === 'decimal') {
                $decimals = $this->getDecimals($column);
                $style->getNumberFormat()->setFormatCode($decimals > 0 ? '#,##0.' . str_repeat('0', $decimals) : '#,##0');
            } elseif (($numberFormat = $this->getNumberFormat($column)) !== null) {
                $style->getNumberFormat()->setFormatCode($numberFormat);
            }
            $style->applyFromArray($this->totals['valueStyle'] ?? self::TOTALS_STYLE);
        }
        ++$this->j;
    }

    /**
     * @param string $attribute
     * @return int column index
     * @throws InvalidConfigException
     */
    protected function getColumnIndex(string $attribute): int
    {
        foreach ($this->columns as $position => $column) {
            if (($column['attribute'] ?? null) === $attribute) {
                return $this->startColumn + $position;
            }
        }
        throw new InvalidConfigException('Totals column "' . $attribute . '" is not in the columns.');
    }
}
