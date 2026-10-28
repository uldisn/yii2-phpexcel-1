<?php

namespace alexgx\phpexcel;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use yii\base\BaseObject;

/**
 * Reads a spreadsheet file into arrays, same result as moonland\phpexcel\Excel::readFile():
 * - file with one sheet: rows
 * - file with several sheets: rows by sheet index, or by sheet name when $setIndexSheetByName is set
 *
 * Rows are keyed by the first row values ($setFirstRecordAsKeys) or by column letters.
 */
class ExcelDataReader extends BaseObject
{
    /**
     * @var bool use the first row values as keys of the following rows, otherwise column letters
     */
    public $setFirstRecordAsKeys = true;

    /**
     * @var bool index sheets by name instead of sheet index
     */
    public $setIndexSheetByName = false;

    /**
     * @var string|null read only the sheet with this name, returns its rows
     */
    public $getOnlySheet;

    /**
     * @var array keep only these row indexes. For several sheets: sheet index => row indexes
     */
    public $getOnlyRecordByIndex = [];

    /**
     * @var array skip these row indexes. For several sheets: sheet index => row indexes
     */
    public $leaveRecordByIndex = [];

    /**
     * @var string[] PhpSpreadsheet reader types allowed to open the file. Keep it narrow for uploaded files.
     */
    public $allowedFormats = [IOFactory::READER_XLSX];

    /**
     * @param string $fileName
     * @return array
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception when the file is not one of $allowedFormats
     */
    public function read(string $fileName): array
    {
        $spreadsheet = IOFactory::load($fileName, 0, $this->allowedFormats);
        try {
            if ($spreadsheet->getSheetCount() === 1) {
                return $this->filterRecords(
                    $this->readSheet($spreadsheet->getActiveSheet()),
                    $this->getOnlyRecordByIndex,
                    $this->leaveRecordByIndex
                );
            }
            if ($this->getOnlySheet !== null) {
                if (!$sheet = $spreadsheet->getSheetByName($this->getOnlySheet)) {
                    return [];
                }
                return $this->filterRecords(
                    $this->readSheet($sheet),
                    $this->getOnlyRecordByIndex,
                    $this->leaveRecordByIndex
                );
            }
            $sheetsData = [];
            foreach ($spreadsheet->getWorksheetIterator() as $sheetIndex => $sheet) {
                $index = $this->setIndexSheetByName ? $sheet->getTitle() : $sheetIndex;
                $sheetsData[$index] = $this->filterRecords(
                    $this->readSheet($sheet),
                    $this->getOnlyRecordByIndex[$index] ?? [],
                    $this->leaveRecordByIndex[$index] ?? []
                );
            }
            return $sheetsData;
        } finally {
            $spreadsheet->disconnectWorksheets();
        }
    }

    /**
     * @param Worksheet $sheet
     * @return array
     */
    protected function readSheet(Worksheet $sheet): array
    {
        $rows = $sheet->toArray(null, true, true, true);
        if (!$this->setFirstRecordAsKeys) {
            return $rows;
        }
        $keys = array_map(static function ($key) {
            return (string)$key;
        }, $rows[1] ?? []);
        unset($rows[1]);
        $records = [];
        foreach ($rows as $values) {
            $records[] = array_combine($keys, $values);
        }
        return $records;
    }

    /**
     * @param array $records
     * @param array $onlyIndexes
     * @param array $skipIndexes
     * @return array
     */
    protected function filterRecords(array $records, array $onlyIndexes, array $skipIndexes): array
    {
        if ($onlyIndexes) {
            $records = array_intersect_key($records, array_flip($onlyIndexes));
        }
        if ($skipIndexes) {
            $records = array_diff_key($records, array_flip($skipIndexes));
        }
        return $records;
    }
}
