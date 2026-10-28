PhpSpreadsheet extension for Yii2
=============

[![Latest Stable Version](https://poser.pugx.org/alexgx/yii2-phpexcel/v/stable.svg)](https://packagist.org/packages/alexgx/yii2-phpexcel) [![Total Downloads](https://poser.pugx.org/alexgx/yii2-phpexcel/downloads.svg)](https://packagist.org/packages/alexgx/yii2-phpexcel) [![Latest Unstable Version](https://poser.pugx.org/alexgx/yii2-phpexcel/v/unstable.svg)](https://packagist.org/packages/alexgx/yii2-phpexcel) [![License](https://poser.pugx.org/alexgx/yii2-phpexcel/license.svg)](https://packagist.org/packages/alexgx/yii2-phpexcel)

%short_description%

Installation
------------
The preferred way to install this extension is through [composer](http://getcomposer.org/download/).

Either run

```
php composer.phar require "alexgx/yii2-phpexcel" "*"
```
or add

```json
"alexgx/yii2-phpexcel" : "*"
```

to the require section of your application's `composer.json` file.

Features
---------------
Requires PhpSpreadsheet ^5.10.

- `PhpExcel` - create, load, send (`responseFile()`) or save (`saveFile()`) a workbook.
  The format comes from the file name extension; `.xlsx` is appended when the name has none.
  PHPExcel era format names (`Excel2007`, `Excel5`, ...) are mapped to PhpSpreadsheet types.
- `ExcelDataWriter` - writes rows to a sheet with explicit headers, cell types and styles.
- `GridDataWriter` - writes models or arrays with GridView like columns (compatible with `moonland\phpexcel\Excel` columns):
  attribute labels as headers, Yii formats, header style, borders, frozen header, auto filter and a totals row.
- `GridExport` - one sheet workbook written by `GridDataWriter`, sent as download or saved to a file.
- `ExcelDataReader` - reads a file into arrays keyed by the first row (same result as `moonland\phpexcel\Excel::readFile()`).
  By default only `.xlsx` files are opened.


Usage
-----
```php

use alexgx\phpexcel\PhpExcel;

$phpExcel = new PhpExcel();
$objPHPExcel = $phpExcel->create();

$objPHPExcel->getProperties()->setCreator("Traiding")
    ->setLastModifiedBy("Uldis Nelsons")
    ->setTitle("Packing list")
    ->setSubject("Packing list");

$activeSheet = $objPHPExcel->setActiveSheetIndex(0);
$activeSheet->setTitle('Packing List')
    ->setCellValue('A1', 'PACKING LIST')
    ->setCellValue('A3', 'VESSEL:')
    ->setCellValue('A4', 'B/L date:')
    ->setCellValue('A5', 'B/L No.:');



$activeSheet->getStyle('A1')->getFont()->setBold(true)->setSize(14);
$activeSheet->getStyle('A3:c9')->getFont()->setBold(true)->setSize(11);

$writer = new ExcelDataWriter();
$writer->setStartRow(11);
$writer->sheet = $activeSheet;
$writer->data = $packingList;

$headerStyles = [
    'font' => [
        'bold' => true
    ]
];
$writer->columns = [
    [
        'attribute' => 'departure_date',
        'header' => 'Departure Date',
        'headerStyles' => $headerStyles,
    ],
    [
        'attribute' => 'car_number',
        'header' => 'Car Number',
        'headerStyles' => $headerStyles,
    ],
    [
        'attribute' => 'delivery_note',
        'header' => 'Delivery Note',
        'headerStyles' => $headerStyles,
    ],
    [
        'attribute' => 'weight',
        'header' => 'Weight',
        'headerStyles' => $headerStyles,
    ],
    [
        'attribute' => 'gtd',
        'header' => 'Gtd',
        'headerStyles' => $headerStyles,
    ]
];

$writer->write();
$phpExcel->responseFile($objPHPExcel, 'packing.xls');


```

### Export models with GridExport

```php
use alexgx\phpexcel\GridExport;

(new GridExport([
    'models' => $dataProvider->getModels(),
    'columns' => [
        ['attribute' => 'invoice_date', 'format' => 'date', 'excelWidth' => 11],
        ['attribute' => 'number', 'format' => 'text'],
        'partner_name',
        ['attribute' => 'amount', 'header' => 'Total Amount', 'format' => ['decimal', 2]],
        ['attribute' => 'pay_amount', 'format' => ['decimal', 2]],
    ],
    'pageOrientationPortrait' => false,
    'pageFitToPage' => true,
    'headerHeight' => 30,
    'totals' => [
        'label' => Yii::t('d3invoices', 'Total'),
        'labelColspan' => 4,                      // label merged over A:D
        'sumColumns' => ['amount', 'pay_amount'], // =SUM() of these columns
    ],
]))->send('Invoices_' . date('ymdHi'));          // or ->save($directory, $fileName), ->getSpreadsheet()
```

Column keys:

| key | description |
|---|---|
| `'attribute'` or `'attribute:format:header'` | string shorthand |
| `attribute` | value path, read with `ArrayHelper::getValue()` (models, arrays, `relation.attribute`) |
| `value` | closure `function ($model, $key, $row, $writer)` or value path |
| `list` | value => label dictionary, unknown values are written as `value!` |
| `header`, `label` | header text; otherwise `headers[attribute]`, then the model attribute label |
| `format` | `date`, `date-time` (Excel date cells), `text` (always text), `['decimal', n]` (number), any `yii\i18n\Formatter` format |
| `type` | `ExcelDataWriter` cell type (`date`, `url` or a PhpSpreadsheet `DataType`) |
| `excelWidth` | column width, otherwise auto size (`'autoSize' => false` disables it) |
| `excelWrap`, `headerStyle`, `visible` | wrap text, header cell style, skip the column |

Values starting with `=` are written as text, never as formulas.

`totals` keys: `label`, `labelColspan`, `sumColumns` (attributes), `labelStyle` and `valueStyle`
(PhpSpreadsheet style arrays, default bold, right aligned). Totals are skipped when there are no rows.

Sheet options (on `GridExport` or `GridDataWriter`): `headerStyle`, `bodyStyle`, `headerHeight`,
`freezeHeader` (default `true`), `autoFilter` (default `true`), `dateFormat` (default `dd/mm/yy`),
`dateTimeFormat` (default `d/m/yy h:mm`). `GridDataWriter` also has `decimalSeparator`,
`thousandSeparator` and `formatter`.

`GridExport` creates `GridDataWriter` through the DI container, so application defaults can be set in the config:

```php
'container' => [
    'definitions' => [
        'alexgx\phpexcel\GridDataWriter' => [
            'dateFormat' => 'dd.mm.yy',
            'decimalSeparator' => '.',
            'thousandSeparator' => '',
        ],
    ],
],
```

### Read a file with ExcelDataReader

```php
use alexgx\phpexcel\ExcelDataReader;

$rows = (new ExcelDataReader(['setFirstRecordAsKeys' => true]))->read($file);
```

- one sheet: rows keyed by the first row values (or column letters when `setFirstRecordAsKeys` is `false`)
- several sheets: rows by sheet index, by sheet name with `setIndexSheetByName`, or one sheet with `getOnlySheet`
- `getOnlyRecordByIndex`, `leaveRecordByIndex` filter rows
- `allowedFormats` (default `['Xlsx']`): other files throw `PhpOffice\PhpSpreadsheet\Reader\Exception`

### Migrating from moonlandsoft/yii2-phpexcel

- `Excel::widget(['models' => ..., 'columns' => ..., 'fileName' => $name, ...])` -> `(new GridExport([...]))->send($name)`;
  `columns`, `headers`, `pageOrientationPortrait`, `pageFitToPage`, `headerHeight`, `headerStyle`, `bodyStyle`,
  `dateFormat`, `freezeHeader`, `autoFilter` keep their meaning.
- `Excel::export([... 'savePath' => $dir, 'asAttachment' => false])` -> `(new GridExport([...]))->save($dir, $name)`.
- `footerRows` with `TableTdCell` / `TableDecimalCell` and `loadDataInExcelStyle` -> `totals`.
  The style arrays are PhpSpreadsheet `applyFromArray()` styles and can be passed as `labelStyle` / `valueStyle`.
- `titleRows` -> write the title cells on your sheet and start the table below them with `GridDataWriter::setStartRow()`.
- `(new Excel())->readFile($file)` -> `(new ExcelDataReader())->read($file)`.

Further Information
-------------------
Please, check the [PHPOffice/PhpSpreadsheet github repo](https://github.com/PHPOffice/phpspreadsheet) documentation for further information about its configuration options.
