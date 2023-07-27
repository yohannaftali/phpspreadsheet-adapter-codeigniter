# PhpOffice - PhpSpreadsheet

PhpSpreadsheet is a library written in pure PHP and offers a set of classes that allow you to read and write various spreadsheet file formats such as Excel and LibreOffice Calc.

## Installation With Composer
- Go To third_party folder

```bash
cd code\application\third_party\
```

- Create PHP Office Directory

```bash
mkdir phpoffice
```

- Go to directory

```bash
cd phpoffice
```

- Install with composer

```bash
composer require phpoffice/phpspreadsheet
```

- If installation failed, check PHP Version, update composer.json and edit php version then run compose update
```
{
    "require": {
        "phpoffice/phpspreadsheet": "^1.28"
    },
    "config": {
        "platform": {
            "php": "7.4"
        }
    }
}
```


## Update PhpOffice
- Use composer update

```bash
composer update
```

- If found problem delete folder vendor and run

```bash
composer install
```

## Create PhpOffice.php adapter in library

code\application\library\PhpOffice.php

```php
<?php defined('BASEPATH') OR exit('No direct script access allowed');

require_once APPPATH . '/third_party/phpoffice/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Writer\Xls;
use PhpOffice\PhpSpreadsheet\Writer\Pdf;
use PhpOffice\PhpSpreadsheet\Writer\Ods;
use PhpOffice\PhpSpreadsheet\Writer\Csv;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Worksheet\PageSetup;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Font;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Shared\Date as SharedDate;


class OfficeStyleBorder extends Border { public function __construct() { parent::__construct(); } }
class OfficeStyleAlignment extends Alignment { public function __construct() { parent::__construct(); } }
class OfficeStyleFill extends Fill { public function __construct() { parent::__construct(); } }
class OfficeStyleColor extends Color { public function __construct() { parent::__construct(); } }
class OfficeStyleFont extends Font { public function __construct() { parent::__construct(); } }
class OfficeStyleNumberFormat extends NumberFormat { public function __construct() { parent::__construct(); } }
class OfficeWorksheetPageSetup extends PageSetup { public function __construct() { parent::__construct(); } }
class OfficeSharedDate extends PhpOffice\PhpSpreadsheet\Shared\Date { }

class PhpOffice extends Spreadsheet
{
    private $PageSetup;
    private $BorderThin;
    private $BorderDouble;
    private $BorderThick;

    public function __construct()
    {
        parent::__construct();

        $this->PageSetup = new PageSetup();
        $this->BorderThin = Border::BORDER_THIN;
        $this->BorderDouble = Border::BORDER_DOUBLE;
        $this->BorderThick = Border::BORDER_THICK;
    }

    public function php_to_excel($date) {
        if($date == '') return '';        
        $excelDate = SharedDate::PHPToExcel($date);
        if(!$excelDate) return $date;
        return $excelDate;
    }

    public function write_xlsx($filename)
    {
        $writer = new Xlsx($this);
        ob_start(); 
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="' . $filename . '"');
        header('Cache-Control: max-age=0');
        $writer->save('php://output');
        $content = ob_get_contents();
        ob_end_clean();
        die($content);
        exit();
    }

    public function write_xls($filename)
    {
        $writer = new Xls($this);
        ob_start(); 
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="' . $filename . '"');
        header('Cache-Control: max-age=0');
        $writer->save('php://output');
        $content = ob_get_contents();
        ob_end_clean();
        die($content);
        exit();
    }

    public function write_ods($filename)
    {
        $writer = new Ods($this);
        ob_start(); 
        header('Content-Type: application/vnd.oasis.opendocument.spreadsheet');
        header('Content-Disposition: attachment;filename="' . $filename . '"');
        header('Cache-Control: max-age=0');
        $writer->save('php://output');
        $content = ob_get_contents();
        ob_end_clean();
        die($content);
        exit();
    }

    public function export($data)
    {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        // Add data to the sheet
        $sheet->fromArray($data);
        // Save the file
        $writer = new Xlsx($spreadsheet);
        $filename = 'example.xlsx';
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="' . $filename . '"');
        header('Cache-Control: max-age=0');
        $writer->save('php://output');
    }

    public function import_xlsx($file_data)
    {
        return $this->import_spreadsheet($file_data, 'xlsx');
    }

    public function import_xlsx_no_header($file_data)
    {
        return $this->import_spreadsheet($file_data, 'xlsx', false);
    }

    public function import_xls($file_data)
    {
        return $this->import_spreadsheet($file_data, 'xls');
    }

    public function import_xls_no_header($file_data)
    {
        return $this->import_spreadsheet($file_data, 'xls', false);
    }

    public function import_spreadsheet($file_data, $extension, $with_header = true, $with_format = true)
    {
        if (!isset($file_data)) {
            return ['error' => 'error: please restart application'];
        }
        $file = $file_data[$extension . '_file'] ?? [];
        $file_path =  $file['tmp_name'] ?? '';
        if($file_path == '') {
            return ['error' => 'file data not found: ' . $extension . '_file'];
        }
        return $this->import($file_path, $with_header, $with_format);
    }

    // Read File and Return as array
    public function import($file_path, $with_header = true, $with_format = true)
    {
        ini_set('auto_detect_line_endings', true);
        $spreasheet = IOFactory::load($file_path);
        $sheet = $spreasheet->getActiveSheet();
        $highestRow = $sheet->getHighestRow();
        $highestColString = $sheet->getHighestColumn();
        $highestCol = Coordinate::columnIndexFromString($highestColString);
        $data = [];
        $fields = [];
        for ($row = 0; $row < $highestRow; $row++) {
            for ($col = 0; $col < $highestCol; $col++) {
                $value = $with_format 
                    ? $sheet->getCellByColumnAndRow($col + 1, $row + 1)->getFormattedValue()
                    : $sheet->getCellByColumnAndRow($col + 1, $row + 1)->getValue();
                $value = trim($value);
                if (!$with_header) {
                    $data[$row][$col] = $value;
                    continue;
                }
                if ($row == 0) {
                    // Read first row for field name
                    $fields[$col] = $value;
                    continue;
                }
                // Read data row
                $field_name = $fields[$col] ?? $col;
                $data[$row - 1][$field_name] = $value;	
            }
        }
        return $data;
    }
}

```


## Load PhpOffice on model
- load

```php
$this->load->library('PhpOffice');
```

## Implementation

```bash
$this->load->library('PhpSpreadsheet');
```

### Init 

```php
				$office = new PhpOffice();
        $office->getProperties()
            ->setCreator($company_name)
            ->setLastModifiedBy($company_name)
            ->setTitle($title)
            ->setSubject($title)
            ->setDescription($title)
            ->setKeywords($title)
            ->setCategory($title);
        $office->getDefaultStyle()
            ->getFont()
            ->setName('Calibri (Body)')
            ->setSize(10);
        $office->setActiveSheetIndex(0);
        $sheet = $office->getActiveSheet();
				$sheet->setTitle($title);
        $sheet->getPageSetup()
            ->setPaperSize(OfficeWorksheetPageSetup::PAPERSIZE_A4)
            ->setOrientation(OfficeWorksheetPageSetup::ORIENTATION_LANDSCAPE)
            ->setScale(59)
            ->setRowsToRepeatAtTopByStartAndEnd(8, 9);
        $sheet->setShowGridlines(false);
        $sheet->getPageMargins()
            ->setLeft(0.5)
            ->setTop(0.5)
            ->setRight(0.5)
            ->setBottom(0.7);
        $sheet->getHeaderFooter()
            ->setOddFooter('&CHalaman &P / &N')
            ->setEvenFooter('&CHalaman &P / &N');

```

### Set dimension

```php
$col   = array('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ', 'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM', 'BN', 'BO', 'BP', 'BQ', 'BR', 'BS', 'BT', 'BU' );
$width = array( 8, 10, 8, 10, 21, 10, 10, 21, 21, 25, 10, 10, 10, 16, 10, 10, 10, 5, 5, 5, 5, 12, 12, 12, 45, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 21, 16, 20, 20, 20, 16, 16, 16, 16, 16, 16, 16, 16, 16, 16, 16, 16, 16, 16, 16, 16, 16 );

foreach ($col as $key => $item) {
    $sheet->getColumnDimension($item)->setWidth($width[$key]);
}
```

### Set border

```php
$border_thin = ['borderStyle' => OfficeStyleBorder::BORDER_THIN];
$border_double = ['borderStyle' => OfficeStyleBorder::BORDER_DOUBLE];

$style_border_thin_top = ['borders' => ['top' => $border_thin]];
$style_border_thin_left = ['borders' => ['left' => $border_thin]];
$style_border_thin_right = ['borders' => ['right' => $border_thin]];
$style_border_thin_bottom = ['borders' => ['bottom' => $border_thin]];
$style_border_double_bottom = ['borders' => ['bottom' => $border_double]];

$sheet->getStyle('C' . $row . ':D' . $row)->applyFromArray($style_border_thin_bottom);
```

### Merge Cells

```php
$sheet->mergeCells('A' . $row . ':B' . $row);
```

### Set Cell Value

```php
$sheet->setCellValue('A' . $row, 'Limit');
```

### To Write, at end of function

```php
return $office->write_xlsx($filename);
```
