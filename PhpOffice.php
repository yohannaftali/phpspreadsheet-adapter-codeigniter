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
