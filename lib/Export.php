<?php

namespace YFormExport;

use PhpOffice\PhpSpreadsheet\Writer\Exception;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class Export
{
    private \rex_yform_manager_table $table;

    /**
     * @throws \rex_sql_exception
     */
    public function __construct() {
        $func = \rex_request('func', 'string', '');

        if ('yform_table_export' === $func && \rex_get('table')) {
            $this->table = \rex_yform_manager_table::get(\rex_get('table', 'string'));
            $this->exportTableSet();
        }
    }

    /**
     * @throws \rex_sql_exception
     * @throws Exception
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    private function exportTableSet(): void {
        $sql = \rex_sql::factory();
        $data = $sql->getArray('SELECT * FROM ' . $this->table->getTableName());

        if ($data === null) {
            exit();
        }


        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        $this->setHeader($sheet, $data[0]);

        $r = 2;
        foreach ($data as $row) {
            $c = 1;
            foreach ($row as $column) {
                $sheet->setCellValue([$c, $r], $column);
                $c++;
            }
            $r++;
        }

        $writer = new Xlsx($spreadsheet);
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment; filename="' . urlencode(time() . '_' . $this->table->getTableName() . '.xlsx') . '"');
        $writer->save('php://output');
        exit();
    }

    /**
     * set table/sheet header
     * @param Worksheet $sheet
     * @param array $data
     * @return void
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    private function setHeader(Worksheet $sheet, array $data): void {
        $i = 1;
        foreach ($data as $name => $value) {
            $label = $name;
            $valueField = $this->table->getValueField($name);

            if ($valueField) {
                $label = $valueField->getLabel();
            }

            $sheet->setCellValue([$i, 1], $label);
            $i++;
        }

        /**
         * fix first row/labels
         */
        $sheet->freezePane([1, 2]);
    }
}