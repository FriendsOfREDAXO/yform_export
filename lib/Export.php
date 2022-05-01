<?php

namespace YFormExport;

use PhpOffice\PhpSpreadsheet\Writer\Exception;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class Export
{
    private \rex_yform_manager_table $table;
    private array $relationsMap;
    private Spreadsheet $spreadsheet;
    private Worksheet $sheet;

    /**
     * @throws Exception
     * @throws \PhpOffice\PhpSpreadsheet\Exception
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
     * set table relations
     * @return void
     */
    private function setRelations(): void {
        $tableRelations = $this->table->getRelations();
        $relations = [];

        if ($tableRelations) {
            foreach ($tableRelations as $column => $field) {
                $relations[$column] = \rex_yform_value_be_manager_relation::getListValues($field->getElement('table'), $field->getElement('field'));
            }

            $i = 2;
            foreach ($this->table->getColumns() as $column) {
                if(array_key_exists($column['name'], $relations)) {
                    $this->relationsMap[$i] = $relations[$column['name']];
                }

                $i++;
            }
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

        /** create the spreadsheet */
        $this->spreadsheet = new Spreadsheet();
        $this->sheet = $this->spreadsheet->getActiveSheet();

        /** set worksheet name */
        $this->sheet->setTitle($this->table->getName());

        /** set relations */
        $this->setRelations();

        /** set header labels */
        $this->setHeader($data[0]);

        $r = 2;
        foreach ($data as $row) {
            $c = 1;
            foreach ($row as $name => $value) {
                /** set relation */
                if(isset($this->relationsMap[$c])) {
                    $value = $this->relationsMap[$c][$value];
                }

                $this->sheet->setCellValue([$c, $r], $value);
                $c++;
            }
            $r++;
        }

        $writer = new Xlsx($this->spreadsheet);
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment; filename="' . urlencode(time() . '_' . $this->table->getTableName() . '.xlsx') . '"');
        $writer->save('php://output');
        exit();
    }

    /**
     * set table/sheet header
     * @param array $data
     * @return void
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    private function setHeader(array $data): void {
        $i = 1;
        foreach ($data as $name => $value) {
            $label = $name;
            $valueField = $this->table->getValueField($name);

            if ($valueField) {
                $label = $valueField->getLabel();
            }

            $this->sheet->setCellValue([$i, 1], $label);
            $i++;
        }

        /**
         * fix first row/labels
         */
        $this->sheet->freezePane([1, 2]);
    }
}