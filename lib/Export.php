<?php

namespace YFormExport;

use PhpOffice\PhpSpreadsheet\Shared\Date;
use PhpOffice\PhpSpreadsheet\Writer\Exception;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class Export
{
    private \rex_yform_manager_table $table;
    private array $relationsMap;
    private array $columnTypes;
    private array $choices;
    private array $checkboxes;
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
     * get available choice names
     * @return void
     */
    private function setChoices(): void {
        $choices = $this->table->getValueFields(['type_name' => 'choice']);

        foreach ($choices as $choice) {
            $name = $choice->getName();
            $this->choices[$name] = $this->resolveChoices($name, $choice);
        }
    }

    /**
     * resolve choice values
     * @param string $columnName
     * @param \rex_yform_manager_field $field
     * @return mixed
     */
    private function resolveChoices(string $columnName, \rex_yform_manager_field $field):array {
        return \rex_yform_value_choice::getListValues([
            'field' => $columnName,
            'choices' => $field->getElement('choices'),
            'params' => [
                'field' => $field->toArray(),
                'fields' => $this->table->getFields(),
            ],
        ]);
    }

    /**
     * get available checkbox names
     * @return void
     */
    private function setCheckboxes(): void {
        $checkboxes = $this->table->getValueFields(['type_name' => 'checkbox']);

        foreach ($checkboxes as $checkbox) {
            $name = $checkbox->getName();
            $this->checkboxes[$name] = $this->resolveCheckboxes($checkbox);
        }
    }

    /**
     * resolve checkbox values
     * @param \rex_yform_manager_field $field
     * @return mixed
     */
    private function resolveCheckboxes(\rex_yform_manager_field $field):array {
        $values = $field->getElement('output_values');
        return $values ? explode(',', $values) : [0,1];
    }

    /**
     * set column types
     * @return void
     */
    private function setColumnTypes(): void {
        $columnTypes = [];
        $types = [
            'datetime',
            'date',
            'choice',
            'checkbox',
            'be_link',
            'be_media',
        ];

        $i = 2;
        foreach ($this->table->getColumns() as $column) {
            $valueField = $this->table->getValueField($column['name']);

            if(!$valueField) {
                continue;
            }

            $valueType = $valueField->getTypeName();
            if (in_array($valueType, $types, true)) {
                $this->columnTypes[$i] = $valueType;
            }
            $i++;
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
                if (array_key_exists($column['name'], $relations)) {
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

        /** set checkboxes */
        $this->setCheckboxes();

        /** set choices */
        $this->setChoices();

        /** set column types to format cells */
        $this->setColumnTypes();

        /** set header labels */
        $this->setHeader($data[0]);

        $r = 2;
        foreach ($data as $row) {
            $c = 1;
            foreach ($row as $name => $value) {
                $column = $this->sheet->getCell([$c, $r])->getColumn();

                /** set relation */
                if (isset($this->relationsMap[$c])) {
                    $value = $this->relationsMap[$c][$value];
                }

                /** format cell */
                if (isset($this->columnTypes[$c])) {
                    switch ($this->columnTypes[$c]) {
                        case 'datetime':
                            $this->sheet->setCellValue([$c, $r], Date::PHPToExcel($value));
                            $this->sheet
                                ->getStyle([$c, $r, $c, $r])
                                ->getNumberFormat()
                                ->setFormatCode('dd/mm/yyyy hh:mm:ss');
                            break;
                        case 'choice':
                            if('' !== $value) {
                                $this->sheet->setCellValue([$c, $r], $this->choices[$name][$value]);
                            }
                            break;
                        case 'checkbox':
                            if('' !== $value) {
                                $this->sheet->setCellValue([$c, $r], $this->checkboxes[$name][$value]);
                            }
                            break;
                        case 'be_link':
                            if('' !== $value) {
                                $article = \rex_article::get($value);
                                if($article) {
                                    $cell = $this->sheet->getCell([$c, $r]);
                                    $this->sheet->setCellValue([$c, $r], $article->getName())
                                        ->getHyperlink($cell->getCoordinate())
                                        ->setUrl(\rex::getServer().$article->getUrl());
                                }
                            }
                            break;
                        case 'be_media':
                            if('' !== $value) {
                                $media = \rex_media::get($value);
                                if($media) {
                                    $cell = $this->sheet->getCell([$c, $r]);
                                    $this->sheet->setCellValue([$c, $r], $media->getFileName())
                                        ->getHyperlink($cell->getCoordinate())
                                        ->setUrl(\rex::getServer().$media->getUrl());
                                }
                            }
                            break;
                        default:
                            $this->sheet->setCellValue([$c, $r], $value);
                            break;
                    }
                }
                else {
                    $this->sheet->setCellValue([$c, $r], $value);
                }

                /** set auto width */
                $this->sheet->getColumnDimension($column)->setAutoSize(true);

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