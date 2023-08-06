<?php

namespace YFormExport;

use rex;
use rex_exception;
use rex_extension_point;
use rex_i18n;
use rex_url;
use rex_view;
use rex_yform_manager_table;

class ExtensionPoints
{
    /**
     * add export button.
     * @throws rex_exception
     */
    public static function YFORM_DATA_LIST_LINKS(rex_extension_point $ep): void // phpcs:ignore
    {
        $linkSets = $ep->getSubject();
        /** @var rex_yform_manager_table $table */
        $table = $ep->getParams()['table'];

        $linkParams = [
            'func' => 'yform_table_export',
            'table' => $table->getTableName(),
            'table_name' => $table->getTableName(),
        ];

        $item = [];
        $item['label'] = '<i class="fa fa-file" aria-hidden="true"></i>&nbsp;&nbsp;' . rex_i18n::msg('yform_manager_export');
        $item['url'] = rex_url::currentBackendPage() . '&' . http_build_query($linkParams);
        $item['attributes']['class'][] = 'btn-info';
        $item['attributes']['id'] = 'yform-export-table';
        //        $item['attributes']['download'] = '';

        /** add export button to table links */
        $linkSets['table_links'][] = $item;
        $ep->setSubject($linkSets);
    }

    /**
     * add messages.
     * @throws rex_exception
     */
    public static function YFORM_MANAGER_DATA_PAGE_HEADER(rex_extension_point $ep): void // phpcs:ignore
    {
        $tableHeader = $ep->getSubject();
        $message = '';

        /**
         * show message if no data is available.
         */
        if (rex::hasProperty('yform_export_data_empty')) {
            $message = rex_view::error(rex_i18n::msg('yform_export_data_empty'));
        }

        $ep->setSubject($tableHeader . $message);
    }
}
