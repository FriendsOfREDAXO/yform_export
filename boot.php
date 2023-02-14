<?php

use YFormExport\Export;

if (rex::isBackend() && 'index.php?page=yform/manager/data_edit' === rex_url::currentBackendPage()) {
    /**
     * register EP, attach export button.
     */
    \rex_extension::register('YFORM_DATA_LIST_LINKS', 'YFormExport\ExtensionPoints::YFORM_DATA_LIST_LINKS');

    new Export();
}
