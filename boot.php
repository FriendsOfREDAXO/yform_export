<?php

use YFormExport\Export;

if (rex::isBackend() && rex_url::currentBackendPage() === 'index.php?page=yform/manager/data_edit') {
    /**
     * register EP, attach export button
     */
    \rex_extension::register('YFORM_DATA_LIST_LINKS','YFormExport\ExtensionPoints::YFORM_DATA_LIST_LINKS');

    new Export();
}
