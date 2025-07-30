<?php

use FriendsOfRedaxo\YFormExport\Export;
use FriendsOfRedaxo\YFormExport\ExtensionPoints;

if (rex::isBackend() && 'index.php?page=yform/manager/data_edit' === rex_url::currentBackendPage()) {
    /**
     * register EP, attach export button.
     */
    \rex_extension::register('YFORM_DATA_LIST_LINKS', ExtensionPoints::YFORM_DATA_LIST_LINKS(...));

    new Export();

    /**
     * register EP, attach messages.
     */
    \rex_extension::register('YFORM_MANAGER_DATA_PAGE_HEADER', ExtensionPoints::YFORM_MANAGER_DATA_PAGE_HEADER(...));
}
