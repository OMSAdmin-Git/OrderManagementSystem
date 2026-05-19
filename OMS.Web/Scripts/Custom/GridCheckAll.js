(function (global) {
    'use strict';

    // 名前空間（グローバル衝突回避）
    var App = global.OMS || (global.OMS = {});
    var Grid = App.Grid || (App.Grid = {});

    /**
     * 全選択／全解除
     * @param {string} gridClientId - GridViewのClientID
     * @param {HTMLInputElement} headerCheckbox - ヘッダーのチェックボックス（thisを渡す）
     * @param {string} itemIdPrefix - 行チェックボックスのID接頭辞（例: 'chkImport'）
     */
    Grid.toggleAll = function (gridClientId, headerCheckbox, itemIdPrefix) {
        var grid = document.getElementById(gridClientId);
        if (!grid || !headerCheckbox) return;

        var inputs = grid.getElementsByTagName('input');
        for (var i = 0; i < inputs.length; i++) {
            var el = inputs[i];
            if (el.type === 'checkbox' && el.id.indexOf(itemIdPrefix) !== -1) {
                el.checked = headerCheckbox.checked;
            }
        }
        // ヘッダーの半チェックは不要（全部ON/OFFした直後は明確）
        headerCheckbox.indeterminate = false;
    };

    /**
     * 行側のチェック変更に応じてヘッダーをON/OFF＆半チェックにする
     * @param {string} gridClientId - GridViewのClientID
     * @param {string} headerCheckboxId - ヘッダーのチェックボックスID
     * @param {string} itemIdPrefix - 行チェックボックスのID接頭辞
     */
    Grid.updateHeader = function (gridClientId, headerCheckboxId, itemIdPrefix) {
        var grid = document.getElementById(gridClientId);
        var header = document.getElementById(headerCheckboxId);
        if (!grid || !header) return;

        var inputs = grid.getElementsByTagName('input');
        var anyRow = false, allChecked = true, anyChecked = false;

        for (var i = 0; i < inputs.length; i++) {
            var el = inputs[i];
            if (el.type === 'checkbox' && el.id.indexOf(itemIdPrefix) !== -1) {
                anyRow = true;
                if (el.checked) anyChecked = true;
                if (!el.checked) allChecked = false;
            }
        }

        if (anyRow) {
            header.checked = allChecked;
            header.indeterminate = !allChecked && anyChecked; // 半チェック表現
        } else {
            header.checked = false;
            header.indeterminate = false;
        }
    };

})(window);
