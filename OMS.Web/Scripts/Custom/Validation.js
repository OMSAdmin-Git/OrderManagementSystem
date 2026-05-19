function validateRequiredFromBtn(button) {
    // ボタンから form と data 属性を取得
    var form = button ? button.form : null;
    var errorLabelId = button ? button.getAttribute('data-error-label-id') : null;
    return validateRequired(form, errorLabelId);
}

function validateRequired(form, errorLabelId) {
    if (!form) return true;

    var lbl = errorLabelId ? document.getElementById(errorLabelId) : null;

    // 赤スタイルを効かせる
    form.classList.add('was-validated');

    // 未入力あり？
    if (!form.checkValidity()) {
        // エラーメッセージ
        if (lbl) {
            var msg = '必須項目が未入力です。赤枠の項目を入力してください。';
            if ('textContent' in lbl) lbl.textContent = msg; else lbl.innerText = msg;
            // 表示を強制（競合対策）
            lbl.style.setProperty('display', 'block', 'important');
            lbl.style.color = '#d32f2f';
            lbl.setAttribute('aria-live', 'polite');
            lbl.setAttribute('role', 'alert');
        }

        // 最初の不正要素へフォーカス
        var firstInvalid = form.querySelector(':invalid');
        if (firstInvalid) firstInvalid.focus();
        return false; // ← ポストバック停止
    }

    // すべてOK → メッセージ消してポストバック
    if (lbl) {
        if ('textContent' in lbl) lbl.textContent = ''; else lbl.innerText = '';
        lbl.style.display = 'none';
    }
    return true;
}
