// 1つの <select> に対してクラスを付け替える
function setFcstColor(selectEl) {
    // 既存の状態をクリア
    selectEl.classList.remove('flag-y', 'flag-n');
    // 値に応じて付与
    if (selectEl.value === 'Y') {
        selectEl.classList.add('flag-y');
    } else if (selectEl.value === 'N') {
        selectEl.classList.add('flag-n');
    }
}

// ページ読み込み時に対象すべてへ初期適用
document.addEventListener('DOMContentLoaded', function () {
    document.querySelectorAll('select.ddl-flag').forEach(function (el) {
        setFcstColor(el);
        // 変更イベントでも色を更新
        el.addEventListener('change', function () { setFcstColor(el); });
    });
});