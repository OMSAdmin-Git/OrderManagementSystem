
/* Enterキー押下によるページバック防止 */
document.addEventListener("keydown", function (e) {
    if (e.key === "Enter" && e.target.tagName === "INPUT") {
        e.preventDefault();
        return false;
    }
});