/* config.js - 業務管理アプリ設定
 *
 * 注意: ここに書いた値はクライアントから丸見えのため秘密情報を入れない。
 * 環境別に切り替えたい場合は config.local.js（.gitignore 済）を併用するか、
 * 初回起動時に「⚙ 設定」UI から localStorage("apiUrl") を上書きする。
 *
 * 本番 GAS の URL を別環境で差し替えたい場合は config.local.js を作成し、
 *   window.APP_CONFIG = Object.assign({}, window.APP_CONFIG, { DEFAULT_API_URL: "..." });
 * を記述した上で HTML に <script src="./config.local.js?v=1"></script> を追加する。
 */
window.APP_CONFIG = {
  DEFAULT_API_URL: "https://script.google.com/macros/s/AKfycbyjhCOYYP0RnAF9AlOYRlyBzjgV_WstLYfR4M343DCTfXLniTe1KNPMU8ScOskFeNcWtA/exec"
};
