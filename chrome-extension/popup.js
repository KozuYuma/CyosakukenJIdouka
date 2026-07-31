const APP_URL = "http://localhost:8501";
const TARGET_COOKIES = ["_sess", "XSRF-TOKEN", "_ga", "_gid"];

document.getElementById("syncBtn").addEventListener("click", () => {
  const status = document.getElementById("status");
  status.textContent = "取得中...";
  status.className = "";

  chrome.cookies.getAll({ domain: "minc.or.jp" }, (cookies) => {
    if (chrome.runtime.lastError) {
      status.textContent = "エラー: " + chrome.runtime.lastError.message;
      status.className = "err";
      return;
    }

    // _sess が含まれているか確認
    const sess = cookies.find(c => c.name === "_sess");
    if (!sess) {
      status.textContent = "⚠ _sess が見つかりません。minc.or.jp にログインしてください。";
      status.className = "err";
      return;
    }

    // 必要な Cookie だけを送信（URL 長短縮のため）
    const payload = cookies
      .filter(c => TARGET_COOKIES.includes(c.name))
      .map(c => ({
        name:    c.name,
        value:   c.value,
        domain:  c.domain,
        path:    c.path,
        secure:  c.secure,
        expires: c.expirationDate || -1
      }));

    // Base64 エンコードしてクエリパラメータで渡す
    const encoded = btoa(unescape(encodeURIComponent(JSON.stringify(payload))));
    const syncUrl = `${APP_URL}/?sync_minc=${encodeURIComponent(encoded)}`;

    chrome.tabs.create({ url: syncUrl }, () => {
      if (chrome.runtime.lastError) {
        status.textContent = "タブ作成エラー: " + chrome.runtime.lastError.message;
        status.className = "err";
      } else {
        status.textContent = "✅ アプリのタブが開きました！";
        status.className = "ok";
      }
    });
  });
});
