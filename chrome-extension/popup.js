// アプリのURL。Render 等にデプロイしたら、ここを公開URLに書き換える。
const APP_URL = "http://localhost:8501";
const TARGET_COOKIES = ["_sess", "XSRF-TOKEN", "_ga", "_gid"];

// 同期用タブを開いてから閉じるまでの待ち時間(ms)。
// Streamlit はページの読み込み完了(complete)より後、WebSocket 経由で
// スクリプトを実行してからクエリパラメータを処理するため、
// complete だけで閉じると同期前に閉じてしまう。少し余裕を持たせる。
const SYNC_WAIT_MS = 2500;

const status = () => document.getElementById("status");

function show(msg, cls = "") {
  status().textContent = msg;
  status().className = cls;
}

/** 同期用のタブを裏で開き、処理が終わったら自動で閉じる。 */
function syncViaHiddenTab(syncUrl) {
  // active:false = 裏で開く。手前のアプリのタブが切り替わらないので、
  // 読み込み済みの楽曲データ(session_state)がそのまま残る。
  chrome.tabs.create({ url: syncUrl, active: false }, (tab) => {
    if (chrome.runtime.lastError) {
      show("タブ作成エラー: " + chrome.runtime.lastError.message, "err");
      return;
    }
    show("同期中...");

    let closed = false;
    const close = () => {
      if (closed) return;
      closed = true;
      chrome.tabs.onUpdated.removeListener(onUpdated);
      chrome.tabs.remove(tab.id, () => {
        void chrome.runtime.lastError;   // 利用者が先に閉じていても無視
        show("✅ 同期しました（アプリのタブはそのままです）", "ok");
      });
    };

    const onUpdated = (id, info) => {
      if (id === tab.id && info.status === "complete") {
        setTimeout(close, SYNC_WAIT_MS);
      }
    };
    chrome.tabs.onUpdated.addListener(onUpdated);

    // complete が来ないまま固まった場合の保険
    setTimeout(close, SYNC_WAIT_MS + 8000);
  });
}

document.getElementById("syncBtn").addEventListener("click", () => {
  show("取得中...");

  chrome.cookies.getAll({ domain: "minc.or.jp" }, async (cookies) => {
    if (chrome.runtime.lastError) {
      show("エラー: " + chrome.runtime.lastError.message, "err");
      return;
    }

    const sess = cookies.find((c) => c.name === "_sess");
    if (!sess) {
      show("⚠ _sess が見つかりません。minc.or.jp にログインしてください。", "err");
      return;
    }

    // アプリが起動していないと、開いたタブがエラーページのまま閉じて
    // 「同期した」と誤って出てしまうので、先に到達できるか確かめる。
    try {
      await fetch(APP_URL, { method: "GET", cache: "no-store" });
    } catch {
      show("⚠ アプリに接続できません。run.bat で起動してください。", "err");
      return;
    }

    // 必要な Cookie だけを送信（URL 長短縮のため）
    const payload = cookies
      .filter((c) => TARGET_COOKIES.includes(c.name))
      .map((c) => ({
        name:    c.name,
        value:   c.value,
        domain:  c.domain,
        path:    c.path,
        secure:  c.secure,
        expires: c.expirationDate || -1,
      }));

    const encoded = btoa(unescape(encodeURIComponent(JSON.stringify(payload))));
    syncViaHiddenTab(`${APP_URL}/?sync_minc=${encodeURIComponent(encoded)}`);
  });
});
