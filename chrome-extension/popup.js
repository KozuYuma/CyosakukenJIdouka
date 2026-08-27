// 同期先の既定値。手元で動かしているアプリに送りたいときは、
// ポップアップの入力欄を http://localhost:8501 に書き換える。
// 書き換えた値は次回も覚えている。
const DEFAULT_APP_URL = "https://cyosakuken-app.onrender.com";
const TARGET_COOKIES = ["_sess", "XSRF-TOKEN", "_ga", "_gid"];

// 同期用タブを開いてから閉じるまでの待ち時間(ms)。
// Streamlit はページの読み込み完了(complete)より後、WebSocket 経由で
// スクリプトを実行してからクエリパラメータを処理するため、
// complete だけで閉じると同期前に閉じてしまう。少し余裕を持たせる。
const SYNC_WAIT_MS = 2500;

// complete が来ないまま固まった場合の保険(ms)。
// 起こす処理は先に済ませてあるので、ここまで待つのは異常時だけ。
const SYNC_GIVEUP_MS = 30000;

// アプリを起こすのに待つ時間(ms)。Render の無料プランは15分使わないと
// 寝てしまい、次の1回は起きるのに30〜60秒かかる。ここで待たずにタブを
// 開くと、起きる前に閉じて「同期しました」と嘘をつくことになる。
const WAKE_TIMEOUT_MS = 90000;

const status = () => document.getElementById("status");
const urlBox = () => document.getElementById("appUrl");

function show(msg, cls = "") {
  status().textContent = msg;
  status().className = cls;
}

/** 入力欄の同期先。末尾の / は落とす。 */
function appUrl() {
  return (urlBox().value.trim() || DEFAULT_APP_URL).replace(/\/+$/, "");
}

// 前回の同期先を思い出す
chrome.storage.local.get("appUrl", (got) => {
  urlBox().value = got.appUrl || DEFAULT_APP_URL;
});

/** アプリが応答するまで待つ。起きたら true。 */
async function wakeApp(base) {
  const deadline = Date.now() + WAKE_TIMEOUT_MS;
  let told = false;
  while (Date.now() < deadline) {
    try {
      // Streamlit の生存確認の口。画面より軽く、寝ていると時間がかかる
      const res = await fetch(`${base}/_stcore/health`, { cache: "no-store" });
      if (res.ok) return true;
    } catch {
      // まだ起きていない
    }
    if (!told) {
      show("アプリを起こしています…（最大90秒）");
      told = true;
    }
    await new Promise((r) => setTimeout(r, 2000));
  }
  return false;
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
    let loaded = false;
    const close = () => {
      if (closed) return;
      closed = true;
      chrome.tabs.onUpdated.removeListener(onUpdated);
      chrome.tabs.remove(tab.id, () => {
        void chrome.runtime.lastError;   // 利用者が先に閉じていても無視
        // 読み込みが終わらないまま時間切れになったときは、届いたか
        // 分からない。分からないことを分かったように出さない
        if (loaded) {
          show("✅ 同期しました（アプリのタブはそのままです）", "ok");
        } else {
          show("⚠ 応答がありませんでした。届いたか分かりません。"
               + "アプリを開いて MINC の状態を確かめてください。", "err");
        }
      });
    };

    const onUpdated = (id, info) => {
      if (id === tab.id && info.status === "complete") {
        loaded = true;
        setTimeout(close, SYNC_WAIT_MS);
      }
    };
    chrome.tabs.onUpdated.addListener(onUpdated);

    setTimeout(close, SYNC_GIVEUP_MS);
  });
}

document.getElementById("syncBtn").addEventListener("click", () => {
  show("取得中...");
  const base = appUrl();
  chrome.storage.local.set({ appUrl: base });

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
    // 寝ているだけのこともあるので、少し待ってやる。
    if (!(await wakeApp(base))) {
      show(`⚠ ${base} に繋がりません。URL を確かめてください`
           + "（手元のアプリなら run.bat で起動）。", "err");
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
    syncViaHiddenTab(`${base}/?sync_minc=${encodeURIComponent(encoded)}`);
  });
});
