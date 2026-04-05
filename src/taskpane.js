/* ==========================================================
 *  電子印鑑 (Hanko / Data-in) — taskpane.js
 *  Canvas で日本式データ印画像を生成し、Excel に挿入する
 * ========================================================== */

// ---------- 定数 ----------
const STAMP_SIZE = 200;            // Canvas 描画サイズ (px)
const INSERT_SIZE_PX = 60;         // Excel 挿入時のサイズ (px)
const STAMP_COLOR = "#c62828";     // 赤色
const TRACKING_COLOR = "#aaaaaa";  // トラッキングID 文字色 (グレー)
const TRACKING_FONT_SIZE = 8;      // トラッキングID 文字サイズ (px)

// ---------- DOM 要素 ----------
let topTextEl, midTextEl, bottomTextEl, btnStamp, statusEl, canvas, ctx;

// ---------- Office 初期化 ----------
Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    initUI();
  }
});

function initUI() {
  topTextEl = document.getElementById("top-text");
  midTextEl = document.getElementById("mid-text");
  bottomTextEl = document.getElementById("bottom-text");
  btnStamp = document.getElementById("btn-stamp");
  statusEl = document.getElementById("status");
  canvas = document.getElementById("preview-canvas");
  ctx = canvas.getContext("2d");

  // 今日の日付を初期表示
  midTextEl.value = formatToday();

  // 入力変更でプレビュー更新
  [topTextEl, midTextEl, bottomTextEl].forEach(function (el) {
    el.addEventListener("input", drawPreview);
  });

  btnStamp.disabled = false;
  btnStamp.addEventListener("click", onStamp);

  drawPreview();
}

// ---------- トラッキングID生成 ----------
function generateTrackingId() {
  // 8桁の短い英数字ID (例: "a3f1b20e")
  var arr = new Uint8Array(4);
  crypto.getRandomValues(arr);
  return Array.from(arr, function (b) {
    return b.toString(16).padStart(2, "0");
  }).join("");
}

// ---------- 日付ヘルパー ----------
function formatToday() {
  var d = new Date();
  var yyyy = d.getFullYear();
  var mm = ("0" + (d.getMonth() + 1)).slice(-2);
  var dd = ("0" + d.getDate()).slice(-2);
  return yyyy + "/" + mm + "/" + dd;
}

// ---------- Canvas 描画 ----------
function drawPreview() {
  drawStamp(
    ctx,
    STAMP_SIZE,
    topTextEl.value,
    midTextEl.value,
    bottomTextEl.value
  );
}

/**
 * Canvas にデータ印を描画する
 * @param {CanvasRenderingContext2D} c
 * @param {number} size  - Canvas の幅・高さ
 * @param {string} top   - 上段テキスト
 * @param {string} mid   - 中段テキスト (日付)
 * @param {string} bottom - 下段テキスト
 * @param {string} [trackingId] - トラッキングID（省略時は描画しない）
 */
function drawStamp(c, size, top, mid, bottom, trackingId) {
  var cx = size / 2;
  var cy = size / 2;
  var radius = size / 2 - 6;       // 余白を確保
  var lineWidth = 3;

  c.clearRect(0, 0, size, size);
  c.strokeStyle = STAMP_COLOR;
  c.fillStyle = STAMP_COLOR;
  c.lineWidth = lineWidth;

  // --- 外枠 (円) ---
  c.beginPath();
  c.arc(cx, cy, radius, 0, Math.PI * 2);
  c.stroke();

  // --- 横線 (3分割) ---
  // 円内を均等に3分割: y座標は中心 ± radius/3
  var divY1 = cy - radius / 3;
  var divY2 = cy + radius / 3;

  drawChordLine(c, cx, cy, radius, divY1);
  drawChordLine(c, cx, cy, radius, divY2);

  // --- テキスト描画 ---
  c.textAlign = "center";
  c.textBaseline = "middle";

  // 上段
  var topY = cy - radius * 2 / 3;
  var topAreaH = radius / 3 * 2 / 1;  // 上端 〜 divY1
  drawFittedText(c, top, cx, (cy - radius + divY1) / 2, radius, topAreaH);

  // 中段 (日付 — 小さめフォント)
  drawFittedText(c, mid, cx, cy, radius, divY2 - divY1, true);

  // 下段
  drawFittedText(c, bottom, cx, (divY2 + cy + radius) / 2, radius, topAreaH);

  // --- トラッキングID (右下に極小グレー文字) ---
  if (trackingId) {
    c.save();
    c.fillStyle = TRACKING_COLOR;
    c.font = TRACKING_FONT_SIZE + "px 'Consolas', 'Courier New', monospace";
    c.textAlign = "right";
    c.textBaseline = "top";
    // 円の右下 45° の接線付近に配置
    var idX = cx + radius * 0.72;
    var idY = cy + radius * 0.72;
    c.fillText(trackingId, idX, idY);
    c.restore();
  }
}

/**
 * 円の弦（水平線）を描画
 */
function drawChordLine(c, cx, cy, r, y) {
  var dy = y - cy;
  var halfChord = Math.sqrt(r * r - dy * dy);
  c.beginPath();
  c.moveTo(cx - halfChord, y);
  c.lineTo(cx + halfChord, y);
  c.stroke();
}

/**
 * 領域に収まるようフォントサイズを自動調整してテキストを描画
 * @param {boolean} isDate 日付行かどうか
 */
function drawFittedText(c, text, x, y, maxWidth, areaHeight, isDate) {
  if (!text) return;

  // 使用可能幅は円弦の幅よりやや狭く
  var usableWidth = maxWidth * 1.4;
  var fontSize = Math.floor(areaHeight * 0.65);
  if (fontSize < 8) fontSize = 8;

  var fontFamily = isDate
    ? "'Consolas', 'Courier New', monospace"
    : "'Meiryo', 'Yu Gothic', 'Hiragino Sans', sans-serif";

  // フォントサイズを縮小しながらフィッティング
  for (; fontSize >= 8; fontSize -= 1) {
    c.font = "bold " + fontSize + "px " + fontFamily;
    if (c.measureText(text).width <= usableWidth) break;
  }

  c.fillStyle = STAMP_COLOR;
  c.fillText(text, x, y);
}

// ---------- 押印 (Excel 挿入) ----------
function onStamp() {
  setStatus("", "");
  btnStamp.disabled = true;

  // トラッキングID を生成
  var trackingId = generateTrackingId();

  // オフスクリーン Canvas で高品質画像を生成
  var offCanvas = document.createElement("canvas");
  offCanvas.width = STAMP_SIZE;
  offCanvas.height = STAMP_SIZE;
  var offCtx = offCanvas.getContext("2d");

  drawStamp(
    offCtx,
    STAMP_SIZE,
    topTextEl.value,
    midTextEl.value,
    bottomTextEl.value,
    trackingId
  );

  // Base64 PNG (data:image/png;base64,... のプレフィックスを除去)
  var dataUrl = offCanvas.toDataURL("image/png");
  var base64 = dataUrl.replace(/^data:image\/png;base64,/, "");

  insertImageToExcel(base64)
    .then(function () {
      setStatus("押印しました (ID: " + trackingId + ")", "success");
    })
    .catch(function (err) {
      console.error(err);
      setStatus("エラー: " + err.message, "error");
    })
    .finally(function () {
      btnStamp.disabled = false;
    });
}

/**
 * Base64 PNG 画像をアクティブセルに挿入する
 */
function insertImageToExcel(base64) {
  return Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = context.workbook.getSelectedRange();
    range.load("top,left");

    return context.sync().then(function () {
      var image = sheet.shapes.addImage(base64);
      image.left = range.left;
      image.top = range.top;

      // 挿入サイズ (ポイント単位: 1px ≈ 0.75pt)
      var sizePt = INSERT_SIZE_PX * 0.75;
      image.width = sizePt;
      image.height = sizePt;
      image.lockAspectRatio = true;

      return context.sync();
    });
  });
}

// ---------- ステータス表示 ----------
function setStatus(msg, type) {
  statusEl.textContent = msg;
  statusEl.className = type ? "status-" + type : "";
}
