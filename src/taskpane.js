/* ==========================================================
 *  電子印鑑 (Hanko / Data-in) — taskpane.js
 *  Canvas で日本式データ印画像を生成し、
 *  Excel / Word / PowerPoint に挿入する
 * ========================================================== */

// ---------- 定数 ----------
const STAMP_SIZE = 200;            // Canvas 描画サイズ (px)
const INSERT_SIZE_PX = 60;         // Excel 挿入時のサイズ (px)
const STAMP_COLOR = "#c62828";     // 赤色
const TRACKING_COLOR = "#aaaaaa";  // トラッキングID 文字色 (グレー)
const TRACKING_FONT_SIZE = 8;      // トラッキングID 文字サイズ (px)

// ---------- DOM 要素 ----------
let topTextEl, midTextEl, bottomTextEl, btnStamp, statusEl, canvas, ctx;
let btnExportCsv, btnClearAudit, auditCountEl;

// 印影スタイル (circle / square / stamp)
function getSelectedShape() {
  var selected = document.querySelector('input[name="shape"]:checked');
  return selected ? selected.value : "circle";
}

// ---------- 証跡 (audit log) localStorage キー ----------
const AUDIT_LOG_KEY = "hanko_audit_log_v1";

// ---------- ホストアプリ ----------
let currentHost = null;

// ---------- Office 初期化 ----------
Office.onReady(function (info) {
  currentHost = info.host;

  // Excel / Word / PowerPoint いずれでも起動
  if (
    info.host === Office.HostType.Excel ||
    info.host === Office.HostType.Word ||
    info.host === Office.HostType.PowerPoint
  ) {
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

  // 印影スタイル変更でプレビュー更新
  document.querySelectorAll('input[name="shape"]').forEach(function (el) {
    el.addEventListener("change", drawPreview);
  });

  // ホスト名を表示
  updateHostLabel();

  btnStamp.disabled = false;
  btnStamp.addEventListener("click", onStamp);

  // 証跡 UI 初期化
  btnExportCsv = document.getElementById("btn-export-csv");
  btnClearAudit = document.getElementById("btn-clear-audit");
  auditCountEl = document.getElementById("audit-count");
  if (btnExportCsv) btnExportCsv.addEventListener("click", onExportAuditCsv);
  if (btnClearAudit) btnClearAudit.addEventListener("click", onClearAudit);
  refreshAuditUi();

  drawPreview();
}

// ---------- ホスト名表示 ----------
function getHostDisplayName() {
  switch (currentHost) {
    case Office.HostType.Excel:      return "Excel";
    case Office.HostType.Word:       return "Word";
    case Office.HostType.PowerPoint: return "PowerPoint";
    default:                         return "Office";
  }
}

function updateHostLabel() {
  var hostLabel = document.getElementById("host-label");
  if (hostLabel) {
    hostLabel.textContent = getHostDisplayName() + " で動作中";
  }
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
    bottomTextEl.value,
    null,
    getSelectedShape()
  );
}

/**
 * Canvas にデータ印を描画する（3 形状対応）
 * @param {CanvasRenderingContext2D} c
 * @param {number} size  - Canvas の幅・高さ
 * @param {string} top   - 上段テキスト
 * @param {string} mid   - 中段テキスト (日付)
 * @param {string} bottom - 下段テキスト
 * @param {string} [trackingId] - トラッキングID（省略時は描画しない）
 * @param {string} [shape] - 印影形状 "circle" / "square" / "stamp"（既定: circle）
 */
function drawStamp(c, size, top, mid, bottom, trackingId, shape) {
  shape = shape || "circle";
  var cx = size / 2;
  var cy = size / 2;
  var lineWidth = 3;

  c.clearRect(0, 0, size, size);
  c.strokeStyle = STAMP_COLOR;
  c.fillStyle = STAMP_COLOR;
  c.lineWidth = lineWidth;

  // 形状ごとに描画ヘルパーを切り替え
  if (shape === "square") {
    drawSquareFrame_(c, size, cx, cy);
    drawSquareBody_(c, size, cx, cy, top, mid, bottom);
  } else if (shape === "stamp") {
    drawStampFrame_(c, size, cx, cy);
    drawStampBody_(c, size, cx, cy, top, mid, bottom);
  } else {
    drawCircleFrame_(c, size, cx, cy);
    drawCircleBody_(c, size, cx, cy, top, mid, bottom);
  }

  // --- トラッキングID (右下に極小グレー文字) ---
  if (trackingId) {
    c.save();
    c.fillStyle = TRACKING_COLOR;
    c.font = TRACKING_FONT_SIZE + "px 'Consolas', 'Courier New', monospace";
    c.textAlign = "right";
    c.textBaseline = "top";
    var pad = size * 0.06;
    c.fillText(trackingId, size - pad, size - pad - TRACKING_FONT_SIZE);
    c.restore();
  }
}

// ── 丸印 (既存ロジック踏襲) ───────────────────────────────────────
function drawCircleFrame_(c, size, cx, cy) {
  var radius = size / 2 - 6;
  c.beginPath();
  c.arc(cx, cy, radius, 0, Math.PI * 2);
  c.stroke();
}
function drawCircleBody_(c, size, cx, cy, top, mid, bottom) {
  var radius = size / 2 - 6;
  var divY1 = cy - radius / 3;
  var divY2 = cy + radius / 3;
  drawChordLine(c, cx, cy, radius, divY1);
  drawChordLine(c, cx, cy, radius, divY2);
  c.textAlign = "center";
  c.textBaseline = "middle";
  var topAreaH = radius / 3 * 2 / 1;
  drawFittedText(c, top,    cx, (cy - radius + divY1) / 2, radius, topAreaH);
  drawFittedText(c, mid,    cx, cy,                         radius, divY2 - divY1, true);
  drawFittedText(c, bottom, cx, (divY2 + cy + radius) / 2,  radius, topAreaH);
}

// ── 角印 (正方形) ─────────────────────────────────────────────────
function drawSquareFrame_(c, size, cx, cy) {
  var half = size / 2 - 6;
  c.strokeRect(cx - half, cy - half, half * 2, half * 2);
}
function drawSquareBody_(c, size, cx, cy, top, mid, bottom) {
  var half = size / 2 - 6;
  var divY1 = cy - half / 3;
  var divY2 = cy + half / 3;
  // 水平区切り (フレーム内一杯)
  c.beginPath(); c.moveTo(cx - half, divY1); c.lineTo(cx + half, divY1); c.stroke();
  c.beginPath(); c.moveTo(cx - half, divY2); c.lineTo(cx + half, divY2); c.stroke();
  c.textAlign = "center";
  c.textBaseline = "middle";
  var areaH = (half * 2) / 3;
  drawFittedText(c, top,    cx, cy - half + areaH / 2, half * 1.4, areaH);
  drawFittedText(c, mid,    cx, cy,                    half * 1.4, areaH, true);
  drawFittedText(c, bottom, cx, cy + half - areaH / 2, half * 1.4, areaH);
}

// ── スタンプ風 (角丸＋影なし、横長気味) ───────────────────────────
function drawStampFrame_(c, size, cx, cy) {
  var w = size / 2 - 4;   // 横幅大きめ
  var h = size / 2.4;     // 縦幅やや小さめ → 横長感
  var r = h * 0.18;        // 角丸半径
  c.beginPath();
  // 角丸矩形
  c.moveTo(cx - w + r, cy - h);
  c.lineTo(cx + w - r, cy - h);
  c.quadraticCurveTo(cx + w, cy - h, cx + w, cy - h + r);
  c.lineTo(cx + w, cy + h - r);
  c.quadraticCurveTo(cx + w, cy + h, cx + w - r, cy + h);
  c.lineTo(cx - w + r, cy + h);
  c.quadraticCurveTo(cx - w, cy + h, cx - w, cy + h - r);
  c.lineTo(cx - w, cy - h + r);
  c.quadraticCurveTo(cx - w, cy - h, cx - w + r, cy - h);
  c.closePath();
  c.stroke();
}
function drawStampBody_(c, size, cx, cy, top, mid, bottom) {
  var w = size / 2 - 4;
  var h = size / 2.4;
  var divY1 = cy - h / 3;
  var divY2 = cy + h / 3;
  c.beginPath(); c.moveTo(cx - w, divY1); c.lineTo(cx + w, divY1); c.stroke();
  c.beginPath(); c.moveTo(cx - w, divY2); c.lineTo(cx + w, divY2); c.stroke();
  c.textAlign = "center";
  c.textBaseline = "middle";
  var areaH = (h * 2) / 3;
  drawFittedText(c, top,    cx, cy - h + areaH / 2, w * 1.8, areaH);
  drawFittedText(c, mid,    cx, cy,                  w * 1.8, areaH, true);
  drawFittedText(c, bottom, cx, cy + h - areaH / 2, w * 1.8, areaH);
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
 */
function drawFittedText(c, text, x, y, maxWidth, areaHeight, isDate) {
  if (!text) return;

  var usableWidth = maxWidth * 1.4;
  var fontSize = Math.floor(areaHeight * 0.65);
  if (fontSize < 8) fontSize = 8;

  var fontFamily = isDate
    ? "'Consolas', 'Courier New', monospace"
    : "'Meiryo', 'Yu Gothic', 'Hiragino Sans', sans-serif";

  for (; fontSize >= 8; fontSize -= 1) {
    c.font = "bold " + fontSize + "px " + fontFamily;
    if (c.measureText(text).width <= usableWidth) break;
  }

  c.fillStyle = STAMP_COLOR;
  c.fillText(text, x, y);
}

// ==========================================================
//  押印 (メイン処理)
// ==========================================================
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
    trackingId,
    getSelectedShape()
  );

  // Base64 PNG
  var dataUrl = offCanvas.toDataURL("image/png");
  var base64 = dataUrl.replace(/^data:image\/png;base64,/, "");

  // ホストに応じた挿入処理を呼び出し
  insertStampImage(base64)
    .then(function () {
      setStatus(
        "文書に押印しました (ID: " + trackingId + ")",
        "success"
      );
      // 証跡に記録（押印成功時のみ）
      appendAuditEntry({
        trackingId: trackingId,
        timestamp: new Date().toISOString(),
        host: getHostDisplayName(),
        shape: getSelectedShape(),
        topText: topTextEl.value || "",
        midText: midTextEl.value || "",
        bottomText: bottomTextEl.value || ""
      });
    })
    .catch(function (err) {
      console.error(err);
      setStatus("エラー: " + err.message, "error");
    })
    .finally(function () {
      btnStamp.disabled = false;
    });
}

// ==========================================================
//  ホスト別 画像挿入ルーター
// ==========================================================
function insertStampImage(base64) {
  switch (currentHost) {
    case Office.HostType.Excel:
      return insertIntoExcel(base64);
    case Office.HostType.Word:
      return insertIntoWord(base64);
    case Office.HostType.PowerPoint:
      return insertIntoPowerPoint(base64);
    default:
      return Promise.reject(
        new Error("未対応のホストアプリケーションです: " + currentHost)
      );
  }
}

// ==========================================================
//  Excel: アクティブセルの位置にシェイプとして挿入
// ==========================================================
function insertIntoExcel(base64) {
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

// ==========================================================
//  Word: カーソル位置にインライン画像として挿入
// ==========================================================
function insertIntoWord(base64) {
  return Word.run(function (context) {
    var selection = context.document.getSelection();

    // Word の insertInlinePictureFromBase64:
    //   第1引数: Base64文字列
    //   第2引数: 挿入位置 ("Replace" | "Start" | "End" | "Before" | "After")
    var picture = selection.insertInlinePictureFromBase64(base64, "Replace");

    // サイズ設定 (ポイント単位)
    var sizePt = INSERT_SIZE_PX * 0.75;
    picture.width = sizePt;
    picture.height = sizePt;
    picture.lockAspectRatio = true;

    return context.sync();
  });
}

// ==========================================================
//  PowerPoint: 共通API で選択スライドに画像を挿入
// ==========================================================
function insertIntoPowerPoint(base64) {
  return new Promise(function (resolve, reject) {
    // Office 共通 API — setSelectedDataAsync + CoercionType.Image
    Office.context.document.setSelectedDataAsync(
      base64,
      {
        coercionType: Office.CoercionType.Image,
        imageWidth: INSERT_SIZE_PX,
        imageHeight: INSERT_SIZE_PX
      },
      function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(new Error(result.error.message));
        }
      }
    );
  });
}

// ---------- ステータス表示 ----------
function setStatus(msg, type) {
  statusEl.textContent = msg;
  statusEl.className = type ? "status-" + type : "";
}

// ==========================================================
//  証跡 (audit log) — localStorage 永続化 + CSV エクスポート
//  法人プラン: クラウド同期・部署一括展開を別途提供予定
// ==========================================================

function loadAuditLog() {
  try {
    var raw = localStorage.getItem(AUDIT_LOG_KEY);
    if (!raw) return [];
    var parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) {
    console.warn("証跡の読み込みに失敗:", e);
    return [];
  }
}

function saveAuditLog(log) {
  try {
    localStorage.setItem(AUDIT_LOG_KEY, JSON.stringify(log));
  } catch (e) {
    // QuotaExceeded など
    console.error("証跡の保存に失敗:", e);
    setStatus("証跡の保存に失敗（容量超過の可能性）", "error");
  }
}

function appendAuditEntry(entry) {
  var log = loadAuditLog();
  log.push(entry);
  saveAuditLog(log);
  refreshAuditUi();
}

function refreshAuditUi() {
  var log = loadAuditLog();
  if (auditCountEl) auditCountEl.textContent = log.length + " 件";
  if (btnExportCsv) btnExportCsv.disabled = log.length === 0;
  if (btnClearAudit) btnClearAudit.disabled = log.length === 0;
}

/**
 * CSV 1セルの文字列を RFC4180 風にエスケープ
 *  - ダブルクォート / カンマ / 改行を含む値はダブルクォートで囲む
 *  - 値中のダブルクォートは "" にエスケープ
 */
function csvCell(value) {
  if (value == null) return "";
  var s = String(value);
  if (/[",\r\n]/.test(s)) {
    return '"' + s.replace(/"/g, '""') + '"';
  }
  return s;
}

function buildAuditCsv(log) {
  var headers = [
    "timestamp_iso8601",
    "tracking_id",
    "host_app",
    "shape",
    "top_text",
    "middle_text",
    "bottom_text"
  ];
  var lines = [headers.join(",")];
  log.forEach(function (e) {
    lines.push([
      csvCell(e.timestamp),
      csvCell(e.trackingId),
      csvCell(e.host),
      csvCell(e.shape || "circle"),
      csvCell(e.topText),
      csvCell(e.midText),
      csvCell(e.bottomText)
    ].join(","));
  });
  // CRLF + 末尾改行（Excel での文字化け回避のため UTF-8 BOM を先頭に）
  return "﻿" + lines.join("\r\n") + "\r\n";
}

function onExportAuditCsv() {
  var log = loadAuditLog();
  if (log.length === 0) {
    setStatus("証跡がありません", "error");
    return;
  }
  var csv = buildAuditCsv(log);
  var blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });

  // ファイル名: hanko_audit_YYYYMMDD_HHMMSS.csv
  var d = new Date();
  var pad = function (n) { return ("0" + n).slice(-2); };
  var fname = "hanko_audit_" +
    d.getFullYear() + pad(d.getMonth() + 1) + pad(d.getDate()) + "_" +
    pad(d.getHours()) + pad(d.getMinutes()) + pad(d.getSeconds()) + ".csv";

  // <a download> でブラウザにダウンロードさせる
  var url = URL.createObjectURL(blob);
  var a = document.createElement("a");
  a.href = url;
  a.download = fname;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  // メモリ解放はやや遅延（保存ダイアログとの競合回避）
  setTimeout(function () { URL.revokeObjectURL(url); }, 5000);

  setStatus("CSV 出力: " + fname + " (" + log.length + " 件)", "success");
}

function onClearAudit() {
  var log = loadAuditLog();
  if (log.length === 0) return;
  var ok = window.confirm(
    "押印履歴 " + log.length + " 件を全て削除します。\n" +
    "この操作は取り消せません。事前に CSV 出力を済ませることを推奨します。\n\n" +
    "本当に削除しますか?"
  );
  if (!ok) return;
  localStorage.removeItem(AUDIT_LOG_KEY);
  refreshAuditUi();
  setStatus("履歴を削除しました", "success");
}
