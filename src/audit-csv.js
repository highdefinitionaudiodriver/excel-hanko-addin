/* ==========================================================
 *  証跡 CSV 生成ロジック（純粋関数）
 *  ブラウザ（taskpane.js から global 経由）と
 *  Node（テスト・require 経由）の両方で利用できる UMD 形式。
 * ========================================================== */
(function (root, factory) {
  var api = factory();
  if (typeof module !== "undefined" && module.exports) {
    module.exports = api;
  }
  if (root) {
    root.HankoAudit = api;
  }
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
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

  /**
   * 押印履歴の配列を CSV 文字列に変換する。
   * Excel での文字化け回避のため UTF-8 BOM を先頭に付与し、
   * 行区切りは CRLF、末尾にも改行を付ける。
   */
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
    (log || []).forEach(function (e) {
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
    return "﻿" + lines.join("\r\n") + "\r\n";
  }

  return { csvCell: csvCell, buildAuditCsv: buildAuditCsv };
});
