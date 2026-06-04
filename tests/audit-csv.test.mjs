import test from "node:test";
import assert from "node:assert/strict";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const { csvCell, buildAuditCsv } = require("../src/audit-csv.js");

test("csvCell leaves plain values untouched", () => {
  assert.equal(csvCell("EMP001"), "EMP001");
  assert.equal(csvCell("circle"), "circle");
});

test("csvCell returns empty string for null/undefined", () => {
  assert.equal(csvCell(null), "");
  assert.equal(csvCell(undefined), "");
});

test("csvCell quotes and escapes values containing commas, quotes, or newlines", () => {
  assert.equal(csvCell("a,b"), '"a,b"');
  assert.equal(csvCell('say "hi"'), '"say ""hi"""');
  assert.equal(csvCell("line1\nline2"), '"line1\nline2"');
  assert.equal(csvCell("cr\r\nlf"), '"cr\r\nlf"');
});

test("buildAuditCsv starts with a UTF-8 BOM", () => {
  const csv = buildAuditCsv([]);
  assert.equal(csv.charCodeAt(0), 0xfeff);
});

test("buildAuditCsv emits the header row even when the log is empty", () => {
  const csv = buildAuditCsv([]);
  const firstLine = csv.replace(/^﻿/, "").split("\r\n")[0];
  assert.equal(
    firstLine,
    "timestamp_iso8601,tracking_id,host_app,shape,top_text,middle_text,bottom_text"
  );
});

test("buildAuditCsv tolerates a null/undefined log", () => {
  assert.ok(buildAuditCsv(undefined).includes("timestamp_iso8601"));
  assert.ok(buildAuditCsv(null).includes("timestamp_iso8601"));
});

test("buildAuditCsv renders an entry and defaults shape to circle", () => {
  const csv = buildAuditCsv([
    {
      timestamp: "2026-06-04T00:00:00.000Z",
      trackingId: "a3f1b20e",
      host: "Excel",
      topText: "承認",
      midText: "2026/06/04",
      bottomText: "山田"
    }
  ]);
  const rows = csv.replace(/^﻿/, "").split("\r\n");
  assert.equal(rows[1], "2026-06-04T00:00:00.000Z,a3f1b20e,Excel,circle,承認,2026/06/04,山田");
});

test("buildAuditCsv escapes a value containing a comma so columns are not split", () => {
  const csv = buildAuditCsv([
    { timestamp: "t", trackingId: "id", host: "Word", shape: "square", topText: "部長, 課長", midText: "", bottomText: "" }
  ]);
  assert.ok(csv.includes('"部長, 課長"'));
});

test("buildAuditCsv terminates with a trailing CRLF", () => {
  const csv = buildAuditCsv([{ timestamp: "t", trackingId: "id", host: "Excel" }]);
  assert.ok(csv.endsWith("\r\n"));
});
