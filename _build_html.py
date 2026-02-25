"""Build the HTML app with embedded Excel template."""
import base64
import os

folder = os.path.dirname(os.path.abspath(__file__))

# Read the template
with open(os.path.join(folder, "代理受領通知書_原本.xlsx"), "rb") as f:
    template_b64 = base64.b64encode(f.read()).decode("ascii")

# HTML content (using triple-quoted string with no single quotes in JS)
html_content = """<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>代理受領通知書 自動作成ツール</title>
<script src="https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js"><\/script>
<script src="https://cdn.jsdelivr.net/npm/file-saver@2.0.5/dist/FileSaver.min.js"><\/script>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:"Segoe UI","Yu Gothic UI","Meiryo",sans-serif;background:#f0f4f8;color:#1a202c;min-height:100vh}
.container{max-width:900px;margin:0 auto;padding:24px 16px}
h1{text-align:center;font-size:1.5rem;padding:20px 0;color:#2d3748}
h1 small{display:block;font-size:.75rem;color:#718096;font-weight:normal;margin-top:4px}
.step{background:#fff;border-radius:12px;padding:24px;margin-bottom:16px;box-shadow:0 1px 3px rgba(0,0,0,.1)}
.step-header{display:flex;align-items:center;gap:10px;margin-bottom:16px}
.step-num{background:#4299e1;color:#fff;width:28px;height:28px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:.85rem;font-weight:bold;flex-shrink:0}
.step-num.done{background:#48bb78}
.step-title{font-size:1.05rem;font-weight:600}
.file-area{border:2px dashed #cbd5e0;border-radius:8px;padding:20px;text-align:center;cursor:pointer;transition:all .2s;margin-bottom:12px;position:relative}
.file-area:hover{border-color:#4299e1;background:#ebf8ff}
.file-area.loaded{border-color:#48bb78;border-style:solid;background:#f0fff4}
.file-area input[type="file"]{position:absolute;inset:0;opacity:0;cursor:pointer}
.file-area .icon{font-size:2rem;margin-bottom:8px}
.file-area .label{font-size:.9rem;color:#718096}
.file-area .loaded-info{font-size:.85rem;color:#2f855a;font-weight:600;margin-top:6px}
.form-grid{display:grid;grid-template-columns:120px 1fr;gap:12px;align-items:center}
.form-grid label{font-size:.9rem;font-weight:600;color:#4a5568}
.form-grid input,.form-grid select{padding:8px 12px;border:1px solid #e2e8f0;border-radius:6px;font-size:.9rem}
.form-grid input:focus,.form-grid select:focus{outline:none;border-color:#4299e1;box-shadow:0 0 0 3px rgba(66,153,225,.2)}
.data-table-wrap{max-height:400px;overflow-y:auto;border:1px solid #e2e8f0;border-radius:8px}
table{width:100%;border-collapse:collapse;font-size:.85rem}
thead{position:sticky;top:0;z-index:1}
th{background:#edf2f7;padding:10px 12px;text-align:left;font-weight:600;color:#4a5568;border-bottom:2px solid #cbd5e0}
td{padding:8px 12px;border-bottom:1px solid #edf2f7}
tr:hover td{background:#f7fafc}
.num{text-align:right;font-variant-numeric:tabular-nums}
.total-row td{font-weight:700;background:#edf2f7!important;border-top:2px solid #cbd5e0}
.btn-generate{display:block;width:100%;padding:14px;background:#4299e1;color:#fff;border:none;border-radius:8px;font-size:1.1rem;font-weight:700;cursor:pointer;transition:all .2s}
.btn-generate:hover:not(:disabled){background:#3182ce;transform:translateY(-1px);box-shadow:0 4px 12px rgba(66,153,225,.3)}
.btn-generate:disabled{background:#a0aec0;cursor:not-allowed}
.status{margin-top:12px;padding:12px 16px;border-radius:8px;font-size:.9rem;display:none}
.status.success{display:block;background:#f0fff4;color:#276749;border:1px solid #c6f6d5}
.status.error{display:block;background:#fff5f5;color:#9b2c2c;border:1px solid #fed7d7}
.status.info{display:block;background:#ebf8ff;color:#2b6cb0;border:1px solid #bee3f8}
.template-badge{display:inline-block;background:#48bb78;color:#fff;padding:2px 8px;border-radius:4px;font-size:.75rem;font-weight:600;margin-left:8px}
.spinner{display:inline-block;width:16px;height:16px;border:2px solid #fff;border-top-color:transparent;border-radius:50%;animation:spin .6s linear infinite;vertical-align:middle;margin-right:6px}
@keyframes spin{to{transform:rotate(360deg)}}
</style>
</head>
<body>
<div class="container">
  <h1>代理受領通知書 自動作成ツール<small>CSVを読み込むだけで通知書Excelを一括生成（テンプレート内蔵）</small></h1>

  <div class="step">
    <div class="step-header">
      <div class="step-num" id="step1-num">1</div>
      <div class="step-title">CSVファイル選択 <span class="template-badge">テンプレート内蔵済</span></div>
    </div>
    <div class="file-area" id="csv-area">
      <input type="file" id="csv-input" accept=".csv,.CSV">
      <div class="icon">&#128196;</div>
      <div class="label">国保連CSVファイル（*.CSV）をクリックまたはドロップで選択</div>
      <div class="loaded-info" id="csv-info" style="display:none"></div>
    </div>
  </div>

  <div class="step">
    <div class="step-header">
      <div class="step-num" id="step2-num">2</div>
      <div class="step-title">基本情報入力</div>
    </div>
    <div class="form-grid">
      <label for="issue-date">発行日</label>
      <input type="text" id="issue-date" placeholder="例: 2026年02月24日">
      <label for="receipt-date">受給日</label>
      <input type="text" id="receipt-date" placeholder="例: 令和7年11月20日">
      <label for="service-type">サービス種別</label>
      <div style="display:flex;gap:8px;">
        <select id="service-select" style="flex:1;">
          <option value="">--- CSVから自動判定 ---</option>
          <option value="共同生活援助">共同生活援助</option>
          <option value="短期入所">短期入所</option>
          <option value="生活介護">生活介護</option>
          <option value="就労継続支援Ａ型">就労継続支援Ａ型</option>
          <option value="就労継続支援Ｂ型">就労継続支援Ｂ型</option>
          <option value="就労移行支援">就労移行支援</option>
          <option value="居宅介護">居宅介護</option>
          <option value="__custom__">その他（手入力）</option>
        </select>
        <input type="text" id="service-custom" placeholder="手入力" style="flex:1;display:none;">
      </div>
    </div>
  </div>

  <div class="step" id="step3" style="display:none;">
    <div class="step-header">
      <div class="step-num" id="step3-num">3</div>
      <div class="step-title">データ確認 (<span id="data-count">0</span>件)</div>
    </div>
    <div class="data-table-wrap">
      <table>
        <thead>
          <tr>
            <th>#</th>
            <th>受給者番号</th>
            <th>氏名</th>
            <th>対象月</th>
            <th class="num">総費用額</th>
            <th class="num">利用者負担額</th>
            <th class="num">代理受領額</th>
            <th class="num">特定障害者給付費</th>
          </tr>
        </thead>
        <tbody id="data-tbody"></tbody>
      </table>
    </div>
  </div>

  <div class="step">
    <button class="btn-generate" id="btn-generate" disabled>Excel出力</button>
    <div class="status" id="status"></div>
  </div>
</div>

<script>
// ==== Embedded template ====
var TEMPLATE_B64 = "TEMPLATE_PLACEHOLDER";

// ==== State ====
var csvData = [];
var templateBuffer = null;

// ==== Service code master (障害福祉サービス種類コード) ====
var SERVICE_NAMES = {
  "11": "\\u5c45\\u5b85\\u4ecb\\u8b77",
  "12": "\\u91cd\\u5ea6\\u8a2a\\u554f\\u4ecb\\u8b77",
  "13": "\\u540c\\u884c\\u63f4\\u8b77",
  "14": "\\u884c\\u52d5\\u63f4\\u8b77",
  "21": "\\u77ed\\u671f\\u5165\\u6240",
  "22": "\\u751f\\u6d3b\\u4ecb\\u8b77",
  "23": "\\u65bd\\u8a2d\\u5165\\u6240\\u652f\\u63f4",
  "31": "\\u5c31\\u52b4\\u79fb\\u884c\\u652f\\u63f4",
  "32": "\\u5c31\\u52b4\\u7d99\\u7d9a\\u652f\\u63f4\\uff21\\u578b",
  "33": "\\u5171\\u540c\\u751f\\u6d3b\\u63f4\\u52a9",
  "34": "\\u5c31\\u52b4\\u5b9a\\u7740\\u652f\\u63f4",
  "35": "\\u5c31\\u52b4\\u9078\\u629e\\u652f\\u63f4",
  "36": "\\u81ea\\u7acb\\u8a13\\u7df4\\uff08\\u751f\\u6d3b\\u8a13\\u7df4\\uff09",
  "43": "\\u5171\\u540c\\u751f\\u6d3b\\u63f4\\u52a9",
  "51": "\\u5c31\\u52b4\\u7d99\\u7d9a\\u652f\\u63f4\\uff22\\u578b",
  "65": "\\u77ed\\u671f\\u5165\\u6240"
};

// ==== Init ====
document.addEventListener("DOMContentLoaded", function() {
  // Decode embedded template
  var binaryStr = atob(TEMPLATE_B64);
  var bytes = new Uint8Array(binaryStr.length);
  for (var i = 0; i < binaryStr.length; i++) bytes[i] = binaryStr.charCodeAt(i);
  templateBuffer = bytes.buffer;

  // Default issue date
  var now = new Date();
  document.getElementById("issue-date").value =
    now.getFullYear() + "\\u5e74" + String(now.getMonth()+1).padStart(2,"0") + "\\u6708" + String(now.getDate()).padStart(2,"0") + "\\u65e5";

  document.getElementById("service-select").addEventListener("change", function(e) {
    document.getElementById("service-custom").style.display = e.target.value === "__custom__" ? "" : "none";
  });
  document.getElementById("csv-input").addEventListener("change", handleCsvFile);
  document.getElementById("btn-generate").addEventListener("click", generateExcel);
});

// ==== CSV Parsing ====
function handleCsvFile(e) {
  var file = e.target.files[0];
  if (!file) return;
  var reader = new FileReader();
  reader.onload = function(ev) {
    try {
      var text = decodeShiftJIS(new Uint8Array(ev.target.result));
      csvData = parseCsv(text);
      document.getElementById("csv-area").classList.add("loaded");
      var info = document.getElementById("csv-info");
      info.style.display = "";
      info.textContent = "\\u2714 " + file.name + " \\u8aad\\u8fbc\\u6e08 (" + csvData.length + "\\u4ef6)";
      document.getElementById("step1-num").classList.add("done");
      renderDataTable();
      updateGenerateButton();
    } catch(err) {
      showStatus("error", "CSV\\u8aad\\u307f\\u8fbc\\u307f\\u30a8\\u30e9\\u30fc: " + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

function decodeShiftJIS(uint8arr) {
  try { return new TextDecoder("shift_jis").decode(uint8arr); }
  catch(e) { return new TextDecoder("utf-8").decode(uint8arr); }
}

function parseCsv(text) {
  var lines = text.split(/\\r?\\n/);
  var users = {};
  for (var li = 0; li < lines.length; li++) {
    var line = lines[li];
    if (!line.trim()) continue;
    var row = parseCSVLine(line);
    if (row.length < 3) continue;
    var recordId = row[2], subType = row.length > 3 ? row[3] : "";

    if (recordId === "J131" && subType === "01") {
      var userId = (row[7] || "").trim();
      var kanaName = (row[9] || "").trim();
      var targetMonth = (row[4] || "").trim();
      var serviceCost = parseInt(row[22], 10) || 0;
      var userBurden  = parseInt(row[23], 10) || 0;
      var proxyAmount = parseInt(row[29], 10) || 0;
      var specGrant   = parseInt(row[35], 10) || 0;

      var key = userId + "_" + targetMonth;
      if (!users[key]) {
        users[key] = { userId: userId, name: kanaName, month: targetMonth,
          serviceCost: serviceCost, userBurden: userBurden,
          proxyAmount: proxyAmount, specGrant: specGrant, serviceCode: "" };
      } else {
        users[key].serviceCost += serviceCost;
        users[key].userBurden  += userBurden;
        users[key].proxyAmount += proxyAmount;
        users[key].specGrant   += specGrant;
      }
    }

    // J131 type 02: service TYPE code (2-digit, e.g. "33" = 共同生活援助)
    if (recordId === "J131" && subType === "02") {
      var uid2 = (row[7] || "").trim();
      var mon2 = (row[4] || "").trim();
      var sc2 = (row[8] || "").trim();
      var k2 = uid2 + "_" + mon2;
      if (users[k2]) users[k2].serviceCode = sc2;
    }
  }
  return Object.values(users);
}

function parseCSVLine(line) {
  var result = [], current = "", inQuotes = false;
  for (var i = 0; i < line.length; i++) {
    var ch = line[i];
    if (inQuotes) {
      if (ch === '"') {
        if (i + 1 < line.length && line[i + 1] === '"') { current += '"'; i++; }
        else inQuotes = false;
      } else current += ch;
    } else {
      if (ch === '"') inQuotes = true;
      else if (ch === ",") { result.push(current); current = ""; }
      else current += ch;
    }
  }
  result.push(current);
  return result;
}

// ==== Data Table ====
function renderDataTable() {
  var tbody = document.getElementById("data-tbody");
  tbody.innerHTML = "";
  if (csvData.length === 0) { document.getElementById("step3").style.display = "none"; return; }
  document.getElementById("step3").style.display = "";
  document.getElementById("data-count").textContent = csvData.length;

  var tC = 0, tB = 0, tP = 0, tS = 0;
  for (var i = 0; i < csvData.length; i++) {
    var d = csvData[i];
    tC += d.serviceCost; tB += d.userBurden; tP += d.proxyAmount; tS += d.specGrant;
    var tr = document.createElement("tr");
    tr.innerHTML = "<td>" + (i+1) + "</td><td>" + esc(d.userId) + "</td><td>" + esc(d.name) +
      "</td><td>" + esc(formatMonth(d.month)) + "</td><td class=num>" + fmtNum(d.serviceCost) +
      "</td><td class=num>" + fmtNum(d.userBurden) + "</td><td class=num>" + fmtNum(d.proxyAmount) +
      "</td><td class=num>" + fmtNum(d.specGrant) + "</td>";
    tbody.appendChild(tr);
  }
  var tt = document.createElement("tr"); tt.className = "total-row";
  tt.innerHTML = "<td></td><td></td><td>\\u5408\\u8a08</td><td></td><td class=num>" + fmtNum(tC) +
    "</td><td class=num>" + fmtNum(tB) + "</td><td class=num>" + fmtNum(tP) +
    "</td><td class=num>" + fmtNum(tS) + "</td>";
  tbody.appendChild(tt);
  document.getElementById("step3-num").classList.add("done");
}

// ==== Excel Generation ====
async function generateExcel() {
  if (csvData.length === 0 || !templateBuffer) return;
  var btn = document.getElementById("btn-generate");
  btn.disabled = true;
  btn.innerHTML = "<span class=spinner></span>\\u751f\\u6210\\u4e2d...";
  showStatus("info", "Excel\\u30d5\\u30a1\\u30a4\\u30eb\\u3092\\u751f\\u6210\\u3057\\u3066\\u3044\\u307e\\u3059...");

  try {
    var issueDate = document.getElementById("issue-date").value;
    var receiptDate = document.getElementById("receipt-date").value;
    var serviceSelect = document.getElementById("service-select").value;
    var serviceCustom = document.getElementById("service-custom").value;
    var serviceNameOverride = serviceSelect === "__custom__" ? serviceCustom : (serviceSelect || "");

    var wb = new ExcelJS.Workbook();
    await wb.xlsx.load(templateBuffer);

    // FIX: Set default font to Yu Gothic (matching template)
    wb.properties = wb.properties || {};

    var userMasters = extractUserMasters(wb);
    var sourceSheet = wb.getWorksheet("\\u539f\\u672c") || wb.worksheets[0];

    // FIX: Resolve H15-H19 company values (some may be formulas referencing other sheets)
    // We need to read from 事業者情報 sheet directly for formula cells
    var companyVals = {};
    var bizSheet = wb.getWorksheet("\\u4e8b\\u696d\\u8005\\u60c5\\u5831");
    for (var r = 15; r <= 19; r++) {
      var cellVal = sourceSheet.getCell("H" + r).value;
      // If cell is a formula object, resolve it manually from 事業者情報
      if (cellVal && typeof cellVal === "object" && cellVal.formula) {
        cellVal = resolveCompanyFormula(cellVal.formula, bizSheet);
      }
      companyVals["H" + r] = cellVal;
    }

    // FIX: Save border styles from template for E26:H28 range (right border fix)
    var borderTemplates = {};
    ["E26","E27","E28","H26","H27","H28"].forEach(function(addr) {
      var c = sourceSheet.getCell(addr);
      if (c.style && c.style.border) {
        borderTemplates[addr] = JSON.parse(JSON.stringify(c.style.border));
      }
    });

    // FIX: Save E8 style for vertical alignment
    var e8Style = sourceSheet.getCell("E8").style ?
      JSON.parse(JSON.stringify(sourceSheet.getCell("E8").style)) : null;

    var sheetCount = 0;
    for (var di = 0; di < csvData.length; di++) {
      var data = csvData[di];
      var userId = String(data.userId).trim();

      // Lookup with ID normalization
      var kanjiName = "", municipality = "";
      var me = userMasters[userId] || userMasters[userId.replace(/^0+/, "")];
      if (me) { kanjiName = me.kanjiName || ""; municipality = me.municipality || ""; }
      var displayName = kanjiName || data.name;

      var safeName = displayName.replace(/[:\\\\\\/\\?\\*\\[\\]]/g, "").substring(0, 31);
      if (!safeName) safeName = "User_" + sheetCount;

      var finalName = safeName, suffix = 2;
      while (wb.getWorksheet(finalName)) { finalName = safeName.substring(0, 28) + "_" + suffix; suffix++; }

      var ns = wb.addWorksheet(finalName);
      copySheetStructure(sourceSheet, ns);

      // === CORRECT cell mapping (matching template layout) ===
      ns.getCell("D7").value = userId;
      ns.getCell("D8").value = displayName;
      ns.getCell("H4").value = "\\u767a\\u884c\\u65e5: " + issueDate;

      // Row 24: month + proxy amount
      var monthStr = data.month;
      if (monthStr && monthStr.length === 6) {
        ns.getCell("C24").value = parseInt(monthStr.substring(4), 10);
      }
      ns.getCell("F24").value = data.proxyAmount;

      // Row 26-28
      ns.getCell("E26").value = municipality;

      var serviceName = serviceNameOverride;
      if (!serviceName) {
        var rawCode = data.serviceCode || "";
        serviceName = SERVICE_NAMES[rawCode] || SERVICE_NAMES[rawCode.substring(0, 2)] || rawCode;
      }
      ns.getCell("E27").value = serviceName;
      ns.getCell("E28").value = receiptDate;

      // Row 29-33: financial data
      ns.getCell("H29").value = data.serviceCost;  // 総支給額（単位）
      ns.getCell("H30").value = data.serviceCost;  // 総サービス費(A)
      ns.getCell("H31").value = data.userBurden;   // 利用者負担額(B)
      ns.getCell("H32").value = data.specGrant;    // 特定障害者給付費(C)
      ns.getCell("H33").value = data.proxyAmount;  // 合計 = A - B + C

      // FIX: Company info - write resolved values (not formulas)
      for (var cr = 15; cr <= 19; cr++) ns.getCell("H" + cr).value = companyVals["H" + cr];

      // FIX: Restore right borders on E26:H28
      ["E26","E27","E28","H26","H27","H28"].forEach(function(addr) {
        if (borderTemplates[addr]) {
          var cell = ns.getCell(addr);
          var st = cell.style ? JSON.parse(JSON.stringify(cell.style)) : {};
          st.border = borderTemplates[addr];
          cell.style = st;
        }
      });

      // FIX: Restore E8 vertical alignment
      if (e8Style) {
        ns.getCell("E8").style = JSON.parse(JSON.stringify(e8Style));
      }

      // FIX: Correct print settings (remove fitToWidth/fitToHeight)
      if (ns.pageSetup) {
        delete ns.pageSetup.fitToWidth;
        delete ns.pageSetup.fitToHeight;
      }

      // FIX: Set default font for all cells to Yu Gothic
      ns.eachRow(function(row) {
        row.eachCell(function(cell) {
          if (cell.style) {
            var s = JSON.parse(JSON.stringify(cell.style));
            if (!s.font) s.font = {};
            if (!s.font.name || s.font.name === "Calibri") s.font.name = "\\u6e38\\u30b4\\u30b7\\u30c3\\u30af";
            cell.style = s;
          }
        });
      });

      sheetCount++;
    }

    // Remove master sheets
    ["\\u539f\\u672c", "\\u4e8b\\u696d\\u8005\\u60c5\\u5831", "\\u53d7\\u7d66\\u8005\\u60c5\\u5831", "Sheet1"].forEach(function(name) {
      var ws = wb.getWorksheet(name);
      if (ws) wb.removeWorksheet(ws.id);
    });

    var buf = await wb.xlsx.writeBuffer();
    var blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    var now = new Date();
    var ts = now.getFullYear() + String(now.getMonth()+1).padStart(2,"0") + String(now.getDate()).padStart(2,"0") +
      "_" + String(now.getHours()).padStart(2,"0") + String(now.getMinutes()).padStart(2,"0") + String(now.getSeconds()).padStart(2,"0");
    saveAs(blob, "\\u4ee3\\u7406\\u53d7\\u9818\\u901a\\u77e5\\u66f8_\\u4e00\\u62ec\\u51fa\\u529b_" + ts + ".xlsx");

    showStatus("success", "\\u2714 " + sheetCount + "\\u540d\\u5206\\u306eExcel\\u3092\\u751f\\u6210\\u3057\\u307e\\u3057\\u305f\\u3002\\u30c0\\u30a6\\u30f3\\u30ed\\u30fc\\u30c9\\u3092\\u78ba\\u8a8d\\u3057\\u3066\\u304f\\u3060\\u3055\\u3044\\u3002");
  } catch(err) {
    showStatus("error", "\\u30a8\\u30e9\\u30fc: " + err.message);
    console.error(err);
  } finally {
    btn.disabled = false;
    btn.textContent = "Excel\\u51fa\\u529b";
    updateGenerateButton();
  }
}

// ==== Resolve company formula (H17-H19 reference 事業者情報 sheet) ====
function resolveCompanyFormula(formula, bizSheet) {
  // Handle formulas like: 事業者情報!C5&"　"&事業者情報!C6
  // or: 事業者情報!C7&""
  if (!bizSheet) return "";
  var parts = formula.split("&");
  var result = "";
  for (var i = 0; i < parts.length; i++) {
    var p = parts[i].trim();
    // Match cell reference like 事業者情報!C5
    var m = p.match(/!([A-Z]+)(\\d+)/);
    if (m) {
      var cellAddr = m[1] + m[2];
      var val = bizSheet.getCell(cellAddr).value;
      result += (val !== null && val !== undefined) ? String(val) : "";
    } else {
      // Literal string like "　" or ""
      var sm = p.match(/^"(.*)"$/);
      if (sm) result += sm[1];
    }
  }
  return result;
}

// ==== Extract user masters with ID normalization ====
function extractUserMasters(wb) {
  var masters = {};
  var ws = wb.getWorksheet("\\u53d7\\u7d66\\u8005\\u60c5\\u5831");
  if (!ws) return masters;

  function addM(row, mc, uc, nc) {
    var muni = getCellText(ws, row, mc);
    var uid = getCellText(ws, row, uc);
    var name = getCellText(ws, row, nc);
    if (uid) {
      masters[uid] = { municipality: muni, kanjiName: name };
      var padded = uid.padStart(10, "0");
      if (padded !== uid) masters[padded] = { municipality: muni, kanjiName: name };
    }
  }

  for (var r = 3; r <= 30; r++) { addM(r, 1, 2, 3); addM(r, 5, 6, 7); }
  return masters;
}

function getCellText(ws, row, col) {
  var val = ws.getCell(row, col).value;
  if (val === null || val === undefined) return "";
  return String(val).trim();
}

// ==== Copy sheet structure ====
function copySheetStructure(src, dst) {
  if (src.columns) {
    src.columns.forEach(function(col, i) {
      if (col && col.width) dst.getColumn(i + 1).width = col.width;
    });
  }
  src.eachRow({ includeEmpty: true }, function(srcRow, rowNum) {
    var dstRow = dst.getRow(rowNum);
    dstRow.height = srcRow.height;
    srcRow.eachCell({ includeEmpty: true }, function(srcCell, colNum) {
      var dstCell = dstRow.getCell(colNum);
      dstCell.value = srcCell.value;
      if (srcCell.style) { try { dstCell.style = JSON.parse(JSON.stringify(srcCell.style)); } catch(e) {} }
    });
    dstRow.commit();
  });
  if (src._merges) {
    Object.keys(src._merges).forEach(function(key) {
      try { var m = src._merges[key]; dst.mergeCells(m.range || m); } catch(e) {}
    });
  }
  if (src.pageSetup) { try { dst.pageSetup = JSON.parse(JSON.stringify(src.pageSetup)); } catch(e) {} }
}

// ==== Helpers ====
function formatMonth(ym) { return (!ym || ym.length !== 6) ? ym : ym.substring(0,4) + "/" + ym.substring(4); }
function fmtNum(n) { return n.toLocaleString("ja-JP"); }
function esc(s) { if (!s) return ""; var el = document.createElement("span"); el.textContent = s; return el.innerHTML; }
function updateGenerateButton() { document.getElementById("btn-generate").disabled = !(csvData.length > 0 && templateBuffer); }
function showStatus(type, msg) { var el = document.getElementById("status"); el.className = "status " + type; el.textContent = msg; }
<\/script>
</body>
</html>"""

# Replace placeholder with actual base64
html_content = html_content.replace("TEMPLATE_PLACEHOLDER", template_b64)

# Fix escaped script tags
html_content = html_content.replace("<\\/script>", "</script>")

output_path = os.path.join(folder, "代理受領通知書.html")
with open(output_path, "w", encoding="utf-8") as f:
    f.write(html_content)

print(f"Done. Output: {output_path}")
print(f"File size: {len(html_content):,} bytes")
print(f"Template base64: {len(template_b64):,} bytes")
