// /* ======================================================
//    DATE & TIME
// ====================================================== */
// const now = new Date();
// const date = now.toISOString().slice(0, 10);
// const time = now.toTimeString().slice(0, 8).replace(/:/g, "-");

// /* ======================================================
//    GLOBAL STATE
// ====================================================== */
// let txtLines = [];
// let errorLines = [];
// let macroCommandName = [];
// let macroCommandDocs = [];

// /* ======================================================
//    EXPECTED HEADER FIELDS
// ====================================================== */
// const EXPECTED_FIELDS = [
//   "passport",
//   "firstname",
//   "lastname",
//   "gender",
//   "title",
//   "dob",
//   "doe",
// ];

// /* ======================================================
//    FLEXIBLE HEADER HANDLING
// ====================================================== */
// function normalizeKey(str) {
//   return String(str).toLowerCase().replace(/[^a-z0-9]/g, "");
// }

// function autoMapHeaders(headers) {
//   const map = {};
//   headers.forEach((header) => {
//     const n = normalizeKey(header);
//     const match = EXPECTED_FIELDS.find((f) => n.includes(f));
//     if (match) map[match] = header;
//   });
//   return map;
// }

// let headerMap = {};

// function getRowValue(row, field) {
//   const header = headerMap[field];
//   return header ? row[header] ?? "" : "";
// }

// /* ======================================================
//    DOM EVENTS
// ====================================================== */
// document.getElementById("excelFile").addEventListener("change", handleFile);
// document.getElementById("downloadBtn").onclick = () => downloadTXT(txtLines);
// document.getElementById("macroNameBtn").onclick = () =>
//   macroNameTXT(macroCommandName);
// document.getElementById("macroDocsBtn").onclick = () =>
//   macroDocsTXT(macroCommandDocs);

// /* ======================================================
//    HEADER PREVIEW
// ====================================================== */
// function renderHeaderPreview(mapped) {
//   const el = document.getElementById("headerPreview");

//   let html = `
//     <h3>📌 Auto Header Mapping Preview</h3>
//     <table border="1" cellpadding="6" cellspacing="0">
//       <tr>
//         <th>Expected Field</th>
//         <th>Excel Header</th>
//         <th>Status</th>
//       </tr>
//   `;

//   EXPECTED_FIELDS.forEach((f) => {
//     html += `
//       <tr>
//         <td>${f}</td>
//         <td>${mapped[f] || "-"}</td>
//         <td>${mapped[f] ? "✅ Matched" : "❌ Missing"}</td>
//       </tr>
//     `;
//   });

//   html += "</table>";
//   el.innerHTML = html;
// }

// /* ======================================================
//    FILE HANDLER
// ====================================================== */
// function handleFile(e) {
//   txtLines = [];
//   errorLines = [];
//   macroCommandName = [];
//   macroCommandDocs = [];

//   document.getElementById("output").textContent = "";
//   document.getElementById("errorOutput").textContent = "";

//   const reader = new FileReader();

//   reader.onload = (evt) => {
//     const workbook = XLSX.read(new Uint8Array(evt.target.result), {
//       type: "array",
//     });

//     const sheet = workbook.Sheets[workbook.SheetNames[0]];
//     const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
//     if (!rows.length) return;

//     /* ---------- AUTO HEADER MAP ---------- */
//     headerMap = autoMapHeaders(Object.keys(rows[0]));
//     renderHeaderPreview(headerMap);

//     /* ---------- REQUIRED CHECK ---------- */
//     const REQUIRED = ["passport", "firstname", "lastname", "gender"];
//     const missing = REQUIRED.filter((f) => !headerMap[f]);
//     if (missing.length) {
//       alert("Missing required columns: " + missing.join(", "));
//       return;
//     }

//     let docsCounter = 2;

//     /* ---------- NAME BLOCK ---------- */
//     rows.forEach((row, idx) => {
//       const passport = getRowValue(row, "passport");
//       const gender = getRowValue(row, "gender").toUpperCase();

//       const validPassport = passport && passport.length !== 8;
//       const validGender = ["M", "F"].includes(gender);

//       if (validPassport && validGender) {
//         const prefix = idx === 0 ? "" : "Σ";
//         txtLines.push(formatName(row, prefix));
//         macroCommandName.push(`WINCMD("${formatName(row, "")}")`);
//         macroCommandName.push("SLEEP(1000)");
//       }
//     });

//     /* ---------- DOCS BLOCK ---------- */
//     rows.forEach((row, idx) => {
//       const passport = getRowValue(row, "passport");
//       const gender = getRowValue(row, "gender").toUpperCase();

//       const validPassport = passport && passport.length !== 8;
//       const validGender = ["M", "F"].includes(gender);

//       if (validPassport && validGender) {
//         txtLines.push(formatDOCS(row, idx, docsCounter));
//         macroCommandDocs.push(`WINCMD("${formatDOCS(row, 0, docsCounter)}")`);
//         macroCommandDocs.push("SLEEP(1000)");
//         docsCounter++;
//       } else {
//         if (!validPassport)
//           errorLines.push(
//             `Passport issue | ${passport} | ${getRowValue(row, "lastname")} | ${getRowValue(row, "firstname")}`
//           );
//         if (!validGender)
//           errorLines.push(
//             `Gender issue | ${passport} | ${getRowValue(row, "lastname")} | ${getRowValue(row, "firstname")}`
//           );
//       }
//     });

//     document.getElementById("output").textContent = txtLines.join("\n");
//     document.getElementById("errorOutput").textContent = errorLines.join("\n");
//     document.getElementById("macroNameCommand").textContent =
//       macroCommandName.join("\n");
//     document.getElementById("macroDocsCommand").textContent =
//       macroCommandDocs.join("\n");

//     document.getElementById("downloadBtn").disabled = false;
//     document.getElementById("macroNameBtn").disabled = false;
//     document.getElementById("macroDocsBtn").disabled = false;
//     document.getElementById("excelFile").value = "";
//   };

//   reader.readAsArrayBuffer(e.target.files[0]);
// }

// /* ======================================================
//    FORMATTERS
// ====================================================== */
// function formatName(row, prefix) {
//   const gender = getRowValue(row, "gender").toUpperCase();

//   // ✅ TITLE LOGIC (EXCEL → FALLBACK)
//   let title = getRowValue(row, "title").toUpperCase();
//   if (!title) {
//     title = gender === "F" ? "MS" : "MR";
//   }

//   const ln = getRowValue(row, "lastname");
//   const fn = getRowValue(row, "firstname");

//   if (ln && fn)
//     return `${prefix}-${ln.toUpperCase().trim()}/${fn.toUpperCase().trim()} ${title}`;
//   if (ln) return `${prefix}-${ln.toUpperCase().trim()}/${title}`;
//   return `${prefix}-${fn.toUpperCase().trim()}/${title}`;
// }

// function formatDOCS(row, idx, counter) {
//   const prefix = idx === 0 ? "" : "Σ";

//   return `${prefix}4DOCS/P/BGD/${getRowValue(row, "passport")}/BGD/${formatDate(
//     getRowValue(row, "dob")
//   )}/${getRowValue(row, "gender")}/${formatDate(
//     getRowValue(row, "doe")
//   )}/${(
//     getRowValue(row, "lastname") || getRowValue(row, "firstname")
//   )
//     .toUpperCase()
//     .trim()}/${getRowValue(row, "firstname") ? getRowValue(row, "firstname").toUpperCase().trim() : "FNU"}-${counter}.1`;
// }

// /* ======================================================
//    DATE HELPERS
// ====================================================== */
// function excelDateToJSDate(v) {
//   if (typeof v === "string") {
//     const d = new Date(v);
//     if (!isNaN(d)) return d;
//   }
//   return new Date((v - 25569) * 86400 * 1000);
// }

// function formatDate(v) {
//   const d = excelDateToJSDate(v);
//   const m = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"];
//   return `${String(d.getDate()).padStart(2, "0")}${m[d.getMonth()]}${String(
//     d.getFullYear()
//   ).slice(-2)}`;
// }

// /* ======================================================
//    DOWNLOADERS
// ====================================================== */
// function downloadFile(lines, name) {
//   const blob = new Blob([lines.join("\n")], { type: "text/plain" });
//   const url = URL.createObjectURL(blob);
//   const a = document.createElement("a");
//   a.href = url;
//   a.download = name;
//   a.click();
//   URL.revokeObjectURL(url);
// }

// function downloadTXT(l) {
//   downloadFile(l, `PNR_DATA_${date}_${time}.txt`);
// }
// function macroNameTXT(l) {
//   downloadFile(l, `MACRO_NAME_${date}_${time}.txt`);
// }
// function macroDocsTXT(l) {
//   downloadFile(l, `MACRO_DOCS_${date}_${time}.txt`);
// }


/* ======================================================
   DATE & TIME
====================================================== */
const now = new Date();
const date = now.toISOString().slice(0, 10);
const time = now.toTimeString().slice(0, 8).replace(/:/g, "-");

/* ======================================================
   GLOBAL STATE
====================================================== */
let txtLines = [];
let errorLines = [];
let macroCommandName = [];
let macroCommandDocs = [];

/* ======================================================
   EXPECTED HEADER FIELDS (ALL REQUIRED)
====================================================== */
const EXPECTED_FIELDS = [
  "passport",
  "firstname",
  "lastname",
  "gender",
  "title",
  "dob",
  "doe",
];

/* ======================================================
   FLEXIBLE HEADER HANDLING
====================================================== */
function normalizeKey(str) {
  return String(str).toLowerCase().replace(/[^a-z0-9]/g, "");
}

function autoMapHeaders(headers) {
  const map = {};
  headers.forEach((header) => {
    const n = normalizeKey(header);
    const match = EXPECTED_FIELDS.find((f) => n.includes(f));
    if (match) map[match] = header;
  });
  return map;
}

let headerMap = {};

function getRowValue(row, field) {
  const header = headerMap[field];
  return header ? row[header] ?? "" : "";
}

/* ======================================================
   DOM EVENTS
====================================================== */
document.getElementById("excelFile").addEventListener("change", handleFile);
document.getElementById("downloadBtn").onclick = () => downloadTXT(txtLines);
document.getElementById("macroNameBtn").onclick = () =>
  macroNameTXT(macroCommandName);
document.getElementById("macroDocsBtn").onclick = () =>
  macroDocsTXT(macroCommandDocs);

/* ======================================================
   HEADER PREVIEW
====================================================== */
function renderHeaderPreview(mapped) {
  const el = document.getElementById("headerPreview");

  let html = `
    <h3>📌 Auto Header Mapping Preview</h3>
    <table border="1" cellpadding="6" cellspacing="0">
      <tr>
        <th>Expected Field</th>
        <th>Excel Header</th>
        <th>Status</th>
      </tr>
  `;

  EXPECTED_FIELDS.forEach((f) => {
    html += `
      <tr>
        <td>${f}</td>
        <td>${mapped[f] || "-"}</td>
        <td>${mapped[f] ? "✅ Matched" : "❌ Missing"}</td>
      </tr>
    `;
  });

  html += "</table>";
  el.innerHTML = html;
}

/* ======================================================
   FILE HANDLER (STRICT MODE)
====================================================== */
function handleFile(e) {
  // Reset everything first
  txtLines = [];
  errorLines = [];
  macroCommandName = [];
  macroCommandDocs = [];

  document.getElementById("output").textContent = "";
  document.getElementById("errorOutput").textContent = "";
  document.getElementById("headerPreview").innerHTML = "";

  // Disable buttons until success
  document.getElementById("downloadBtn").disabled = true;
  document.getElementById("macroNameBtn").disabled = true;
  document.getElementById("macroDocsBtn").disabled = true;

  const reader = new FileReader();

  reader.onload = (evt) => {
    const workbook = XLSX.read(new Uint8Array(evt.target.result), {
      type: "array",
    });

    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    if (!rows.length) return;

    /* ---------- AUTO HEADER MAP ---------- */
    headerMap = autoMapHeaders(Object.keys(rows[0]));
    renderHeaderPreview(headerMap);

    /* ❌ HARD STOP IF ANY EXPECTED FIELD IS MISSING */
    const missingFields = EXPECTED_FIELDS.filter((f) => !headerMap[f]);

    if (missingFields.length > 0) {
      alert(
        "Upload stopped.\nMissing required columns:\n\n" +
          missingFields.join(", ")
      );

      // Reset file input so user must re-upload
      document.getElementById("excelFile").value = "";
      return; // ⛔ STOP ENTIRE FUNCTION
    }

    /* ======================================================
       PROCESSING STARTS ONLY IF ALL HEADERS EXIST
    ====================================================== */

    let docsCounter = 2;

    /* ---------- NAME BLOCK ---------- */
    rows.forEach((row, idx) => {
      const passport = getRowValue(row, "passport");
      const gender = getRowValue(row, "gender").toUpperCase();

      const validPassport = passport && passport.length !== 8;
      const validGender = ["M", "F"].includes(gender);

      if (validPassport && validGender) {
        const prefix = idx === 0 ? "" : "Σ";
        const nameLine = formatName(row, prefix);
        txtLines.push(nameLine);
        macroCommandName.push(`WINCMD("${formatName(row, "")}")`);
        macroCommandName.push("SLEEP(1000)");
      }
    });

    /* ---------- DOCS BLOCK ---------- */
    rows.forEach((row, idx) => {
      const passport = getRowValue(row, "passport");
      const gender = getRowValue(row, "gender").toUpperCase();

      const validPassport = passport && passport.length !== 8;
      const validGender = ["M", "F"].includes(gender);

      if (validPassport && validGender) {
        txtLines.push(formatDOCS(row, idx, docsCounter));
        macroCommandDocs.push(`WINCMD("${formatDOCS(row, 0, docsCounter)}")`);
        macroCommandDocs.push("SLEEP(1000)");
        docsCounter++;
      }
    });

    /* ---------- OUTPUT ---------- */
    document.getElementById("output").textContent = txtLines.join("\n");
    document.getElementById("macroNameCommand").textContent =
      macroCommandName.join("\n");
    document.getElementById("macroDocsCommand").textContent =
      macroCommandDocs.join("\n");

    // Enable buttons ONLY after success
    document.getElementById("downloadBtn").disabled = false;
    document.getElementById("macroNameBtn").disabled = false;
    document.getElementById("macroDocsBtn").disabled = false;

    document.getElementById("excelFile").value = "";
  };

  reader.readAsArrayBuffer(e.target.files[0]);
}

/* ======================================================
   FORMATTERS
====================================================== */
function formatName(row, prefix) {
  const gender = getRowValue(row, "gender").toUpperCase();

  // Use TITLE column, fallback to gender
  let title = getRowValue(row, "title").toUpperCase();
  if (!title) title = gender === "F" ? "MS" : "MR";

  const ln = getRowValue(row, "lastname");
  const fn = getRowValue(row, "firstname");

  if (ln && fn)
    return `${prefix}-${ln.toUpperCase().trim()}/${fn.toUpperCase().trim()} ${title}`;
  if (ln) return `${prefix}-${ln.toUpperCase().trim()}/${title}`;
  return `${prefix}-${fn.toUpperCase().trim()}/${title}`;
}

function formatDOCS(row, idx, counter) {
  const prefix = idx === 0 ? "" : "Σ";

  return `${prefix}4DOCS/P/BGD/${getRowValue(row, "passport")}/BGD/${formatDate(
    getRowValue(row, "dob")
  )}/${getRowValue(row, "gender")}/${formatDate(
    getRowValue(row, "doe")
  )}/${(
    getRowValue(row, "lastname") || getRowValue(row, "firstname")
  )
    .toUpperCase()
    .trim()}/${getRowValue(row, "firstname")
    ? getRowValue(row, "firstname").toUpperCase().trim()
    : "FNU"}-${counter}.1`;
}

/* ======================================================
   DATE HELPERS
====================================================== */
function excelDateToJSDate(v) {
  if (typeof v === "string") {
    const d = new Date(v);
    if (!isNaN(d)) return d;
  }
  return new Date((v - 25569) * 86400 * 1000);
}

function formatDate(v) {
  const d = excelDateToJSDate(v);
  const m = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"];
  return `${String(d.getDate()).padStart(2, "0")}${m[d.getMonth()]}${String(
    d.getFullYear()
  ).slice(-2)}`;
}

/* ======================================================
   DOWNLOADERS
====================================================== */
function downloadFile(lines, name) {
  const blob = new Blob([lines.join("\n")], { type: "text/plain" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  a.click();
  URL.revokeObjectURL(url);
}

function downloadTXT(l) {
  downloadFile(l, `PNR_DATA_${date}_${time}.txt`);
}
function macroNameTXT(l) {
  downloadFile(l, `MACRO_NAME_${date}_${time}.txt`);
}
function macroDocsTXT(l) {
  downloadFile(l, `MACRO_DOCS_${date}_${time}.txt`);
}
