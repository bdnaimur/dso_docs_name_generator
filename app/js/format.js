/* ================= DATE & TIME ================= */
const now = new Date();
const date = now.toISOString().slice(0, 10); // YYYY-MM-DD
const time = now.toTimeString().slice(0, 8).replace(/:/g, "-"); // HH-MM-SS

/* ================= GLOBAL ARRAYS ================= */
let txtLines = [];
let errorLines = [];
let macroCommandName = [];
let macroCommandDocs = [];

/* ================= FLEXIBLE KEY HANDLING ================= */
function normalizeKey(str) {
  return String(str)
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function getRowValue(row, searchKey) {
  const normalizedSearch = normalizeKey(searchKey);
  const foundKey = Object.keys(row).find((k) =>
    normalizeKey(k).includes(normalizedSearch),
  );
  return foundKey ? row[foundKey] : "";
}

/* ================= EVENTS ================= */
document.getElementById("excelFile").addEventListener("change", handleFile);

document.getElementById("downloadBtn").addEventListener("click", () => {
  downloadTXT(txtLines);
});

document.getElementById("macroNameBtn").addEventListener("click", () => {
  macroNameTXT(macroCommandName);
});

document.getElementById("macroDocsBtn").addEventListener("click", () => {
  macroDocsTXT(macroCommandDocs);
});

/* ================= FILE HANDLER ================= */
function handleFile(e) {
  txtLines = [];
  errorLines = [];
  macroCommandName = [];
  macroCommandDocs = [];

  document.getElementById("output").textContent = "";
  document.getElementById("errorOutput").textContent = "";

  const reader = new FileReader();

  reader.onload = (evt) => {
    const workbook = XLSX.read(new Uint8Array(evt.target.result), {
      type: "array",
    });

    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    let docsCounter = 2;

    /* ========= NAME BLOCK ========= */
    rows.forEach((row, idx) => {
      const passportNo = getRowValue(row, "passport");
      //   const gender = getRowValue(row, "gender").toUpperCase();
      const title = getRowValue(row, "title")?.toUpperCase();

      console.log("title", title);

      const isValidPassport = passportNo && passportNo.length !== 8;
      const isGenderExist = ["MR", "MRS", "MSTR", "MISS", "MS"].includes(title);

      if (isValidPassport && isGenderExist) {
        const prefix = idx === 0 ? "" : "Σ";
        txtLines.push(formatName(row, prefix));
        macroCommandName.push(`WINCMD("${formatName(row, "")}")`);
        macroCommandName.push(`SLEEP(1000)`);
      }
    });

    /* ========= DOCS BLOCK ========= */
    rows.forEach((row, idx) => {
      const passportNo = getRowValue(row, "passport");
      const gender = getRowValue(row, "gender").toUpperCase();

      const isValidPassport = passportNo && passportNo.length !== 8;
      const isGenderExist = ["M", "F"].includes(gender);

      if (isValidPassport && isGenderExist) {
        txtLines.push(formatDOCS(row, idx, docsCounter));
        macroCommandDocs.push(`WINCMD("${formatDOCS(row, 0, docsCounter)}")`);
        macroCommandDocs.push(`SLEEP(1000)`);
        docsCounter++;
      } else {
        if (!isValidPassport) {
          errorLines.push(
            `Passport issue | ${passportNo} | ${getRowValue(row, "lastname")} | ${getRowValue(row, "firstname")}`,
          );
        }

        if (!isGenderExist) {
          errorLines.push(
            `Gender issue | ${passportNo} | ${getRowValue(row, "lastname")} | ${getRowValue(row, "firstname")}`,
          );
        }
      }
    });

    document.getElementById("output").textContent = txtLines.join("\n");
    document.getElementById("errorOutput").textContent = errorLines.join("\n");
    document.getElementById("macroNameCommand").textContent =
      macroCommandName.join("\n");
    document.getElementById("macroDocsCommand").textContent =
      macroCommandDocs.join("\n");

    document.getElementById("downloadBtn").disabled = false;
    document.getElementById("macroNameBtn").disabled = false;
    document.getElementById("macroDocsBtn").disabled = false;

    if (txtLines.length > 0) {
      document.getElementById("excelFile").value = "";
    }
  };

  reader.readAsArrayBuffer(e.target.files[0]);
}

/* ================= FORMAT NAME ================= */
function formatName(row, prefix) {
  let title = getRowValue(row, "title")?.toUpperCase();
  const gender = getRowValue(row, "gender").toUpperCase();
  if (!title) {
    title = gender === "F" ? "MS" : "MR";
  }

  const lastName = getRowValue(row, "lastname");
  const firstName = getRowValue(row, "firstname");

  if (lastName && firstName) {
    return `${prefix}-${lastName.trim().toUpperCase()}/${firstName.trim().toUpperCase()} ${title}`;
  } else if (lastName) {
    return `${prefix}-${lastName.trim().toUpperCase()}/${title}`;
  } else {
    return `${prefix}-${firstName.trim().toUpperCase()}/${title}`;
  }
}

/* ================= FORMAT DOCS ================= */
function formatDOCS(row, idx, docsCounter) {
  const prefix = idx === 0 ? "" : "Σ";

  const passportNo = getRowValue(row, "passport");
  const gender = getRowValue(row, "gender");
  const dob = getRowValue(row, "dob");
  const doe = getRowValue(row, "doe");
  const lastName = getRowValue(row, "lastname");
  const firstName = getRowValue(row, "firstname");

  if (lastName && firstName) {
    return `${prefix}4DOCS/P/BGD/${passportNo}/BGD/${formatDate(
      dob,
    )}/${gender}/${formatDate(
      doe,
    )}/${lastName.toUpperCase().trim()}/${firstName.toUpperCase().trim()}-${docsCounter}.1`;
  } else {
    return `${prefix}4DOCS/P/BGD/${passportNo}/BGD/${formatDate(
      dob,
    )}/${gender}/${formatDate(
      doe,
    )}/${(lastName || firstName).toUpperCase().trim()}/FNU-${docsCounter}.1`;
  }
}

/* ================= DATE HELPERS ================= */
function excelDateToJSDate(excelDate) {
  if (typeof excelDate === "string") {
    const d = new Date(excelDate);
    if (!isNaN(d)) return d;
  }
  return new Date((excelDate - 25569) * 86400 * 1000);
}

function formatDate(excelValue) {
  const d = excelDateToJSDate(excelValue);
  const months = [
    "JAN",
    "FEB",
    "MAR",
    "APR",
    "MAY",
    "JUN",
    "JUL",
    "AUG",
    "SEP",
    "OCT",
    "NOV",
    "DEC",
  ];
  return `${String(d.getDate()).padStart(2, "0")}${months[d.getMonth()]}${String(d.getFullYear()).slice(-2)}`;
}

/* ================= DOWNLOADERS ================= */
function downloadTXT(lines) {
  downloadFile(lines, `PNR_DATA_${date}_${time}.txt`);
}

function macroNameTXT(lines) {
  downloadFile(lines, `MACRO_NAME_${date}_${time}.txt`);
}

function macroDocsTXT(lines) {
  downloadFile(lines, `MACRO_DOCS_${date}_${time}.txt`);
}

function downloadFile(lines, filename) {
  const blob = new Blob([lines.join("\n")], { type: "text/plain" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}
