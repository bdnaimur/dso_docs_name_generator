const now = new Date();

const date = now.toISOString().slice(0, 10); // YYYY-MM-DD
const time = now.toTimeString().slice(0, 8).replace(/:/g, "-"); // HH-MM-SS

let txtLines = [];
let errorLines = [];
let macroCommandName = [];
let macroCommandDocs = [];
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
function handleFile(e) {
  document.getElementById("output").innerHTML = "";
  const reader = new FileReader();
  reader.onload = (evt) => {
    const workbook = XLSX.read(new Uint8Array(evt.target.result), {
      type: "array",
    });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    console.log(rows);

    const BLOCK_SIZE = 15;
    // let txtLines = [];
    let docsCounter = 2; // starts from 2.1

    // for (let i = 0; i < rows.length; i += BLOCK_SIZE) {
    // for (let i = 0; i < rows.length; i++) {
      const block = rows.slice(0, rows.length);

      block.forEach((row, idx) => {
        const isValidPassport = row?.PassportNo.length !== 8;
        const isGenderExist =
          row?.Gender &&
          (row?.Gender.toUpperCase() == "M" ||
            row?.Gender.toUpperCase() == "F");
        // console.log("isGender", isGenderExist);

        if (isValidPassport && isGenderExist) {
          const prefix = idx === 0 ? "" : "Σ";
          txtLines.push(formatName(row, prefix));
          macroCommandName.push(`WINCMD("${formatName(row, "")}")`);
          macroCommandName.push(`SLEEP(1000)`);
        }
      });

      block.forEach((row, idx) => {
        const isValidPassport = row?.PassportNo.length !== 8;
        const isGenderExist =
          row?.Gender &&
          (row?.Gender.toUpperCase() == "M" ||
            row?.Gender.toUpperCase() == "F");
        if (isValidPassport && isGenderExist) {
          macroCommandDocs.push(`WINCMD("${formatDOCS(row, 0, docsCounter)}")`);
          macroCommandDocs.push(`SLEEP(1000)`);
          txtLines.push(formatDOCS(row, idx, docsCounter));
          docsCounter++; // 🔴 NEVER RESET
        } else {
          isValidPassport &&
            errorLines.push(
              `Passport issue | ${row.PassportNo} | ${row.LastName} | ${row.FirstName}`,
            );

          !isGenderExist &&
            errorLines.push(
              `Gender issue | ${row.PassportNo} | ${row.LastName} | ${row.FirstName}`,
            );
        }
      });
    // }

    document.getElementById("output").textContent = txtLines.join("\n");
    // console.log(errorLines);
    document.getElementById("errorOutput").textContent = errorLines.join("\n");
    // document.getElementById("macroCommand").textContent = macroCommandName.join("\n");
    document.getElementById("macroNameCommand").textContent =
      macroCommandName.join("\n");
    document.getElementById("macroDocsCommand").textContent =
      macroCommandDocs.join("\n");

    //  document.getElementById("output").innerHTML = `<br/>`
    //  document.getElementById("output").textContent = errorLines.join("\n")
    document.getElementById("downloadBtn").disabled = false;
    document.getElementById("macroNameBtn").disabled = false;
    document.getElementById("macroDocsBtn").disabled = false;

    if (txtLines.length > 0) {
      document.getElementById("excelFile").value = "";
    }
  };
  reader.readAsArrayBuffer(e.target.files[0]);
}

function formatName(row, index) {
  const title = row.Gender === "F" ? "MS" : "MR";

  // First line of every 15-line block → no Σ
//   const isFirstInBlock = index % 15 === 0;
//   const prefix = isFirstInBlock ? "" : "Σ";
  const prefix = index == 0 ? "" : "Σ";

  // console.log(row.FirstName);
  if (row?.LastName && row?.FirstName) {
    return `${prefix}-${row.LastName.trim().toUpperCase()}/${row.FirstName.trim().toUpperCase()} ${title}`;
  } else if (row?.LastName) {
    return `${prefix}-${row.LastName.trim().toUpperCase()}/${title}`;
  } else {
    return `${prefix}-${row.FirstName.trim().toUpperCase()}/${title}`;
  }
}

// function formatDOCS(row, counter) {
//   return `4DOCS/P/BGD/${row.PassportNo}/BGD/${row.DOB}/${row.Gender}/${formatDate(row.DOE)}/${row.LastName.toUpperCase()}/${row.FirstName.toUpperCase()}-${counter}.1`;
// }
function formatDOCS(row, idx, docsCounter) {
  const prefix = idx === 0 ? "" : "Σ";

  if (row?.LastName && row?.FirstName) {
    return `${prefix}4DOCS/P/BGD/${row.PassportNo}/BGD/${formatDate(
      row.DOB,
    )}/${row.Gender}/${formatDate(
      row.DOE,
    )}/${row?.LastName?.toUpperCase().trim()}/${row?.FirstName?.toUpperCase().trim()}-${docsCounter}.1`;
  } else {
    return `${prefix}4DOCS/P/BGD/${row.PassportNo}/BGD/${formatDate(
      row.DOB,
    )}/${row.Gender}/${formatDate(row.DOE)}/${
      row?.LastName?.toUpperCase().trim() || row?.FirstName?.toUpperCase()
    }/${"FNU"}-${docsCounter}.1`;
  }
}

function excelDateToJSDate(excelDate) {
  // If already text like 25-Feb-33
  if (typeof excelDate === "string") {
    const d = new Date(excelDate);
    if (!isNaN(d)) return d;
  }

  // If Excel serial number
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

  return `${String(d.getDate()).padStart(2, "0")}${
    months[d.getMonth()]
  }${String(d.getFullYear()).slice(-2)}`;
}

function downloadTXT(lines) {
  const blob = new Blob([lines.join("\n")], { type: "text/plain" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `PNR_DATA_${date}_${time}.txt`;
  a.click();
  URL.revokeObjectURL(url);
  console.log(errorLines);
}

function macroNameTXT(lines) {
  const blob = new Blob([lines.join("\n")], { type: "text/plain" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `MACRO_NAME_${date}_${time}.txt`;
  a.click();
  URL.revokeObjectURL(url);
  console.log(errorLines);
}
function macroDocsTXT(lines) {
  const blob = new Blob([lines.join("\n")], { type: "text/plain" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `MACRO_DOCS_${date}_${time}.txt`;
  a.click();
  URL.revokeObjectURL(url);
  console.log(errorLines);
}
