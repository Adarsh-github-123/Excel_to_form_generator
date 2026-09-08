let rows = [];
const fields = ["Name", "Address", "Phone Number", "Email"];
const samplePath = "name.xlsx";

function escapeHtml(s) {
  return String(s ?? "").replace(
    /[&<>"']/g,
    (c) =>
      ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[
        c
      ])
  );
}
function normalizeRows(data) {
  return data
    .map((original) => {
      const out = {};
      const keys = Object.keys(original);
      fields.forEach((field) => {
        const key = keys.find(
          (k) => k.trim().toLowerCase() === field.toLowerCase()
        );
        out[field] = key !== undefined ? String(original[key] ?? "") : "";
      });
      return out;
    })
    .filter((r) => Object.values(r).some((v) => v !== ""));
}
function render() {
  const box = document.getElementById("forms"),
    name = document.getElementById("formName").value.trim() || "Student Form";
  document.getElementById("countBadge").textContent = `${rows.length} form${
    rows.length === 1 ? "" : "s"
  }`;
  document.getElementById("downloadAllBtn").disabled = !rows.length;
  document.getElementById("status").textContent = rows.length
    ? `${rows.length} form${rows.length === 1 ? "" : "s"} ready to review.`
    : "No data loaded.";
  if (!rows.length) {
    box.innerHTML =
      '<div class="empty">Upload an Excel file to generate forms.</div>';
    return;
  }
  box.innerHTML = rows
    .map(
      (row, i) => `
    <div class="form-card">
      <div class="form-banner"><h3>${escapeHtml(name)}</h3></div>
      <div class="form-body">
        ${fields
          .map(
            (f) => `<div class="field"><label>${escapeHtml(f)}</label>
          <input value="${escapeHtml(
            row[f]
          )}" data-i="${i}" data-field="${escapeHtml(f)}"></div>`
          )
          .join("")}
        <div class="card-actions"><button class="primary individual-btn" data-index="${i}">Download this form</button></div>
      </div>
    </div>`
    )
    .join("");
  box.querySelectorAll("input").forEach((input) =>
    input.addEventListener("input", (e) => {
      rows[+e.target.dataset.i][e.target.dataset.field] = e.target.value;
    })
  );
  box
    .querySelectorAll(".individual-btn")
    .forEach((btn) =>
      btn.addEventListener("click", () => downloadSingle(+btn.dataset.index))
    );
}
async function readWorkbookFile(file) {
  const buffer = await file.arrayBuffer();
  const wb = XLSX.read(buffer, { type: "array" });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  return normalizeRows(XLSX.utils.sheet_to_json(sheet, { defval: "" }));
}
async function loadSample() {
  try {
    const response = await fetch(samplePath);
    if (!response.ok) throw new Error("Sample file not found");
    const buffer = await response.arrayBuffer();
    const wb = XLSX.read(buffer, { type: "array" });
    rows = normalizeRows(
      XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" })
    );
    render();
  } catch (e) {
    document.getElementById("status").textContent =
      "Sample Excel could not be loaded. Upload an Excel file to begin.";
    render();
  }
}
document.getElementById("excelFile").addEventListener("change", async (e) => {
  const file = e.target.files[0];
  if (!file) return;
  try {
    rows = await readWorkbookFile(file);
    render();
  } catch (err) {
    alert("Could not read the Excel file. Please check its format.");
  }
});
document.getElementById("sampleBtn").addEventListener("click", loadSample);
document.getElementById("formName").addEventListener("input", render);

const menu = document.getElementById("downloadMenu");
document
  .getElementById("downloadAllBtn")
  .addEventListener("click", () => menu.classList.toggle("show"));
document.addEventListener("click", (e) => {
  if (!e.target.closest(".download-menu")) menu.classList.remove("show");
});
document.getElementById("downloadPdf").addEventListener("click", () => {
  menu.classList.remove("show");
  downloadAll();
});
document.getElementById("downloadIndividual").addEventListener("click", () => {
  menu.classList.remove("show");
  rows.forEach((_, i) => downloadSingle(i));
});

function createDoc(row) {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF();
  const title =
    document.getElementById("formName").value.trim() || "Student Form";
  doc.setFillColor(91, 91, 214);
  doc.roundedRect(15, 15, 180, 28, 5, 5, "F");
  doc.setTextColor(255, 255, 255);
  doc.setFontSize(20);
  doc.setFont(undefined, "bold");
  doc.text(title, 25, 33);
  doc.setTextColor(35, 43, 58);
  let y = 65;
  fields.forEach((f) => {
    doc.setFontSize(10);
    doc.setFont(undefined, "bold");
    doc.setTextColor(91, 98, 116);
    doc.text(f.toUpperCase(), 25, y);
    doc.setFontSize(12);
    doc.setFont(undefined, "normal");
    doc.setTextColor(31, 41, 55);
    const lines = doc.splitTextToSize(String(row[f] ?? ""), 155);
    doc.text(lines, 25, y + 8);
    y += Math.max(25, lines.length * 7 + 17);
  });
  return doc;
}
function safeName() {
  return (
    document.getElementById("formName").value.trim() || "Student_Form"
  ).replace(/[^a-z0-9]+/gi, "_");
}
function downloadSingle(i) {
  if (!rows[i]) return;
  createDoc(rows[i]).save(`${safeName()}_${i + 1}.pdf`);
}
function downloadAll() {
  if (!rows.length) return;
  const doc = createDoc(rows[0]);
  for (let i = 1; i < rows.length; i++) {
    doc.addPage();
    const page = createDoc(rows[i]);
    const count = page.internal.getNumberOfPages();
    doc.setPage(1); /* content copied below */
    // Rebuild current page content using the same renderer.
    doc.setFillColor(91, 91, 214);
    doc.roundedRect(15, 15, 180, 28, 5, 5, "F");
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(20);
    doc.setFont(undefined, "bold");
    doc.text(
      document.getElementById("formName").value.trim() || "Student Form",
      25,
      33
    );
    doc.setTextColor(35, 43, 58);
    let y = 65;
    fields.forEach((f) => {
      doc.setFontSize(10);
      doc.setFont(undefined, "bold");
      doc.setTextColor(91, 98, 116);
      doc.text(f.toUpperCase(), 25, y);
      doc.setFontSize(12);
      doc.setFont(undefined, "normal");
      doc.setTextColor(31, 41, 55);
      const lines = doc.splitTextToSize(String(rows[i][f] ?? ""), 155);
      doc.text(lines, 25, y + 8);
      y += Math.max(25, lines.length * 7 + 17);
    });
  }
  doc.save(`${safeName()}_All_Forms.pdf`);
}
loadSample();
