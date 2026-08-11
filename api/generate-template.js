import ExcelJS from "exceljs";

export const config = { maxDuration: 30 };

// Country detection from location string
function getCountry(loc) {
  const l = loc.toLowerCase();
  const map = [
    [["uk","london","manchester","birmingham","edinburgh","bristol","england","glasgow","leeds","sheffield","newcastle","southampton","basildon","yeovil","luton"], "United Kingdom"],
    [["canada","toronto","vancouver","montreal","calgary","ottawa"], "Canada"],
    [["germany","berlin","munich","frankfurt","hamburg","düsseldorf","cologne","stuttgart"], "Germany"],
    [["france","paris","lyon","marseille","toulouse","bordeaux"], "France"],
    [["australia","sydney","melbourne","brisbane","perth","adelaide"], "Australia"],
    [["india","bangalore","bengaluru","mumbai","delhi","hyderabad","chennai","pune"], "India"],
    [["sweden","stockholm","gothenburg","malmö","västerås"], "Sweden"],
    [["remote"], "Remote"],
  ];
  for (const [keys, country] of map) {
    if (keys.some(k => l.includes(k))) return country;
  }
  return "United States";
}

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  try {
    const { jobs = [], locations = [], clientName = "[CLIENT NAME]", deadline = "[DEADLINE DATE]" } = req.body;

    if (!jobs.length || !locations.length) {
      return res.status(400).json({ error: "Jobs and locations are required" });
    }

    const wb = new ExcelJS.Workbook();
    wb.creator = "KellyOCG SupplierRate";
    const ws = wb.addWorksheet("Rate Survey");

    // ── Column widths ──
    ws.columns = [
      { key: "a", width: 20 },
      { key: "b", width: 22 },
      { key: "c", width: 40 },
      { key: "d", width: 28 },
      { key: "e", width: 28 },
      { key: "f", width: 28 },
    ];

    // ── Colors ──
    const NAVY       = "FF0F2340";
    const AMBER      = "FFFFF2CC";
    const GREEN_FILL = "FFE2EFDA";
    const HDR_BLUE   = "FF1F4E79";
    const LIGHT_GREY = "FFF2F2F2";
    const WHITE      = "FFFFFFFF";
    const BORDER_CLR = { style: "thin", color: { argb: "FFBFBFBF" } };

    function thinBorder() {
      return { top: BORDER_CLR, bottom: BORDER_CLR, left: BORDER_CLR, right: BORDER_CLR };
    }

    // ── ROW 1 — Banner ──
    ws.getRow(1).height = 36;
    ws.mergeCells("A1:D1");
    ws.mergeCells("E1:F1");

    const r1a = ws.getCell("A1");
    r1a.value = "KellyOCG  ·  MRA Supplier Rate Survey";
    r1a.font = { bold: true, color: { argb: "FFFFFFFF" }, size: 14, name: "Arial" };
    r1a.fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } };
    r1a.alignment = { horizontal: "left", vertical: "middle", indent: 1 };

    const r1e = ws.getCell("E1");
    r1e.value = `Client: ${clientName}`;
    r1e.font = { bold: true, color: { argb: "FFFFFFFF" }, size: 11, name: "Arial" };
    r1e.fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } };
    r1e.alignment = { horizontal: "right", vertical: "middle", indent: 1 };

    // ── ROW 2 — Instructions bar ──
    ws.getRow(2).height = 18;
    ws.mergeCells("A2:F2");
    const r2 = ws.getCell("A2");
    r2.value = "Please complete all green-shaded fields and return to ratecards@kellyocg.com";
    r2.font = { italic: true, color: { argb: "FF404040" }, size: 9, name: "Arial" };
    r2.fill = { type: "pattern", pattern: "solid", fgColor: { argb: AMBER } };
    r2.alignment = { horizontal: "left", vertical: "middle", indent: 1 };

    // ── ROWS 3-7 — Instructions + Supplier fields ──
    const instrLines = [
      "Instructions",
      "Please enter your feedback on competitive pay and bill hourly rates – in local currency.",
      "Please provide only a competitive rate, and refrain from providing a range.",
      "If your organization does not have competitive rate information or does not fill a specific job title, leave blank",
      "Return the completed workbook to KellyOCG at ratecards@kellyocg.com",
    ];
    const supplierLabels = ["Supplier Name:", "Point of Contact (name):", "Supplier E-mail address:"];

    for (let i = 0; i < 5; i++) {
      const rowNum = i + 3;
      ws.getRow(rowNum).height = 16;
      ws.mergeCells(`A${rowNum}:D${rowNum}`);

      const lCell = ws.getCell(`A${rowNum}`);
      lCell.value = instrLines[i];
      lCell.font = i === 0
        ? { bold: true, color: { argb: "FF0F2340" }, size: 10, name: "Arial" }
        : { color: { argb: "FF404040" }, size: 9, name: "Arial" };
      lCell.alignment = { vertical: "middle", indent: 1, wrapText: true };

      if (i < 3) {
        const lblCell = ws.getCell(`E${rowNum}`);
        lblCell.value = supplierLabels[i];
        lblCell.font = { bold: true, color: { argb: "FF1F1F1F" }, size: 9, name: "Arial" };
        lblCell.alignment = { horizontal: "right", vertical: "middle" };

        const valCell = ws.getCell(`F${rowNum}`);
        valCell.value = "[Enter Here]";
        valCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: GREEN_FILL } };
        valCell.border = thinBorder();
        valCell.font = { color: { argb: "FF808080" }, size: 9, name: "Arial", italic: true };
        valCell.alignment = { vertical: "middle", indent: 1 };
      }
    }

    // ── ROW 8 — spacer ──
    ws.getRow(8).height = 8;

    // ── ROW 9 — Return by ──
    ws.getRow(9).height = 16;
    const r9e = ws.getCell("E9");
    r9e.value = "Return by:";
    r9e.font = { bold: true, color: { argb: "FF0F2340" }, size: 9, name: "Arial" };
    r9e.alignment = { horizontal: "right", vertical: "middle" };

    const r9f = ws.getCell("F9");
    r9f.value = deadline;
    r9f.font = { bold: true, color: { argb: "FFC00000" }, size: 9, name: "Arial" };
    r9f.alignment = { horizontal: "left", vertical: "middle", indent: 1 };

    // ── ROW 10 — spacer ──
    ws.getRow(10).height = 8;

    // ── ROW 11 — Column headers ──
    ws.getRow(11).height = 40;
    const colHeaders = [
      "Country / Region",
      "Location",
      "Job Title\n(as shown in VMS)",
      "Level",
      "Recommended Hourly\nPay Rate\n(local currency)",
      "Recommended Hourly\nBill Rate\n(local currency)",
    ];
    ["A","B","C","D","E","F"].forEach((col, i) => {
      const cell = ws.getCell(`${col}11`);
      cell.value = colHeaders[i];
      cell.font = { bold: true, color: { argb: "FFFFFFFF" }, size: 9, name: "Arial" };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: HDR_BLUE } };
      cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
      cell.border = { top: { style: "thin", color: { argb: "FFFFFFFF" } }, bottom: { style: "thin", color: { argb: "FFFFFFFF" } }, left: { style: "thin", color: { argb: "FFFFFFFF" } }, right: { style: "thin", color: { argb: "FFFFFFFF" } } };
    });

    // ── DATA ROWS ──
    let rowNum = 12;
    let idx = 0;
    for (const loc of locations) {
      const country = getCountry(loc);
      for (const job of jobs) {
        const bg = idx % 2 === 0 ? LIGHT_GREY : WHITE;
        ws.getRow(rowNum).height = 15;

        const cells = [
          { col: "A", val: country, bold: false },
          { col: "B", val: loc, bold: false },
          { col: "C", val: job.fullTitle || job.title, bold: true },  // fullTitle preserves original
          { col: "D", val: job.level || "", bold: false },
        ];

        for (const { col, val, bold } of cells) {
          const cell = ws.getCell(`${col}${rowNum}`);
          cell.value = val;
          cell.font = { bold, size: 9, name: "Arial", color: { argb: "FF1F1F1F" } };
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: bg } };
          cell.alignment = { vertical: "middle", indent: 1 };
          cell.border = thinBorder();
        }

        // Pay Rate — green input cell
        const payCell = ws.getCell(`E${rowNum}`);
        payCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: GREEN_FILL } };
        payCell.border = thinBorder();
        payCell.alignment = { horizontal: "center", vertical: "middle" };
        payCell.numFmt = "#,##0.00";

        // Bill Rate — green input cell
        const billCell = ws.getCell(`F${rowNum}`);
        billCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: GREEN_FILL } };
        billCell.border = thinBorder();
        billCell.alignment = { horizontal: "center", vertical: "middle" };
        billCell.numFmt = "#,##0.00";

        rowNum++;
        idx++;
      }
    }

    // Freeze header row
    ws.views = [{ state: "frozen", ySplit: 11 }];

    // Stream the file
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", "attachment; filename=KellyOCG_MRA_Rate_Survey_Template.xlsx");

    await wb.xlsx.write(res);
    res.end();

  } catch (err) {
    console.error("Excel generation error:", err);
    res.status(500).json({ error: err.message });
  }
}
