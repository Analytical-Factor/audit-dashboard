(function () {
  "use strict";

  const MIME_TYPE =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  const ISO_DATE_PATTERN = /^\d{4}-\d{2}-\d{2}$/;
  const HEADER_FILL = "FF0A3553";
  const DIFFERENCE_FILL = "FFFFF0DC";

  function requireObject(value, fieldName) {
    if (!value || typeof value !== "object" || Array.isArray(value)) {
      throw new Error(`${fieldName} is missing from the report response.`);
    }
    return value;
  }

  function requireText(value, fieldName) {
    const text = String(value ?? "").trim();
    if (!text) throw new Error(`${fieldName} is missing from the report response.`);
    return text;
  }

  function requireIsoDate(value, fieldName) {
    const text = requireText(value, fieldName);
    if (!ISO_DATE_PATTERN.test(text)) {
      throw new Error(`${fieldName} is invalid in the report response.`);
    }
    const date = new Date(`${text}T00:00:00.000Z`);
    if (Number.isNaN(date.getTime()) || date.toISOString().slice(0, 10) !== text) {
      throw new Error(`${fieldName} is invalid in the report response.`);
    }
    return text;
  }

  function requireInteger(value, fieldName) {
    const number = Number(value);
    if (!Number.isSafeInteger(number)) {
      throw new Error(`${fieldName} is invalid in the report response.`);
    }
    return number;
  }

  function optionalNumber(value, fieldName) {
    if (value === null || value === undefined || value === "") return null;
    const number = Number(value);
    if (!Number.isFinite(number)) {
      throw new Error(`${fieldName} is invalid in the report response.`);
    }
    return Object.is(number, -0) ? 0 : number;
  }

  function normalizePayload(payload) {
    const body = requireObject(payload, "Report data");
    const report = requireObject(body.report, "Report metadata");
    const counts = requireObject(report.counts, "Report counts");
    if (!Array.isArray(body.rows) || body.rows.length === 0) {
      throw new Error("The report response contains no comparison rows.");
    }

    const extractedAt = new Date(requireText(report.extractedAt, "Extracted at"));
    if (Number.isNaN(extractedAt.getTime())) {
      throw new Error("Extracted at is invalid in the report response.");
    }

    const normalizedReport = {
      historyEnd: requireIsoDate(report.historyEnd, "History end"),
      periodStart: requireIsoDate(report.periodStart, "Period start"),
      periodEnd: requireIsoDate(report.periodEnd, "Period end"),
      periodEndExclusive: requireIsoDate(
        report.periodEndExclusive,
        "Exclusive period end"
      ),
      extractedAt,
      demantraSource: requireText(report.demantraSource, "Demantra source"),
      counts: {
        afDpCombinations: requireInteger(
          counts.afDpCombinations,
          "AF DP combination count"
        ),
        demantraCombinations: requireInteger(
          counts.demantraCombinations,
          "Demantra combination count"
        ),
        commonCombinations: requireInteger(
          counts.commonCombinations,
          "Common combination count"
        ),
        reportRows: requireInteger(counts.reportRows, "Report row count")
      }
    };

    const rows = body.rows.map((row, index) => {
      const value = requireObject(row, `Report row ${index + 1}`);
      if (typeof value.hasDifference !== "boolean") {
        throw new Error(`hasDifference is invalid in report row ${index + 1}.`);
      }
      const normalized = {
        itemPartId: requireInteger(value.itemPartId, `itemPartId in row ${index + 1}`),
        item: requireText(value.item, `item in row ${index + 1}`),
        emrAbcClass: requireText(
          value.emrAbcClass,
          `emrAbcClass in row ${index + 1}`
        ),
        adjustedForecast: optionalNumber(
          value.adjustedForecast,
          `adjustedForecast in row ${index + 1}`
        ),
        afDpOverrideFixed: optionalNumber(
          value.afDpOverrideFixed,
          `afDpOverrideFixed in row ${index + 1}`
        ),
        percentDiff: optionalNumber(
          value.percentDiff,
          `percentDiff in row ${index + 1}`
        ),
        qtyDiff: optionalNumber(value.qtyDiff, `qtyDiff in row ${index + 1}`),
        hasDifference: value.hasDifference
      };
      const expectedDifference = normalized.qtyDiff !== null && normalized.qtyDiff !== 0;
      if (normalized.hasDifference !== expectedDifference) {
        throw new Error(`Difference status is inconsistent in report row ${index + 1}.`);
      }
      return normalized;
    });

    if (normalizedReport.demantraSource !== "oracle") {
      throw new Error("The report response was not generated from Oracle.");
    }
    if (normalizedReport.counts.reportRows !== rows.length) {
      throw new Error("The report row count does not match the returned data.");
    }
    return { report: normalizedReport, rows };
  }

  function longMonthLabel(isoDate) {
    const [year, month] = isoDate.split("-").map(Number);
    return `${new Date(Date.UTC(year, month - 1, 1)).toLocaleString("en-US", {
      month: "long",
      timeZone: "UTC"
    })} ${year}`;
  }

  function compactMonthLabel(isoDate) {
    const [year, month] = isoDate.split("-").map(Number);
    return `${new Date(Date.UTC(year, month - 1, 1)).toLocaleString("en-US", {
      month: "long",
      timeZone: "UTC"
    })}${String(year).slice(-2)}`;
  }

  function extractDayLabel(extractedAt) {
    return extractedAt.toLocaleDateString("en-US", {
      month: "long",
      day: "numeric",
      timeZone: "UTC"
    });
  }

  function discrepancyText(row) {
    const totalDifference = Math.abs(row.qtyDiff).toFixed(4);
    const percentDifference = row.percentDiff === null
      ? "N/A"
      : `${(Math.abs(row.percentDiff) * 100).toFixed(2)}%`;
    return `Part #${row.item}: Total difference of ${totalDifference} (${percentDifference}) between the AF DP Tool and Demantra.`;
  }

  function reportTitle(report) {
    return `AF DP Tool vs Demantra - ${compactMonthLabel(report.periodStart)}-${compactMonthLabel(report.periodEnd)} Comparison - Aligned to Synced Logic - ${extractDayLabel(report.extractedAt)}`;
  }

  function buildFilename(report) {
    return `${reportTitle(report)}.xlsx`;
  }

  function styleHeader(row) {
    row.height = 32;
    row.font = { bold: true, color: { argb: "FFFFFFFF" } };
    row.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: HEADER_FILL }
    };
    row.alignment = { vertical: "middle", wrapText: true };
  }

  function addGuide(workbook, report, rows) {
    const guide = workbook.addWorksheet("Guide", {
      views: [{ showGridLines: false }]
    });
    guide.columns = [{ width: 28 }, { width: 110 }];
    guide.mergeCells("A1:B1");
    guide.getCell("A1").value = reportTitle(report);
    styleHeader(guide.getRow(1));

    const discrepancies = rows
      .filter(row => row.hasDifference)
      .slice()
      .sort((left, right) =>
        left.itemPartId - right.itemPartId || left.item.localeCompare(right.item)
      );
    const discrepancyCount = discrepancies.length;
    const resultText = discrepancyCount === 0
      ? "The values are as expected. No differences were found between the AF DP Tool and Demantra when applying the sync logic described above."
      : `The values are as expected. ${discrepancyCount} ${discrepancyCount === 1 ? "difference was" : "differences were"} found between the AF DP Tool and Demantra when applying the sync logic described above.`;
    const guideRows = [
      ["Guide:", ""],
      ["", `This file compares the Demantra AF_DP_OVERRIDE values against the Adjusted Forecast column in the AF DP Tool for the period ${longMonthLabel(report.periodStart)} through ${longMonthLabel(report.periodEnd)}.`],
      [],
      ["", "The Demantra AF_DP_OVERRIDE values were evaluated according to the sync logic:"],
      ["Sync Logic:", "Demantra values should be overridden only for months where the Adjusted Forecast value in the AF DP Tool is not null."],
      [],
      ["Calculation Method for column \"AF_DP_OVERRIDE (Demantra, Fixed)\":", ""],
      ["", "For each item and month:"],
      ["", "• Check whether the Adjusted Forecast value in the AF DP Tool is null."],
      ["", "• If the value is null, the existing AF_DP_OVERRIDE value in Demantra is considered to be from an old sync and should not be included in the comparison or sum calculations."],
      ["", "• Only months where the Adjusted Forecast value is not null are included in the calculation."],
      [],
      ["Result:", resultText],
      ...(discrepancyCount > 0
        ? [["Discrepancies Identified:", ""], ...discrepancies.map(row => ["", discrepancyText(row)])]
        : [])
    ];
    guideRows.forEach(values => guide.addRow(values));
    guide.mergeCells("A2:B2");
    guide.mergeCells("A8:B8");
    if (discrepancyCount > 0) guide.mergeCells("A15:B15");
    const sectionHeadings = ["A2", "A6", "A8", "A14"];
    if (discrepancyCount > 0) sectionHeadings.push("A15");
    sectionHeadings.forEach(address => {
      guide.getCell(address).font = { bold: true, color: { argb: HEADER_FILL } };
    });
    const bold = { bold: true };
    const setRichText = (address, fragments) => {
      guide.getCell(address).value = { richText: fragments.map(([text, isBold]) => ({
        text,
        ...(isBold ? { font: bold } : {})
      })) };
    };
    setRichText("B3", [
      ["This file compares the ", false], ["Demantra AF_DP_OVERRIDE", true],
      [" values against the ", false], ["Adjusted Forecast", true],
      [" column in the ", false], ["AF DP Tool", true],
      [" for the period ", false],
      [`${longMonthLabel(report.periodStart)} through ${longMonthLabel(report.periodEnd)}`, true],
      [".", false]
    ]);
    setRichText("B5", [["The Demantra ", false], ["AF_DP_OVERRIDE", true], [" values were evaluated according to the sync logic:", false]]);
    setRichText("B6", [
      ["Demantra values should be overridden only for months where the ", false],
      ["Adjusted Forecast", true], [" value in the ", false], ["AF DP Tool", true],
      [" is ", false], ["not null", true], [".", false]
    ]);
    setRichText("B10", [["• Check whether the ", false], ["Adjusted Forecast", true], [" value in the ", false], ["AF DP Tool", true], [" is null.", false]]);
    setRichText("B11", [["• If the value is null, the existing ", false], ["AF_DP_OVERRIDE", true], [" value in Demantra is considered to be from an old sync and should not be included in the comparison or sum calculations.", false]]);
    setRichText("B12", [["• Only months where the ", false], ["Adjusted Forecast", true], [" value in the ", false], ["AF DP Tool", true], [" is not null are included in the calculation.", false]]);
    guide.eachRow((row, rowNumber) => {
      if (rowNumber > 1) row.alignment = { vertical: "top", wrapText: true };
    });
  }

  function addComparison(workbook, rows) {
    const sheet = workbook.addWorksheet("Override Comparison", {
      views: [{ state: "frozen", ySplit: 1 }]
    });
    sheet.columns = [
      { header: "item_part_id", key: "itemPartId", width: 16 },
      { header: "item", key: "item", width: 24 },
      { header: "emr_abc_class", key: "emrAbcClass", width: 16 },
      { header: "Adjusted Forecast (AF DP Tool)", key: "adjustedForecast", width: 30, style: { numFmt: "0.00" } },
      { header: "AF_DP_OVERRIDE (Demantra, Fixed)", key: "afDpOverrideFixed", width: 34, style: { numFmt: "0.00" } },
      { header: "% Diff", key: "percentDiff", width: 14, style: { numFmt: "0.00%" } },
      { header: "QTY Diff", key: "qtyDiff", width: 16, style: { numFmt: "0.00" } }
    ];
    styleHeader(sheet.getRow(1));

    rows
      .slice()
      .sort((left, right) =>
        left.itemPartId - right.itemPartId || left.item.localeCompare(right.item)
      )
      .forEach(value => {
        const row = sheet.addRow(value);
        if (value.hasDifference) {
          row.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: DIFFERENCE_FILL }
          };
        }
      });
    sheet.autoFilter = {
      from: "A1",
      to: `G${Math.max(sheet.rowCount, 1)}`
    };
  }

  async function buildWorkbook(payload) {
    if (!window.ExcelJS) {
      throw new Error("The Excel report library could not be loaded.");
    }
    const normalized = normalizePayload(payload);
    const workbook = new window.ExcelJS.Workbook();
    workbook.creator = "AF Demand Planning";
    workbook.created = normalized.report.extractedAt;
    workbook.modified = normalized.report.extractedAt;
    addGuide(workbook, normalized.report, normalized.rows);
    addComparison(workbook, normalized.rows);
    return { workbook, report: normalized.report };
  }

  function downloadBuffer(buffer, filename) {
    const objectUrl = URL.createObjectURL(
      new Blob([buffer], { type: MIME_TYPE })
    );
    const link = document.createElement("a");
    link.href = objectUrl;
    link.download = filename;
    link.hidden = true;
    document.body.appendChild(link);
    link.click();
    link.remove();
    window.setTimeout(() => URL.revokeObjectURL(objectUrl), 0);
  }

  async function download(payload) {
    const { workbook, report } = await buildWorkbook(payload);
    const buffer = await workbook.xlsx.writeBuffer();
    downloadBuffer(buffer, buildFilename(report));
  }

  window.AfdpOverrideReport = { buildWorkbook, download };
})();
