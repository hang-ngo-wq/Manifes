/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import ExcelJS from "exceljs";
import { BillingCalculation } from "../types";

export async function exportShipperExcelReport(
  shipperName: string,
  rows: BillingCalculation[],
  exchangeRate: number,
  month: number,
  year: number
): Promise<void> {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("BẢNG KÊ HOÁ ĐƠN");

  // Show Grid Lines explicitly
  worksheet.views = [{ showGridLines: true }];

  // 1. Spacing rows 1 & 2
  worksheet.addRow([]);
  worksheet.addRow([]);

  // 2. Report Main Title (Exactly on Row 3, centered, merged across columns A to N)
  worksheet.mergeCells("A3:N3");
  const titleCell = worksheet.getCell("A3");
  titleCell.value = `BẢNG KÊ HOÁ ĐƠN THÁNG ${month} NĂM ${year} - ${shipperName.toUpperCase()}`;
  titleCell.font = { name: "Arial", size: 16, bold: true, color: { argb: "FF000000" } };
  titleCell.alignment = { vertical: "middle", horizontal: "center" };
  worksheet.getRow(3).height = 32;

  // Row 4 spacing
  worksheet.addRow([]);

  // 3. Row 5: NGÀY XUẤT and TỔNG SỐ BILL THEO MNF
  // We explicitly configure row 5 height and styling
  const row5 = worksheet.getRow(5);
  row5.height = 20;

  // B5: NGÀY XUẤT, highlighted in yellow
  const dateLabelCell = worksheet.getCell("B5");
  dateLabelCell.value = "NGÀY XUẤT";
  dateLabelCell.font = { name: "Arial", size: 10, bold: true };
  dateLabelCell.alignment = { vertical: "middle", horizontal: "center" };
  dateLabelCell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFFFF200" }, // Bright yellow
  };

  // K5: TỔNG SỐ BILL THEO, highlighted in yellow
  const billLabelCell = worksheet.getCell("K5");
  billLabelCell.value = "TỔNG SỐ BILL THEO";
  billLabelCell.font = { name: "Arial", size: 10, bold: true };
  billLabelCell.alignment = { vertical: "middle", horizontal: "right" };
  billLabelCell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFFFF200" }, // Bright yellow
  };

  // L5: MNF
  const mnfCell = worksheet.getCell("L5");
  mnfCell.value = "MNF";
  mnfCell.font = { name: "Arial", size: 10, bold: true };
  mnfCell.alignment = { vertical: "middle", horizontal: "left" };

  // Row 6 spacing
  worksheet.addRow([]);

  // 4. Headers definition (Row 7)
  const headers = [
    { name: "NO", key: "no", width: 6, align: "center" },
    { name: "ETD", key: "etd", width: 14, align: "center" },
    { name: "ETA", key: "eta", width: 14, align: "center" },
    { name: "MAWB", key: "mawb", width: 22, align: "center" },
    { name: "BILL", key: "bill", width: 18, align: "center" },
    { name: "ROUTE", key: "route", width: 14, align: "center" },
    { name: "WEIGHT", key: "weight", width: 12, align: "right" },
    { name: "Unit price\n(USD/kg)", key: "unitPrice", width: 15, align: "right" },
    { name: "Freight charge\n(VAT 0%)", key: "freightCharge", width: 16, align: "right" },
    { name: "Hanlding charge", key: "handlingCharge", width: 16, align: "right" },
    { name: "Warehous\ne charge", key: "warehouseCharge", width: 14, align: "right" },
    { name: "Oth.Chg", key: "otherCharge", width: 12, align: "right" },
    { name: "TOTAL", key: "total", width: 16, align: "right" },
    { name: "Mark\n(HAWB)", key: "markHawb", width: 15, align: "center" },
  ];

  const headerRowNumber = 7;
  const headerRow = worksheet.getRow(headerRowNumber);
  headerRow.height = 32;

  headers.forEach((h, index) => {
    const colIndex = index + 1;
    const cell = headerRow.getCell(colIndex);
    cell.value = h.name;
    cell.font = { name: "Arial", size: 9, bold: true, color: { argb: "FF000000" } };
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFDDEBF7" }, // Soft, clean light blue from the image
    };
    cell.border = {
      top: { style: "thin", color: { argb: "FF000000" } },
      bottom: { style: "thin", color: { argb: "FF000000" } },
      left: { style: "thin", color: { argb: "FF000000" } },
      right: { style: "thin", color: { argb: "FF000000" } },
    };
  });

  // Apply widths
  headers.forEach((h, index) => {
    worksheet.getColumn(index + 1).width = h.width;
  });

  // 5. Populate Data Rows starting from row 8
  const dataStartRow = 8;
  const borderThin = {
    top: { style: "thin" as ExcelJS.BorderStyle, color: { argb: "FF7F7F7F" } },
    bottom: { style: "thin" as ExcelJS.BorderStyle, color: { argb: "FF7F7F7F" } },
    left: { style: "thin" as ExcelJS.BorderStyle, color: { argb: "FF7F7F7F" } },
    right: { style: "thin" as ExcelJS.BorderStyle, color: { argb: "FF7F7F7F" } },
  };

  rows.forEach((row, i) => {
    const rNumber = dataStartRow + i;
    const worksheetRow = worksheet.getRow(rNumber);
    worksheetRow.height = 20;

    // Formatting cell values
    worksheetRow.getCell(1).value = i + 1; // NO
    worksheetRow.getCell(2).value = row.etd || row.date || ""; // ETD
    worksheetRow.getCell(3).value = row.eta || ""; // ETA
    worksheetRow.getCell(4).value = row.mawbNo; // MAWB
    worksheetRow.getCell(5).value = row.bill || row.hawb || ""; // BILL
    worksheetRow.getCell(6).value = row.route; // ROUTE
    worksheetRow.getCell(7).value = row.weight; // WEIGHT
    worksheetRow.getCell(8).value = row.unitPrice; // Unit Price

    // Freight charge (I) formula: G * H
    worksheetRow.getCell(9).value = {
      formula: `G${rNumber}*H${rNumber}`,
      result: row.freightCharge
    };

    worksheetRow.getCell(10).value = row.handlingCharge; // Handling charge
    worksheetRow.getCell(11).value = row.warehouseCharge; // Warehouse charge
    worksheetRow.getCell(12).value = row.otherCharge > 0 ? row.otherCharge : ""; // Other charge (leave empty if 0)

    // TOTAL (M) formula: I + J + K + L
    worksheetRow.getCell(13).value = {
      formula: `I${rNumber}+J${rNumber}+K${rNumber}+IF(ISNUMBER(L${rNumber}),L${rNumber},0)`,
      result: row.totalUsd
    };

    worksheetRow.getCell(14).value = row.markHawb || ""; // Mark HAWB

    // Format all columns in row
    for (let col = 1; col <= 14; col++) {
      const cell = worksheetRow.getCell(col);
      cell.font = { name: "Arial", size: 9 };
      cell.border = borderThin;
      cell.alignment = {
        vertical: "middle",
        horizontal: headers[col - 1].align as "center" | "right" | "left"
      };

      if (col === 6) {
        // ROUTE orange/peach fill
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFFCE4D6" } // Soft orange-peach standard table accent
        };
      } else if (col === 7) {
        cell.numFmt = "0.0";
      } else if (col === 8) {
        cell.numFmt = "$#,##0.00";
      } else if (col === 9) {
        cell.numFmt = "#,##0.00";
      } else if (col === 10) {
        cell.numFmt = "$#,##0.00";
      } else if (col === 11) {
        cell.numFmt = "#,##0.5"; // Format like 2 or 2.0 depending on value, clean
      } else if (col === 12) {
        cell.numFmt = "#,##0.00";
      } else if (col === 13) {
        cell.numFmt = "#,##0.00";
        cell.font = { name: "Arial", size: 9, bold: true };
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFFFF200" } // Highlighted solid yellow as total
        };
      }
    }
  });

  // 6. Template Padding (rows with empty values but functional formulas, up to row 27 minimum)
  const minLastRow = 27;
  const paddingRowsCount = Math.max(rows.length, 14); // show at least 14 billing rows as in template (8 to 21)
  let paddedEndRow = dataStartRow + paddingRowsCount - 1;
  if (paddedEndRow < minLastRow) paddedEndRow = minLastRow;

  for (let rNumber = dataStartRow + rows.length; rNumber <= paddedEndRow; rNumber++) {
    const worksheetRow = worksheet.getRow(rNumber);
    worksheetRow.height = 20;

    // Formulas so user can fill manually in Excel
    worksheetRow.getCell(9).value = {
      formula: `G${rNumber}*H${rNumber}`,
      result: 0
    };
    worksheetRow.getCell(13).value = {
      formula: `I${rNumber}+J${rNumber}+K${rNumber}+IF(ISNUMBER(L${rNumber}),L${rNumber},0)`,
      result: 0
    };

    for (let col = 1; col <= 14; col++) {
      const cell = worksheetRow.getCell(col);
      cell.font = { name: "Arial", size: 9 };
      cell.border = borderThin;
      cell.alignment = {
        vertical: "middle",
        horizontal: headers[col - 1].align as "center" | "right" | "left"
      };

      if (col === 6) {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFFCE4D6" } // Soft orange-peach
        };
      } else if (col === 7) {
        cell.numFmt = "0.0";
      } else if (col === 8) {
        cell.numFmt = "$#,##0.00";
      } else if (col === 9) {
        cell.numFmt = `_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)`; // Dash for accounting zeros
      } else if (col === 10) {
        cell.numFmt = "$#,##0.00";
      } else if (col === 11) {
        cell.numFmt = `_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)`;
      } else if (col === 12) {
        cell.numFmt = `_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)`;
      } else if (col === 13) {
        cell.numFmt = `_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)`;
        cell.font = { name: "Arial", size: 9, bold: true };
        
        // Highlight first 14 rows total column like in the image (rows 8 to 21 are yellow)
        if (rNumber <= 21) {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFFF200" } // Yellow total highlight
          };
        }
      }
    }
  }

  // 7. SUMMATION ROW ("TOTAL / TỔNG CỘNG") right after padded rows
  const grandTotalRowNumber = paddedEndRow + 1;
  const footerRow = worksheet.getRow(grandTotalRowNumber);
  footerRow.height = 24;

  worksheet.mergeCells(`A${grandTotalRowNumber}:F${grandTotalRowNumber}`);
  const sumLabelCell = footerRow.getCell(1);
  sumLabelCell.value = "TOTAL / TỔNG CỘNG";
  sumLabelCell.font = { name: "Arial", size: 10, bold: true, color: { argb: "FF0F5E3D" } };
  sumLabelCell.alignment = { vertical: "middle", horizontal: "center" };

  for (let c = 1; c <= 6; c++) {
    footerRow.getCell(c).border = {
      top: { style: "medium", color: { argb: "FF0F5E3D" } },
      bottom: { style: "medium", color: { argb: "FF0F5E3D" } },
      left: { style: "thin", color: { argb: "FFCCCCCC" } },
      right: { style: "thin", color: { argb: "FFCCCCCC" } },
    };
    footerRow.getCell(c).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFD1EAE5" }, // Mint green background
    };
  }

  const totalsCols = [
    { col: 7, letter: "G", fmt: "0.0" },          // WEIGHT
    { col: 8, letter: "H", fmt: "-" },            // UNIT PRICE (no total of units)
    { col: 9, letter: "I", fmt: "#,##0.00" },     // FREIGHT Charge
    { col: 10, letter: "J", fmt: "$#,##0.00" },   // HANDLING Charge
    { col: 11, letter: "K", fmt: "#,##0.00" },    // WAREHOUSE Charge
    { col: 12, letter: "L", fmt: "#,##0.00" },    // OTHER Charge
    { col: 13, letter: "M", fmt: "#,##0.00" },    // TOTAL Sum
  ];

  totalsCols.forEach((tf) => {
    const cell = footerRow.getCell(tf.col);
    if (tf.letter === "H") {
      cell.value = "-";
    } else {
      cell.value = {
        formula: `SUM(${tf.letter}8:${tf.letter}${grandTotalRowNumber - 1})`,
        result: rows.reduce((acc, r) => {
          if (tf.letter === "G") return acc + r.weight;
          if (tf.letter === "I") return acc + r.freightCharge;
          if (tf.letter === "J") return acc + r.handlingCharge;
          if (tf.letter === "K") return acc + r.warehouseCharge;
          if (tf.letter === "L") return acc + r.otherCharge;
          if (tf.letter === "M") return acc + r.totalUsd;
          return acc;
        }, 0)
      };
    }

    cell.font = { name: "Arial", size: 9, bold: true };
    cell.alignment = { vertical: "middle", horizontal: tf.letter === "H" ? "center" : "right" };
    cell.border = {
      top: { style: "medium", color: { argb: "FF0F5E3D" } },
      bottom: { style: "medium", color: { argb: "FF0F5E3D" } },
      left: { style: "thin", color: { argb: "FFCCCCCC" } },
      right: { style: "thin", color: { argb: "FFCCCCCC" } },
    };
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFD1EAE5" }, // Soft Mint
    };

    if (tf.fmt !== "-") {
      cell.numFmt = tf.fmt;
    }

    // Yellow highlights total column value cell
    if (tf.col === 13) {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFFF200" },
      };
      cell.font = { name: "Arial", size: 10, bold: true, color: { argb: "FF990000" } };
    }
  });

  // Highlight column N cell on the footer row too
  footerRow.getCell(14).border = {
    top: { style: "medium", color: { argb: "FF0F5E3D" } },
    bottom: { style: "medium", color: { argb: "FF0F5E3D" } },
    left: { style: "thin", color: { argb: "FFCCCCCC" } },
    right: { style: "thin", color: { argb: "FFCCCCCC" } },
  };
  footerRow.getCell(14).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFD1EAE5" },
  };

  // 8. GRAND TOTAL CONVERTED TO VND (Using double borders)
  const vndRowNumber = grandTotalRowNumber + 2;
  worksheet.mergeCells(`A${vndRowNumber}:F${vndRowNumber}`);
  const vndLabelCell = worksheet.getCell(`A${vndRowNumber}`);
  vndLabelCell.value = "TỔNG TIỀN THANH TOÁN QUY ĐỔI (TOTAL VALUE IN VND):";
  vndLabelCell.font = { name: "Arial", size: 10, bold: true, color: { argb: "FF1E3A8A" } };
  vndLabelCell.alignment = { vertical: "middle", horizontal: "right" };

  worksheet.mergeCells(`G${vndRowNumber}:M${vndRowNumber}`);
  const vndValCell = worksheet.getCell(`G${vndRowNumber}`);
  
  const totalUsdCellRef = `M${grandTotalRowNumber}`;
  const sumTotalUsd = rows.reduce((acc, r) => acc + r.totalUsd, 0);
  const calculatedVnd = Math.round(sumTotalUsd * exchangeRate);

  vndValCell.value = {
    formula: `${totalUsdCellRef}*${exchangeRate}`,
    result: calculatedVnd
  };
  
  vndValCell.font = { name: "Arial", size: 11, bold: true, color: { argb: "FFB91C1C" } }; // Deep Red
  vndValCell.alignment = { vertical: "middle", horizontal: "center" };
  vndValCell.numFmt = `#,##0" VND"`;
  
  worksheet.getRow(vndRowNumber).height = 26;
  for (let c = 1; c <= 13; c++) {
    const cCell = worksheet.getCell(vndRowNumber, c);
    cCell.border = {
      top: { style: "medium", color: { argb: "FF1E3A8A" } },
      bottom: { style: "double", color: { argb: "FF1E3A8A" } },
      left: c === 1 || c === 7 ? { style: "medium", color: { argb: "FF1E3A8A" } } : undefined,
      right: c === 6 || c === 13 ? { style: "medium", color: { argb: "FF1E3A8A" } } : undefined,
    };
    cCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFFDE2" }, // Rich cream yellow
    };
  }

  // 9. SIGNATURES BLOCKS (Prepared & Approved)
  const sigRowStart = vndRowNumber + 3;
  worksheet.getRow(sigRowStart).height = 18;
  
  worksheet.mergeCells(`B${sigRowStart}:D${sigRowStart}`);
  const sigMaker = worksheet.getCell(`B${sigRowStart}`);
  sigMaker.value = "Người Lập Biểu / Prepared By";
  sigMaker.font = { name: "Arial", size: 9, bold: true, italic: true };
  sigMaker.alignment = { horizontal: "center" };

  worksheet.mergeCells(`K${sigRowStart}:M${sigRowStart}`);
  const sigApprv = worksheet.getCell(`K${sigRowStart}`);
  sigApprv.value = "Giám Đốc / Approved By";
  sigApprv.font = { name: "Arial", size: 9, bold: true };
  sigApprv.alignment = { horizontal: "center" };

  const sigNameRow = sigRowStart + 4;
  worksheet.mergeCells(`B${sigNameRow}:D${sigNameRow}`);
  const sigMakerName = worksheet.getCell(`B${sigNameRow}`);
  sigMakerName.value = "(Ký, ghi rõ họ tên)";
  sigMakerName.font = { name: "Arial", size: 8, italic: true, color: { argb: "FF7F7F7F" } };
  sigMakerName.alignment = { horizontal: "center" };

  worksheet.mergeCells(`K${sigNameRow}:M${sigNameRow}`);
  const sigApprvName = worksheet.getCell(`K${sigNameRow}`);
  sigApprvName.value = "(Ký tên, đóng dấu)";
  sigApprvName.font = { name: "Arial", size: 8, italic: true, color: { argb: "FF7F7F7F" } };
  sigApprvName.alignment = { horizontal: "center" };

  // Write out and trigger download in browser
  const buffer = await workbook.xlsx.writeBuffer();
  const fileType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  const blob = new Blob([buffer], { type: fileType });
  const url = window.URL.createObjectURL(blob);
  
  const a = document.createElement("a");
  a.href = url;
  a.download = `Bang_Ke_Hoa_Don_Thang_${String(month).padStart(2, "0")}_${year}_${shipperName.replace(/\s+/g, "_")}.xlsx`;
  document.body.appendChild(a);
  a.click();
  
  // Cleanup
  setTimeout(() => {
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);
  }, 100);
}
