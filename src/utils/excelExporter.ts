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
  const worksheet = workbook.addWorksheet("Bao Cao Tinh Tien");

  // Show Grid Lines explicitly
  worksheet.views = [{ showGridLines: true }];

  // 1. Company Header (Prestige Logistics look)
  worksheet.mergeCells("A1:K1");
  const companyTitle = worksheet.getCell("A1");
  companyTitle.value = "GLOBAL EXPRESS LOGISTICS CO., LTD.";
  companyTitle.font = { name: "Arial", size: 10, bold: true, color: { argb: "FF4F4F4F" } };
  companyTitle.alignment = { vertical: "middle", horizontal: "left" };

  worksheet.mergeCells("A2:K2");
  const companyAddr = worksheet.getCell("A2");
  companyAddr.value = "SGN Cargo Terminal, Tan Son Nhat Airport, Ho Chi Minh City, Vietnam";
  companyAddr.font = { name: "Arial", size: 8, italic: true, color: { argb: "FF7F7F7F" } };
  companyAddr.alignment = { vertical: "middle", horizontal: "left" };

  // Spacing
  worksheet.addRow([]);

  // 2. Report Main Title
  worksheet.mergeCells("A4:K4");
  const titleCell = worksheet.getCell("A4");
  titleCell.value = `BẢNG KÊ HÓA ĐƠN THÁNG ${String(month).padStart(2, "0")} NĂM ${year} - ${shipperName.toUpperCase()}`;
  titleCell.font = { name: "Arial", size: 14, bold: true, color: { argb: "FF0A3622" } }; // Deep Forest Green
  titleCell.alignment = { vertical: "middle", horizontal: "center" };
  worksheet.getRow(4).height = 28;

  // 3. Exchange Rate Meta Info
  worksheet.mergeCells("A5:K5");
  const rateCell = worksheet.getCell("A5");
  rateCell.value = `Tỷ giá áp dụng thanh toán (Exchange Rate): ${exchangeRate.toLocaleString("vi-VN")} VND/USD`;
  rateCell.font = { name: "Arial", size: 10, italic: true, bold: true, color: { argb: "FF1E3A8A" } }; // Navy blue text
  rateCell.alignment = { vertical: "middle", horizontal: "center" };
  worksheet.getRow(5).height = 18;

  worksheet.addRow([]); // Blank line before table header

  // 4. Headers definitions
  const headers = [
    { name: "NO.", key: "no", width: 6, align: "center" },
    { name: "MAWB NO", key: "mawb", width: 25, align: "center" },
    { name: "HAWB", key: "hawb", width: 15, align: "center" },
    { name: "ROUTE", key: "route", width: 16, align: "center" },
    { name: "WEIGHT (KG)", key: "weight", width: 14, align: "right" },
    { name: "UNIT PRICE (USD)", key: "unit", width: 16, align: "right" },
    { name: "FREIGHT CHARGE (USD)", key: "freight", width: 24, align: "right" },
    { name: "HANDLING CHARGE (USD)", key: "handling", width: 22, align: "right" },
    { name: "WAREHOUSE CHARGE (USD)", key: "warehouse", width: 24, align: "right" },
    { name: "OTH.CHG (USD)", key: "other", width: 15, align: "right" },
    { name: "TOTAL (USD)", key: "total", width: 18, align: "right" },
  ];

  // Set visual column headers
  const headerRowNumber = 7;
  const headerRow = worksheet.getRow(headerRowNumber);
  headerRow.height = 26;

  headers.forEach((h, index) => {
    const colIndex = index + 1;
    const cell = headerRow.getCell(colIndex);
    cell.value = h.name;
    cell.font = { name: "Arial", size: 9, bold: true, color: { argb: "FF0F5E3D" } }; // Dark green
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFD1EAE5" }, // Mint green background
    };
    cell.border = {
      top: { style: "medium", color: { argb: "FF4D7C0F" } },
      bottom: { style: "medium", color: { argb: "FF4D7C0F" } },
      left: { style: "thin", color: { argb: "FFB3D8D1" } },
      right: { style: "thin", color: { argb: "FFB3D8D1" } },
    };
  });

  // Adjust columns configuration wide sizing
  headers.forEach((h, index) => {
    worksheet.getColumn(index + 1).width = h.width;
  });

  // 5. Populate Data Rows
  let dataStartRow = 8;
  const borderThin = {
    top: { style: "thin" as ExcelJS.BorderStyle, color: { argb: "FFCCCCCC" } },
    bottom: { style: "thin" as ExcelJS.BorderStyle, color: { argb: "FFCCCCCC" } },
    left: { style: "thin" as ExcelJS.BorderStyle, color: { argb: "FFCCCCCC" } },
    right: { style: "thin" as ExcelJS.BorderStyle, color: { argb: "FFCCCCCC" } },
  };

  rows.forEach((row, i) => {
    const rNumber = dataStartRow + i;
    const worksheetRow = worksheet.getRow(rNumber);
    worksheetRow.height = 20;

    // Col cells data
    const cellsVal = [
      i + 1,                                       // NO.
      row.mawbNo,                                  // MAWB NO
      row.hawb,                                    // HAWB
      row.route,                                   // ROUTE
      row.weight,                                  // WEIGHT (KG)
      row.unitPrice,                               // UNIT PRICE
      row.freightCharge,                           // FREIGHT CHARGE
      row.handlingCharge,                          // HANDLING
      row.warehouseCharge,                         // WAREHOUSE
      row.otherCharge,                             // OTHER
      row.totalUsd                                 // TOTAL (USD)
    ];

    cellsVal.forEach((val, cIndex) => {
      const cell = worksheetRow.getCell(cIndex + 1);
      cell.value = val;
      cell.font = { name: "Arial", size: 9 };
      cell.alignment = {
        vertical: "middle",
        horizontal: headers[cIndex].align as "center" | "right" | "left",
      };
      cell.border = borderThin;

      // Formatting numbers
      if (cIndex === 4) {
        // WEIGHT (KG)
        cell.numFmt = "#,##0.0";
      } else if (cIndex === 5) {
        // UNIT PRICE (USD)
        cell.numFmt = "$#,##0.00";
      } else if (cIndex >= 6 && cIndex <= 9) {
        // CHARGES (USD)
        cell.numFmt = "$#,##0.00";
      } else if (cIndex === 10) {
        // TOTAL (USD)
        cell.numFmt = "$#,##0.00";
        cell.font = { name: "Arial", size: 9, bold: true, color: { argb: "FF854D0E" } }; // golden bold text
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFFDF9C3" }, // Highlighting Yellow Background for TOTAL Column
        };
      }
    });
  });

  const grandTotalRowNumber = dataStartRow + rows.length;
  const footerRow = worksheet.getRow(grandTotalRowNumber);
  footerRow.height = 24;

  // 6. SUMMATION ROW ("TOTAL")
  // Merge A to D for "TOTAL" label
  worksheet.mergeCells(`A${grandTotalRowNumber}:D${grandTotalRowNumber}`);
  const sumLabelCell = footerRow.getCell(1);
  sumLabelCell.value = "TOTAL / TỔNG CỘNG";
  sumLabelCell.font = { name: "Arial", size: 9, bold: true, color: { argb: "FF0F5E3D" } };
  sumLabelCell.alignment = { vertical: "middle", horizontal: "center" };

  // Set borders for merged cells
  for (let c = 1; c <= 4; c++) {
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

  // Formulas for totals
  const totalFormulaCols = [
    { col: 5, letter: "E", fmt: "#,##0.0" },      // WEIGHT
    { col: 6, letter: "F", fmt: "-" },            // UNIT PRICE (no total of units)
    { col: 7, letter: "G", fmt: "$#,##0.00" },    // FREIGHT
    { col: 8, letter: "H", fmt: "$#,##0.00" },    // HANDLING
    { col: 9, letter: "I", fmt: "$#,##0.00" },    // WAREHOUSE
    { col: 10, letter: "J", fmt: "$#,##0.00" },   // OTHER
    { col: 11, letter: "K", fmt: "$#,##0.00" },   // TOTAL USD
  ];

  totalFormulaCols.forEach((tf) => {
    const cell = footerRow.getCell(tf.col);
    if (tf.letter === "F") {
      cell.value = "-";
    } else {
      cell.value = {
        formula: `SUM(${tf.letter}8:${tf.letter}${grandTotalRowNumber - 1})`,
        result: rows.reduce((acc, r) => {
          if (tf.letter === "E") return acc + r.weight;
          if (tf.letter === "G") return acc + r.freightCharge;
          if (tf.letter === "H") return acc + r.handlingCharge;
          if (tf.letter === "I") return acc + r.warehouseCharge;
          if (tf.letter === "J") return acc + r.otherCharge;
          if (tf.letter === "K") return acc + r.totalUsd;
          return acc;
        }, 0)
      };
    }

    cell.font = { name: "Arial", size: 9, bold: true };
    cell.alignment = { vertical: "middle", horizontal: "right" };
    cell.border = {
      top: { style: "medium", color: { argb: "FF0F5E3D" } },
      bottom: { style: "medium", color: { argb: "FF0F5E3D" } },
      left: { style: "thin", color: { argb: "FFCCCCCC" } },
      right: { style: "thin", color: { argb: "FFCCCCCC" } },
    };
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFD1EAE5" }, // Mint green background
    };

    if (tf.fmt !== "-") {
      cell.numFmt = tf.fmt;
    }

    // Double highlighted border or yellow background for Grand Total USD
    if (tf.col === 11) {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFDF9C3" }, // Yellow accent for final sum
      };
      cell.font = { name: "Arial", size: 10, bold: true, color: { argb: "FF990000" } };
    }
  });


  // 7. VƯỢT TRỘI: TỔNG TIỀN QUY ĐỔI VND (GRAND TOTAL CONVERTED TO VND)
  const vndRowNumber = grandTotalRowNumber + 2;
  worksheet.mergeCells(`A${vndRowNumber}:F${vndRowNumber}`);
  const vndLabelCell = worksheet.getCell(`A${vndRowNumber}`);
  vndLabelCell.value = "TỔNG TIỀN THANH TOÁN QUY ĐỔI (TOTAL VALUE IN VND):";
  vndLabelCell.font = { name: "Arial", size: 10, bold: true, color: { argb: "FF1E3A8A" } };
  vndLabelCell.alignment = { vertical: "middle", horizontal: "right" };

  worksheet.mergeCells(`G${vndRowNumber}:K${vndRowNumber}`);
  const vndValCell = worksheet.getCell(`G${vndRowNumber}`);
  
  // Formula: Total USD * Exchange Rate
  const totalUsdCellRef = `K${grandTotalRowNumber}`;
  const sumTotalUsd = rows.reduce((acc, r) => acc + r.totalUsd, 0);
  const calculatedVnd = Math.round(sumTotalUsd * exchangeRate);

  vndValCell.value = {
    formula: `${totalUsdCellRef}*${exchangeRate}`,
    result: calculatedVnd
  };
  
  vndValCell.font = { name: "Arial", size: 11, bold: true, color: { argb: "FFB91C1C" } }; // Deep Red text
  vndValCell.alignment = { vertical: "middle", horizontal: "center" };
  vndValCell.numFmt = `#,##0" VND"`;
  
  // Style boundaries for VND Converted Row
  worksheet.getRow(vndRowNumber).height = 26;
  for (let c = 1; c <= 11; c++) {
    const cCell = worksheet.getCell(vndRowNumber, c);
    cCell.border = {
      top: { style: "medium", color: { argb: "FF1E3A8A" } },
      bottom: { style: "double", color: { argb: "FF1E3A8A" } },
      left: c === 1 || c === 7 ? { style: "medium", color: { argb: "FF1E3A8A" } } : undefined,
      right: c === 6 || c === 11 ? { style: "medium", color: { argb: "FF1E3A8A" } } : undefined,
    };
    cCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FEF9C3" }, // Soft golden yellow highlight
    };
  }


  // 8. SIGNATURE BLOCK (GIÁM ĐỐC / KẾ TOÁN / NGƯỜI LẬP BIỂU)
  const sigRowStart = vndRowNumber + 3;
  worksheet.getRow(sigRowStart).height = 18;
  
  worksheet.mergeCells(`A${sigRowStart}:C${sigRowStart}`);
  const sigMaker = worksheet.getCell(`A${sigRowStart}`);
  sigMaker.value = "Người Lập Biểu / Prepared By";
  sigMaker.font = { name: "Arial", size: 9, bold: true, italic: true };
  sigMaker.alignment = { horizontal: "center" };

  worksheet.mergeCells(`H${sigRowStart}:K${sigRowStart}`);
  const sigApprv = worksheet.getCell(`H${sigRowStart}`);
  sigApprv.value = "Giám Đốc / Approved By";
  sigApprv.font = { name: "Arial", size: 9, bold: true };
  sigApprv.alignment = { horizontal: "center" };

  const sigNameRow = sigRowStart + 4;
  worksheet.mergeCells(`A${sigNameRow}:C${sigNameRow}`);
  const sigMakerName = worksheet.getCell(`A${sigNameRow}`);
  sigMakerName.value = "(Ký, ghi rõ họ tên)";
  sigMakerName.font = { name: "Arial", size: 8, italic: true, color: { argb: "FF7F7F7F" } };
  sigMakerName.alignment = { horizontal: "center" };

  worksheet.mergeCells(`H${sigNameRow}:K${sigNameRow}`);
  const sigApprvName = worksheet.getCell(`H${sigNameRow}`);
  sigApprvName.value = "(Ký tên, đóng dấu)";
  sigApprvName.font = { name: "Arial", size: 8, italic: true, color: { argb: "FF7F7F7F" } };
  sigApprvName.alignment = { horizontal: "center" };


  // 9. Write and trigger save
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
