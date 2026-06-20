/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useRef, useMemo } from "react";
import {
  UploadCloud,
  FileSpreadsheet,
  AlertTriangle,
  CheckCircle2,
  Settings,
  HelpCircle,
  FileText,
  Trash2,
  ChevronDown,
  ChevronUp,
  RefreshCw,
  Plus,
  Play
} from "lucide-react";
import ExcelJS from "exceljs";
import { ManifestRow } from "../types";

interface ManifestUploaderProps {
  onImport: (rows: ManifestRow[], overwrite: boolean) => void;
  shipperOptions: string[];
}

interface ParsedRow {
  rowNumber: number;
  values: any[];
}

interface HeaderMapping {
  mawbNoIdx: number;
  hawbIdx: number;
  shipperIdx: number;
  consigneeIdx: number;
  wtIdx: number;
  rwtIdx: number;
  vwIdx: number;
  etdIdx?: number;
  etaIdx?: number;
  billIdx?: number;
  markHawbIdx?: number;
  dateIdx?: number;
}

export default function ManifestUploader({ onImport, shipperOptions }: ManifestUploaderProps) {
  const [isDragging, setIsDragging] = useState(false);
  const [loading, setLoading] = useState(false);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);
  const [parsedRows, setParsedRows] = useState<ParsedRow[]>([]);
  const [columnNames, setColumnNames] = useState<string[]>([]);
  const [fileName, setFileName] = useState<string | null>(null);
  const [fileSizeStr, setFileSizeStr] = useState<string>("");
  const [fileType, setFileType] = useState<"excel" | "csv" | "txt" | null>(null);
  const [showConfig, setShowConfig] = useState(false);

  // Column mapping states
  const [mapping, setMapping] = useState<HeaderMapping>({
    mawbNoIdx: -1,
    hawbIdx: -1,
    shipperIdx: -1,
    consigneeIdx: -1,
    wtIdx: -1,
    rwtIdx: -1,
    vwIdx: -1,
    etdIdx: -1,
    etaIdx: -1,
    billIdx: -1,
    markHawbIdx: -1,
    dateIdx: -1,
  });

  // Default replacement values for missing fields during parse
  const [defaultShipper, setDefaultShipper] = useState("DYM VIETNAM CO., LTD");
  const [defaultConsignee, setDefaultConsignee] = useState("CLIENT KOREA");
  const [defaultWT, setDefaultWT] = useState(100.0);
  const [defaultRWT, setDefaultRWT] = useState(100.0);
  const [defaultVW, setDefaultVW] = useState(90.0);

  const fileInputRef = useRef<HTMLInputElement>(null);

  // Helper mapping values dictionary
  const mappingKeys: { key: keyof HeaderMapping; label: string; desc: string; required: boolean }[] = [
    { key: "mawbNoIdx", label: "Mã MAWB NO", desc: "Mã vận đơn chính, dùng phân tuyến HAN", required: true },
    { key: "hawbIdx", label: "Số HAWB", desc: "Mã phụ, dùng tách phí Warehouse ngầm định", required: true },
    { key: "shipperIdx", label: "Tên Shipper", desc: "Khách hàng gửi hàng", required: false },
    { key: "consigneeIdx", label: "Người Nhận (Consignee)", desc: "Khách nhận hàng", required: false },
    { key: "wtIdx", label: "Trọng Thô W/T (kg)", desc: "Trọng lượng vật lý thô thực tế", required: false },
    { key: "rwtIdx", label: "Tính Cước R.W/T (kg)", desc: "Trọng lượng dùng nhân đơn giá cước", required: false },
    { key: "vwIdx", label: "Thể Tích V.W (kg)", desc: "Trọng lượng quy đổi theo thể tích", required: false },
    { key: "etdIdx", label: "ETD (Giờ đi)", desc: "Ngày và giờ khởi hành bay thực tế", required: false },
    { key: "etaIdx", label: "ETA (Giờ đến)", desc: "Ngày và giờ hạ cánh thực tế tại ICN", required: false },
    { key: "billIdx", label: "Invoice BILL No.", desc: "Mã hóa đơn/Biên nhận cước nội bộ", required: false },
    { key: "markHawbIdx", label: "Kí hiệu Mark (HAWB)", desc: "Ghi chú/Kí hiệu nhận dạng phân lô hàng", required: false },
    { key: "dateIdx", label: "Ngày tạo (Date)", desc: "Ngày tạo của Manifest (Thường lấy để gán cho ETD ròng)", required: false },
  ];

  // Helper to read and format file size
  const formatFileSize = (bytes: number) => {
    if (bytes === 0) return "0 Bytes";
    const k = 1024;
    const sizes = ["Bytes", "KB", "MB"];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + " " + sizes[i];
  };

  // Helper: Try parsing cells safely
  const parseCellText = (cellVal: any): string => {
    if (cellVal === null || cellVal === undefined) return "";
    if (typeof cellVal === "object") {
       if ("result" in cellVal) return String(cellVal.result || "");
       if ("text" in cellVal) return String(cellVal.text || "");
       if ("richText" in cellVal && Array.isArray(cellVal.richText)) {
         return cellVal.richText.map((rt: any) => rt.text || "").join("");
       }
       return JSON.stringify(cellVal);
    }
    return String(cellVal).trim();
  };

  const parseCellNumber = (cellVal: any, fallbackValue: number): number => {
    if (cellVal === null || cellVal === undefined) return fallbackValue;
    if (typeof cellVal === "number") return cellVal;
    
    // Clean string from currencies or commas
    const text = parseCellText(cellVal);
    const cleaned = text.replace(/[^0-9.-]/g, "");
    const num = parseFloat(cleaned);
    return isNaN(num) ? fallbackValue : num;
  };

  // 1. Core Logic: Search for optimal headers row
  const autoDetectHeaders = (rows: ParsedRow[]): { detectedMapping: HeaderMapping; headerIndex: number; headers: string[] } => {
    let bestHeaderRowIndex = -1;
    let maxMatchedCount = 0;
    
    const keywords = {
      mawbNo: ["mawb", "master bill", "mã mawb", "vận đơn chính", "mawb no", "mawb_no"],
      hawb: ["hawb", "house bill", "mã hawb", "vận đơn phụ", "hawb no", "hawb_no", "awb"],
      shipper: ["shipper", "gửi", "người gửi", "company", "khách hàng", "shipper name"],
      consignee: ["consignee", "nhận", "người nhận", "cnee", "consignee name"],
      wt: ["w/t", "wt", "weight", "gross", "trọng thô", "trọng lượng thực", "thực tế", "gross weight"],
      rwt: ["r.w/t", "rwt", "chargeable", "tính cước", "trọng lượng cước", "r.w.t", "chargeable weight"],
      vw: ["v.w", "vw", "volume", "thể tích", "trọng lượng thể tích", "v.w.t", "volume weight"],
      etd: ["etd", "departure", "khởi hành", "ngày đi", "etd_date"],
      eta: ["eta", "arrival", "đến nhật", "ngày đến", "eta_date"],
      bill: ["bill", "bill hố", "mã hóa đơn", "số bill", "bill no", "invoice"],
      markHawb: ["mark", "mark (hawb)", "mark hawb", "ghi chú hawb", "kí hiệu", "hawb mark"],
      date: ["date", "ngày", "ngay", "ngày tạo", "ngày bay", "ngay bay", "ngay tao", "manifest date"]
    };

    let optimalMapping: HeaderMapping = {
      mawbNoIdx: -1,
      hawbIdx: -1,
      shipperIdx: -1,
      consigneeIdx: -1,
      wtIdx: -1,
      rwtIdx: -1,
      vwIdx: -1,
      etdIdx: -1,
      etaIdx: -1,
      billIdx: -1,
      markHawbIdx: -1,
      dateIdx: -1
    };

    // Scan up to first 12 rows to find headers
    const scanRowsCount = Math.min(rows.length, 12);
    for (let r = 0; r < scanRowsCount; r++) {
      const row = rows[r];
      let matches = 0;
      let tempMapping = {
        mawbNoIdx: -1,
        hawbIdx: -1,
        shipperIdx: -1,
        consigneeIdx: -1,
        wtIdx: -1,
        rwtIdx: -1,
        vwIdx: -1,
        etdIdx: -1,
        etaIdx: -1,
        billIdx: -1,
        markHawbIdx: -1,
        dateIdx: -1
      };

      row.values.forEach((cell, idx) => {
        const valStr = parseCellText(cell).toLowerCase();
        if (!valStr) return;

        if (keywords.mawbNo.some(k => valStr === k || valStr.includes(k))) {
          tempMapping.mawbNoIdx = idx;
          matches++;
        } else if (keywords.hawb.some(k => valStr === k || valStr.includes(k))) {
          tempMapping.hawbIdx = idx;
          matches++;
        } else if (keywords.shipper.some(k => valStr === k || valStr.includes(k))) {
          tempMapping.shipperIdx = idx;
          matches++;
        } else if (keywords.consignee.some(k => valStr === k || valStr.includes(k))) {
          tempMapping.consigneeIdx = idx;
          matches++;
        } else if (keywords.rwt.some(k => valStr === k || valStr.includes(k))) {
          tempMapping.rwtIdx = idx;
          matches++;
        } else if (keywords.wt.some(k => valStr === k || valStr.includes(k))) {
          tempMapping.wtIdx = idx;
          matches++;
        } else if (keywords.vw.some(k => valStr === k || valStr.includes(k))) {
          tempMapping.vwIdx = idx;
          matches++;
        } else if (keywords.etd.some(k => valStr === k || valStr.includes(k))) {
          tempMapping.etdIdx = idx;
          matches++;
        } else if (keywords.eta.some(k => valStr === k || valStr.includes(k))) {
          tempMapping.etaIdx = idx;
          matches++;
        } else if (keywords.bill.some(k => valStr === k || valStr.includes(k))) {
          tempMapping.billIdx = idx;
          matches++;
        } else if (keywords.markHawb.some(k => valStr === k || valStr.includes(k))) {
          tempMapping.markHawbIdx = idx;
          matches++;
        } else if (keywords.date.some(k => valStr === k || valStr.includes(k))) {
          tempMapping.dateIdx = idx;
          matches++;
        }
      });

      if (matches > maxMatchedCount) {
        maxMatchedCount = matches;
        bestHeaderRowIndex = r;
        optimalMapping = tempMapping;
      }
    }

    // Determine custom column headers array
    let detectedHeaders: string[] = [];
    if (bestHeaderRowIndex !== -1) {
      detectedHeaders = rows[bestHeaderRowIndex].values.map((v, i) => parseCellText(v) || `Cột ${i + 1}`);
    } else {
      // Create letters like A, B, C, D if no header row found
      const maxColCount = rows.reduce((max, row) => Math.max(max, row.values.length), 0);
      detectedHeaders = Array.from({ length: maxColCount }, (_, i) => String.fromCharCode(65 + (i % 26)) + (i >= 26 ? Math.floor(i / 26) : ""));
      
      // Fallback: Bind manually to A, B, C, D, E, F, G
      optimalMapping = {
        mawbNoIdx: maxColCount >= 1 ? 0 : -1,
        hawbIdx: maxColCount >= 2 ? 1 : -1,
        shipperIdx: maxColCount >= 3 ? 2 : -1,
        consigneeIdx: maxColCount >= 4 ? 3 : -1,
        wtIdx: maxColCount >= 5 ? 4 : -1,
        rwtIdx: maxColCount >= 6 ? 5 : -1,
        vwIdx: maxColCount >= 7 ? 6 : -1,
      };
    }

    return {
      detectedMapping: optimalMapping,
      headerIndex: bestHeaderRowIndex,
      headers: detectedHeaders
    };
  };

  // Process selected file
  const handleFileProcess = async (file: File) => {
    setLoading(true);
    setErrorMsg(null);
    setFileName(file.name);
    setFileSizeStr(formatFileSize(file.size));
    
    const fileExt = file.name.split(".").pop()?.toLowerCase();
    
    try {
      if (fileExt === "xlsx" || fileExt === "xls") {
        setFileType("excel");
        const reader = new FileReader();
        reader.onload = async (e) => {
          try {
            const buffer = e.target?.result as ArrayBuffer;
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(buffer);
            const sheet = workbook.worksheets[0];
            
            if (!sheet) {
              throw new Error("Tệp tin Excel không chứa Worksheet nào.");
            }

            const tempParsedList: ParsedRow[] = [];
            sheet.eachRow({ includeEmpty: false }, (row, rowNum) => {
              const rowValues: any[] = [];
              const lastCellCol = Math.max(sheet.columnCount, 15); // scan up to column limit
              
              for (let c = 1; c <= lastCellCol; c++) {
                rowValues.push(row.getCell(c).value);
              }
              
              if (rowValues.some(v => v !== null && v !== undefined && v !== "")) {
                tempParsedList.push({
                  rowNumber: rowNum,
                  values: rowValues
                });
              }
            });

            if (tempParsedList.length === 0) {
              throw new Error("Không tìm thấy dữ liệu dòng nào trong tệp tin Excel.");
            }

            const result = autoDetectHeaders(tempParsedList);
            setParsedRows(tempParsedList);
            setColumnNames(result.headers);
            setMapping(result.detectedMapping);
            setShowConfig(true);
          } catch (err: any) {
            setErrorMsg(`Lỗi đọc tệp Excel: ${err.message || err}`);
          } finally {
            setLoading(false);
          }
        };
        reader.readAsArrayBuffer(file);

      } else if (fileExt === "csv" || fileExt === "txt") {
        setFileType(fileExt === "csv" ? "csv" : "txt");
        const reader = new FileReader();
        reader.onload = (e) => {
          try {
            const text = e.target?.result as string;
            if (!text) {
              throw new Error("Không có nội dung văn bản nào để đọc.");
            }

            const lines = text.split(/\r?\n/);
            const tempParsedList: ParsedRow[] = [];
            let iCounter = 1;

            lines.forEach((line) => {
              if (!line.trim()) return;
              
              // Guess delimiter: Tab, Semicolon, Comma
              let delimiter = ",";
              if (line.includes("\t")) {
                delimiter = "\t";
              } else if (line.includes(";")) {
                delimiter = ";";
              }

              const rowValues: any[] = [];
              let currentVal = "";
              let escape = false;

              for (let index = 0; index < line.length; index++) {
                const char = line[index];
                if (char === '"') {
                  escape = !escape;
                } else if (char === delimiter && !escape) {
                  rowValues.push(currentVal.trim().replace(/^"|"$/g, ''));
                  currentVal = "";
                } else {
                  currentVal += char;
                }
              }
              rowValues.push(currentVal.trim().replace(/^"|"$/g, ''));

              if (rowValues.some(v => v !== "")) {
                tempParsedList.push({
                  rowNumber: iCounter++,
                  values: rowValues
                });
              }
            });

            if (tempParsedList.length === 0) {
              throw new Error("Không có dòng hữu ích nào được parse từ tệp dạng text.");
            }

            const result = autoDetectHeaders(tempParsedList);
            setParsedRows(tempParsedList);
            setColumnNames(result.headers);
            setMapping(result.detectedMapping);
            setShowConfig(true);
          } catch (err: any) {
            setErrorMsg(`Lỗi đọc tệp văn bản: ${err.message || err}`);
          } finally {
            setLoading(false);
          }
        };
        reader.readAsText(file, "UTF-8");
      } else {
        throw new Error("Định dạng tệp không hỗ trợ! Vui lòng chỉ kéo thả hoặc chọn tệp .xlsx, .xls, .csv, hoặc .txt.");
      }
    } catch (err: any) {
      setErrorMsg(err.message || "Đã xảy ra lỗi không xác định.");
      setLoading(false);
    }
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const files = e.dataTransfer.files;
    if (files && files.length > 0) {
      handleFileProcess(files[0]);
    }
  };

  const handleSelectFileClick = () => {
    fileInputRef.current?.click();
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (files && files.length > 0) {
      handleFileProcess(files[0]);
    }
  };

  const handleMappingChange = (key: keyof HeaderMapping, val: number) => {
    setMapping((prev) => ({
      ...prev,
      [key]: val,
    }));
  };

  const handleReset = () => {
    setParsedRows([]);
    setColumnNames([]);
    setFileName(null);
    setFileType(null);
    setShowConfig(false);
    setErrorMsg(null);
    if (fileInputRef.current) fileInputRef.current.value = "";
  };

  // 2. Computed Live mapped rows based on current matching configuration
  const mappedManifestItems = useMemo<ManifestRow[]>(() => {
    if (parsedRows.length === 0) return [];

    // Find first data row (usually row following the headers, or from start if no header mapped)
    const headerRowIndex = parsedRows.findIndex((r, idx) => {
      // Find if this is the header row
      return idx < 12 && r.values.every((v, i) => parseCellText(v) === columnNames[i]);
    });

    const startDataRowIdx = headerRowIndex !== -1 ? headerRowIndex + 1 : 0;
    const finalItems: ManifestRow[] = [];

    for (let rIdx = startDataRowIdx; rIdx < parsedRows.length; rIdx++) {
      const r = parsedRows[rIdx];
      const vals = r.values;

      const mawb = mapping.mawbNoIdx >= 0 && mapping.mawbNoIdx < vals.length
        ? parseCellText(vals[mapping.mawbNoIdx])
        : "";
      const hawb = mapping.hawbIdx >= 0 && mapping.hawbIdx < vals.length
        ? parseCellText(vals[mapping.hawbIdx])
        : "";

      // Skip lines where MAWB or HAWB is extremely blank
      if (!mawb && !hawb) continue;

      const shipper = mapping.shipperIdx >= 0 && mapping.shipperIdx < vals.length && parseCellText(vals[mapping.shipperIdx])
        ? parseCellText(vals[mapping.shipperIdx])
        : defaultShipper;

      const consignee = mapping.consigneeIdx >= 0 && mapping.consigneeIdx < vals.length && parseCellText(vals[mapping.consigneeIdx])
        ? parseCellText(vals[mapping.consigneeIdx])
        : defaultConsignee;

      const wt = mapping.wtIdx >= 0 && mapping.wtIdx < vals.length
        ? parseCellNumber(vals[mapping.wtIdx], defaultWT)
        : defaultWT;

      const rwt = mapping.rwtIdx >= 0 && mapping.rwtIdx < vals.length
        ? parseCellNumber(vals[mapping.rwtIdx], defaultRWT)
        : defaultRWT;

      const vw = mapping.vwIdx >= 0 && mapping.vwIdx < vals.length
        ? parseCellNumber(vals[mapping.vwIdx], defaultVW)
        : defaultVW;

      const etd = mapping.etdIdx !== undefined && mapping.etdIdx >= 0 && mapping.etdIdx < vals.length
        ? parseCellText(vals[mapping.etdIdx])
        : "";

      const eta = mapping.etaIdx !== undefined && mapping.etaIdx >= 0 && mapping.etaIdx < vals.length
        ? parseCellText(vals[mapping.etaIdx])
        : "";

      const bill = mapping.billIdx !== undefined && mapping.billIdx >= 0 && mapping.billIdx < vals.length
        ? parseCellText(vals[mapping.billIdx])
        : "";

      const markHawb = mapping.markHawbIdx !== undefined && mapping.markHawbIdx >= 0 && mapping.markHawbIdx < vals.length
        ? parseCellText(vals[mapping.markHawbIdx])
        : "";

      const parsedDate = mapping.dateIdx !== undefined && mapping.dateIdx >= 0 && mapping.dateIdx < vals.length && parseCellText(vals[mapping.dateIdx])
        ? parseCellText(vals[mapping.dateIdx])
        : new Date().toISOString().split("T")[0];

      finalItems.push({
        id: `M_IMP_${Date.now()}_${rIdx}`,
        mawbNo: mawb || `994-UNASSIGNED-${rIdx}`,
        hawb: hawb || `HW-WH-${rIdx}`,
        shipper: shipper.trim(),
        consignee: consignee.trim(),
        wt: Math.max(0, wt),
        rwt: Math.max(0, rwt),
        vw: Math.max(0, vw),
        date: parsedDate,
        etd,
        eta,
        bill,
        markHawb
      });
    }

    return finalItems;
  }, [
    parsedRows,
    mapping,
    columnNames,
    defaultShipper,
    defaultConsignee,
    defaultWT,
    defaultRWT,
    defaultVW
  ]);

  const handleApplyImport = (overwrite: boolean) => {
    if (mappedManifestItems.length === 0) {
      alert("Không có dòng vận đơn nào hợp lệ để nhập vào hệ thống.");
      return;
    }
    
    onImport(mappedManifestItems, overwrite);
    
    // Nice feedback and reset uploader state
    alert(`Đã nhập thành công ${mappedManifestItems.length} vận đơn tệp ${fileName} vào hệ thống!`);
    handleReset();
  };

  return (
    <div id="manifest-uploader-container" className="bg-[#0b101c] rounded-2xl border border-slate-800 p-5 shadow-xl space-y-4">
      {/* HEADER SECTION */}
      <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3 border-b border-slate-800 pb-3">
        <div className="flex items-center gap-2">
          <div className="p-2 bg-gradient-to-br from-emerald-500/10 to-teal-500/10 text-emerald-400 rounded-lg border border-emerald-500/15">
            <UploadCloud className="w-4 h-4" />
          </div>
          <div>
            <h4 className="font-bold text-xs text-slate-100 uppercase tracking-widest font-mono">
              Tải Manifest Nâng Cao (Smart Drop)
            </h4>
            <p className="text-[10px] text-slate-400">Hỗ trợ Excel (.xlsx, .xls) & File dạng văn bản (.csv, .txt)</p>
          </div>
        </div>
        
        {fileName && (
          <button
            onClick={handleReset}
            className="text-[10px] bg-slate-900 border border-slate-800 rounded px-2.5 py-1 text-rose-400 hover:bg-rose-950/20 hover:text-rose-350 cursor-pointer flex items-center gap-1.5 transition-all font-mono"
          >
            <Trash2 className="w-3 h-3" /> Hủy Tệp Tin
          </button>
        )}
      </div>

      {/* CORE DRAG DROP ZONE */}
      {!showConfig && (
        <div
          onDragOver={handleDragOver}
          onDragLeave={handleDragLeave}
          onDrop={handleDrop}
          onClick={handleSelectFileClick}
          className={`relative border-2 border-dashed rounded-xl p-8 text-center cursor-pointer transition-all ${
            isDragging
              ? "border-emerald-400 bg-emerald-500/5 shadow-inner shadow-emerald-950/20"
              : "border-slate-800 hover:border-emerald-500/40 bg-[#070b13]/60 hover:bg-[#070b13]"
          }`}
        >
          {/* Invisible file input */}
          <input
            type="file"
            ref={fileInputRef}
            onChange={handleFileChange}
            accept=".xlsx, .xls, .csv, .txt"
            className="hidden"
          />

          <div className="space-y-3.5 flex flex-col items-center">
            {loading ? (
              <div className="p-4 bg-emerald-500/10 text-emerald-400 rounded-full animate-spin border-t-2 border-emerald-400">
                <RefreshCw className="w-6 h-6" />
              </div>
            ) : (
              <div className={`p-4 bg-slate-900 text-slate-400 rounded-full border border-slate-800 transition-colors ${isDragging ? "text-emerald-400 bg-slate-950" : ""}`}>
                <FileSpreadsheet className="w-7 h-7" />
              </div>
            )}

            <div>
              <p className="text-xs font-semibold text-slate-200">
                {loading ? "Đang xử lý phân tích cấu trúc dữ liệu..." : "Kéo thả file Manifest của bạn vào đây"}
              </p>
              <p className="text-[10px] text-slate-400 mt-1">hoặc click để tìm kiếm file trên ổ đĩa máy tính</p>
            </div>

            {/* Badges for types of files supported */}
            <div className="flex gap-1.5 pt-1.5 select-none pointer-events-none">
              <span className="text-[9px] font-mono font-bold bg-slate-950 border border-slate-850 px-1.5 py-0.5 rounded text-emerald-450">.XLSX</span>
              <span className="text-[9px] font-mono font-bold bg-slate-950 border border-slate-850 px-1.5 py-0.5 rounded text-emerald-450">.XLS</span>
              <span className="text-[9px] font-mono font-bold bg-slate-950 border border-slate-850 px-1.5 py-0.5 rounded text-cyan-400">.CSV</span>
              <span className="text-[9px] font-mono font-bold bg-slate-950 border border-slate-850 px-1.5 py-0.5 rounded text-amber-400">.TXT</span>
            </div>
          </div>
        </div>
      )}

      {/* ERROR FEEDBACK */}
      {errorMsg && (
        <div className="p-3.5 bg-rose-500/10 border border-rose-500/20 text-rose-300 rounded-xl text-xs flex gap-2.5">
          <AlertTriangle className="w-4 h-4 text-rose-400 shrink-0 mt-0.5" />
          <div className="space-y-1">
            <span className="font-bold font-mono">XẢY RA LỖI PHÂN TÍCH:</span>
            <p className="leading-relaxed">{errorMsg}</p>
          </div>
        </div>
      )}

      {/* INTERACTIVE COLUMN CONFIGURATION MATCHER */}
      {showConfig && parsedRows.length > 0 && (
        <div className="space-y-4 animate-fade-in">
          
          {/* File summary bar */}
          <div className="bg-[#070b13] border border-slate-850 rounded-xl p-3 flex flex-wrap justify-between items-center gap-3">
            <div className="flex items-center gap-2">
              <FileText className="w-4 h-4 text-emerald-400" />
              <div>
                <p className="text-xs font-mono font-bold text-slate-200">{fileName}</p>
                <p className="text-[9px] text-slate-500 font-mono">Dung lượng: {fileSizeStr} | Tổng {parsedRows.length} dòng thô phân tích</p>
              </div>
            </div>
            <div className="flex items-center gap-1.5">
              <span className="text-[10px] font-bold uppercase tracking-wider bg-emerald-500/10 text-emerald-400 border border-emerald-500/20 px-2 py-0.5 rounded">
                {mappedManifestItems.length} Vận Đơn Khả Dụng
              </span>
            </div>
          </div>

          {/* Configuration Grid */}
          <div className="bg-[#0e1628]/60 border border-slate-800/80 rounded-xl p-4 space-y-4">
            
            {/* Title / Info */}
            <div className="flex items-center gap-1.5 border-b border-slate-800 pb-2">
              <Settings className="w-3.5 h-3.5 text-emerald-400" />
              <span className="text-[10px] font-extrabold uppercase tracking-widest text-[#94a3b8] font-mono">Thiết lập Column Mapping (Ánh Xạ Cột)</span>
            </div>

            {/* Mapping Selectors */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-x-5 gap-y-3 text-xs">
              {mappingKeys.map((item) => {
                const currentVal = mapping[item.key];
                return (
                  <div key={item.key} className="flex flex-col space-y-1">
                    <div className="flex justify-between items-center">
                      <label className="font-semibold text-slate-300 flex items-center gap-1">
                        {item.label}
                        {item.required && <span className="text-rose-450 font-bold">*</span>}
                      </label>
                      <span className="text-[9px] text-slate-500 italic max-w-[160px] truncate" title={item.desc}>{item.desc}</span>
                    </div>

                    <select
                      value={currentVal}
                      onChange={(e) => handleMappingChange(item.key, parseInt(e.target.value))}
                      className="w-full bg-[#070b13] border border-slate-850 rounded p-2 outline-hidden text-slate-200 font-mono text-[11px] focus:border-emerald-500 cursor-pointer"
                    >
                      <option value={-1}>-- Bỏ qua / Sử dụng giá trị mặc định --</option>
                      {columnNames.map((colName, index) => {
                        // Gather sample values for this column to make mapping super predictive
                        const samples: string[] = [];
                        for (let i = 1; i <= Math.min(parsedRows.length, 4); i++) {
                          const val = parsedRows[i]?.values[index];
                          if (val !== undefined && val !== null && val !== "") {
                            samples.push(String(val).substring(0, 16));
                          }
                        }
                        const sampleStr = samples.length > 0 ? ` (Mẫu: ${samples.join(", ")})` : "";
                        return (
                          <option key={index} value={index}>
                            Cột {index + 1}: {colName.substring(0, 20)} {sampleStr}
                          </option>
                        );
                      })}
                    </select>
                  </div>
                );
              })}
            </div>

            {/* Advanced fallback config drawer/parameters */}
            <div className="pt-3 border-t border-slate-800 space-y-3">
              <div className="flex items-center gap-1.5">
                <HelpCircle className="w-3.5 h-3.5 text-slate-400" />
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-widest font-mono">Giá Trị Mặc Định Khi Cột Trống/Không Chọn Mapped</span>
              </div>

              <div className="grid grid-cols-2 sm:grid-cols-5 gap-2.5 text-2xs font-mono">
                {/* Fallback Shipper */}
                <div className="space-y-1">
                  <span className="text-slate-500 block uppercase">Shipper</span>
                  <select
                    value={defaultShipper}
                    onChange={(e) => setDefaultShipper(e.target.value)}
                    className="w-full bg-[#070b13] border border-slate-850 rounded p-1 text-slate-300"
                  >
                    {shipperOptions.map(sh => (
                      <option key={sh} value={sh}>{sh}</option>
                    ))}
                  </select>
                </div>

                {/* Fallback Consignee */}
                <div className="space-y-1">
                  <span className="text-slate-500 block uppercase">Consignee</span>
                  <input
                    type="text"
                    value={defaultConsignee}
                    onChange={(e) => setDefaultConsignee(e.target.value)}
                    className="w-full bg-[#070b13] border border-slate-850 rounded p-1 h-6 text-slate-300 text-center"
                  />
                </div>

                {/* Fallback WT */}
                <div className="space-y-1">
                  <span className="text-slate-500 block uppercase">W/T (kg)</span>
                  <input
                    type="number"
                    value={defaultWT}
                    onChange={(e) => setDefaultWT(Math.max(0, parseFloat(e.target.value) || 0))}
                    className="w-full bg-[#070b13] border border-slate-850 rounded p-1 h-6 text-slate-300 text-center"
                  />
                </div>

                {/* Fallback RWT */}
                <div className="space-y-1">
                  <span className="text-slate-500 block uppercase">R.W/T (kg)</span>
                  <input
                    type="number"
                    value={defaultRWT}
                    onChange={(e) => setDefaultRWT(Math.max(0, parseFloat(e.target.value) || 0))}
                    className="w-full bg-[#070b13] border border-slate-850 rounded p-1 h-6 text-[#10b981] font-bold text-center"
                  />
                </div>

                {/* Fallback VW */}
                <div className="space-y-1">
                  <span className="text-slate-500 block">V.W (kg)</span>
                  <input
                    type="number"
                    value={defaultVW}
                    onChange={(e) => setDefaultVW(Math.max(0, parseFloat(e.target.value) || 0))}
                    className="w-full bg-[#070b13] border border-slate-850 rounded p-1 h-6 text-slate-300 text-center"
                  />
                </div>
              </div>
            </div>

          </div>

          {/* TABLE PREVIEW */}
          <div className="space-y-2">
            <span className="text-[10px] font-bold text-slate-400 uppercase tracking-widest font-mono">
              Live Preview: 5 Dòng Vận Đơn Đầu Tiên Để Nhập
            </span>
            <div className="border border-slate-850 rounded-xl overflow-hidden bg-[#070b13]/55">
              <table className="w-full text-left border-collapse text-[10px] font-mono">
                <thead>
                  <tr className="bg-slate-950 text-slate-400 border-b border-slate-850">
                    <th className="py-2 px-3">STT</th>
                    <th className="py-2 px-2 text-emerald-400">MAWB NO</th>
                    <th className="py-2 px-2 text-emerald-400">HAWB</th>
                    <th className="py-2 px-2">SHIPPER</th>
                    <th className="py-2 px-2">CONSIGNEE</th>
                    <th className="py-2 px-2 text-right">W/T</th>
                    <th className="py-2 px-3 text-right text-emerald-300 bg-emerald-500/10 font-bold border-l border-slate-850">R.W/T (*)</th>
                    <th className="py-2 px-2 text-right">V.W</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-850">
                  {mappedManifestItems.length === 0 ? (
                    <tr>
                      <td colSpan={8} className="py-6 text-center text-slate-500 text-2xs italic">
                        Chưa nhận diện thành công dòng nào. Vui lòng ánh xạ cột MAWB NO và HAWB ở phía trên!
                      </td>
                    </tr>
                  ) : (
                    mappedManifestItems.slice(0, 5).map((row, index) => (
                      <tr key={index} className="hover:bg-slate-900/30 text-slate-300 transition-colors">
                        <td className="py-1.5 px-3 text-slate-500">{index + 1}</td>
                        <td className="py-1.5 px-2 text-slate-100 font-bold">{row.mawbNo}</td>
                        <td className="py-1.5 px-2 text-slate-300">{row.hawb}</td>
                        <td className="py-1.5 px-2 truncate max-w-[120px]">{row.shipper}</td>
                        <td className="py-1.5 px-2 truncate max-w-[120px]">{row.consignee}</td>
                        <td className="py-1.5 px-2 text-right">{row.wt.toFixed(1)}</td>
                        <td className="py-1.5 px-3 text-right text-emerald-450 bg-emerald-500/5 font-bold border-l border-slate-850">{row.rwt.toFixed(1)}</td>
                        <td className="py-1.5 px-2 text-right">{row.vw.toFixed(1)}</td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>

              {mappedManifestItems.length > 5 && (
                <div className="p-2 border-t border-slate-850 bg-slate-950/20 text-center text-slate-500 text-2xs font-bold uppercase tracking-widest font-mono">
                  và {mappedManifestItems.length - 5} dòng vận đơn khác...
                </div>
              )}
            </div>
          </div>

          {/* FINAL IMPORT ACTION BUTTON GROUP */}
          <div className="flex flex-col sm:flex-row gap-3 pt-2.5 border-t border-slate-800">
            <button
              onClick={() => handleApplyImport(false)}
              className="flex-1 flex items-center justify-center gap-2 px-4 py-2 bg-slate-800 hover:bg-slate-750 text-xs font-bold text-slate-200 border border-slate-700 hover:border-slate-600 rounded-lg cursor-pointer transition-all active:scale-[0.99]"
            >
              <Plus className="w-4 h-4 text-emerald-400" />
              <span>Cộng Dồn / Bổ Sung (Append to Current)</span>
            </button>

            <button
              onClick={() => handleApplyImport(true)}
              className="flex-1 flex items-center justify-center gap-2 px-4 py-2 bg-gradient-to-r from-emerald-500 to-teal-600 hover:from-emerald-600 hover:to-teal-700 text-xs font-bold text-slate-950 rounded-lg shadow-md shadow-emerald-500/5 cursor-pointer transition-all active:scale-[0.99] hover:scale-[1.01]"
            >
              <Play className="w-4 h-4 text-slate-950" />
              <span>Ghi Đè Hoàn Toàn (Overwrite Data)</span>
            </button>
          </div>

        </div>
      )}
    </div>
  );
}
