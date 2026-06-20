/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useMemo } from "react";
import {
  FileSpreadsheet,
  Database,
  SlidersHorizontal,
  ChevronRight,
  Plus,
  RefreshCw,
  TrendingUp,
  Download,
  CheckCircle2,
  Table,
  Cpu,
  Trash2,
  Calendar,
  DollarSign
} from "lucide-react";
import { ManifestRow, BillingCalculation, ShipperSummary } from "./types";
import { INITIAL_MANIFEST_DATA, DEMO_UNIT_PRICES } from "./data/demoData";
import { exportShipperExcelReport } from "./utils/excelExporter";
import DatabaseDesign from "./components/DatabaseDesign";
import PythonScriptView from "./components/PythonScriptView";
import ManifestUploader from "./components/ManifestUploader";

export default function App() {
  // --- Active Tab State ---
  const [activeTab, setActiveTab] = useState<"calculator" | "developer">("calculator");
  const [developerSubTab, setDeveloperSubTab] = useState<"database" | "python">("python");

  // --- Core Application States ---
  const [manifests, setManifests] = useState<ManifestRow[]>(INITIAL_MANIFEST_DATA);
  const [unitPrices, setUnitPrices] = useState<Record<string, number>>(DEMO_UNIT_PRICES);
  const [exchangeRate, setExchangeRate] = useState<number>(26160);
  const [reportMonth, setReportMonth] = useState<number>(6);
  const [reportYear, setReportYear] = useState<number>(2026);

  // Derive unique list of shippers from current active manifests
  const uniqueShippers = useMemo(() => {
    const set = new Set(manifests.map((m) => m.shipper));
    return Array.from(set);
  }, [manifests]);

  // Selected Shipper for interactive detailed billing report (Hình 2 visualization)
  const [selectedShipper, setSelectedShipper] = useState<string>("DYM VIETNAM CO., LTD");

  // Keep track of custom Warehouse charges edited by user.
  // Defaults will be parsed dynamically from the HAWB key numbers if not present.
  const [warehouseCharges, setWarehouseCharges] = useState<Record<string, number>>({});
  const [otherCharges, setOtherCharges] = useState<Record<string, number>>({});

  // Reset to original demo dataset state
  const handleResetData = () => {
    setManifests(INITIAL_MANIFEST_DATA);
    setUnitPrices(DEMO_UNIT_PRICES);
    setExchangeRate(26160);
    setWarehouseCharges({});
    setOtherCharges({});
    if (!INITIAL_MANIFEST_DATA.some(m => m.shipper === selectedShipper)) {
      setSelectedShipper(INITIAL_MANIFEST_DATA[0].shipper);
    }
  };

  // --- Dynamic calculations mapping and logic ---
  // Every time manifest data, unit price dict, or other charges changes,
  // we rebuild the computed billing report in real-time. (No useEffect loops, pure reactivity!)
  const computedBillingRows: BillingCalculation[] = useMemo(() => {
    return manifests.map((m) => {
      // a. Core Route Rule: Chứa "HAN" -> SGN-HAN-ICN, Ngược lại -> SGN-ICN
      const isHan = m.mawbNo.toUpperCase().includes("HAN");
      const route = isHan ? "SGN-HAN-ICN" : "SGN-ICN";

      // b. Weight Rule: Lấy chính xác giá trị R.W/T từ Hình 1
      const weight = m.rwt;

      // c. Handling Charge Rule: Cố định 10.00 USD cho tất cả các dòng
      const handlingCharge = 10.00;

      // d. Unit price: dynamic user-authored pricing
      const unitPrice = unitPrices[m.id] !== undefined ? unitPrices[m.id] : 3.00;

      // e. Warehouse charge reference parser from HAWB
      let warehouseChargeInUSD = warehouseCharges[m.id];
      if (warehouseChargeInUSD === undefined) {
        // Parse trailing numeric value in HAWB string if any (e.g. "HW-WH-45" -> 45)
        const parsedNum = m.hawb.match(/\d+/);
        warehouseChargeInUSD = parsedNum ? parseFloat(parsedNum[0]) : 15.00;
      }

      // Other custom charges
      const otherCharge = otherCharges[m.id] || 0.00;

      // g. Math computations
      const freightCharge = weight * unitPrice;
      const totalUsd = freightCharge + handlingCharge + warehouseChargeInUSD + otherCharge;

      return {
        id: m.id,
        manifestId: m.id,
        mawbNo: m.mawbNo,
        hawb: m.hawb,
        shipper: m.shipper,
        route,
        weight,
        unitPrice,
        freightCharge,
        handlingCharge,
        warehouseCharge: warehouseChargeInUSD,
        otherCharge,
        totalUsd
      };
    });
  }, [manifests, unitPrices, warehouseCharges, otherCharges]);

  // Filter computed rows for the selected shipper (corresponds to active Hình 2 sheet)
  const shipperBillingRows = useMemo(() => {
    return computedBillingRows.filter(
      (row) => row.shipper.toUpperCase() === selectedShipper.toUpperCase()
    );
  }, [computedBillingRows, selectedShipper]);

  // Aggregate stats totals for the selected shipper's report
  const selectedShipperTotals = useMemo(() => {
    let totalWeight = 0;
    let totalFreightUSD = 0;
    let totalHandlingUSD = 0;
    let totalWarehouseUSD = 0;
    let totalOtherUSD = 0;
    let totalUSD = 0;

    shipperBillingRows.forEach((row) => {
      totalWeight += row.weight;
      totalFreightUSD += row.freightCharge;
      totalHandlingUSD += row.handlingCharge;
      totalWarehouseUSD += row.warehouseCharge;
      totalOtherUSD += row.otherCharge;
      totalUSD += row.totalUsd;
    });

    const totalVND = Math.round(totalUSD * exchangeRate);

    return {
      totalWeight,
      totalFreightUSD,
      totalHandlingUSD,
      totalWarehouseUSD,
      totalOtherUSD,
      totalUSD,
      totalVND
    };
  }, [shipperBillingRows, exchangeRate]);

  // State elements for adding new temporary Manifest lines
  const [newMawb, setNewMawb] = useState("");
  const [newHawb, setNewHawb] = useState("");
  const [newShipper, setNewShipper] = useState("DYM VIETNAM CO., LTD");
  const [newConsignee, setNewConsignee] = useState("");
  const [newWT, setNewWT] = useState("120");
  const [newRWT, setNewRWT] = useState("115");
  const [newVW, setNewVW] = useState("110");
  const [showAddForm, setShowAddForm] = useState(false);

  // Submit hander to insert manifest item
  const handleAddManifestRow = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newMawb || !newHawb) {
      alert("Vui lòng nhập số MAWB và HAWB để phân loại!");
      return;
    }

    const newId = `M_GEN_${Date.now()}`;
    const newRow: ManifestRow = {
      id: newId,
      mawbNo: newMawb,
      hawb: newHawb,
      shipper: newShipper.trim(),
      consignee: newConsignee || "CLIENT KOREA",
      wt: parseFloat(newWT) || 100.0,
      rwt: parseFloat(newRWT) || 100.0,
      vw: parseFloat(newVW) || 90.0,
      date: new Date().toISOString().split("T")[0]
    };

    setManifests((prev) => [...prev, newRow]);
    
    // Set a default unit price for newly added item
    setUnitPrices((prev) => ({
      ...prev,
      [newId]: 3.20
    }));

    // Reset local inputs
    setNewMawb("");
    setNewHawb("");
    setNewConsignee("");
    setShowAddForm(false);
  };

  // Remove a manifest row
  const handleDeleteRow = (id: string) => {
    setManifests((prev) => prev.filter((m) => m.id !== id));
  };

  // Import handler for Drag & Drop files (Excel, CSV, TXT)
  const handleDataImported = (newRows: ManifestRow[], overwrite: boolean) => {
    if (overwrite) {
      setManifests(newRows);
      
      const newUnitPrices: Record<string, number> = {};
      newRows.forEach((row) => {
        newUnitPrices[row.id] = 3.20;
      });
      setUnitPrices(newUnitPrices);

      // Reset dynamic warehouse charges & other charges to restart fresh
      setWarehouseCharges({});
      setOtherCharges({});

      if (newRows.length > 0) {
        // Auto-select first shipper of the new dataset
        const firstShipper = newRows[0].shipper || "DYM VIETNAM CO., LTD";
        setSelectedShipper(firstShipper);
      }
    } else {
      setManifests((prev) => [...prev, ...newRows]);
      
      setUnitPrices((prev) => {
        const merged = { ...prev };
        newRows.forEach((row) => {
          if (merged[row.id] === undefined) {
            merged[row.id] = 3.20;
          }
        });
        return merged;
      });

      if (newRows.length > 0) {
        // Auto-select the first shipper of the newly appended dataset
        const firstNewShipper = newRows[0].shipper || "DYM VIETNAM CO., LTD";
        setSelectedShipper(firstNewShipper);
      }
    }
  };

  // Update dynamic values (Unit Price, Warehouse charge, and other charge) on-the-fly
  const handleUpdateUnitPrice = (id: string, val: string) => {
    const rateNum = parseFloat(val) || 0;
    setUnitPrices((prev) => ({
      ...prev,
      [id]: rateNum
    }));
  };

  const handleUpdateWarehouseCharge = (id: string, val: string) => {
    const cost = parseFloat(val) || 0;
    setWarehouseCharges((prev) => ({
      ...prev,
      [id]: cost
    }));
  };

  const handleUpdateOtherCharge = (id: string, val: string) => {
    const chg = parseFloat(val) || 0;
    setOtherCharges((prev) => ({
      ...prev,
      [id]: chg
    }));
  };

  // Trigger spreadsheet file download
  const handleExportXLSX = async () => {
    if (shipperBillingRows.length === 0) {
      alert("Không có dòng dữ liệu nào để xuất báo cáo!");
      return;
    }
    try {
      await exportShipperExcelReport(
        selectedShipper,
        shipperBillingRows,
        exchangeRate,
        reportMonth,
        reportYear
      );
    } catch (err: any) {
      console.error(err);
      alert("Đã xảy ra lỗi khi xuất Excel: " + err.message);
    }
  };

  return (
    <div className="min-h-screen bg-[#070b13] text-[#cbd5e1] font-sans selection:bg-emerald-500/20 selection:text-emerald-300">
      {/* GLOBAL LOGISTICS HEADER */}
      <header className="bg-[#0b0f19]/90 border-b border-slate-850 sticky top-0 z-40 backdrop-blur-md shadow-lg shadow-emerald-950/5">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="flex flex-col md:flex-row justify-between items-center py-4 md:h-16 gap-4">
            {/* Branding Logo */}
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 bg-gradient-to-br from-emerald-500 to-teal-600 rounded-xl flex items-center justify-center text-[#070b13] shadow-md shadow-emerald-500/20">
                <FileSpreadsheet className="w-5 h-5 text-slate-900 font-extrabold" />
              </div>
              <div>
                <h1 className="font-display font-bold text-sm sm:text-base text-transparent bg-clip-text bg-gradient-to-r from-emerald-400 via-teal-350 to-cyan-400 tracking-wider leading-none mb-1">
                  PRESTIGE LOGISTICS
                </h1>
                <p className="text-[10px] uppercase tracking-widest font-semibold text-emerald-400/90 font-mono">
                  Phân tích Manifest & Tính tiền tự động
                </p>
              </div>
            </div>

            {/* Navigation Tabs */}
            <div className="flex bg-slate-950 p-1 rounded-xl border border-slate-850">
              <button
                onClick={() => setActiveTab("calculator")}
                className={`flex items-center gap-2 px-4 py-2 text-xs font-semibold rounded-lg transition-all cursor-pointer ${
                  activeTab === "calculator"
                    ? "bg-[#10b981]/15 text-[#10b981] border border-[#10b981]/20 shadow-sm"
                    : "text-slate-400 hover:text-slate-200 hover:bg-slate-900/50"
                }`}
              >
                <Table className="w-3.5 h-3.5" />
                <span>Bàn Tính Interactive (Hình 1 & 2)</span>
              </button>
              <button
                onClick={() => setActiveTab("developer")}
                className={`flex items-center gap-2 px-4 py-2 text-xs font-semibold rounded-lg transition-all cursor-pointer ${
                  activeTab === "developer"
                    ? "bg-slate-800 text-white shadow-sm border border-slate-700"
                    : "text-slate-400 hover:text-slate-200 hover:bg-slate-900/50"
                }`}
              >
                <Cpu className="w-3.5 h-3.5" />
                <span>Giải Pháp Python & Database</span>
              </button>
            </div>
          </div>
        </div>
      </header>

      {/* CORE WORKSPACE */}
      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
        
        {/* TAB 1: WORKSPACE TÍNH CỬA INTERACTIVE */}
        {activeTab === "calculator" && (
          <div className="space-y-8 animate-fade-in">
            
            {/* SECTION 1: DONG ĐIỀU KHIỂN & CONFIG TOÀN CỤC */}
            <div className="bg-[#0b101d] rounded-2xl border border-slate-850 p-5 shadow-xl relative overflow-hidden">
              <div className="absolute top-0 right-0 w-64 h-64 bg-emerald-500/5 blur-3xl pointer-events-none rounded-full"></div>
              
              <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 mb-4 pb-4 border-b border-slate-800">
                <div className="flex items-center gap-2.5">
                  <div className="p-2 bg-emerald-500/10 text-emerald-400 rounded-lg border border-emerald-500/15">
                    <SlidersHorizontal className="w-4 h-4" />
                  </div>
                  <div>
                    <h3 className="font-semibold text-sm text-slate-100">Bảng Điều Khiển Toàn Cục</h3>
                    <p className="text-[11px] text-slate-400">Thiết lập tỷ giá hối đoái, mốc thời gian và chọn Shipper cần trích xuất biên lai</p>
                  </div>
                </div>
                <button
                  onClick={handleResetData}
                  className="flex items-center gap-1.5 px-3 py-1.5 border border-slate-800 text-[11px] font-semibold text-slate-400 hover:text-slate-200 hover:bg-slate-900 rounded-lg cursor-pointer transition-colors"
                >
                  <RefreshCw className="w-3 h-3 text-emerald-500" />
                  Nhập Lại Dữ Liệu Gốc Mẫu
                </button>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
                {/* 1. Chọn Shipper */}
                <div className="space-y-1">
                  <label className="text-[10px] font-bold uppercase tracking-wider text-slate-500">Chọn Shipper (Khách Hàng)</label>
                  <select
                    value={selectedShipper}
                    onChange={(e) => setSelectedShipper(e.target.value)}
                    className="w-full text-xs font-semibold rounded-lg bg-[#070b13] border border-slate-800 p-2.5 outline-hidden focus:border-emerald-500 text-slate-200 transition-all cursor-pointer"
                  >
                    {uniqueShippers.map((sh) => (
                      <option key={sh} value={sh} className="bg-[#0b101d] text-slate-200">
                        {sh}
                      </option>
                    ))}
                  </select>
                </div>

                {/* 2. Căn Tỷ giá */}
                <div className="space-y-1">
                  <div className="flex justify-between items-center">
                    <label className="text-[10px] font-bold uppercase tracking-wider text-slate-500">Tỷ Giá Quy Đổi (USD ➔ VND)</label>
                    <span className="text-[9px] bg-emerald-500/10 text-emerald-450 px-1 py-0.5 rounded font-bold font-mono">Exchange Rate</span>
                  </div>
                  <div className="relative">
                    <input
                      type="number"
                      value={exchangeRate}
                      onChange={(e) => setExchangeRate(parseInt(e.target.value) || 0)}
                      placeholder="26160"
                      className="w-full text-xs font-mono font-bold rounded-lg bg-[#070b13] border border-slate-800 pl-8 pr-12 py-2.5 outline-hidden focus:border-emerald-500 text-[#38bdf8] transition-all"
                    />
                    <DollarSign className="w-3.5 h-3.5 text-[#38bdf8] absolute left-2.5 top-3" />
                    <span className="absolute right-2.5 top-2.5 text-[10px] font-bold text-slate-500">VND/USD</span>
                  </div>
                </div>

                {/* 3. Chọn Tháng Báo Cáo */}
                <div className="space-y-1">
                  <label className="text-[10px] font-bold uppercase tracking-wider text-slate-500">Tháng Báo Cáo</label>
                  <div className="flex gap-2">
                    <select
                      value={reportMonth}
                      onChange={(e) => setReportMonth(parseInt(e.target.value))}
                      className="w-1/2 text-xs font-semibold rounded-lg bg-[#070b13] border border-slate-800 p-2.5 outline-hidden focus:border-emerald-500 text-slate-200"
                    >
                      {Array.from({ length: 12 }, (_, i) => i + 1).map((m) => (
                        <option key={m} value={m} className="bg-[#0b101d] text-slate-200">
                          Tháng {String(m).padStart(2, "0")}
                        </option>
                      ))}
                    </select>

                    <input
                      type="number"
                      value={reportYear}
                      onChange={(e) => setReportYear(parseInt(e.target.value) || 2026)}
                      className="w-1/2 text-xs font-mono font-bold rounded-lg bg-[#070b13] border border-slate-800 p-2.5 outline-hidden focus:border-emerald-500 text-slate-200 text-center"
                    />
                  </div>
                </div>

                {/* 4. Action Export EXCEL */}
                <div className="flex items-end">
                  <button
                    onClick={handleExportXLSX}
                    className="w-full flex items-center justify-center gap-2 px-4 py-2.5 bg-gradient-to-r from-emerald-500 to-teal-600 hover:from-emerald-600 hover:to-teal-700 text-xs font-bold text-slate-950 rounded-lg shadow-md shadow-emerald-950/20 cursor-pointer transition-all hover:scale-[1.01]"
                  >
                    <Download className="w-4 h-4 text-slate-950" />
                    <span>Xuất Excel Đẹp Hệt Hình 2</span>
                  </button>
                </div>
              </div>
            </div>

            {/* TWO GRAPHICAL PANELS ACCORDING TO USER FLOW */}
            <div className="grid grid-cols-1 lg:grid-cols-12 gap-8">
              
              {/* PANEL LEFT (MÔ PHỎNG HÌNH 1): KHO TÀI LIỆU LOGISTICS MANIFEST CHỦ */}
              <div className="lg:col-span-5 space-y-4">
                <ManifestUploader
                  onImport={handleDataImported}
                  shipperOptions={uniqueShippers}
                />

                <div className="bg-[#0b101d] rounded-2xl border border-slate-850 shadow-xl overflow-hidden">
                  <div className="p-4 bg-slate-900/60 border-b border-slate-850 flex justify-between items-center">
                    <div className="flex items-center gap-2">
                      <div className="w-1.5 h-3.5 bg-cyan-500 rounded-sm"></div>
                      <h4 className="font-bold text-xs text-slate-200 uppercase tracking-wider font-mono">
                        Sổ Đăng Ký Manifest Gốc (Hình 1)
                      </h4>
                    </div>
                    
                    <button
                      onClick={() => setShowAddForm(!showAddForm)}
                      className="flex items-center gap-1 text-[10px] font-extrabold text-cyan-400 hover:text-cyan-350 border border-cyan-500/20 hover:border-cyan-500/40 bg-cyan-950/20 px-2.5 py-1 rounded-md transition-colors cursor-pointer"
                    >
                      <Plus className="w-3 h-3" />
                      Thêm Vận Đơn
                    </button>
                  </div>

                  {/* Add shipment form popup/inline */}
                  {showAddForm && (
                    <form onSubmit={handleAddManifestRow} className="p-4 bg-[#070b13]/85 border-b border-slate-850 text-xs space-y-3 animate-fade-in">
                      <h5 className="font-semibold text-slate-300 text-[11px] uppercase tracking-wider mb-2">Thêm Mới Vận Đơn (Manifest Line)</h5>
                      <div className="grid grid-cols-2 gap-3">
                        <div className="space-y-1">
                          <label className="text-[9px] font-semibold text-slate-400">Mã MAWB NO (*chứa HAN để đổi tuyến)</label>
                          <input
                            type="text"
                            placeholder="Ví dụ: 994-11827402 HAN"
                            value={newMawb}
                            onChange={(e) => setNewMawb(e.target.value)}
                            className="w-full font-mono bg-[#0b101d] border border-slate-800 rounded p-1.5 text-xs text-slate-100 placeholder-slate-650 focus:ring-1 focus:ring-cyan-500 focus:outline-hidden"
                            required
                          />
                        </div>
                        <div className="space-y-1">
                          <label className="text-[9px] font-semibold text-slate-400">Số HAWB (*để lấy WH charge từ số đuôi)</label>
                          <input
                            type="text"
                            placeholder="Ví dụ: HW-WH-30"
                            value={newHawb}
                            onChange={(e) => setNewHawb(e.target.value)}
                            className="w-full font-mono bg-[#0b101d] border border-slate-800 rounded p-1.5 text-xs text-slate-100 placeholder-slate-650 focus:ring-1 focus:ring-cyan-500 focus:outline-hidden"
                            required
                          />
                        </div>
                      </div>

                      <div className="grid grid-cols-2 gap-3">
                        <div className="space-y-1">
                          <label className="text-[9px] font-semibold text-slate-400">Tên Shipper</label>
                          <select
                            value={newShipper}
                            onChange={(e) => setNewShipper(e.target.value)}
                            className="w-full bg-[#0b101d] border border-slate-800 rounded p-1.5 text-xs text-slate-200 cursor-pointer"
                          >
                            <option value="DYM VIETNAM CO., LTD">DYM VIETNAM CO., LTD</option>
                            <option value="SAMSUNG ELECTRONICS SEV">SAMSUNG ELECTRONICS SEV</option>
                            <option value="FOXCONN VIETNAM CO., LTD">FOXCONN VIETNAM CO., LTD</option>
                          </select>
                        </div>
                        <div className="space-y-1">
                          <label className="text-[9px] font-semibold text-slate-400">Người Nhận (Consignee)</label>
                          <input
                            type="text"
                            placeholder="NIPPON EXPRESS"
                            value={newConsignee}
                            onChange={(e) => setNewConsignee(e.target.value)}
                            className="w-full bg-[#0b101d] border border-slate-800 rounded p-1.5 text-xs text-slate-100 placeholder-slate-600"
                          />
                        </div>
                      </div>

                      <div className="grid grid-cols-3 gap-2">
                        <div className="space-y-1">
                          <label className="text-[9px] font-semibold text-slate-400">W/T (Trọng thô)</label>
                          <input
                            type="number"
                            step="any"
                            value={newWT}
                            onChange={(e) => setNewWT(e.target.value)}
                            className="w-full bg-[#0b101d] border border-slate-800 rounded p-1.5 text-xs text-center font-mono text-slate-100"
                          />
                        </div>
                        <div className="space-y-1">
                          <label className="text-[9px] font-semibold text-slate-400">R.W/T (Tính cước *)</label>
                          <input
                            type="number"
                            step="any"
                            value={newRWT}
                            onChange={(e) => setNewRWT(e.target.value)}
                            className="w-full bg-[#112240] border border-cyan-500/20 font-bold text-cyan-400 rounded p-1.5 text-xs text-center font-mono"
                          />
                        </div>
                        <div className="space-y-1">
                          <label className="text-[9px] font-semibold text-slate-400">V.W (Thể tích)</label>
                          <input
                            type="number"
                            step="any"
                            value={newVW}
                            onChange={(e) => setNewVW(e.target.value)}
                            className="w-full bg-[#0b101d] border border-slate-800 rounded p-1.5 text-xs text-center font-mono text-slate-100"
                          />
                        </div>
                      </div>

                      <div className="flex justify-end gap-2 pt-1.5 border-t border-slate-850">
                        <button
                          type="button"
                          onClick={() => setShowAddForm(false)}
                          className="px-3 py-1 bg-slate-800 text-slate-300 hover:bg-slate-700 hover:text-white rounded text-[11px] cursor-pointer"
                        >
                          Hủy bỏ
                        </button>
                        <button
                          type="submit"
                          className="px-4 py-1 bg-gradient-to-r from-cyan-500 to-emerald-500 text-slate-950 font-bold rounded text-[11px] cursor-pointer"
                        >
                          Lưu Lô Hàng
                        </button>
                      </div>
                    </form>
                  )}

                  {/* DATA LIST HÌNH 1 */}
                  <div className="overflow-x-auto">
                    <table className="w-full text-left border-collapse text-[11px]">
                      <thead>
                        <tr className="bg-slate-950 text-slate-400 border-b border-slate-850 font-bold">
                          <th className="py-2.5 px-3 font-mono font-medium">MAWB NO</th>
                          <th className="py-2.5 px-2 font-mono font-medium">HAWB</th>
                          <th className="py-2.5 px-2">SHIPPER / CNEE</th>
                          <th className="py-2.5 px-2 text-right font-mono">W/T</th>
                          <th className="py-2.5 px-3 text-right text-emerald-400 bg-[#0e2c26] font-bold border-l border-slate-850 font-mono">R.W/T (*)</th>
                          <th className="py-2.5 px-2 text-right font-mono">V.W</th>
                          <th className="py-2.5 px-2 text-center">Xóa</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-850">
                        {manifests.map((m) => {
                          const isSelectedShipper = m.shipper.toUpperCase() === selectedShipper.toUpperCase();
                          return (
                            <tr
                              key={m.id}
                              className={`hover:bg-slate-900/30 transition-colors ${
                                isSelectedShipper ? "bg-emerald-500/10 text-emerald-300 font-medium" : "text-slate-400"
                              }`}
                            >
                              <td className="py-2.5 px-3 font-mono font-semibold text-slate-200">
                                {m.mawbNo}
                                {m.mawbNo.toUpperCase().includes("HAN") && (
                                  <span className="ml-1 px-1 bg-teal-500/15 text-teal-400 rounded text-[9px] font-bold border border-teal-500/20">
                                    HAN Route
                                  </span>
                                )}
                              </td>
                              <td className="py-2.5 px-2 font-mono text-slate-350">{m.hawb}</td>
                              <td className="py-2.5 px-2 truncate max-w-[120px]" title={`${m.shipper} ➔ ${m.consignee}`}>
                                <span className="block text-slate-200 font-medium truncate">{m.shipper}</span>
                                <span className="block text-[10px] text-slate-500 truncate">➔ {m.consignee}</span>
                              </td>
                              <td className="py-2.5 px-2 text-right font-mono">{m.wt.toFixed(1)}</td>
                              <td className="py-2.5 px-3 text-right font-mono text-emerald-450 bg-emerald-500/5 font-bold border-l border-slate-850">
                                {m.rwt.toFixed(1)}
                              </td>
                              <td className="py-2.5 px-2 text-right font-mono">{m.vw.toFixed(1)}</td>
                              <td className="py-2.5 px-2 text-center">
                                <button
                                  onClick={() => handleDeleteRow(m.id)}
                                  className="p-1 text-slate-500 hover:text-rose-400 rounded hover:bg-rose-500/10 cursor-pointer"
                                  title="Xóa dòng vận đơn khỏi kho"
                                >
                                  <Trash2 className="w-3.5 h-3.5" />
                                </button>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </div>

                {/* LOGISTICS CARD MINT TIP */}
                <div className="bg-[#10b981]/10 rounded-xl p-4 border border-[#10b981]/20 text-xs text-emerald-300 leading-relaxed flex gap-3">
                  <CheckCircle2 className="w-5 h-5 text-emerald-400 shrink-0 mt-0.5" />
                  <div>
                    <span className="font-bold block mb-0.5">Sự Phụ Thuộc Số Liệu (Tính Toán Tự Động):</span>
                    Khi người dùng điều chỉnh cột <strong className="font-extrabold text-emerald-450">R.W/T</strong> tại bảng Manifest gốc bên trên, trọng lượng tính cước sẽ ngay lập tức được truyền và tính tiền tại hóa đơn bên cạnh thời gian thực!
                  </div>
                </div>
              </div>


              {/* PANEL RIGHT (MÔ PHỎNG CHI TIẾT HÌNH 2): BẢNG KÊ HÓA ĐƠN THÁNG EXCEL LIVE PREVIEW */}
              <div className="lg:col-span-7 space-y-4">
                <div className="bg-[#0b101d] rounded-2xl border border-slate-850 shadow-xl overflow-hidden relative">
                  
                  {/* Ledger title metadata background */}
                  <div className="p-3 bg-gradient-to-r from-slate-900/80 to-slate-950/40 border-b border-slate-850 flex flex-col md:flex-row justify-between md:items-center gap-2">
                    <span className="text-[10px] uppercase font-bold text-emerald-400 tracking-widest flex items-center gap-1.5 font-mono">
                      <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-400" /> Live Excel Ledger (Hình 2 Sheet)
                    </span>
                    <span className="text-[10px] font-mono text-slate-400">
                      Shipper: <strong className="text-emerald-300 font-semibold uppercase">{selectedShipper}</strong>
                    </span>
                  </div>

                  {/* Dynamic Visual Paper Format Sheet */}
                  <div className="p-5 sm:p-7 overflow-x-auto bg-[#070b13]">
                    {/* Invoice Meta-Header */}
                    <div className="mb-6 text-center">
                      <h4 className="text-sm font-bold text-slate-100 leading-snug tracking-widest uppercase border-b border-slate-850 pb-2.5 font-mono">
                        BẢNG KÊ HÓA ĐƠN THÁNG {String(reportMonth).padStart(2, "0")} NĂM {reportYear}
                        <span className="block text-emerald-400 text-xs mt-1.5 font-extrabold font-mono tracking-wide">
                          - [ SHIPPER: {selectedShipper.toUpperCase()} ] -
                        </span>
                      </h4>
                      <p className="text-[10px] italic font-semibold text-cyan-400 mt-2 font-mono">
                        Tỷ giá bảo đảm (Exchange Rate): <strong className="font-mono text-[11px] font-bold underline bg-cyan-950/45 text-cyan-300 border border-cyan-500/25 px-1.5 py-0.5 rounded ml-1">{exchangeRate.toLocaleString("vi-VN")} VND/USD</strong>
                      </p>
                    </div>

                    {/* HÌNH 2 MAIN TABLE */}
                    {shipperBillingRows.length === 0 ? (
                      <div className="py-12 text-center text-slate-500 text-xs space-y-2">
                        <TrendingUp className="w-8 h-8 text-slate-600 mx-auto stroke-1" />
                        <p>Không tìm thấy lô hàng nào trong kho ứng với Shipper <strong>{selectedShipper}</strong></p>
                        <p className="text-[10px] text-slate-500">Hãy thêm vận đơn hoặc đổi tên Shipper gốc ở cột trái.</p>
                      </div>
                    ) : (
                      <div className="border border-slate-850 rounded-lg overflow-hidden">
                        <table className="w-full text-left border-collapse text-[10px]">
                          <thead>
                            <tr className="bg-[#0b101c] text-emerald-400 font-semibold font-mono text-center border-b border-emerald-500/30">
                              <th className="py-2.5 px-1 border-r border-slate-850 w-8">NO.</th>
                              <th className="py-2.5 px-1.5 border-r border-slate-850">MAWB NO</th>
                              <th className="py-2.5 px-1.5 border-r border-slate-850 w-16">HAWB</th>
                              <th className="py-2.5 px-1.5 border-r border-slate-850 w-20">ROUTE</th>
                              <th className="py-2.5 px-1.5 border-r border-slate-850 text-right">WEIGHT (kg)</th>
                              <th className="py-2.5 px-2 border-r border-slate-850 text-right w-24">UNIT PRICE ($)</th>
                              <th className="py-2.5 px-1.5 border-r border-slate-850 text-right">FREIGHT</th>
                              <th className="py-2.5 px-1.5 border-r border-slate-850 text-right w-16">HANDLING</th>
                              <th className="py-2.5 px-1.5 border-r border-slate-850 text-right w-18">WAREHOUSE</th>
                              <th className="py-2.5 px-1.5 border-r border-slate-850 text-right w-16">OTH.CHG</th>
                              <th className="py-2.5 px-2 text-right font-bold text-emerald-300 bg-emerald-900/10">TOTAL (USD)</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-850">
                            {shipperBillingRows.map((row, i) => (
                              <tr key={row.id} className="hover:bg-slate-900/20 transition-colors">
                                {/* NO */}
                                <td className="py-2 px-1 text-center font-mono text-slate-500 border-r border-slate-850 bg-slate-950/20">{i + 1}</td>
                                
                                {/* MAWB */}
                                <td className="py-2 px-1.5 border-r border-slate-850 font-mono text-slate-300">{row.mawbNo}</td>
                                
                                {/* HAWB */}
                                <td className="py-2 px-1.5 border-r border-slate-850 font-mono text-slate-100 font-semibold">{row.hawb}</td>
                                
                                {/* ROUTE */}
                                <td className="py-2 px-1.5 border-r border-slate-850 text-center">
                                  <span className={`px-1.5 py-0.5 rounded text-[9px] font-bold ${
                                    row.route === "SGN-HAN-ICN"
                                      ? "bg-amber-500/10 text-amber-400 border border-amber-500/20"
                                      : "bg-blue-500/10 text-blue-400 border border-blue-500/20"
                                  }`}>
                                    {row.route}
                                  </span>
                                </td>

                                {/* WEIGHT */}
                                <td className="py-2 px-1.5 border-r border-slate-850 text-right font-mono font-semibold text-[#cbd5e1] bg-slate-950/10">
                                  {row.weight.toLocaleString("en-US", { minimumFractionDigits: 1, maximumFractionDigits: 1 })}
                                </td>

                                {/* UNIT PRICE (Dynamic input *) */}
                                <td className="py-1 px-1 border-r border-slate-850 text-right">
                                  <div className="relative">
                                    <input
                                      type="number"
                                      step="0.01"
                                      className="w-full text-right bg-amber-500/10 border border-amber-500/20 rounded font-mono font-bold text-xs p-1 text-amber-350 focus:bg-[#070b13] focus:outline-hidden focus:border-amber-400 transition-all focus:ring-1 focus:ring-amber-950"
                                      value={row.unitPrice}
                                      onChange={(e) => handleUpdateUnitPrice(row.id, e.target.value)}
                                    />
                                    <span className="absolute left-1 top-1.5 text-[9px] text-amber-500">$</span>
                                  </div>
                                </td>

                                {/* FREIGHT CHARGE */}
                                <td className="py-2 px-1.5 border-r border-slate-850 text-right font-mono text-slate-400">
                                  ${row.freightCharge.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                                </td>

                                {/* HANDLING CHARGE */}
                                <td className="py-2 px-1.5 border-r border-slate-850 text-right font-mono text-slate-400 bg-slate-950/10">
                                  ${row.handlingCharge.toFixed(2)}
                                </td>

                                {/* WAREHOUSE CHARGE (parsed dynamic input *) */}
                                <td className="py-1 px-1 border-r border-slate-850 text-right">
                                  <div className="relative">
                                    <input
                                      type="number"
                                      step="any"
                                      className="w-full text-right bg-slate-950/50 hover:bg-[#0b101d] border border-slate-800 focus:border-emerald-500 rounded font-mono text-xs p-1 text-slate-300"
                                      value={row.warehouseCharge}
                                      onChange={(e) => handleUpdateWarehouseCharge(row.id, e.target.value)}
                                    />
                                    <span className="absolute left-1 top-1.5 text-[8px] text-slate-500">$</span>
                                  </div>
                                </td>

                                {/* OTH.CHG */}
                                <td className="py-1 px-1 border-r border-slate-850 text-right">
                                  <div className="relative">
                                    <input
                                      type="number"
                                      step="any"
                                      className="w-full text-right bg-slate-950/50 hover:bg-[#0b101d] border border-slate-800 focus:border-emerald-500 rounded font-mono text-xs p-1 text-slate-300"
                                      value={row.otherCharge}
                                      onChange={(e) => handleUpdateOtherCharge(row.id, e.target.value)}
                                    />
                                    <span className="absolute left-1 top-1.5 text-[8px] text-slate-500">$</span>
                                  </div>
                                </td>

                                {/* ROW TOTAL TO MÀU VÀNG (Yellow Highlight *) */}
                                <td className="py-2 px-2.5 text-right font-mono font-bold text-amber-300 bg-amber-500/10">
                                  ${row.totalUsd.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                                </td>
                              </tr>
                            ))}

                            {/* TOTAL SUMMATION FOOTER ROW (Màu mint viền đậm *) */}
                            <tr className="bg-[#0b101c] text-emerald-400 font-bold border-t-2 border-b-2 border-emerald-500/30">
                              <td colSpan={4} className="py-3 px-2 text-center border-r border-slate-850 uppercase tracking-widest text-[9px] font-mono">
                                TOTAL / TỔNG CỘNG
                              </td>
                              
                              {/* SUM WEIGHT */}
                              <td className="py-3 px-1.5 text-right font-mono border-r border-slate-850 text-slate-200">
                                {selectedShipperTotals.totalWeight.toLocaleString("en-US", { minimumFractionDigits: 1, maximumFractionDigits: 1 })}
                              </td>

                              {/* UNIT PRICE SUM PLACEHOLDER */}
                              <td className="py-3 px-1.5 text-center font-mono border-r border-slate-850 text-slate-500">-</td>

                              {/* SUM FREIGHT CHARGE */}
                              <td className="py-3 px-1.5 text-right font-mono border-r border-slate-850">
                                ${selectedShipperTotals.totalFreightUSD.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                              </td>

                              {/* SUM HANDLING CHARGE */}
                              <td className="py-3 px-1.5 text-right font-mono border-r border-slate-850">
                                ${selectedShipperTotals.totalHandlingUSD.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                              </td>

                              {/* SUM WAREHOUSE CHARGE */}
                              <td className="py-3 px-1.5 text-right font-mono border-r border-slate-850">
                                ${selectedShipperTotals.totalWarehouseUSD.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                              </td>

                              {/* SUM OTHER */}
                              <td className="py-3 px-1.5 text-right font-mono border-r border-slate-850">
                                ${selectedShipperTotals.totalOtherUSD.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                              </td>

                              {/* SUM GRAND TOTAL USD WITH YELLOW HIGHLIGHT AND RED COLOR */}
                              <td className="py-3 px-2.5 text-right font-mono text-emerald-300 bg-emerald-500/10 font-black text-[11px]">
                                ${selectedShipperTotals.totalUSD.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                              </td>
                            </tr>
                          </tbody>
                        </table>
                      </div>
                    )}

                    {/* VƯỢT TRỘI: ROW TỔNG TIỀN QUY ĐỔI SANG VND (Yellow + Bold + Double Border) */}
                    {shipperBillingRows.length > 0 && (
                      <div className="mt-6 border border-emerald-500/20 bg-emerald-500/5 rounded-xl p-4 flex flex-col sm:flex-row justify-between items-center gap-4 shadow-md relative overflow-hidden">
                        {/* Decorative watermark */}
                        <div className="absolute right-2 bottom-0 opacity-[0.02] text-cyan-400 font-extrabold text-7xl select-none pointer-events-none font-mono">VND</div>
                        
                        <div className="text-center sm:text-left z-10">
                          <span className="block text-[11px] font-bold text-emerald-400 uppercase tracking-widest font-mono">
                            TỔNG TIỀN QUY TIỀN THỰC TẾ (GRAND TOTAL CONVERTED):
                          </span>
                          <span className="text-[10px] text-slate-400 italic block mt-0.5">
                            Công thức: [Tổng tiền USD] x [Tỷ giá {exchangeRate.toLocaleString("vi-VN")} VND/USD]
                          </span>
                        </div>
                        
                        <div className="text-center sm:text-right bg-[#070b13]/80 border border-slate-850 px-4 py-2 rounded-lg z-10">
                          <span className="text-xs font-mono font-semibold text-slate-400 block md:inline-block mr-1">
                            ${selectedShipperTotals.totalUSD.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} × {exchangeRate.toLocaleString("vi-VN")} =
                          </span>
                          <span className="text-lg font-mono font-bold text-emerald-400 block md:inline-block">
                            {selectedShipperTotals.totalVND.toLocaleString("vi-VN")} VND
                          </span>
                        </div>
                      </div>
                    )}

                    {/* Footer Signature Sample block from Hình 2 */}
                    {shipperBillingRows.length > 0 && (
                      <div className="mt-8 pt-4 border-t border-slate-100 grid grid-cols-2 text-center text-[10px] text-slate-400">
                        <div>
                          <p className="font-bold text-slate-700">Người Lập Biểu / Prepared By</p>
                          <p className="italic mt-0.5">(Ký, ghi rõ họ tên)</p>
                          <div className="h-16"></div>
                          <p className="text-slate-500 font-semibold font-mono">Planners Team</p>
                        </div>
                        <div>
                          <p className="font-bold text-slate-700">Giám Đốc / Approved By</p>
                          <p className="italic mt-0.5">(Ký tên, đóng dấu)</p>
                          <div className="h-16"></div>
                          <p className="text-slate-500 font-semibold">Prestige Board Director</p>
                        </div>
                      </div>
                    )}

                  </div>
                </div>
              </div>

            </div>

          </div>
        )}


        {/* TAB 2: DEVELOPER DESK (GIẢI TRÌNH KIẾN TRÚC & MÃ NGUỒN PYTHON) */}
        {activeTab === "developer" && (
          <div className="space-y-6 animate-fade-in text-slate-300">
            {/* Header subtabs */}
            <div className="flex border-b border-slate-800 gap-6 mb-2">
              <button
                onClick={() => setDeveloperSubTab("python")}
                className={`pb-3 text-xs font-bold font-mono transition-all relative cursor-pointer ${
                  developerSubTab === "python"
                    ? "text-emerald-400 border-b-2 border-emerald-400 font-extrabold"
                    : "text-slate-500 hover:text-slate-350"
                }`}
              >
                Mã Nguồn Python (Pandas & OpenPyXL)
              </button>
              <button
                onClick={() => setDeveloperSubTab("database")}
                className={`pb-3 text-xs font-bold font-mono transition-all relative cursor-pointer ${
                  developerSubTab === "database"
                    ? "text-emerald-400 border-b-2 border-emerald-400 font-extrabold"
                    : "text-slate-500 hover:text-slate-350"
                }`}
              >
                Kiến Trúc Lưu Trữ (SQLite Schema)
              </button>
            </div>

            {/* Subtab contents */}
            {developerSubTab === "database" ? <DatabaseDesign /> : <PythonScriptView />}
          </div>
        )}

      </main>

      {/* FOOTER */}
      <footer className="border-t border-slate-850/80 py-6 mt-16 text-center text-xs text-slate-500 bg-[#0b0f19]/40 font-mono">
        <p>© 2026 Prestige Logistics Automation Hub. Phát triển dựa trên giải pháp kết xuất Excel nâng cao.</p>
      </footer>
    </div>
  );
}
