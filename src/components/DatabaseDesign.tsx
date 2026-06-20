/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState } from "react";
import { Database, Copy, Check, Table2, KeyRound } from "lucide-react";

export default function DatabaseDesign() {
  const [copied, setCopied] = useState(false);

  const sqlSchema = `-- ==========================================
-- KIẾN TRÚC CƠ SỞ DỮ LIỆU LOGISTICS MANIFEST
-- Hệ thống lưu trữ, áp đơn giá và tính tiền tự động
-- Hệ quản trị khuyến nghị: SQLite / PostgreSQL
-- ==========================================

-- 1. Bảng lưu trữ thông tin Shipper (Khách hàng gửi hàng)
CREATE TABLE IF NOT EXISTS shippers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL UNIQUE,          -- Tên khách hàng (Ví dụ: DYM VIETNAM CO., LTD)
    tax_code TEXT,                      -- Mã số thuế dùng xuất hóa đơn
    address TEXT,                       -- Địa chỉ shipper
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- 2. Bảng lưu trữ cấu hình Đơn giá cơ bản (Billing Rates) theo Shipper và Tuyến đường
CREATE TABLE IF NOT EXISTS billing_rates (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    shipper_id INTEGER NOT NULL,
    route TEXT NOT NULL,                -- 'SGN-HAN-ICN' hoặc 'SGN-ICN'
    unit_price_usd REAL NOT NULL,       -- Đơn giá (USD / kg)
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (shipper_id) REFERENCES shippers(id) ON DELETE CASCADE,
    UNIQUE(shipper_id, route)           -- Đảm bảo mỗi Khách hàng chỉ có 1 rate cho mỗi Tuyến đường
);

-- 3. Bảng lưu trữ tài liệu Manifest gốc (Tương đương dữ liệu ở Hình 1)
CREATE TABLE IF NOT EXISTS manifests (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    mawb_no TEXT NOT NULL,              -- Mã vận đơn chủ (Có chữ 'HAN' hoặc không)
    hawb TEXT NOT NULL,                 -- Mã vận đơn phụ (Ví dụ: HW-WH-45)
    shipper_id INTEGER NOT NULL,        -- Khóa ngoại liên kết bảng shippers
    consignee TEXT,                     -- Người nhận hàng
    wt REAL NOT NULL,                   -- Trọng lượng thô (W/T)
    rwt REAL NOT NULL,                  -- Trọng lượng tính cước (R.W/T)
    vw REAL NOT NULL,                   -- Trọng lượng thể tích (V.W)
    manifest_date DATE NOT NULL,        -- Ngày bay / Nhập manifest
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (shipper_id) REFERENCES shippers(id)
);

-- 4. Bảng tính toán Phí hóa đơn chi tiết (Tương đương dữ liệu tính toán cho Hình 2)
CREATE TABLE IF NOT EXISTS billing_calculations (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    manifest_id INTEGER NOT NULL UNIQUE, -- Liên kết 1-1 với dòng Manifest gốc
    route TEXT NOT NULL,                 -- Tuyến đường tự động phân tích ('SGN-HAN-ICN' | 'SGN-ICN')
    chargeable_weight REAL NOT NULL,     -- WEIGHT lấy từ R.W/T
    unit_price_usd REAL NOT NULL,        -- Đơn giá áp dụng tại thời điểm tính tiền
    freight_charge_usd REAL NOT NULL,    -- Tiền vận chuyển: Weight * Unit Price
    handling_charge_usd REAL NOT NULL DEFAULT 10.00, -- Phí xử lý cố định (10.00 USD)
    warehouse_charge_usd REAL NOT NULL,  -- Phí kho bãi trích từ HAWB liên quan
    other_charge_usd REAL NOT NULL DEFAULT 0.00, -- Phí khác
    total_usd REAL NOT NULL,             -- Tổng cộng USD
    exchange_rate REAL NOT NULL,         -- Tỷ giá quy đổi áp dụng (Ví dụ: 26160)
    total_vnd REAL NOT NULL,             -- Tổng tiền quy đổi (VND) = Total USD * Exchange Rate
    calculated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (manifest_id) REFERENCES manifests(id) ON DELETE CASCADE
);

-- CHỈ MỤC TỐI ƯU HÓA TRUY VẤN
CREATE INDEX IF NOT EXISTS idx_manifest_shipper ON manifests(shipper_id);
CREATE INDEX IF NOT EXISTS idx_billing_manifest ON billing_calculations(manifest_id);
`;

  const handleCopy = () => {
    navigator.clipboard.writeText(sqlSchema);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  return (
    <div className="bg-slate-900 text-slate-100 rounded-2xl p-6 shadow-xl border border-slate-800">
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 mb-6 pb-6 border-b border-slate-800">
        <div className="flex items-center gap-3">
          <div className="p-2.5 bg-sky-500/10 rounded-xl text-sky-400">
            <Database className="w-6 h-6" />
          </div>
          <div>
            <h3 className="font-semibold text-lg text-white">Kiến Trúc Dữ Liệu Quan Hệ (Database Schema)</h3>
            <p className="text-xs text-slate-400">Thiết kế bảng chuẩn hóa để tối ưu hóa việc quản lý và tự động hóa quy trình</p>
          </div>
        </div>
        <button
          onClick={handleCopy}
          className="flex items-center gap-2 px-3.5 py-1.5 bg-slate-800 hover:bg-slate-700 active:bg-slate-600 text-xs font-medium text-slate-200 hover:text-white rounded-lg transition-colors border border-slate-700"
        >
          {copied ? (
            <>
              <Check className="w-3.5 h-3.5 text-green-400" />
              <span className="text-green-400">Đã sao chép!</span>
            </>
          ) : (
            <>
              <Copy className="w-3.5 h-3.5" />
              <span>Sao chép SQL</span>
            </>
          )}
        </button>
      </div>

      {/* Visual Diagram Representation */}
      <h4 className="text-sm font-semibold text-slate-200 mb-4 flex items-center gap-2">
        <Table2 className="w-4 h-4 text-emerald-400" /> Sơ đồ liên kết thực thể (ERD Concept)
      </h4>
      <div className="grid grid-cols-1 md:grid-cols-4 gap-4 mb-8">
        <div className="bg-slate-950 p-4 rounded-xl border border-rose-500/30">
          <div className="text-xs font-bold text-rose-400 mb-2 border-b border-rose-500/20 pb-1">shippers (Khách hàng)</div>
          <ul className="space-y-1 text-[11px] font-mono text-slate-300">
            <li className="flex items-center gap-1"><KeyRound className="w-3 h-3 text-yellow-500" />id (PK)</li>
            <li>name (Text, Unique)</li>
            <li>tax_code (Text)</li>
            <li>address (Text)</li>
          </ul>
        </div>

        <div className="bg-slate-950 p-4 rounded-xl border border-sky-500/30">
          <div className="text-xs font-bold text-sky-400 mb-2 border-b border-sky-500/20 pb-1">billing_rates (Đơn Giá)</div>
          <ul className="space-y-1 text-[11px] font-mono text-slate-300">
            <li className="flex items-center gap-1"><KeyRound className="w-3 h-3 text-yellow-500" />id (PK)</li>
            <li className="text-sky-300 opacity-90">shipper_id (FK)</li>
            <li>route (Text)</li>
            <li>unit_price_usd (Real)</li>
          </ul>
        </div>

        <div className="bg-slate-950 p-4 rounded-xl border border-emerald-500/30">
          <div className="text-xs font-bold text-emerald-400 mb-2 border-b border-emerald-500/20 pb-1">manifests (Hình 1)</div>
          <ul className="space-y-1 text-[11px] font-mono text-slate-300">
            <li className="flex items-center gap-1"><KeyRound className="w-3 h-3 text-yellow-500" />id (PK)</li>
            <li>mawb_no (Text)</li>
            <li>hawb (Text)</li>
            <li className="text-rose-300 opacity-95">shipper_id (FK)</li>
            <li>consignee (Text)</li>
            <li>wt, rwt, vw (Real)</li>
          </ul>
        </div>

        <div className="bg-slate-950 p-4 rounded-xl border border-amber-500/30">
          <div className="text-xs font-bold text-amber-400 mb-2 border-b border-amber-500/20 pb-1">calculations (Hình 2)</div>
          <ul className="space-y-1 text-[11px] font-mono text-slate-300">
            <li className="flex items-center gap-1"><KeyRound className="w-3 h-3 text-yellow-500" />id (PK)</li>
            <li className="text-emerald-300 opacity-95">manifest_id (FK)</li>
            <li>route (SGN-HAN-ICN/SGN-ICN)</li>
            <li>chargeable_weight</li>
            <li>unit_price_usd</li>
            <li>freight_charge_usd</li>
            <li>handling_charge_usd (10)</li>
            <li>warehouse_charge_usd</li>
            <li>total_usd, total_vnd</li>
          </ul>
        </div>
      </div>

      {/* Database explanations */}
      <div className="bg-slate-950 rounded-xl p-4 mb-6 text-sm text-slate-300 space-y-3 border border-slate-800">
        <h5 className="font-semibold text-white">💡 Ưu điểm của kiến trúc cơ sở dữ liệu này:</h5>
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4 text-xs">
          <div className="p-3 bg-slate-900 rounded-lg border border-slate-800">
            <span className="font-semibold text-sky-400 block mb-1">1. Tránh trùng lặp dữ liệu (Normalization)</span>
            Dữ liệu Shipper và Đơn giá (Rates) được lưu riêng biệt. Khi có nhiều lô hàng của một Shipper, hệ thống chỉ cần lưu liên kết <code className="text-rose-300">shipper_id</code>, không lặp lại chuỗi văn bản dài.
          </div>
          <div className="p-3 bg-slate-900 rounded-lg border border-slate-800">
            <span className="font-semibold text-emerald-400 block mb-1">2. Kiểm soát lịch sử tài chính</span>
            Bảng <code className="text-amber-400">billing_calculations</code> lưu lại <code className="text-amber-400">unit_price_usd</code> và <code className="text-emerald-400">exchange_rate</code> chính xác tại THỜI ĐIỂM tạo hóa đơn, giúp lưu trữ lịch sử báo cáo kể cả khi bảng giá thay đổi về sau.
          </div>
        </div>
      </div>

      {/* Code Area */}
      <h4 className="text-sm font-semibold text-slate-200 mb-2 flex items-center gap-2">SQL DDL Commands</h4>
      <div className="bg-slate-950 rounded-xl p-4 overflow-x-auto border border-slate-800 text-xs leading-relaxed max-h-96">
        <pre className="font-mono text-emerald-400">
          <code>{sqlSchema}</code>
        </pre>
      </div>
    </div>
  );
}
