/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

export interface ManifestRow {
  id: string;
  mawbNo: string;
  hawb: string;
  shipper: string;
  consignee: string;
  wt: number;     // Gross weight (W/T)
  rwt: number;    // Chargeable weight (R.W/T)
  vw: number;     // Volume weight (V.W)
  date?: string;   // Date of manifest
  etd?: string;    // Estimated Time of Departure
  eta?: string;    // Estimated Time of Arrival
  bill?: string;   // Bill number
  markHawb?: string; // Mark (HAWB)
}

export interface BillingCalculation {
  id: string;
  manifestId: string;
  mawbNo: string;
  hawb: string;
  shipper: string;
  route: "SGN-HAN-ICN" | "SGN-ICN";
  weight: number;          // R.W/T
  unitPrice: number;       // USD / kg (user inputable)
  freightCharge: number;   // Weight * Unit Price
  handlingCharge: number;  // Fixed 10.00 USD
  warehouseCharge: number; // Derived from HAWB reference numeric or manually edited
  otherCharge: number;     // Oth.Chg (editable, default 0)
  totalUsd: number;        // Freight + Handling + Warehouse + Other
  etd?: string;
  eta?: string;
  bill?: string;
  markHawb?: string;
  date?: string;
}

export interface ShipperSummary {
  shipperName: string;
  totalWeight: number;
  totalFreightUSD: number;
  totalHandlingUSD: number;
  totalWarehouseUSD: number;
  totalOtherUSD: number;
  totalUSD: number;
  totalVND: number;
}
