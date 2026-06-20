/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { ManifestRow } from "../types";

export const INITIAL_MANIFEST_DATA: ManifestRow[] = [
  {
    id: "M001",
    mawbNo: "994-33069691 T/S HAN",
    hawb: "HW-WH-45",
    shipper: "DYM VIETNAM CO., LTD",
    consignee: "NIPPON EXPRESS KOREA",
    wt: 165.5,
    rwt: 158.0,
    vw: 140.0,
    date: "2026-06-15",
    etd: "15/06/2026",
    eta: "16/06/2026",
    bill: "800004204251",
    markHawb: "HW-WH-45"
  },
  {
    id: "M002",
    mawbNo: "TS HAN 994-38098642",
    hawb: "HW-WH-25",
    shipper: "DYM VIETNAM CO., LTD",
    consignee: "KUEHNE + NAGEL KOREA",
    wt: 85.0,
    rwt: 92.5,
    vw: 92.5,
    date: "2026-06-15",
    etd: "15/06/2026",
    eta: "16/06/2026",
    bill: "800004204252",
    markHawb: "HW-WH-25"
  },
  {
    id: "M004",
    mawbNo: "180-18407734 TCS",
    hawb: "HW-WH-12",
    shipper: "GNG",
    consignee: "CJ LOGISTICS SEOUL",
    wt: 56.0,
    rwt: 56.0,
    vw: 50.0,
    date: "2026-06-01",
    etd: "01/06/2026",
    eta: "2/6/2026",
    bill: "800004204257",
    markHawb: ""
  }
];

export const DEMO_UNIT_PRICES: Record<string, number> = {
  "M001": 3.40,
  "M002": 3.40,
  "M004": 3.20
};

export const DEMO_WAREHOUSE_CHARGES: Record<string, number> = {
  "M001": 45,
  "M002": 25,
  "M004": 2
};

