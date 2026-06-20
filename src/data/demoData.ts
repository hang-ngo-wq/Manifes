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
    date: "2026-06-15"
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
    date: "2026-06-15"
  },
  {
    id: "M003",
    mawbNo: "180-18407734 TCS",
    hawb: "HW-WH-80",
    shipper: "SAMSUNG ELECTRONICS SEV",
    consignee: "SAMSUNG HQ SEOUL",
    wt: 520.0,
    rwt: 537.2,
    vw: 490.0,
    date: "2026-06-16"
  },
  {
    id: "M004",
    mawbNo: "180-18407734 TCS",
    hawb: "HW-WH-12",
    shipper: "DYM VIETNAM CO., LTD",
    consignee: "CJ LOGISTICS SEOUL",
    wt: 12.4,
    rwt: 15.0,
    vw: 10.0,
    date: "2026-06-16"
  },
  {
    id: "M005",
    mawbNo: "994-33069691 T/S HAN",
    hawb: "HW-WH-60",
    shipper: "SAMSUNG ELECTRONICS SEV",
    consignee: "SAMSUNG HQ SEOUL",
    wt: 310.0,
    rwt: 315.0,
    vw: 280.0,
    date: "2026-06-17"
  },
  {
    id: "M006",
    mawbNo: "120-45920391 RE-ROUTED",
    hawb: "HW-WH-35",
    shipper: "FOXCONN VIETNAM CO., LTD",
    consignee: "FOXCONN GLOBAL INC",
    wt: 185.0,
    rwt: 190.0,
    vw: 175.0,
    date: "2026-06-18"
  },
  {
    id: "M007",
    mawbNo: "TS HAN 994-44029103",
    hawb: "HW-WH-110",
    shipper: "FOXCONN VIETNAM CO., LTD",
    consignee: "FOXCONN GLOBAL INC",
    wt: 450.0,
    rwt: 460.0,
    vw: 430.0,
    date: "2026-06-19"
  }
];

export const DEMO_UNIT_PRICES: Record<string, number> = {
  "M001": 3.40,
  "M002": 3.40,
  "M003": 2.85,
  "M004": 4.10,
  "M005": 2.95,
  "M006": 3.20,
  "M007": 3.10
};
