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
    id: "M004",
    mawbNo: "180-18407734 TCS",
    hawb: "HW-WH-12",
    shipper: "DYM VIETNAM CO., LTD",
    consignee: "CJ LOGISTICS SEOUL",
    wt: 12.4,
    rwt: 15.0,
    vw: 10.0,
    date: "2026-06-16"
  }
];

export const DEMO_UNIT_PRICES: Record<string, number> = {
  "M001": 3.40,
  "M002": 3.40,
  "M004": 4.10
};
