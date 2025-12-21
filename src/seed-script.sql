BEGIN TRANSACTION;

-- =====================================================
-- 1. Drop existing tables
-- =====================================================
DROP TABLE IF EXISTS PolygonData;
DROP TABLE IF EXISTS Cells;
DROP TABLE IF EXISTS Pages;
DROP TABLE IF EXISTS RawJson;

-- =====================================================
-- 2. Recreate Pages table
-- =====================================================
CREATE TABLE Pages (
    PageID        INTEGER PRIMARY KEY AUTOINCREMENT,
    PageKey       TEXT NOT NULL,        -- "1", "1.1", "1.2"
    PdfPageNumber INTEGER NOT NULL,     -- 1
    Name          TEXT NOT NULL,        -- PgNo-1, PgNo-1.1
    UNIQUE (PageKey)
);

-- =====================================================
-- 3. Recreate Cells table (Excel-aware)
-- =====================================================
CREATE TABLE Cells (
    CellPK       INTEGER PRIMARY KEY AUTOINCREMENT,
    PageID       INTEGER NOT NULL,
    CellID       TEXT NOT NULL,
    CellName     TEXT NOT NULL,
    Value        TEXT,
    ExcelRow     INTEGER NOT NULL,
    ExcelColumn  INTEGER NOT NULL,
    UNIQUE (PageID, CellID),
    UNIQUE (PageID, ExcelRow, ExcelColumn),
    FOREIGN KEY (PageID) REFERENCES Pages(PageID)
);

-- =====================================================
-- 4. Recreate PolygonData table
-- =====================================================
CREATE TABLE PolygonData (
    PolygonPointID INTEGER PRIMARY KEY AUTOINCREMENT,
    CellPK         INTEGER NOT NULL,
    PointOrder     INTEGER NOT NULL,
    X              REAL NOT NULL,
    Y              REAL NOT NULL,
    FOREIGN KEY (CellPK) REFERENCES Cells(CellPK),
    UNIQUE (CellPK, PointOrder)
);

-- =====================================================
-- 5. Temp table for JSON input
-- =====================================================
CREATE TEMP TABLE RawJson (
    json TEXT NOT NULL
);

-- =====================================================
-- 5.5 Table for file data
-- =====================================================
CREATE TABLE DataFiles (
    DataFileID          INTEGER PRIMARY KEY AUTOINCREMENT,
    FileName            TEXT NOT NULL,
    XmlPartName         TEXT NOT NULL,
    RawFileSize         INTEGER NOT NULL,
    CompressedFileSize  INTEGER NOT NULL,
    CreatedUtc          TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now')),

    UNIQUE (FileName),

    CHECK (length(trim(FileName)) > 0),
    CHECK (length(trim(XmlPartName)) > 0),
    CHECK (RawFileSize >= 0),
    CHECK (CompressedFileSize >= 0)
);

CREATE INDEX IF NOT EXISTS IX_DataFiles_XmlPartName ON DataFiles(XmlPartName);

-- =====================================================
-- 6. Paste JSON here ONCE
-- =====================================================
INSERT INTO RawJson (json) VALUES (
'
  {
  "Pages": [
    {
      "1": {
        "Cells": [
          {
            "CellId": "0,0",
            "CellName": "A1",
            "Value": "",
            "Polygon": [
              {
                "X": 0.2491,
                "Y": 0.8214
              },
              {
                "X": 4.4647,
                "Y": 0.8214
              },
              {
                "X": 4.4647,
                "Y": 1.2928
              },
              {
                "X": 0.2409,
                "Y": 1.2928
              }
            ]
          },
          {
            "CellId": "0,1",
            "CellName": "B1",
            "Value": "First quarter ended March 31,",
            "Polygon": [
              {
                "X": 4.4647,
                "Y": 0.8214
              },
              {
                "X": 8.2414,
                "Y": 0.8297
              },
              {
                "X": 8.2414,
                "Y": 1.1357
              },
              {
                "X": 4.4647,
                "Y": 1.1357
              }
            ]
          },
          {
            "CellId": "1,1",
            "CellName": "B2",
            "Value": "2022",
            "Polygon": [
              {
                "X": 4.4647,
                "Y": 1.1357
              },
              {
                "X": 6.9742,
                "Y": 1.1357
              },
              {
                "X": 6.9742,
                "Y": 1.2928
              },
              {
                "X": 4.4647,
                "Y": 1.2928
              }
            ]
          },
          {
            "CellId": "1,2",
            "CellName": "C2",
            "Value": "2021",
            "Polygon": [
              {
                "X": 6.9742,
                "Y": 1.1357
              },
              {
                "X": 8.2414,
                "Y": 1.1357
              },
              {
                "X": 8.2414,
                "Y": 1.3011
              },
              {
                "X": 6.9742,
                "Y": 1.2928
              }
            ]
          },
          {
            "CellId": "2,0",
            "CellName": "A3",
            "Value": "Total Segment Adjusted EBITDA",
            "Polygon": [
              {
                "X": 0.2409,
                "Y": 1.2928
              },
              {
                "X": 4.4647,
                "Y": 1.2928
              },
              {
                "X": 4.4647,
                "Y": 1.4583
              },
              {
                "X": 0.2409,
                "Y": 1.4583
              }
            ]
          },
          {
            "CellId": "2,1",
            "CellName": "B3",
            "Value": "$ 1,013",
            "Polygon": [
              {
                "X": 4.4647,
                "Y": 1.2928
              },
              {
                "X": 6.9742,
                "Y": 1.2928
              },
              {
                "X": 6.9742,
                "Y": 1.4583
              },
              {
                "X": 4.4647,
                "Y": 1.4583
              }
            ]
          },
          {
            "CellId": "2,2",
            "CellName": "C3",
            "Value": "$ 569",
            "Polygon": [
              {
                "X": 6.9742,
                "Y": 1.2928
              },
              {
                "X": 8.2414,
                "Y": 1.3011
              },
              {
                "X": 8.2414,
                "Y": 1.4583
              },
              {
                "X": 6.9742,
                "Y": 1.4583
              }
            ]
          },
          {
            "CellId": "3,0",
            "CellName": "A4",
            "Value": "Unallocated amounts:",
            "Polygon": [
              {
                "X": 0.2409,
                "Y": 1.4583
              },
              {
                "X": 4.4647,
                "Y": 1.4583
              },
              {
                "X": 4.4647,
                "Y": 1.6154
              },
              {
                "X": 0.2409,
                "Y": 1.6154
              }
            ]
          },
          {
            "CellId": "3,1",
            "CellName": "B4",
            "Value": "",
            "Polygon": [
              {
                "X": 4.4647,
                "Y": 1.4583
              },
              {
                "X": 6.9742,
                "Y": 1.4583
              },
              {
                "X": 6.9742,
                "Y": 1.6154
              },
              {
                "X": 4.4647,
                "Y": 1.6154
              }
            ]
          },
          {
            "CellId": "3,2",
            "CellName": "C4",
            "Value": "",
            "Polygon": [
              {
                "X": 6.9742,
                "Y": 1.4583
              },
              {
                "X": 8.2414,
                "Y": 1.4583
              },
              {
                "X": 8.2414,
                "Y": 1.6154
              },
              {
                "X": 6.9742,
                "Y": 1.6154
              }
            ]
          },
          {
            "CellId": "4,0",
            "CellName": "A5",
            "Value": "Transformation(1)",
            "Polygon": [
              {
                "X": 0.2409,
                "Y": 1.6154
              },
              {
                "X": 4.4647,
                "Y": 1.6154
              },
              {
                "X": 4.4647,
                "Y": 1.7808
              },
              {
                "X": 0.2409,
                "Y": 1.7808
              }
            ]
          },
          {
            "CellId": "4,1",
            "CellName": "B5",
            "Value": "(14)",
            "Polygon": [
              {
                "X": 4.4647,
                "Y": 1.6154
              },
              {
                "X": 6.9742,
                "Y": 1.6154
              },
              {
                "X": 6.9742,
                "Y": 1.7808
              },
              {
                "X": 4.4647,
                "Y": 1.7808
              }
            ]
          },
          {
            "CellId": "4,2",
            "CellName": "C5",
            "Value": "(11)",
            "Polygon": [
              {
                "X": 6.9742,
                "Y": 1.6154
              },
              {
                "X": 8.2414,
                "Y": 1.6154
              },
              {
                "X": 8.2414,
                "Y": 1.7808
              },
              {
                "X": 6.9742,
                "Y": 1.7808
              }
            ]
          },
          {
            "CellId": "5,0",
            "CellName": "A6",
            "Value": "Intersegment eliminations",
            "Polygon": [
              {
                "X": 0.2409,
                "Y": 1.7808
              },
              {
                "X": 4.4647,
                "Y": 1.7808
              },
              {
                "X": 4.4647,
                "Y": 1.9297
              },
              {
                "X": 0.2409,
                "Y": 1.9297
              }
            ]
          },
          {
            "CellId": "5,1",
            "CellName": "B6",
            "Value": "102",
            "Polygon": [
              {
                "X": 4.4647,
                "Y": 1.7808
              },
              {
                "X": 6.9742,
                "Y": 1.7808
              },
              {
                "X": 6.9742,
                "Y": 1.9297
              },
              {
                "X": 4.4647,
                "Y": 1.9297
              }
            ]
          },
          {
            "CellId": "5,2",
            "CellName": "C6",
            "Value": "(7)",
            "Polygon": [
              {
                "X": 6.9742,
                "Y": 1.7808
              },
              {
                "X": 8.2414,
                "Y": 1.7808
              },
              {
                "X": 8.2414,
                "Y": 1.9297
              },
              {
                "X": 6.9742,
                "Y": 1.9297
              }
            ]
          },
          {
            "CellId": "6,0",
            "CellName": "A7",
            "Value": "Corporate expenses(2)",
            "Polygon": [
              {
                "X": 0.2409,
                "Y": 1.9297
              },
              {
                "X": 4.4647,
                "Y": 1.9297
              },
              {
                "X": 4.4647,
                "Y": 2.0951
              },
              {
                "X": 0.2409,
                "Y": 2.0951
              }
            ]
          },
          {
            "CellId": "6,1",
            "CellName": "B7",
            "Value": "(29)",
            "Polygon": [
              {
                "X": 4.4647,
                "Y": 1.9297
              },
              {
                "X": 6.9742,
                "Y": 1.9297
              },
              {
                "X": 6.9742,
                "Y": 2.0951
              },
              {
                "X": 4.4647,
                "Y": 2.0951
              }
            ]
          },
          {
            "CellId": "6,2",
            "CellName": "C7",
            "Value": "(26)",
            "Polygon": [
              {
                "X": 6.9742,
                "Y": 1.9297
              },
              {
                "X": 8.2414,
                "Y": 1.9297
              },
              {
                "X": 8.2414,
                "Y": 2.0951
              },
              {
                "X": 6.9742,
                "Y": 2.0951
              }
            ]
          },
          {
            "CellId": "7,0",
            "CellName": "A8",
            "Value": "Provision for depreciation, depletion, and amortization",
            "Polygon": [
              {
                "X": 0.2409,
                "Y": 2.0951
              },
              {
                "X": 4.4647,
                "Y": 2.0951
              },
              {
                "X": 4.4647,
                "Y": 2.2523
              },
              {
                "X": 0.2409,
                "Y": 2.2523
              }
            ]
          },
          {
            "CellId": "7,1",
            "CellName": "B8",
            "Value": "(160)",
            "Polygon": [
              {
                "X": 4.4647,
                "Y": 2.0951
              },
              {
                "X": 6.9742,
                "Y": 2.0951
              },
              {
                "X": 6.9742,
                "Y": 2.2523
              },
              {
                "X": 4.4647,
                "Y": 2.2523
              }
            ]
          },
          {
            "CellId": "7,2",
            "CellName": "C8",
            "Value": "(182)",
            "Polygon": [
              {
                "X": 6.9742,
                "Y": 2.0951
              },
              {
                "X": 8.2414,
                "Y": 2.0951
              },
              {
                "X": 8.2414,
                "Y": 2.2523
              },
              {
                "X": 6.9742,
                "Y": 2.2523
              }
            ]
          },
          {
            "CellId": "8,0",
            "CellName": "A9",
            "Value": "Restructuring and other charges, net (D)",
            "Polygon": [
              {
                "X": 0.2409,
                "Y": 2.2523
              },
              {
                "X": 4.4647,
                "Y": 2.2523
              },
              {
                "X": 4.4647,
                "Y": 2.426
              },
              {
                "X": 0.2409,
                "Y": 2.426
              }
            ]
          },
          {
            "CellId": "8,1",
            "CellName": "B9",
            "Value": "(125)",
            "Polygon": [
              {
                "X": 4.4647,
                "Y": 2.2523
              },
              {
                "X": 6.9742,
                "Y": 2.2523
              },
              {
                "X": 6.9742,
                "Y": 2.426
              },
              {
                "X": 4.4647,
                "Y": 2.426
              }
            ]
          },
          {
            "CellId": "8,2",
            "CellName": "C9",
            "Value": "(7)",
            "Polygon": [
              {
                "X": 6.9742,
                "Y": 2.2523
              },
              {
                "X": 8.2414,
                "Y": 2.2523
              },
              {
                "X": 8.2414,
                "Y": 2.426
              },
              {
                "X": 6.9742,
                "Y": 2.426
              }
            ]
          },
          {
            "CellId": "9,0",
            "CellName": "A10",
            "Value": "Interest expense",
            "Polygon": [
              {
                "X": 0.2409,
                "Y": 2.426
              },
              {
                "X": 4.4647,
                "Y": 2.426
              },
              {
                "X": 4.4647,
                "Y": 2.5749
              },
              {
                "X": 0.2409,
                "Y": 2.5749
              }
            ]
          },
          {
            "CellId": "9,1",
            "CellName": "B10",
            "Value": "(25)",
            "Polygon": [
              {
                "X": 4.4647,
                "Y": 2.426
              },
              {
                "X": 6.9742,
                "Y": 2.426
              },
              {
                "X": 6.9742,
                "Y": 2.5749
              },
              {
                "X": 4.4647,
                "Y": 2.5749
              }
            ]
          },
          {
            "CellId": "9,2",
            "CellName": "C10",
            "Value": "(42)",
            "Polygon": [
              {
                "X": 6.9742,
                "Y": 2.426
              },
              {
                "X": 8.2414,
                "Y": 2.426
              },
              {
                "X": 8.2414,
                "Y": 2.5749
              },
              {
                "X": 6.9742,
                "Y": 2.5749
              }
            ]
          },
          {
            "CellId": "10,0",
            "CellName": "A11",
            "Value": "Other income, net (P)",
            "Polygon": [
              {
                "X": 0.2409,
                "Y": 2.5749
              },
              {
                "X": 4.4647,
                "Y": 2.5749
              },
              {
                "X": 4.4647,
                "Y": 2.732
              },
              {
                "X": 0.2409,
                "Y": 2.732
              }
            ]
          },
          {
            "CellId": "10,1",
            "CellName": "B11",
            "Value": "14",
            "Polygon": [
              {
                "X": 4.4647,
                "Y": 2.5749
              },
              {
                "X": 6.9742,
                "Y": 2.5749
              },
              {
                "X": 6.9742,
                "Y": 2.732
              },
              {
                "X": 4.4647,
                "Y": 2.732
              }
            ]
          },
          {
            "CellId": "10,2",
            "CellName": "C11",
            "Value": "24",
            "Polygon": [
              {
                "X": 6.9742,
                "Y": 2.5749
              },
              {
                "X": 8.2414,
                "Y": 2.5749
              },
              {
                "X": 8.2414,
                "Y": 2.732
              },
              {
                "X": 6.9742,
                "Y": 2.732
              }
            ]
          },
          {
            "CellId": "11,0",
            "CellName": "A12",
            "Value": "Other(3)",
            "Polygon": [
              {
                "X": 0.2409,
                "Y": 2.732
              },
              {
                "X": 4.4647,
                "Y": 2.732
              },
              {
                "X": 4.4647,
                "Y": 2.8892
              },
              {
                "X": 0.2409,
                "Y": 2.8892
              }
            ]
          },
          {
            "CellId": "11,1",
            "CellName": "B12",
            "Value": "(13)",
            "Polygon": [
              {
                "X": 4.4647,
                "Y": 2.732
              },
              {
                "X": 6.9742,
                "Y": 2.732
              },
              {
                "X": 6.9742,
                "Y": 2.8892
              },
              {
                "X": 4.4647,
                "Y": 2.8892
              }
            ]
          },
          {
            "CellId": "11,2",
            "CellName": "C12",
            "Value": "(6)",
            "Polygon": [
              {
                "X": 6.9742,
                "Y": 2.732
              },
              {
                "X": 8.2414,
                "Y": 2.732
              },
              {
                "X": 8.2414,
                "Y": 2.8892
              },
              {
                "X": 6.9742,
                "Y": 2.8892
              }
            ]
          },
          {
            "CellId": "12,0",
            "CellName": "A13",
            "Value": "Consolidated income before income taxes",
            "Polygon": [
              {
                "X": 0.2409,
                "Y": 2.8892
              },
              {
                "X": 4.4647,
                "Y": 2.8892
              },
              {
                "X": 4.4647,
                "Y": 3.0629
              },
              {
                "X": 0.2409,
                "Y": 3.0629
              }
            ]
          },
          {
            "CellId": "12,1",
            "CellName": "B13",
            "Value": "763",
            "Polygon": [
              {
                "X": 4.4647,
                "Y": 2.8892
              },
              {
                "X": 6.9742,
                "Y": 2.8892
              },
              {
                "X": 6.9742,
                "Y": 3.0629
              },
              {
                "X": 4.4647,
                "Y": 3.0629
              }
            ]
          },
          {
            "CellId": "12,2",
            "CellName": "C13",
            "Value": "312",
            "Polygon": [
              {
                "X": 6.9742,
                "Y": 2.8892
              },
              {
                "X": 8.2414,
                "Y": 2.8892
              },
              {
                "X": 8.2414,
                "Y": 3.0629
              },
              {
                "X": 6.9742,
                "Y": 3.0629
              }
            ]
          },
          {
            "CellId": "13,0",
            "CellName": "A14",
            "Value": "Provision for income taxes",
            "Polygon": [
              {
                "X": 0.2409,
                "Y": 3.0629
              },
              {
                "X": 4.4647,
                "Y": 3.0629
              },
              {
                "X": 4.4647,
                "Y": 3.2118
              },
              {
                "X": 0.2409,
                "Y": 3.2118
              }
            ]
          },
          {
            "CellId": "13,1",
            "CellName": "B14",
            "Value": "(210)",
            "Polygon": [
              {
                "X": 4.4647,
                "Y": 3.0629
              },
              {
                "X": 6.9742,
                "Y": 3.0629
              },
              {
                "X": 6.9742,
                "Y": 3.2118
              },
              {
                "X": 4.4647,
                "Y": 3.2118
              }
            ]
          },
          {
            "CellId": "13,2",
            "CellName": "C14",
            "Value": "(93)",
            "Polygon": [
              {
                "X": 6.9742,
                "Y": 3.0629
              },
              {
                "X": 8.2414,
                "Y": 3.0629
              },
              {
                "X": 8.2414,
                "Y": 3.2118
              },
              {
                "X": 6.9742,
                "Y": 3.2118
              }
            ]
          },
          {
            "CellId": "14,0",
            "CellName": "A15",
            "Value": "Net income attributable to noncontrolling interest",
            "Polygon": [
              {
                "X": 0.2409,
                "Y": 3.2118
              },
              {
                "X": 4.4647,
                "Y": 3.2118
              },
              {
                "X": 4.4647,
                "Y": 3.3772
              },
              {
                "X": 0.2326,
                "Y": 3.3772
              }
            ]
          },
          {
            "CellId": "14,1",
            "CellName": "B15",
            "Value": "(84)",
            "Polygon": [
              {
                "X": 4.4647,
                "Y": 3.2118
              },
              {
                "X": 6.9742,
                "Y": 3.2118
              },
              {
                "X": 6.9742,
                "Y": 3.3772
              },
              {
                "X": 4.4647,
                "Y": 3.3772
              }
            ]
          },
          {
            "CellId": "14,2",
            "CellName": "C15",
            "Value": "(44)",
            "Polygon": [
              {
                "X": 6.9742,
                "Y": 3.2118
              },
              {
                "X": 8.2414,
                "Y": 3.2118
              },
              {
                "X": 8.2414,
                "Y": 3.3772
              },
              {
                "X": 6.9742,
                "Y": 3.3772
              }
            ]
          },
          {
            "CellId": "15,0",
            "CellName": "A16",
            "Value": "Consolidated net income attributable to Alcoa Corporation",
            "Polygon": [
              {
                "X": 0.2326,
                "Y": 3.3772
              },
              {
                "X": 4.4647,
                "Y": 3.3772
              },
              {
                "X": 4.4647,
                "Y": 3.5509
              },
              {
                "X": 0.2326,
                "Y": 3.5509
              }
            ]
          },
          {
            "CellId": "15,1",
            "CellName": "B16",
            "Value": "$ 469",
            "Polygon": [
              {
                "X": 4.4647,
                "Y": 3.3772
              },
              {
                "X": 6.9742,
                "Y": 3.3772
              },
              {
                "X": 6.9742,
                "Y": 3.5509
              },
              {
                "X": 4.4647,
                "Y": 3.5509
              }
            ]
          },
          {
            "CellId": "15,2",
            "CellName": "C16",
            "Value": "$ 175",
            "Polygon": [
              {
                "X": 6.9742,
                "Y": 3.3772
              },
              {
                "X": 8.2414,
                "Y": 3.3772
              },
              {
                "X": 8.2414,
                "Y": 3.5509
              },
              {
                "X": 6.9742,
                "Y": 3.5509
              }
            ]
          }
        ]
      }
    },
    {
      "1.1": {
        "Cells": [
          {
            "CellId": "0,0",
            "CellName": "A1",
            "Value": "",
            "Polygon": [
              {
                "X": 0.1922,
                "Y": 4.8417
              },
              {
                "X": 3.1042,
                "Y": 4.8417
              },
              {
                "X": 3.1042,
                "Y": 5.3304
              },
              {
                "X": 0.1922,
                "Y": 5.3138
              }
            ]
          },
          {
            "CellId": "0,1",
            "CellName": "B1",
            "Value": "First quarter ended March 31,",
            "Polygon": [
              {
                "X": 3.1042,
                "Y": 4.8417
              },
              {
                "X": 8.2499,
                "Y": 4.85
              },
              {
                "X": 8.2499,
                "Y": 5.1647
              },
              {
                "X": 3.1042,
                "Y": 5.1647
              }
            ]
          },
          {
            "CellId": "1,1",
            "CellName": "B2",
            "Value": "2022",
            "Polygon": [
              {
                "X": 3.1042,
                "Y": 5.1647
              },
              {
                "X": 6.4795,
                "Y": 5.1647
              },
              {
                "X": 6.4795,
                "Y": 5.3304
              },
              {
                "X": 3.1042,
                "Y": 5.3304
              }
            ]
          },
          {
            "CellId": "1,3",
            "CellName": "D2",
            "Value": "2021",
            "Polygon": [
              {
                "X": 6.4795,
                "Y": 5.1647
              },
              {
                "X": 8.2499,
                "Y": 5.1647
              },
              {
                "X": 8.2499,
                "Y": 5.3304
              },
              {
                "X": 6.4795,
                "Y": 5.3304
              }
            ]
          },
          {
            "CellId": "2,0",
            "CellName": "A3",
            "Value": "Primary aluminum",
            "Polygon": [
              {
                "X": 0.1922,
                "Y": 5.3138
              },
              {
                "X": 3.1042,
                "Y": 5.3304
              },
              {
                "X": 3.1042,
                "Y": 5.4878
              },
              {
                "X": 0.1922,
                "Y": 5.4878
              }
            ]
          },
          {
            "CellId": "2,1",
            "CellName": "B3",
            "Value": "$",
            "Polygon": [
              {
                "X": 3.1042,
                "Y": 5.3304
              },
              {
                "X": 5.8177,
                "Y": 5.3304
              },
              {
                "X": 5.8177,
                "Y": 5.4878
              },
              {
                "X": 3.1042,
                "Y": 5.4878
              }
            ]
          },
          {
            "CellId": "2,2",
            "CellName": "C3",
            "Value": "2,447",
            "Polygon": [
              {
                "X": 5.8177,
                "Y": 5.3304
              },
              {
                "X": 6.4795,
                "Y": 5.3304
              },
              {
                "X": 6.4795,
                "Y": 5.4878
              },
              {
                "X": 5.8177,
                "Y": 5.4878
              }
            ]
          },
          {
            "CellId": "2,3",
            "CellName": "D3",
            "Value": "$ 1,727",
            "Polygon": [
              {
                "X": 6.4795,
                "Y": 5.3304
              },
              {
                "X": 8.2499,
                "Y": 5.3304
              },
              {
                "X": 8.2499,
                "Y": 5.4961
              },
              {
                "X": 6.4795,
                "Y": 5.4878
              }
            ]
          },
          {
            "CellId": "3,0",
            "CellName": "A4",
            "Value": "Alumina",
            "Polygon": [
              {
                "X": 0.1922,
                "Y": 5.4878
              },
              {
                "X": 3.1042,
                "Y": 5.4878
              },
              {
                "X": 3.1042,
                "Y": 5.6452
              },
              {
                "X": 0.1922,
                "Y": 5.6452
              }
            ]
          },
          {
            "CellId": "3,1",
            "CellName": "B4",
            "Value": "",
            "Polygon": [
              {
                "X": 3.1042,
                "Y": 5.4878
              },
              {
                "X": 5.8177,
                "Y": 5.4878
              },
              {
                "X": 5.8177,
                "Y": 5.6452
              },
              {
                "X": 3.1042,
                "Y": 5.6452
              }
            ]
          },
          {
            "CellId": "3,2",
            "CellName": "C4",
            "Value": "850",
            "Polygon": [
              {
                "X": 5.8177,
                "Y": 5.4878
              },
              {
                "X": 6.4795,
                "Y": 5.4878
              },
              {
                "X": 6.4795,
                "Y": 5.6452
              },
              {
                "X": 5.8177,
                "Y": 5.6452
              }
            ]
          },
          {
            "CellId": "3,3",
            "CellName": "D4",
            "Value": "760",
            "Polygon": [
              {
                "X": 6.4795,
                "Y": 5.4878
              },
              {
                "X": 8.2499,
                "Y": 5.4961
              },
              {
                "X": 8.2499,
                "Y": 5.6452
              },
              {
                "X": 6.4795,
                "Y": 5.6452
              }
            ]
          },
          {
            "CellId": "4,0",
            "CellName": "A5",
            "Value": "Energy",
            "Polygon": [
              {
                "X": 0.1922,
                "Y": 5.6452
              },
              {
                "X": 3.1042,
                "Y": 5.6452
              },
              {
                "X": 3.1042,
                "Y": 5.8026
              },
              {
                "X": 0.1922,
                "Y": 5.8026
              }
            ]
          },
          {
            "CellId": "4,1",
            "CellName": "B5",
            "Value": "",
            "Polygon": [
              {
                "X": 3.1042,
                "Y": 5.6452
              },
              {
                "X": 5.8177,
                "Y": 5.6452
              },
              {
                "X": 5.8177,
                "Y": 5.8026
              },
              {
                "X": 3.1042,
                "Y": 5.8026
              }
            ]
          },
          {
            "CellId": "4,2",
            "CellName": "C5",
            "Value": "41",
            "Polygon": [
              {
                "X": 5.8177,
                "Y": 5.6452
              },
              {
                "X": 6.4795,
                "Y": 5.6452
              },
              {
                "X": 6.4795,
                "Y": 5.8026
              },
              {
                "X": 5.8177,
                "Y": 5.8026
              }
            ]
          },
          {
            "CellId": "4,3",
            "CellName": "D5",
            "Value": "39",
            "Polygon": [
              {
                "X": 6.4795,
                "Y": 5.6452
              },
              {
                "X": 8.2499,
                "Y": 5.6452
              },
              {
                "X": 8.2499,
                "Y": 5.8109
              },
              {
                "X": 6.4795,
                "Y": 5.8026
              }
            ]
          },
          {
            "CellId": "5,0",
            "CellName": "A6",
            "Value": "Bauxite",
            "Polygon": [
              {
                "X": 0.1922,
                "Y": 5.8026
              },
              {
                "X": 3.1042,
                "Y": 5.8026
              },
              {
                "X": 3.1042,
                "Y": 5.96
              },
              {
                "X": 0.1922,
                "Y": 5.96
              }
            ]
          },
          {
            "CellId": "5,1",
            "CellName": "B6",
            "Value": "",
            "Polygon": [
              {
                "X": 3.1042,
                "Y": 5.8026
              },
              {
                "X": 5.8177,
                "Y": 5.8026
              },
              {
                "X": 5.8177,
                "Y": 5.9683
              },
              {
                "X": 3.1042,
                "Y": 5.96
              }
            ]
          },
          {
            "CellId": "5,2",
            "CellName": "C6",
            "Value": "28",
            "Polygon": [
              {
                "X": 5.8177,
                "Y": 5.8026
              },
              {
                "X": 6.4795,
                "Y": 5.8026
              },
              {
                "X": 6.4795,
                "Y": 5.9683
              },
              {
                "X": 5.8177,
                "Y": 5.9683
              }
            ]
          },
          {
            "CellId": "5,3",
            "CellName": "D6",
            "Value": "52",
            "Polygon": [
              {
                "X": 6.4795,
                "Y": 5.8026
              },
              {
                "X": 8.2499,
                "Y": 5.8109
              },
              {
                "X": 8.2499,
                "Y": 5.9683
              },
              {
                "X": 6.4795,
                "Y": 5.9683
              }
            ]
          },
          {
            "CellId": "6,0",
            "CellName": "A7",
            "Value": "Flat-rolled aluminum(1)",
            "Polygon": [
              {
                "X": 0.1922,
                "Y": 5.96
              },
              {
                "X": 3.1042,
                "Y": 5.96
              },
              {
                "X": 3.1042,
                "Y": 6.1256
              },
              {
                "X": 0.1922,
                "Y": 6.1256
              }
            ]
          },
          {
            "CellId": "6,1",
            "CellName": "B7",
            "Value": "",
            "Polygon": [
              {
                "X": 3.1042,
                "Y": 5.96
              },
              {
                "X": 5.8177,
                "Y": 5.9683
              },
              {
                "X": 5.8177,
                "Y": 6.1256
              },
              {
                "X": 3.1042,
                "Y": 6.1256
              }
            ]
          },
          {
            "CellId": "6,2",
            "CellName": "C7",
            "Value": "—",
            "Polygon": [
              {
                "X": 5.8177,
                "Y": 5.9683
              },
              {
                "X": 6.4795,
                "Y": 5.9683
              },
              {
                "X": 6.4795,
                "Y": 6.1256
              },
              {
                "X": 5.8177,
                "Y": 6.1256
              }
            ]
          },
          {
            "CellId": "6,3",
            "CellName": "D7",
            "Value": "320",
            "Polygon": [
              {
                "X": 6.4795,
                "Y": 5.9683
              },
              {
                "X": 8.2499,
                "Y": 5.9683
              },
              {
                "X": 8.2499,
                "Y": 6.1339
              },
              {
                "X": 6.4795,
                "Y": 6.1256
              }
            ]
          },
          {
            "CellId": "7,0",
            "CellName": "A8",
            "Value": "Other(2)",
            "Polygon": [
              {
                "X": 0.1922,
                "Y": 6.1256
              },
              {
                "X": 3.1042,
                "Y": 6.1256
              },
              {
                "X": 3.1042,
                "Y": 6.283
              },
              {
                "X": 0.1922,
                "Y": 6.283
              }
            ]
          },
          {
            "CellId": "7,1",
            "CellName": "B8",
            "Value": "",
            "Polygon": [
              {
                "X": 3.1042,
                "Y": 6.1256
              },
              {
                "X": 5.8177,
                "Y": 6.1256
              },
              {
                "X": 5.8259,
                "Y": 6.2913
              },
              {
                "X": 3.1042,
                "Y": 6.283
              }
            ]
          },
          {
            "CellId": "7,2",
            "CellName": "C8",
            "Value": "(73)",
            "Polygon": [
              {
                "X": 5.8177,
                "Y": 6.1256
              },
              {
                "X": 6.4795,
                "Y": 6.1256
              },
              {
                "X": 6.4795,
                "Y": 6.2913
              },
              {
                "X": 5.8259,
                "Y": 6.2913
              }
            ]
          },
          {
            "CellId": "7,3",
            "CellName": "D8",
            "Value": "(28)",
            "Polygon": [
              {
                "X": 6.4795,
                "Y": 6.1256
              },
              {
                "X": 8.2499,
                "Y": 6.1339
              },
              {
                "X": 8.2499,
                "Y": 6.2913
              },
              {
                "X": 6.4795,
                "Y": 6.2913
              }
            ]
          },
          {
            "CellId": "8,0",
            "CellName": "A9",
            "Value": "",
            "Polygon": [
              {
                "X": 0.1922,
                "Y": 6.283
              },
              {
                "X": 3.1042,
                "Y": 6.283
              },
              {
                "X": 3.1042,
                "Y": 6.4487
              },
              {
                "X": 0.1922,
                "Y": 6.4404
              }
            ]
          },
          {
            "CellId": "8,1",
            "CellName": "B9",
            "Value": "$",
            "Polygon": [
              {
                "X": 3.1042,
                "Y": 6.283
              },
              {
                "X": 5.8259,
                "Y": 6.2913
              },
              {
                "X": 5.8259,
                "Y": 6.4487
              },
              {
                "X": 3.1042,
                "Y": 6.4487
              }
            ]
          },
          {
            "CellId": "8,2",
            "CellName": "C9",
            "Value": "3,293",
            "Polygon": [
              {
                "X": 5.8259,
                "Y": 6.2913
              },
              {
                "X": 6.4795,
                "Y": 6.2913
              },
              {
                "X": 6.4795,
                "Y": 6.4487
              },
              {
                "X": 5.8259,
                "Y": 6.4487
              }
            ]
          },
          {
            "CellId": "8,3",
            "CellName": "D9",
            "Value": "$ 2,870",
            "Polygon": [
              {
                "X": 6.4795,
                "Y": 6.2913
              },
              {
                "X": 8.2499,
                "Y": 6.2913
              },
              {
                "X": 8.2499,
                "Y": 6.457
              },
              {
                "X": 6.4795,
                "Y": 6.4487
              }
            ]
          }
        ]
      }
    },
    {
      "2": {
        "Cells": [
          {
            "CellId": "0,0",
            "CellName": "A1",
            "Value": "",
            "Polygon": [
              {
                "X": 0.2335,
                "Y": 0.9841
              },
              {
                "X": 4.1466,
                "Y": 0.9923
              },
              {
                "X": 4.1383,
                "Y": 1.4728
              },
              {
                "X": 0.2335,
                "Y": 1.4645
              }
            ]
          },
          {
            "CellId": "0,1",
            "CellName": "B1",
            "Value": "First quarter ended March 31,",
            "Polygon": [
              {
                "X": 4.1466,
                "Y": 0.9923
              },
              {
                "X": 8.2499,
                "Y": 0.9923
              },
              {
                "X": 8.2499,
                "Y": 1.3071
              },
              {
                "X": 4.1383,
                "Y": 1.3071
              }
            ]
          },
          {
            "CellId": "1,1",
            "CellName": "B2",
            "Value": "2022",
            "Polygon": [
              {
                "X": 4.1383,
                "Y": 1.3071
              },
              {
                "X": 6.9345,
                "Y": 1.3071
              },
              {
                "X": 6.9345,
                "Y": 1.4728
              },
              {
                "X": 4.1383,
                "Y": 1.4728
              }
            ]
          },
          {
            "CellId": "1,2",
            "CellName": "C2",
            "Value": "2021",
            "Polygon": [
              {
                "X": 6.9345,
                "Y": 1.3071
              },
              {
                "X": 8.2499,
                "Y": 1.3071
              },
              {
                "X": 8.2499,
                "Y": 1.4728
              },
              {
                "X": 6.9345,
                "Y": 1.4728
              }
            ]
          },
          {
            "CellId": "2,0",
            "CellName": "A3",
            "Value": "Net income attributable to Alcoa Corporation",
            "Polygon": [
              {
                "X": 0.2335,
                "Y": 1.4645
              },
              {
                "X": 4.1383,
                "Y": 1.4728
              },
              {
                "X": 4.1383,
                "Y": 1.6302
              },
              {
                "X": 0.2335,
                "Y": 1.6302
              }
            ]
          },
          {
            "CellId": "2,1",
            "CellName": "B3",
            "Value": "$ 469",
            "Polygon": [
              {
                "X": 4.1383,
                "Y": 1.4728
              },
              {
                "X": 6.9345,
                "Y": 1.4728
              },
              {
                "X": 6.9345,
                "Y": 1.6302
              },
              {
                "X": 4.1383,
                "Y": 1.6302
              }
            ]
          },
          {
            "CellId": "2,2",
            "CellName": "C3",
            "Value": "$ 175",
            "Polygon": [
              {
                "X": 6.9345,
                "Y": 1.4728
              },
              {
                "X": 8.2499,
                "Y": 1.4728
              },
              {
                "X": 8.2499,
                "Y": 1.6385
              },
              {
                "X": 6.9345,
                "Y": 1.6302
              }
            ]
          },
          {
            "CellId": "3,0",
            "CellName": "A4",
            "Value": "Average shares outstanding - basic",
            "Polygon": [
              {
                "X": 0.2335,
                "Y": 1.6302
              },
              {
                "X": 4.1383,
                "Y": 1.6302
              },
              {
                "X": 4.1383,
                "Y": 1.7959
              },
              {
                "X": 0.2335,
                "Y": 1.7959
              }
            ]
          },
          {
            "CellId": "3,1",
            "CellName": "B4",
            "Value": "184",
            "Polygon": [
              {
                "X": 4.1383,
                "Y": 1.6302
              },
              {
                "X": 6.9345,
                "Y": 1.6302
              },
              {
                "X": 6.9345,
                "Y": 1.7959
              },
              {
                "X": 4.1383,
                "Y": 1.7959
              }
            ]
          },
          {
            "CellId": "3,2",
            "CellName": "C4",
            "Value": "186",
            "Polygon": [
              {
                "X": 6.9345,
                "Y": 1.6302
              },
              {
                "X": 8.2499,
                "Y": 1.6385
              },
              {
                "X": 8.2499,
                "Y": 1.7959
              },
              {
                "X": 6.9345,
                "Y": 1.7959
              }
            ]
          },
          {
            "CellId": "4,0",
            "CellName": "A5",
            "Value": "Effect of dilutive securities:",
            "Polygon": [
              {
                "X": 0.2335,
                "Y": 1.7959
              },
              {
                "X": 4.1383,
                "Y": 1.7959
              },
              {
                "X": 4.1383,
                "Y": 1.9533
              },
              {
                "X": 0.2252,
                "Y": 1.9533
              }
            ]
          },
          {
            "CellId": "4,1",
            "CellName": "B5",
            "Value": "",
            "Polygon": [
              {
                "X": 4.1383,
                "Y": 1.7959
              },
              {
                "X": 6.9345,
                "Y": 1.7959
              },
              {
                "X": 6.9345,
                "Y": 1.9533
              },
              {
                "X": 4.1383,
                "Y": 1.9533
              }
            ]
          },
          {
            "CellId": "4,2",
            "CellName": "C5",
            "Value": "",
            "Polygon": [
              {
                "X": 6.9345,
                "Y": 1.7959
              },
              {
                "X": 8.2499,
                "Y": 1.7959
              },
              {
                "X": 8.2499,
                "Y": 1.9533
              },
              {
                "X": 6.9345,
                "Y": 1.9533
              }
            ]
          },
          {
            "CellId": "5,0",
            "CellName": "A6",
            "Value": "Stock options",
            "Polygon": [
              {
                "X": 0.2252,
                "Y": 1.9533
              },
              {
                "X": 4.1383,
                "Y": 1.9533
              },
              {
                "X": 4.1383,
                "Y": 2.119
              },
              {
                "X": 0.2252,
                "Y": 2.119
              }
            ]
          },
          {
            "CellId": "5,1",
            "CellName": "B6",
            "Value": "—",
            "Polygon": [
              {
                "X": 4.1383,
                "Y": 1.9533
              },
              {
                "X": 6.9345,
                "Y": 1.9533
              },
              {
                "X": 6.9345,
                "Y": 2.119
              },
              {
                "X": 4.1383,
                "Y": 2.119
              }
            ]
          },
          {
            "CellId": "5,2",
            "CellName": "C6",
            "Value": "—",
            "Polygon": [
              {
                "X": 6.9345,
                "Y": 1.9533
              },
              {
                "X": 8.2499,
                "Y": 1.9533
              },
              {
                "X": 8.2499,
                "Y": 2.119
              },
              {
                "X": 6.9345,
                "Y": 2.119
              }
            ]
          },
          {
            "CellId": "6,0",
            "CellName": "A7",
            "Value": "Stock units",
            "Polygon": [
              {
                "X": 0.2252,
                "Y": 2.119
              },
              {
                "X": 4.1383,
                "Y": 2.119
              },
              {
                "X": 4.1383,
                "Y": 2.2847
              },
              {
                "X": 0.2252,
                "Y": 2.2847
              }
            ]
          },
          {
            "CellId": "6,1",
            "CellName": "B7",
            "Value": "4",
            "Polygon": [
              {
                "X": 4.1383,
                "Y": 2.119
              },
              {
                "X": 6.9345,
                "Y": 2.119
              },
              {
                "X": 6.9345,
                "Y": 2.2847
              },
              {
                "X": 4.1383,
                "Y": 2.2847
              }
            ]
          },
          {
            "CellId": "6,2",
            "CellName": "C7",
            "Value": "3",
            "Polygon": [
              {
                "X": 6.9345,
                "Y": 2.119
              },
              {
                "X": 8.2499,
                "Y": 2.119
              },
              {
                "X": 8.2499,
                "Y": 2.2847
              },
              {
                "X": 6.9345,
                "Y": 2.2847
              }
            ]
          },
          {
            "CellId": "7,0",
            "CellName": "A8",
            "Value": "Average shares outstanding - diluted",
            "Polygon": [
              {
                "X": 0.2252,
                "Y": 2.2847
              },
              {
                "X": 4.1383,
                "Y": 2.2847
              },
              {
                "X": 4.1383,
                "Y": 2.4586
              },
              {
                "X": 0.2252,
                "Y": 2.4586
              }
            ]
          },
          {
            "CellId": "7,1",
            "CellName": "B8",
            "Value": "188",
            "Polygon": [
              {
                "X": 4.1383,
                "Y": 2.2847
              },
              {
                "X": 6.9345,
                "Y": 2.2847
              },
              {
                "X": 6.9345,
                "Y": 2.4586
              },
              {
                "X": 4.1383,
                "Y": 2.4586
              }
            ]
          },
          {
            "CellId": "7,2",
            "CellName": "C8",
            "Value": "189",
            "Polygon": [
              {
                "X": 6.9345,
                "Y": 2.2847
              },
              {
                "X": 8.2499,
                "Y": 2.2847
              },
              {
                "X": 8.2499,
                "Y": 2.4669
              },
              {
                "X": 6.9345,
                "Y": 2.4586
              }
            ]
          }
        ]
      }
    },
    {
      "3": {
        "Cells": [
          {
            "CellId": "0,0",
            "CellName": "A1",
            "Value": "",
            "Polygon": [
              {
                "X": 0.2532,
                "Y": 1.1417
              },
              {
                "X": 3.6214,
                "Y": 1.1334
              },
              {
                "X": 3.6214,
                "Y": 1.7806
              },
              {
                "X": 0.2532,
                "Y": 1.7723
              }
            ]
          },
          {
            "CellId": "0,1",
            "CellName": "B1",
            "Value": "Alcoa Corporation",
            "Polygon": [
              {
                "X": 3.6214,
                "Y": 1.1334
              },
              {
                "X": 6.2927,
                "Y": 1.1334
              },
              {
                "X": 6.2927,
                "Y": 1.3325
              },
              {
                "X": 3.6214,
                "Y": 1.3325
              }
            ]
          },
          {
            "CellId": "0,3",
            "CellName": "D1",
            "Value": "Noncontrolling interest",
            "Polygon": [
              {
                "X": 6.2927,
                "Y": 1.1334
              },
              {
                "X": 8.2506,
                "Y": 1.1334
              },
              {
                "X": 8.2506,
                "Y": 1.3325
              },
              {
                "X": 6.2927,
                "Y": 1.3325
              }
            ]
          },
          {
            "CellId": "1,1",
            "CellName": "B2",
            "Value": "First quarter ended March 31,",
            "Polygon": [
              {
                "X": 3.6214,
                "Y": 1.3325
              },
              {
                "X": 6.2927,
                "Y": 1.3325
              },
              {
                "X": 6.2927,
                "Y": 1.6229
              },
              {
                "X": 3.6214,
                "Y": 1.6229
              }
            ]
          },
          {
            "CellId": "1,3",
            "CellName": "D2",
            "Value": "First quarter ended March 31,",
            "Polygon": [
              {
                "X": 6.2927,
                "Y": 1.3325
              },
              {
                "X": 8.2506,
                "Y": 1.3325
              },
              {
                "X": 8.2506,
                "Y": 1.6229
              },
              {
                "X": 6.2927,
                "Y": 1.6229
              }
            ]
          },
          {
            "CellId": "2,1",
            "CellName": "B3",
            "Value": "2022",
            "Polygon": [
              {
                "X": 3.6214,
                "Y": 1.6229
              },
              {
                "X": 5.2889,
                "Y": 1.6229
              },
              {
                "X": 5.2889,
                "Y": 1.7806
              },
              {
                "X": 3.6214,
                "Y": 1.7806
              }
            ]
          },
          {
            "CellId": "2,2",
            "CellName": "C3",
            "Value": "2021",
            "Polygon": [
              {
                "X": 5.2889,
                "Y": 1.6229
              },
              {
                "X": 6.2927,
                "Y": 1.6229
              },
              {
                "X": 6.2927,
                "Y": 1.7723
              },
              {
                "X": 5.2889,
                "Y": 1.7806
              }
            ]
          },
          {
            "CellId": "2,3",
            "CellName": "D3",
            "Value": "2022",
            "Polygon": [
              {
                "X": 6.2927,
                "Y": 1.6229
              },
              {
                "X": 7.28,
                "Y": 1.6229
              },
              {
                "X": 7.28,
                "Y": 1.7723
              },
              {
                "X": 6.2927,
                "Y": 1.7723
              }
            ]
          },
          {
            "CellId": "2,4",
            "CellName": "E3",
            "Value": "2021",
            "Polygon": [
              {
                "X": 7.28,
                "Y": 1.6229
              },
              {
                "X": 8.2506,
                "Y": 1.6229
              },
              {
                "X": 8.2506,
                "Y": 1.7723
              },
              {
                "X": 7.28,
                "Y": 1.7723
              }
            ]
          },
          {
            "CellId": "3,0",
            "CellName": "A4",
            "Value": "Pension and other postretirement benefits (K)",
            "Polygon": [
              {
                "X": 0.2532,
                "Y": 1.7723
              },
              {
                "X": 3.6214,
                "Y": 1.7806
              },
              {
                "X": 3.6214,
                "Y": 1.9465
              },
              {
                "X": 0.2532,
                "Y": 1.9382
              }
            ]
          },
          {
            "CellId": "3,1",
            "CellName": "B4",
            "Value": "",
            "Polygon": [
              {
                "X": 3.6214,
                "Y": 1.7806
              },
              {
                "X": 5.2889,
                "Y": 1.7806
              },
              {
                "X": 5.2889,
                "Y": 1.9465
              },
              {
                "X": 3.6214,
                "Y": 1.9465
              }
            ]
          },
          {
            "CellId": "3,2",
            "CellName": "C4",
            "Value": "",
            "Polygon": [
              {
                "X": 5.2889,
                "Y": 1.7806
              },
              {
                "X": 6.2927,
                "Y": 1.7723
              },
              {
                "X": 6.2927,
                "Y": 1.9465
              },
              {
                "X": 5.2889,
                "Y": 1.9465
              }
            ]
          },
          {
            "CellId": "3,3",
            "CellName": "D4",
            "Value": "",
            "Polygon": [
              {
                "X": 6.2927,
                "Y": 1.7723
              },
              {
                "X": 7.28,
                "Y": 1.7723
              },
              {
                "X": 7.28,
                "Y": 1.9382
              },
              {
                "X": 6.2927,
                "Y": 1.9465
              }
            ]
          },
          {
            "CellId": "3,4",
            "CellName": "E4",
            "Value": "",
            "Polygon": [
              {
                "X": 7.28,
                "Y": 1.7723
              },
              {
                "X": 8.2506,
                "Y": 1.7723
              },
              {
                "X": 8.2506,
                "Y": 1.9465
              },
              {
                "X": 7.28,
                "Y": 1.9382
              }
            ]
          },
          {
            "CellId": "4,0",
            "CellName": "A5",
            "Value": "Balance at beginning of period",
            "Polygon": [
              {
                "X": 0.2532,
                "Y": 1.9382
              },
              {
                "X": 3.6214,
                "Y": 1.9465
              },
              {
                "X": 3.6214,
                "Y": 2.0959
              },
              {
                "X": 0.2532,
                "Y": 2.0959
              }
            ]
          },
          {
            "CellId": "4,1",
            "CellName": "B5",
            "Value": "$ (882)",
            "Polygon": [
              {
                "X": 3.6214,
                "Y": 1.9465
              },
              {
                "X": 5.2889,
                "Y": 1.9465
              },
              {
                "X": 5.2889,
                "Y": 2.0959
              },
              {
                "X": 3.6214,
                "Y": 2.0959
              }
            ]
          },
          {
            "CellId": "4,2",
            "CellName": "C5",
            "Value": "$ (2,536)",
            "Polygon": [
              {
                "X": 5.2889,
                "Y": 1.9465
              },
              {
                "X": 6.2927,
                "Y": 1.9465
              },
              {
                "X": 6.2927,
                "Y": 2.0959
              },
              {
                "X": 5.2889,
                "Y": 2.0959
              }
            ]
          },
          {
            "CellId": "4,3",
            "CellName": "D5",
            "Value": "$ (13)",
            "Polygon": [
              {
                "X": 6.2927,
                "Y": 1.9465
              },
              {
                "X": 7.28,
                "Y": 1.9382
              },
              {
                "X": 7.28,
                "Y": 2.0959
              },
              {
                "X": 6.2927,
                "Y": 2.0959
              }
            ]
          },
          {
            "CellId": "4,4",
            "CellName": "E5",
            "Value": "$ (67)",
            "Polygon": [
              {
                "X": 7.28,
                "Y": 1.9382
              },
              {
                "X": 8.2506,
                "Y": 1.9465
              },
              {
                "X": 8.2506,
                "Y": 2.0959
              },
              {
                "X": 7.28,
                "Y": 2.0959
              }
            ]
          },
          {
            "CellId": "5,0",
            "CellName": "A6",
            "Value": "Other comprehensive (loss) income:",
            "Polygon": [
              {
                "X": 0.2532,
                "Y": 2.0959
              },
              {
                "X": 3.6214,
                "Y": 2.0959
              },
              {
                "X": 3.6214,
                "Y": 2.2701
              },
              {
                "X": 0.2532,
                "Y": 2.2618
              }
            ]
          },
          {
            "CellId": "5,1",
            "CellName": "B6",
            "Value": "",
            "Polygon": [
              {
                "X": 3.6214,
                "Y": 2.0959
              },
              {
                "X": 5.2889,
                "Y": 2.0959
              },
              {
                "X": 5.2889,
                "Y": 2.2701
              },
              {
                "X": 3.6214,
                "Y": 2.2701
              }
            ]
          },
          {
            "CellId": "5,2",
            "CellName": "C6",
            "Value": "",
            "Polygon": [
              {
                "X": 5.2889,
                "Y": 2.0959
              },
              {
                "X": 6.2927,
                "Y": 2.0959
              },
              {
                "X": 6.2927,
                "Y": 2.2701
              },
              {
                "X": 5.2889,
                "Y": 2.2701
              }
            ]
          },
          {
            "CellId": "5,3",
            "CellName": "D6",
            "Value": "",
            "Polygon": [
              {
                "X": 6.2927,
                "Y": 2.0959
              },
              {
                "X": 7.28,
                "Y": 2.0959
              },
              {
                "X": 7.28,
                "Y": 2.2701
              },
              {
                "X": 6.2927,
                "Y": 2.2701
              }
            ]
          },
          {
            "CellId": "5,4",
            "CellName": "E6",
            "Value": "",
            "Polygon": [
              {
                "X": 7.28,
                "Y": 2.0959
              },
              {
                "X": 8.2506,
                "Y": 2.0959
              },
              {
                "X": 8.2506,
                "Y": 2.2701
              },
              {
                "X": 7.28,
                "Y": 2.2701
              }
            ]
          },
          {
            "CellId": "6,0",
            "CellName": "A7",
            "Value": "Unrecognized net actuarial loss and prior service cost/benefit",
            "Polygon": [
              {
                "X": 0.2532,
                "Y": 2.2618
              },
              {
                "X": 3.6214,
                "Y": 2.2701
              },
              {
                "X": 3.6214,
                "Y": 2.5688
              },
              {
                "X": 0.2532,
                "Y": 2.5688
              }
            ]
          },
          {
            "CellId": "6,1",
            "CellName": "B7",
            "Value": "(7)",
            "Polygon": [
              {
                "X": 3.6214,
                "Y": 2.2701
              },
              {
                "X": 5.2889,
                "Y": 2.2701
              },
              {
                "X": 5.2889,
                "Y": 2.5688
              },
              {
                "X": 3.6214,
                "Y": 2.5688
              }
            ]
          },
          {
            "CellId": "6,2",
            "CellName": "C7",
            "Value": "69",
            "Polygon": [
              {
                "X": 5.2889,
                "Y": 2.2701
              },
              {
                "X": 6.2927,
                "Y": 2.2701
              },
              {
                "X": 6.2927,
                "Y": 2.5688
              },
              {
                "X": 5.2889,
                "Y": 2.5688
              }
            ]
          },
          {
            "CellId": "6,3",
            "CellName": "D7",
            "Value": "—",
            "Polygon": [
              {
                "X": 6.2927,
                "Y": 2.2701
              },
              {
                "X": 7.28,
                "Y": 2.2701
              },
              {
                "X": 7.28,
                "Y": 2.5688
              },
              {
                "X": 6.2927,
                "Y": 2.5688
              }
            ]
          },
          {
            "CellId": "6,4",
            "CellName": "E7",
            "Value": "—",
            "Polygon": [
              {
                "X": 7.28,
                "Y": 2.2701
              },
              {
                "X": 8.2506,
                "Y": 2.2701
              },
              {
                "X": 8.2506,
                "Y": 2.5688
              },
              {
                "X": 7.28,
                "Y": 2.5688
              }
            ]
          },
          {
            "CellId": "7,0",
            "CellName": "A8",
            "Value": "Tax benefit(2)",
            "Polygon": [
              {
                "X": 0.2532,
                "Y": 2.5688
              },
              {
                "X": 3.6214,
                "Y": 2.5688
              },
              {
                "X": 3.6214,
                "Y": 2.7347
              },
              {
                "X": 0.2532,
                "Y": 2.7347
              }
            ]
          },
          {
            "CellId": "7,1",
            "CellName": "B8",
            "Value": "1",
            "Polygon": [
              {
                "X": 3.6214,
                "Y": 2.5688
              },
              {
                "X": 5.2889,
                "Y": 2.5688
              },
              {
                "X": 5.2889,
                "Y": 2.7347
              },
              {
                "X": 3.6214,
                "Y": 2.7347
              }
            ]
          },
          {
            "CellId": "7,2",
            "CellName": "C8",
            "Value": "2",
            "Polygon": [
              {
                "X": 5.2889,
                "Y": 2.5688
              },
              {
                "X": 6.2927,
                "Y": 2.5688
              },
              {
                "X": 6.2927,
                "Y": 2.7347
              },
              {
                "X": 5.2889,
                "Y": 2.7347
              }
            ]
          },
          {
            "CellId": "7,3",
            "CellName": "D8",
            "Value": "—",
            "Polygon": [
              {
                "X": 6.2927,
                "Y": 2.5688
              },
              {
                "X": 7.28,
                "Y": 2.5688
              },
              {
                "X": 7.28,
                "Y": 2.7347
              },
              {
                "X": 6.2927,
                "Y": 2.7347
              }
            ]
          },
          {
            "CellId": "7,4",
            "CellName": "E8",
            "Value": "—",
            "Polygon": [
              {
                "X": 7.28,
                "Y": 2.5688
              },
              {
                "X": 8.2506,
                "Y": 2.5688
              },
              {
                "X": 8.2506,
                "Y": 2.7347
              },
              {
                "X": 7.28,
                "Y": 2.7347
              }
            ]
          },
          {
            "CellId": "8,0",
            "CellName": "A9",
            "Value": "Total Other comprehensive (loss) income before reclassifications, net of tax",
            "Polygon": [
              {
                "X": 0.2532,
                "Y": 2.7347
              },
              {
                "X": 3.6214,
                "Y": 2.7347
              },
              {
                "X": 3.6131,
                "Y": 3.05
              },
              {
                "X": 0.2532,
                "Y": 3.05
              }
            ]
          },
          {
            "CellId": "8,1",
            "CellName": "B9",
            "Value": "(6)",
            "Polygon": [
              {
                "X": 3.6214,
                "Y": 2.7347
              },
              {
                "X": 5.2889,
                "Y": 2.7347
              },
              {
                "X": 5.2889,
                "Y": 3.05
              },
              {
                "X": 3.6131,
                "Y": 3.05
              }
            ]
          },
          {
            "CellId": "8,2",
            "CellName": "C9",
            "Value": "71",
            "Polygon": [
              {
                "X": 5.2889,
                "Y": 2.7347
              },
              {
                "X": 6.2927,
                "Y": 2.7347
              },
              {
                "X": 6.2927,
                "Y": 3.05
              },
              {
                "X": 5.2889,
                "Y": 3.05
              }
            ]
          },
          {
            "CellId": "8,3",
            "CellName": "D9",
            "Value": "—",
            "Polygon": [
              {
                "X": 6.2927,
                "Y": 2.7347
              },
              {
                "X": 7.28,
                "Y": 2.7347
              },
              {
                "X": 7.28,
                "Y": 3.0583
              },
              {
                "X": 6.2927,
                "Y": 3.05
              }
            ]
          },
          {
            "CellId": "8,4",
            "CellName": "E9",
            "Value": "—",
            "Polygon": [
              {
                "X": 7.28,
                "Y": 2.7347
              },
              {
                "X": 8.2506,
                "Y": 2.7347
              },
              {
                "X": 8.2506,
                "Y": 3.0417
              },
              {
                "X": 7.28,
                "Y": 3.0583
              }
            ]
          },
          {
            "CellId": "9,0",
            "CellName": "A10",
            "Value": "Amortization of net actuarial loss and prior service cost/benefit(1)",
            "Polygon": [
              {
                "X": 0.2532,
                "Y": 3.05
              },
              {
                "X": 3.6131,
                "Y": 3.05
              },
              {
                "X": 3.6131,
                "Y": 3.3653
              },
              {
                "X": 0.2532,
                "Y": 3.3653
              }
            ]
          },
          {
            "CellId": "9,1",
            "CellName": "B10",
            "Value": "28",
            "Polygon": [
              {
                "X": 3.6131,
                "Y": 3.05
              },
              {
                "X": 5.2889,
                "Y": 3.05
              },
              {
                "X": 5.2889,
                "Y": 3.3736
              },
              {
                "X": 3.6131,
                "Y": 3.3653
              }
            ]
          },
          {
            "CellId": "9,2",
            "CellName": "C10",
            "Value": "61",
            "Polygon": [
              {
                "X": 5.2889,
                "Y": 3.05
              },
              {
                "X": 6.2927,
                "Y": 3.05
              },
              {
                "X": 6.2927,
                "Y": 3.3736
              },
              {
                "X": 5.2889,
                "Y": 3.3736
              }
            ]
          },
          {
            "CellId": "9,3",
            "CellName": "D10",
            "Value": "1",
            "Polygon": [
              {
                "X": 6.2927,
                "Y": 3.05
              },
              {
                "X": 7.28,
                "Y": 3.0583
              },
              {
                "X": 7.28,
                "Y": 3.3736
              },
              {
                "X": 6.2927,
                "Y": 3.3736
              }
            ]
          },
          {
            "CellId": "9,4",
            "CellName": "E10",
            "Value": "1",
            "Polygon": [
              {
                "X": 7.28,
                "Y": 3.0583
              },
              {
                "X": 8.2506,
                "Y": 3.0417
              },
              {
                "X": 8.2506,
                "Y": 3.3736
              },
              {
                "X": 7.28,
                "Y": 3.3736
              }
            ]
          },
          {
            "CellId": "10,0",
            "CellName": "A11",
            "Value": "Tax expense(2)",
            "Polygon": [
              {
                "X": 0.2532,
                "Y": 3.3653
              },
              {
                "X": 3.6131,
                "Y": 3.3653
              },
              {
                "X": 3.6131,
                "Y": 3.5312
              },
              {
                "X": 0.2532,
                "Y": 3.5312
              }
            ]
          },
          {
            "CellId": "10,1",
            "CellName": "B11",
            "Value": "—",
            "Polygon": [
              {
                "X": 3.6131,
                "Y": 3.3653
              },
              {
                "X": 5.2889,
                "Y": 3.3736
              },
              {
                "X": 5.2889,
                "Y": 3.5395
              },
              {
                "X": 3.6131,
                "Y": 3.5312
              }
            ]
          },
          {
            "CellId": "10,2",
            "CellName": "C11",
            "Value": "(1)",
            "Polygon": [
              {
                "X": 5.2889,
                "Y": 3.3736
              },
              {
                "X": 6.2927,
                "Y": 3.3736
              },
              {
                "X": 6.2927,
                "Y": 3.5395
              },
              {
                "X": 5.2889,
                "Y": 3.5395
              }
            ]
          },
          {
            "CellId": "10,3",
            "CellName": "D11",
            "Value": "—",
            "Polygon": [
              {
                "X": 6.2927,
                "Y": 3.3736
              },
              {
                "X": 7.28,
                "Y": 3.3736
              },
              {
                "X": 7.28,
                "Y": 3.5395
              },
              {
                "X": 6.2927,
                "Y": 3.5395
              }
            ]
          },
          {
            "CellId": "10,4",
            "CellName": "E11",
            "Value": "—",
            "Polygon": [
              {
                "X": 7.28,
                "Y": 3.3736
              },
              {
                "X": 8.2506,
                "Y": 3.3736
              },
              {
                "X": 8.2506,
                "Y": 3.5395
              },
              {
                "X": 7.28,
                "Y": 3.5395
              }
            ]
          },
          {
            "CellId": "11,0",
            "CellName": "A12",
            "Value": "Total amount reclassified from Accumulated other comprehensive loss, net of tax(6)",
            "Polygon": [
              {
                "X": 0.2532,
                "Y": 3.5312
              },
              {
                "X": 3.6131,
                "Y": 3.5312
              },
              {
                "X": 3.6131,
                "Y": 3.8465
              },
              {
                "X": 0.2532,
                "Y": 3.8465
              }
            ]
          },
          {
            "CellId": "11,1",
            "CellName": "B12",
            "Value": "28",
            "Polygon": [
              {
                "X": 3.6131,
                "Y": 3.5312
              },
              {
                "X": 5.2889,
                "Y": 3.5395
              },
              {
                "X": 5.2889,
                "Y": 3.8465
              },
              {
                "X": 3.6131,
                "Y": 3.8465
              }
            ]
          },
          {
            "CellId": "11,2",
            "CellName": "C12",
            "Value": "60",
            "Polygon": [
              {
                "X": 5.2889,
                "Y": 3.5395
              },
              {
                "X": 6.2927,
                "Y": 3.5395
              },
              {
                "X": 6.2927,
                "Y": 3.8465
              },
              {
                "X": 5.2889,
                "Y": 3.8465
              }
            ]
          },
          {
            "CellId": "11,3",
            "CellName": "D12",
            "Value": "1",
            "Polygon": [
              {
                "X": 6.2927,
                "Y": 3.5395
              },
              {
                "X": 7.28,
                "Y": 3.5395
              },
              {
                "X": 7.28,
                "Y": 3.8465
              },
              {
                "X": 6.2927,
                "Y": 3.8465
              }
            ]
          },
          {
            "CellId": "11,4",
            "CellName": "E12",
            "Value": "1",
            "Polygon": [
              {
                "X": 7.28,
                "Y": 3.5395
              },
              {
                "X": 8.2506,
                "Y": 3.5395
              },
              {
                "X": 8.2506,
                "Y": 3.8465
              },
              {
                "X": 7.28,
                "Y": 3.8465
              }
            ]
          },
          {
            "CellId": "12,0",
            "CellName": "A13",
            "Value": "Total Other comprehensive income",
            "Polygon": [
              {
                "X": 0.2532,
                "Y": 3.8465
              },
              {
                "X": 3.6131,
                "Y": 3.8465
              },
              {
                "X": 3.6131,
                "Y": 4.0208
              },
              {
                "X": 0.2532,
                "Y": 4.0208
              }
            ]
          },
          {
            "CellId": "12,1",
            "CellName": "B13",
            "Value": "22",
            "Polygon": [
              {
                "X": 3.6131,
                "Y": 3.8465
              },
              {
                "X": 5.2889,
                "Y": 3.8465
              },
              {
                "X": 5.2889,
                "Y": 4.0208
              },
              {
                "X": 3.6131,
                "Y": 4.0208
              }
            ]
          },
          {
            "CellId": "12,2",
            "CellName": "C13",
            "Value": "131",
            "Polygon": [
              {
                "X": 5.2889,
                "Y": 3.8465
              },
              {
                "X": 6.2927,
                "Y": 3.8465
              },
              {
                "X": 6.2927,
                "Y": 4.0208
              },
              {
                "X": 5.2889,
                "Y": 4.0208
              }
            ]
          },
          {
            "CellId": "12,3",
            "CellName": "D13",
            "Value": "1",
            "Polygon": [
              {
                "X": 6.2927,
                "Y": 3.8465
              },
              {
                "X": 7.28,
                "Y": 3.8465
              },
              {
                "X": 7.28,
                "Y": 4.0208
              },
              {
                "X": 6.2927,
                "Y": 4.0208
              }
            ]
          },
          {
            "CellId": "12,4",
            "CellName": "E13",
            "Value": "1",
            "Polygon": [
              {
                "X": 7.28,
                "Y": 3.8465
              },
              {
                "X": 8.2506,
                "Y": 3.8465
              },
              {
                "X": 8.2506,
                "Y": 4.0208
              },
              {
                "X": 7.28,
                "Y": 4.0208
              }
            ]
          },
          {
            "CellId": "13,0",
            "CellName": "A14",
            "Value": "Balance at end of period",
            "Polygon": [
              {
                "X": 0.2532,
                "Y": 4.0208
              },
              {
                "X": 3.6131,
                "Y": 4.0208
              },
              {
                "X": 3.6131,
                "Y": 4.1784
              },
              {
                "X": 0.2532,
                "Y": 4.1784
              }
            ]
          },
          {
            "CellId": "13,1",
            "CellName": "B14",
            "Value": "$ (860)",
            "Polygon": [
              {
                "X": 3.6131,
                "Y": 4.0208
              },
              {
                "X": 5.2889,
                "Y": 4.0208
              },
              {
                "X": 5.2889,
                "Y": 4.1867
              },
              {
                "X": 3.6131,
                "Y": 4.1784
              }
            ]
          },
          {
            "CellId": "13,2",
            "CellName": "C14",
            "Value": "$ (2,405)",
            "Polygon": [
              {
                "X": 5.2889,
                "Y": 4.0208
              },
              {
                "X": 6.2927,
                "Y": 4.0208
              },
              {
                "X": 6.2927,
                "Y": 4.1867
              },
              {
                "X": 5.2889,
                "Y": 4.1867
              }
            ]
          },
          {
            "CellId": "13,3",
            "CellName": "D14",
            "Value": "$ (12)",
            "Polygon": [
              {
                "X": 6.2927,
                "Y": 4.0208
              },
              {
                "X": 7.28,
                "Y": 4.0208
              },
              {
                "X": 7.28,
                "Y": 4.1784
              },
              {
                "X": 6.2927,
                "Y": 4.1867
              }
            ]
          },
          {
            "CellId": "13,4",
            "CellName": "E14",
            "Value": "$ (66)",
            "Polygon": [
              {
                "X": 7.28,
                "Y": 4.0208
              },
              {
                "X": 8.2506,
                "Y": 4.0208
              },
              {
                "X": 8.2506,
                "Y": 4.1784
              },
              {
                "X": 7.28,
                "Y": 4.1784
              }
            ]
          },
          {
            "CellId": "14,0",
            "CellName": "A15",
            "Value": "Foreign currency translation",
            "Polygon": [
              {
                "X": 0.2532,
                "Y": 4.1784
              },
              {
                "X": 3.6131,
                "Y": 4.1784
              },
              {
                "X": 3.6131,
                "Y": 4.502
              },
              {
                "X": 0.2449,
                "Y": 4.502
              }
            ]
          },
          {
            "CellId": "14,1",
            "CellName": "B15",
            "Value": "",
            "Polygon": [
              {
                "X": 3.6131,
                "Y": 4.1784
              },
              {
                "X": 5.2889,
                "Y": 4.1867
              },
              {
                "X": 5.2889,
                "Y": 4.502
              },
              {
                "X": 3.6131,
                "Y": 4.502
              }
            ]
          },
          {
            "CellId": "14,2",
            "CellName": "C15",
            "Value": "",
            "Polygon": [
              {
                "X": 5.2889,
                "Y": 4.1867
              },
              {
                "X": 6.2927,
                "Y": 4.1867
              },
              {
                "X": 6.2927,
                "Y": 4.502
              },
              {
                "X": 5.2889,
                "Y": 4.502
              }
            ]
          },
          {
            "CellId": "14,3",
            "CellName": "D15",
            "Value": "",
            "Polygon": [
              {
                "X": 6.2927,
                "Y": 4.1867
              },
              {
                "X": 7.28,
                "Y": 4.1784
              },
              {
                "X": 7.28,
                "Y": 4.502
              },
              {
                "X": 6.2927,
                "Y": 4.502
              }
            ]
          },
          {
            "CellId": "14,4",
            "CellName": "E15",
            "Value": "",
            "Polygon": [
              {
                "X": 7.28,
                "Y": 4.1784
              },
              {
                "X": 8.2506,
                "Y": 4.1784
              },
              {
                "X": 8.2506,
                "Y": 4.502
              },
              {
                "X": 7.28,
                "Y": 4.502
              }
            ]
          },
          {
            "CellId": "15,0",
            "CellName": "A16",
            "Value": "Balance at beginning of period",
            "Polygon": [
              {
                "X": 0.2449,
                "Y": 4.502
              },
              {
                "X": 3.6131,
                "Y": 4.502
              },
              {
                "X": 3.6131,
                "Y": 4.6679
              },
              {
                "X": 0.2449,
                "Y": 4.6679
              }
            ]
          },
          {
            "CellId": "15,1",
            "CellName": "B16",
            "Value": "$ (2,614)",
            "Polygon": [
              {
                "X": 3.6131,
                "Y": 4.502
              },
              {
                "X": 5.2889,
                "Y": 4.502
              },
              {
                "X": 5.2889,
                "Y": 4.6679
              },
              {
                "X": 3.6131,
                "Y": 4.6679
              }
            ]
          },
          {
            "CellId": "15,2",
            "CellName": "C16",
            "Value": "$ (2,385)",
            "Polygon": [
              {
                "X": 5.2889,
                "Y": 4.502
              },
              {
                "X": 6.2927,
                "Y": 4.502
              },
              {
                "X": 6.2927,
                "Y": 4.6679
              },
              {
                "X": 5.2889,
                "Y": 4.6679
              }
            ]
          },
          {
            "CellId": "15,3",
            "CellName": "D16",
            "Value": "$ (937)",
            "Polygon": [
              {
                "X": 6.2927,
                "Y": 4.502
              },
              {
                "X": 7.28,
                "Y": 4.502
              },
              {
                "X": 7.28,
                "Y": 4.6679
              },
              {
                "X": 6.2927,
                "Y": 4.6679
              }
            ]
          },
          {
            "CellId": "15,4",
            "CellName": "E16",
            "Value": "$ (844)",
            "Polygon": [
              {
                "X": 7.28,
                "Y": 4.502
              },
              {
                "X": 8.2506,
                "Y": 4.502
              },
              {
                "X": 8.2506,
                "Y": 4.6679
              },
              {
                "X": 7.28,
                "Y": 4.6679
              }
            ]
          },
          {
            "CellId": "16,0",
            "CellName": "A17",
            "Value": "Other comprehensive income (loss)",
            "Polygon": [
              {
                "X": 0.2449,
                "Y": 4.6679
              },
              {
                "X": 3.6131,
                "Y": 4.6679
              },
              {
                "X": 3.6131,
                "Y": 4.8256
              },
              {
                "X": 0.2449,
                "Y": 4.8173
              }
            ]
          },
          {
            "CellId": "16,1",
            "CellName": "B17",
            "Value": "326",
            "Polygon": [
              {
                "X": 3.6131,
                "Y": 4.6679
              },
              {
                "X": 5.2889,
                "Y": 4.6679
              },
              {
                "X": 5.2889,
                "Y": 4.8339
              },
              {
                "X": 3.6131,
                "Y": 4.8256
              }
            ]
          },
          {
            "CellId": "16,2",
            "CellName": "C17",
            "Value": "(176)",
            "Polygon": [
              {
                "X": 5.2889,
                "Y": 4.6679
              },
              {
                "X": 6.2927,
                "Y": 4.6679
              },
              {
                "X": 6.2927,
                "Y": 4.8339
              },
              {
                "X": 5.2889,
                "Y": 4.8339
              }
            ]
          },
          {
            "CellId": "16,3",
            "CellName": "D17",
            "Value": "98",
            "Polygon": [
              {
                "X": 6.2927,
                "Y": 4.6679
              },
              {
                "X": 7.28,
                "Y": 4.6679
              },
              {
                "X": 7.28,
                "Y": 4.8256
              },
              {
                "X": 6.2927,
                "Y": 4.8339
              }
            ]
          },
          {
            "CellId": "16,4",
            "CellName": "E17",
            "Value": "(60)",
            "Polygon": [
              {
                "X": 7.28,
                "Y": 4.6679
              },
              {
                "X": 8.2506,
                "Y": 4.6679
              },
              {
                "X": 8.2506,
                "Y": 4.8339
              },
              {
                "X": 7.28,
                "Y": 4.8256
              }
            ]
          },
          {
            "CellId": "17,0",
            "CellName": "A18",
            "Value": "Balance at end of period",
            "Polygon": [
              {
                "X": 0.2449,
                "Y": 4.8173
              },
              {
                "X": 3.6131,
                "Y": 4.8256
              },
              {
                "X": 3.6131,
                "Y": 4.9998
              },
              {
                "X": 0.2449,
                "Y": 4.9998
              }
            ]
          },
          {
            "CellId": "17,1",
            "CellName": "B18",
            "Value": "$ (2,288)",
            "Polygon": [
              {
                "X": 3.6131,
                "Y": 4.8256
              },
              {
                "X": 5.2889,
                "Y": 4.8339
              },
              {
                "X": 5.2889,
                "Y": 4.9998
              },
              {
                "X": 3.6131,
                "Y": 4.9998
              }
            ]
          },
          {
            "CellId": "17,2",
            "CellName": "C18",
            "Value": "$ (2,561)",
            "Polygon": [
              {
                "X": 5.2889,
                "Y": 4.8339
              },
              {
                "X": 6.2927,
                "Y": 4.8339
              },
              {
                "X": 6.2927,
                "Y": 4.9998
              },
              {
                "X": 5.2889,
                "Y": 4.9998
              }
            ]
          },
          {
            "CellId": "17,3",
            "CellName": "D18",
            "Value": "$ (839)",
            "Polygon": [
              {
                "X": 6.2927,
                "Y": 4.8339
              },
              {
                "X": 7.28,
                "Y": 4.8256
              },
              {
                "X": 7.28,
                "Y": 4.9998
              },
              {
                "X": 6.2927,
                "Y": 4.9998
              }
            ]
          },
          {
            "CellId": "17,4",
            "CellName": "E18",
            "Value": "$ (904)",
            "Polygon": [
              {
                "X": 7.28,
                "Y": 4.8256
              },
              {
                "X": 8.2506,
                "Y": 4.8339
              },
              {
                "X": 8.2506,
                "Y": 4.9998
              },
              {
                "X": 7.28,
                "Y": 4.9998
              }
            ]
          },
          {
            "CellId": "18,0",
            "CellName": "A19",
            "Value": "Cash flow hedges (L)",
            "Polygon": [
              {
                "X": 0.2449,
                "Y": 4.9998
              },
              {
                "X": 3.6131,
                "Y": 4.9998
              },
              {
                "X": 3.6131,
                "Y": 5.3151
              },
              {
                "X": 0.2449,
                "Y": 5.3151
              }
            ]
          },
          {
            "CellId": "18,1",
            "CellName": "B19",
            "Value": "",
            "Polygon": [
              {
                "X": 3.6131,
                "Y": 4.9998
              },
              {
                "X": 5.2889,
                "Y": 4.9998
              },
              {
                "X": 5.2889,
                "Y": 5.3151
              },
              {
                "X": 3.6131,
                "Y": 5.3151
              }
            ]
          },
          {
            "CellId": "18,2",
            "CellName": "C19",
            "Value": "",
            "Polygon": [
              {
                "X": 5.2889,
                "Y": 4.9998
              },
              {
                "X": 6.2927,
                "Y": 4.9998
              },
              {
                "X": 6.2927,
                "Y": 5.3151
              },
              {
                "X": 5.2889,
                "Y": 5.3151
              }
            ]
          },
          {
            "CellId": "18,3",
            "CellName": "D19",
            "Value": "",
            "Polygon": [
              {
                "X": 6.2927,
                "Y": 4.9998
              },
              {
                "X": 7.28,
                "Y": 4.9998
              },
              {
                "X": 7.28,
                "Y": 5.3151
              },
              {
                "X": 6.2927,
                "Y": 5.3151
              }
            ]
          },
          {
            "CellId": "18,4",
            "CellName": "E19",
            "Value": "",
            "Polygon": [
              {
                "X": 7.28,
                "Y": 4.9998
              },
              {
                "X": 8.2506,
                "Y": 4.9998
              },
              {
                "X": 8.2506,
                "Y": 5.3151
              },
              {
                "X": 7.28,
                "Y": 5.3151
              }
            ]
          },
          {
            "CellId": "19,0",
            "CellName": "A20",
            "Value": "Balance at beginning of period",
            "Polygon": [
              {
                "X": 0.2449,
                "Y": 5.3151
              },
              {
                "X": 3.6131,
                "Y": 5.3151
              },
              {
                "X": 3.6131,
                "Y": 5.4727
              },
              {
                "X": 0.2449,
                "Y": 5.4727
              }
            ]
          },
          {
            "CellId": "19,1",
            "CellName": "B20",
            "Value": "$ (1,096)",
            "Polygon": [
              {
                "X": 3.6131,
                "Y": 5.3151
              },
              {
                "X": 5.2889,
                "Y": 5.3151
              },
              {
                "X": 5.2889,
                "Y": 5.4727
              },
              {
                "X": 3.6131,
                "Y": 5.4727
              }
            ]
          },
          {
            "CellId": "19,2",
            "CellName": "C20",
            "Value": "$ (708)",
            "Polygon": [
              {
                "X": 5.2889,
                "Y": 5.3151
              },
              {
                "X": 6.2927,
                "Y": 5.3151
              },
              {
                "X": 6.2927,
                "Y": 5.4727
              },
              {
                "X": 5.2889,
                "Y": 5.4727
              }
            ]
          },
          {
            "CellId": "19,3",
            "CellName": "D20",
            "Value": "$ (1)",
            "Polygon": [
              {
                "X": 6.2927,
                "Y": 5.3151
              },
              {
                "X": 7.28,
                "Y": 5.3151
              },
              {
                "X": 7.28,
                "Y": 5.4727
              },
              {
                "X": 6.2927,
                "Y": 5.4727
              }
            ]
          },
          {
            "CellId": "19,4",
            "CellName": "E20",
            "Value": "$ (1)",
            "Polygon": [
              {
                "X": 7.28,
                "Y": 5.3151
              },
              {
                "X": 8.2506,
                "Y": 5.3151
              },
              {
                "X": 8.2506,
                "Y": 5.4727
              },
              {
                "X": 7.28,
                "Y": 5.4727
              }
            ]
          },
          {
            "CellId": "20,0",
            "CellName": "A21",
            "Value": "Other comprehensive (loss) income:",
            "Polygon": [
              {
                "X": 0.2449,
                "Y": 5.4727
              },
              {
                "X": 3.6131,
                "Y": 5.4727
              },
              {
                "X": 3.6131,
                "Y": 5.6387
              },
              {
                "X": 0.2449,
                "Y": 5.6387
              }
            ]
          },
          {
            "CellId": "20,1",
            "CellName": "B21",
            "Value": "",
            "Polygon": [
              {
                "X": 3.6131,
                "Y": 5.4727
              },
              {
                "X": 5.2889,
                "Y": 5.4727
              },
              {
                "X": 5.2889,
                "Y": 5.6387
              },
              {
                "X": 3.6131,
                "Y": 5.6387
              }
            ]
          },
          {
            "CellId": "20,2",
            "CellName": "C21",
            "Value": "",
            "Polygon": [
              {
                "X": 5.2889,
                "Y": 5.4727
              },
              {
                "X": 6.2927,
                "Y": 5.4727
              },
              {
                "X": 6.2927,
                "Y": 5.6387
              },
              {
                "X": 5.2889,
                "Y": 5.6387
              }
            ]
          },
          {
            "CellId": "20,3",
            "CellName": "D21",
            "Value": "",
            "Polygon": [
              {
                "X": 6.2927,
                "Y": 5.4727
              },
              {
                "X": 7.28,
                "Y": 5.4727
              },
              {
                "X": 7.28,
                "Y": 5.647
              },
              {
                "X": 6.2927,
                "Y": 5.6387
              }
            ]
          },
          {
            "CellId": "20,4",
            "CellName": "E21",
            "Value": "",
            "Polygon": [
              {
                "X": 7.28,
                "Y": 5.4727
              },
              {
                "X": 8.2506,
                "Y": 5.4727
              },
              {
                "X": 8.2506,
                "Y": 5.647
              },
              {
                "X": 7.28,
                "Y": 5.647
              }
            ]
          },
          {
            "CellId": "21,0",
            "CellName": "A22",
            "Value": "Net change from periodic revaluations",
            "Polygon": [
              {
                "X": 0.2449,
                "Y": 5.6387
              },
              {
                "X": 3.6131,
                "Y": 5.6387
              },
              {
                "X": 3.6131,
                "Y": 5.7963
              },
              {
                "X": 0.2449,
                "Y": 5.7963
              }
            ]
          },
          {
            "CellId": "21,1",
            "CellName": "B22",
            "Value": "(1,063)",
            "Polygon": [
              {
                "X": 3.6131,
                "Y": 5.6387
              },
              {
                "X": 5.2889,
                "Y": 5.6387
              },
              {
                "X": 5.2889,
                "Y": 5.7963
              },
              {
                "X": 3.6131,
                "Y": 5.7963
              }
            ]
          },
          {
            "CellId": "21,2",
            "CellName": "C22",
            "Value": "(303)",
            "Polygon": [
              {
                "X": 5.2889,
                "Y": 5.6387
              },
              {
                "X": 6.2927,
                "Y": 5.6387
              },
              {
                "X": 6.2927,
                "Y": 5.7963
              },
              {
                "X": 5.2889,
                "Y": 5.7963
              }
            ]
          },
          {
            "CellId": "21,3",
            "CellName": "D22",
            "Value": "1",
            "Polygon": [
              {
                "X": 6.2927,
                "Y": 5.6387
              },
              {
                "X": 7.28,
                "Y": 5.647
              },
              {
                "X": 7.28,
                "Y": 5.7963
              },
              {
                "X": 6.2927,
                "Y": 5.7963
              }
            ]
          },
          {
            "CellId": "21,4",
            "CellName": "E22",
            "Value": "(10)",
            "Polygon": [
              {
                "X": 7.28,
                "Y": 5.647
              },
              {
                "X": 8.2506,
                "Y": 5.647
              },
              {
                "X": 8.2506,
                "Y": 5.7963
              },
              {
                "X": 7.28,
                "Y": 5.7963
              }
            ]
          },
          {
            "CellId": "22,0",
            "CellName": "A23",
            "Value": "Tax benefit(2)",
            "Polygon": [
              {
                "X": 0.2449,
                "Y": 5.7963
              },
              {
                "X": 3.6131,
                "Y": 5.7963
              },
              {
                "X": 3.6131,
                "Y": 5.9623
              },
              {
                "X": 0.2449,
                "Y": 5.9623
              }
            ]
          },
          {
            "CellId": "22,1",
            "CellName": "B23",
            "Value": "153",
            "Polygon": [
              {
                "X": 3.6131,
                "Y": 5.7963
              },
              {
                "X": 5.2889,
                "Y": 5.7963
              },
              {
                "X": 5.2889,
                "Y": 5.9706
              },
              {
                "X": 3.6131,
                "Y": 5.9623
              }
            ]
          },
          {
            "CellId": "22,2",
            "CellName": "C23",
            "Value": "56",
            "Polygon": [
              {
                "X": 5.2889,
                "Y": 5.7963
              },
              {
                "X": 6.2927,
                "Y": 5.7963
              },
              {
                "X": 6.2927,
                "Y": 5.9706
              },
              {
                "X": 5.2889,
                "Y": 5.9706
              }
            ]
          },
          {
            "CellId": "22,3",
            "CellName": "D23",
            "Value": "—",
            "Polygon": [
              {
                "X": 6.2927,
                "Y": 5.7963
              },
              {
                "X": 7.28,
                "Y": 5.7963
              },
              {
                "X": 7.28,
                "Y": 5.9706
              },
              {
                "X": 6.2927,
                "Y": 5.9706
              }
            ]
          },
          {
            "CellId": "22,4",
            "CellName": "E23",
            "Value": "3",
            "Polygon": [
              {
                "X": 7.28,
                "Y": 5.7963
              },
              {
                "X": 8.2506,
                "Y": 5.7963
              },
              {
                "X": 8.2506,
                "Y": 5.9706
              },
              {
                "X": 7.28,
                "Y": 5.9706
              }
            ]
          },
          {
            "CellId": "23,0",
            "CellName": "A24",
            "Value": "Total Other comprehensive (loss) income before reclassifications, net of tax",
            "Polygon": [
              {
                "X": 0.2449,
                "Y": 5.9623
              },
              {
                "X": 3.6131,
                "Y": 5.9623
              },
              {
                "X": 3.6131,
                "Y": 6.2775
              },
              {
                "X": 0.2449,
                "Y": 6.2775
              }
            ]
          },
          {
            "CellId": "23,1",
            "CellName": "B24",
            "Value": "(910)",
            "Polygon": [
              {
                "X": 3.6131,
                "Y": 5.9623
              },
              {
                "X": 5.2889,
                "Y": 5.9706
              },
              {
                "X": 5.2889,
                "Y": 6.2775
              },
              {
                "X": 3.6131,
                "Y": 6.2775
              }
            ]
          },
          {
            "CellId": "23,2",
            "CellName": "C24",
            "Value": "(247)",
            "Polygon": [
              {
                "X": 5.2889,
                "Y": 5.9706
              },
              {
                "X": 6.2927,
                "Y": 5.9706
              },
              {
                "X": 6.2927,
                "Y": 6.2775
              },
              {
                "X": 5.2889,
                "Y": 6.2775
              }
            ]
          },
          {
            "CellId": "23,3",
            "CellName": "D24",
            "Value": "1",
            "Polygon": [
              {
                "X": 6.2927,
                "Y": 5.9706
              },
              {
                "X": 7.28,
                "Y": 5.9706
              },
              {
                "X": 7.28,
                "Y": 6.2775
              },
              {
                "X": 6.2927,
                "Y": 6.2775
              }
            ]
          },
          {
            "CellId": "23,4",
            "CellName": "E24",
            "Value": "(7)",
            "Polygon": [
              {
                "X": 7.28,
                "Y": 5.9706
              },
              {
                "X": 8.2506,
                "Y": 5.9706
              },
              {
                "X": 8.2506,
                "Y": 6.2775
              },
              {
                "X": 7.28,
                "Y": 6.2775
              }
            ]
          },
          {
            "CellId": "24,0",
            "CellName": "A25",
            "Value": "Net amount reclassified to earnings:",
            "Polygon": [
              {
                "X": 0.2449,
                "Y": 6.2775
              },
              {
                "X": 3.6131,
                "Y": 6.2775
              },
              {
                "X": 3.6131,
                "Y": 6.4186
              },
              {
                "X": 0.2449,
                "Y": 6.4186
              }
            ]
          },
          {
            "CellId": "24,1",
            "CellName": "B25",
            "Value": "",
            "Polygon": [
              {
                "X": 3.6131,
                "Y": 6.2775
              },
              {
                "X": 5.2889,
                "Y": 6.2775
              },
              {
                "X": 5.2889,
                "Y": 6.4186
              },
              {
                "X": 3.6131,
                "Y": 6.4186
              }
            ]
          },
          {
            "CellId": "24,2",
            "CellName": "C25",
            "Value": "",
            "Polygon": [
              {
                "X": 5.2889,
                "Y": 6.2775
              },
              {
                "X": 6.2927,
                "Y": 6.2775
              },
              {
                "X": 6.2927,
                "Y": 6.4186
              },
              {
                "X": 5.2889,
                "Y": 6.4186
              }
            ]
          },
          {
            "CellId": "24,3",
            "CellName": "D25",
            "Value": "",
            "Polygon": [
              {
                "X": 6.2927,
                "Y": 6.2775
              },
              {
                "X": 7.28,
                "Y": 6.2775
              },
              {
                "X": 7.28,
                "Y": 6.4186
              },
              {
                "X": 6.2927,
                "Y": 6.4186
              }
            ]
          },
          {
            "CellId": "24,4",
            "CellName": "E25",
            "Value": "",
            "Polygon": [
              {
                "X": 7.28,
                "Y": 6.2775
              },
              {
                "X": 8.2506,
                "Y": 6.2775
              },
              {
                "X": 8.2506,
                "Y": 6.4269
              },
              {
                "X": 7.28,
                "Y": 6.4186
              }
            ]
          },
          {
            "CellId": "25,0",
            "CellName": "A26",
            "Value": "Aluminum contracts(3)",
            "Polygon": [
              {
                "X": 0.2449,
                "Y": 6.4186
              },
              {
                "X": 3.6131,
                "Y": 6.4186
              },
              {
                "X": 3.6131,
                "Y": 6.6011
              },
              {
                "X": 0.2449,
                "Y": 6.6011
              }
            ]
          },
          {
            "CellId": "25,1",
            "CellName": "B26",
            "Value": "110",
            "Polygon": [
              {
                "X": 3.6131,
                "Y": 6.4186
              },
              {
                "X": 5.2889,
                "Y": 6.4186
              },
              {
                "X": 5.2889,
                "Y": 6.6094
              },
              {
                "X": 3.6131,
                "Y": 6.6011
              }
            ]
          },
          {
            "CellId": "25,2",
            "CellName": "C26",
            "Value": "41",
            "Polygon": [
              {
                "X": 5.2889,
                "Y": 6.4186
              },
              {
                "X": 6.2927,
                "Y": 6.4186
              },
              {
                "X": 6.2927,
                "Y": 6.6094
              },
              {
                "X": 5.2889,
                "Y": 6.6094
              }
            ]
          },
          {
            "CellId": "25,3",
            "CellName": "D26",
            "Value": "—",
            "Polygon": [
              {
                "X": 6.2927,
                "Y": 6.4186
              },
              {
                "X": 7.28,
                "Y": 6.4186
              },
              {
                "X": 7.28,
                "Y": 6.6094
              },
              {
                "X": 6.2927,
                "Y": 6.6094
              }
            ]
          },
          {
            "CellId": "25,4",
            "CellName": "E26",
            "Value": "—",
            "Polygon": [
              {
                "X": 7.28,
                "Y": 6.4186
              },
              {
                "X": 8.2506,
                "Y": 6.4269
              },
              {
                "X": 8.2506,
                "Y": 6.6094
              },
              {
                "X": 7.28,
                "Y": 6.6094
              }
            ]
          },
          {
            "CellId": "26,0",
            "CellName": "A27",
            "Value": "Financial contracts(4)",
            "Polygon": [
              {
                "X": 0.2449,
                "Y": 6.6011
              },
              {
                "X": 3.6131,
                "Y": 6.6011
              },
              {
                "X": 3.6131,
                "Y": 6.7671
              },
              {
                "X": 0.2449,
                "Y": 6.7671
              }
            ]
          },
          {
            "CellId": "26,1",
            "CellName": "B27",
            "Value": "—",
            "Polygon": [
              {
                "X": 3.6131,
                "Y": 6.6011
              },
              {
                "X": 5.2889,
                "Y": 6.6094
              },
              {
                "X": 5.2889,
                "Y": 6.7671
              },
              {
                "X": 3.6131,
                "Y": 6.7671
              }
            ]
          },
          {
            "CellId": "26,2",
            "CellName": "C27",
            "Value": "9",
            "Polygon": [
              {
                "X": 5.2889,
                "Y": 6.6094
              },
              {
                "X": 6.2927,
                "Y": 6.6094
              },
              {
                "X": 6.2927,
                "Y": 6.7671
              },
              {
                "X": 5.2889,
                "Y": 6.7671
              }
            ]
          },
          {
            "CellId": "26,3",
            "CellName": "D27",
            "Value": "—",
            "Polygon": [
              {
                "X": 6.2927,
                "Y": 6.6094
              },
              {
                "X": 7.28,
                "Y": 6.6094
              },
              {
                "X": 7.28,
                "Y": 6.7671
              },
              {
                "X": 6.2927,
                "Y": 6.7671
              }
            ]
          },
          {
            "CellId": "26,4",
            "CellName": "E27",
            "Value": "6",
            "Polygon": [
              {
                "X": 7.28,
                "Y": 6.6094
              },
              {
                "X": 8.2506,
                "Y": 6.6094
              },
              {
                "X": 8.2506,
                "Y": 6.7671
              },
              {
                "X": 7.28,
                "Y": 6.7671
              }
            ]
          },
          {
            "CellId": "27,0",
            "CellName": "A28",
            "Value": "Interest rate contracts(5)",
            "Polygon": [
              {
                "X": 0.2449,
                "Y": 6.7671
              },
              {
                "X": 3.6131,
                "Y": 6.7671
              },
              {
                "X": 3.6131,
                "Y": 6.9247
              },
              {
                "X": 0.2449,
                "Y": 6.9247
              }
            ]
          },
          {
            "CellId": "27,1",
            "CellName": "B28",
            "Value": "4",
            "Polygon": [
              {
                "X": 3.6131,
                "Y": 6.7671
              },
              {
                "X": 5.2889,
                "Y": 6.7671
              },
              {
                "X": 5.2889,
                "Y": 6.933
              },
              {
                "X": 3.6131,
                "Y": 6.9247
              }
            ]
          },
          {
            "CellId": "27,2",
            "CellName": "C28",
            "Value": "3",
            "Polygon": [
              {
                "X": 5.2889,
                "Y": 6.7671
              },
              {
                "X": 6.2927,
                "Y": 6.7671
              },
              {
                "X": 6.2927,
                "Y": 6.9247
              },
              {
                "X": 5.2889,
                "Y": 6.933
              }
            ]
          },
          {
            "CellId": "27,3",
            "CellName": "D28",
            "Value": "—",
            "Polygon": [
              {
                "X": 6.2927,
                "Y": 6.7671
              },
              {
                "X": 7.28,
                "Y": 6.7671
              },
              {
                "X": 7.28,
                "Y": 6.933
              },
              {
                "X": 6.2927,
                "Y": 6.9247
              }
            ]
          },
          {
            "CellId": "27,4",
            "CellName": "E28",
            "Value": "—",
            "Polygon": [
              {
                "X": 7.28,
                "Y": 6.7671
              },
              {
                "X": 8.2506,
                "Y": 6.7671
              },
              {
                "X": 8.2506,
                "Y": 6.933
              },
              {
                "X": 7.28,
                "Y": 6.933
              }
            ]
          },
          {
            "CellId": "28,0",
            "CellName": "A29",
            "Value": "Foreign exchange contracts(3)",
            "Polygon": [
              {
                "X": 0.2449,
                "Y": 6.9247
              },
              {
                "X": 3.6131,
                "Y": 6.9247
              },
              {
                "X": 3.6131,
                "Y": 7.0741
              },
              {
                "X": 0.2449,
                "Y": 7.0741
              }
            ]
          },
          {
            "CellId": "28,1",
            "CellName": "B29",
            "Value": "—",
            "Polygon": [
              {
                "X": 3.6131,
                "Y": 6.9247
              },
              {
                "X": 5.2889,
                "Y": 6.933
              },
              {
                "X": 5.2889,
                "Y": 7.0741
              },
              {
                "X": 3.6131,
                "Y": 7.0741
              }
            ]
          },
          {
            "CellId": "28,2",
            "CellName": "C29",
            "Value": "(1)",
            "Polygon": [
              {
                "X": 5.2889,
                "Y": 6.933
              },
              {
                "X": 6.2927,
                "Y": 6.9247
              },
              {
                "X": 6.2845,
                "Y": 7.0741
              },
              {
                "X": 5.2889,
                "Y": 7.0741
              }
            ]
          },
          {
            "CellId": "28,3",
            "CellName": "D29",
            "Value": "—",
            "Polygon": [
              {
                "X": 6.2927,
                "Y": 6.9247
              },
              {
                "X": 7.28,
                "Y": 6.933
              },
              {
                "X": 7.28,
                "Y": 7.0741
              },
              {
                "X": 6.2845,
                "Y": 7.0741
              }
            ]
          },
          {
            "CellId": "28,4",
            "CellName": "E29",
            "Value": "—",
            "Polygon": [
              {
                "X": 7.28,
                "Y": 6.933
              },
              {
                "X": 8.2506,
                "Y": 6.933
              },
              {
                "X": 8.2506,
                "Y": 7.0824
              },
              {
                "X": 7.28,
                "Y": 7.0741
              }
            ]
          },
          {
            "CellId": "29,0",
            "CellName": "A30",
            "Value": "Sub-total",
            "Polygon": [
              {
                "X": 0.2449,
                "Y": 7.0741
              },
              {
                "X": 3.6131,
                "Y": 7.0741
              },
              {
                "X": 3.6048,
                "Y": 7.2483
              },
              {
                "X": 0.2449,
                "Y": 7.2483
              }
            ]
          },
          {
            "CellId": "29,1",
            "CellName": "B30",
            "Value": "114",
            "Polygon": [
              {
                "X": 3.6131,
                "Y": 7.0741
              },
              {
                "X": 5.2889,
                "Y": 7.0741
              },
              {
                "X": 5.2889,
                "Y": 7.2566
              },
              {
                "X": 3.6048,
                "Y": 7.2483
              }
            ]
          },
          {
            "CellId": "29,2",
            "CellName": "C30",
            "Value": "52",
            "Polygon": [
              {
                "X": 5.2889,
                "Y": 7.0741
              },
              {
                "X": 6.2845,
                "Y": 7.0741
              },
              {
                "X": 6.2845,
                "Y": 7.2566
              },
              {
                "X": 5.2889,
                "Y": 7.2566
              }
            ]
          },
          {
            "CellId": "29,3",
            "CellName": "D30",
            "Value": "—",
            "Polygon": [
              {
                "X": 6.2845,
                "Y": 7.0741
              },
              {
                "X": 7.28,
                "Y": 7.0741
              },
              {
                "X": 7.28,
                "Y": 7.2566
              },
              {
                "X": 6.2845,
                "Y": 7.2566
              }
            ]
          },
          {
            "CellId": "29,4",
            "CellName": "E30",
            "Value": "6",
            "Polygon": [
              {
                "X": 7.28,
                "Y": 7.0741
              },
              {
                "X": 8.2506,
                "Y": 7.0824
              },
              {
                "X": 8.2506,
                "Y": 7.2566
              },
              {
                "X": 7.28,
                "Y": 7.2566
              }
            ]
          },
          {
            "CellId": "30,0",
            "CellName": "A31",
            "Value": "Tax expense(2)",
            "Polygon": [
              {
                "X": 0.2449,
                "Y": 7.2483
              },
              {
                "X": 3.6048,
                "Y": 7.2483
              },
              {
                "X": 3.6048,
                "Y": 7.4059
              },
              {
                "X": 0.2449,
                "Y": 7.4059
              }
            ]
          },
          {
            "CellId": "30,1",
            "CellName": "B31",
            "Value": "(34)",
            "Polygon": [
              {
                "X": 3.6048,
                "Y": 7.2483
              },
              {
                "X": 5.2889,
                "Y": 7.2566
              },
              {
                "X": 5.2889,
                "Y": 7.4059
              },
              {
                "X": 3.6048,
                "Y": 7.4059
              }
            ]
          },
          {
            "CellId": "30,2",
            "CellName": "C31",
            "Value": "(9)",
            "Polygon": [
              {
                "X": 5.2889,
                "Y": 7.2566
              },
              {
                "X": 6.2845,
                "Y": 7.2566
              },
              {
                "X": 6.2845,
                "Y": 7.4142
              },
              {
                "X": 5.2889,
                "Y": 7.4059
              }
            ]
          },
          {
            "CellId": "30,3",
            "CellName": "D31",
            "Value": "—",
            "Polygon": [
              {
                "X": 6.2845,
                "Y": 7.2566
              },
              {
                "X": 7.28,
                "Y": 7.2566
              },
              {
                "X": 7.28,
                "Y": 7.4142
              },
              {
                "X": 6.2845,
                "Y": 7.4142
              }
            ]
          },
          {
            "CellId": "30,4",
            "CellName": "E31",
            "Value": "(2)",
            "Polygon": [
              {
                "X": 7.28,
                "Y": 7.2566
              },
              {
                "X": 8.2506,
                "Y": 7.2566
              },
              {
                "X": 8.2506,
                "Y": 7.4142
              },
              {
                "X": 7.28,
                "Y": 7.4142
              }
            ]
          },
          {
            "CellId": "31,0",
            "CellName": "A32",
            "Value": "Total amount reclassified from Accumulated other comprehensive loss, net of tax(6)",
            "Polygon": [
              {
                "X": 0.2449,
                "Y": 7.4059
              },
              {
                "X": 3.6048,
                "Y": 7.4059
              },
              {
                "X": 3.6048,
                "Y": 7.8955
              },
              {
                "X": 0.2449,
                "Y": 7.8955
              }
            ]
          },
          {
            "CellId": "31,1",
            "CellName": "B32",
            "Value": "80",
            "Polygon": [
              {
                "X": 3.6048,
                "Y": 7.4059
              },
              {
                "X": 5.2889,
                "Y": 7.4059
              },
              {
                "X": 5.2889,
                "Y": 7.8955
              },
              {
                "X": 3.6048,
                "Y": 7.8955
              }
            ]
          },
          {
            "CellId": "31,2",
            "CellName": "C32",
            "Value": "43",
            "Polygon": [
              {
                "X": 5.2889,
                "Y": 7.4059
              },
              {
                "X": 6.2845,
                "Y": 7.4142
              },
              {
                "X": 6.2845,
                "Y": 7.8955
              },
              {
                "X": 5.2889,
                "Y": 7.8955
              }
            ]
          },
          {
            "CellId": "31,3",
            "CellName": "D32",
            "Value": "—",
            "Polygon": [
              {
                "X": 6.2845,
                "Y": 7.4142
              },
              {
                "X": 7.28,
                "Y": 7.4142
              },
              {
                "X": 7.28,
                "Y": 7.8955
              },
              {
                "X": 6.2845,
                "Y": 7.8955
              }
            ]
          },
          {
            "CellId": "31,4",
            "CellName": "E32",
            "Value": "4",
            "Polygon": [
              {
                "X": 7.28,
                "Y": 7.4142
              },
              {
                "X": 8.2506,
                "Y": 7.4142
              },
              {
                "X": 8.2506,
                "Y": 7.9038
              },
              {
                "X": 7.28,
                "Y": 7.8955
              }
            ]
          },
          {
            "CellId": "32,0",
            "CellName": "A33",
            "Value": "Total Other comprehensive (loss) income",
            "Polygon": [
              {
                "X": 0.2449,
                "Y": 7.8955
              },
              {
                "X": 3.6048,
                "Y": 7.8955
              },
              {
                "X": 3.6048,
                "Y": 8.0614
              },
              {
                "X": 0.2449,
                "Y": 8.0614
              }
            ]
          },
          {
            "CellId": "32,1",
            "CellName": "B33",
            "Value": "(830)",
            "Polygon": [
              {
                "X": 3.6048,
                "Y": 7.8955
              },
              {
                "X": 5.2889,
                "Y": 7.8955
              },
              {
                "X": 5.2889,
                "Y": 8.0614
              },
              {
                "X": 3.6048,
                "Y": 8.0614
              }
            ]
          },
          {
            "CellId": "32,2",
            "CellName": "C33",
            "Value": "(204)",
            "Polygon": [
              {
                "X": 5.2889,
                "Y": 7.8955
              },
              {
                "X": 6.2845,
                "Y": 7.8955
              },
              {
                "X": 6.2845,
                "Y": 8.0614
              },
              {
                "X": 5.2889,
                "Y": 8.0614
              }
            ]
          },
          {
            "CellId": "32,3",
            "CellName": "D33",
            "Value": "1",
            "Polygon": [
              {
                "X": 6.2845,
                "Y": 7.8955
              },
              {
                "X": 7.28,
                "Y": 7.8955
              },
              {
                "X": 7.28,
                "Y": 8.0697
              },
              {
                "X": 6.2845,
                "Y": 8.0614
              }
            ]
          },
          {
            "CellId": "32,4",
            "CellName": "E33",
            "Value": "(3)",
            "Polygon": [
              {
                "X": 7.28,
                "Y": 7.8955
              },
              {
                "X": 8.2506,
                "Y": 7.9038
              },
              {
                "X": 8.2506,
                "Y": 8.0697
              },
              {
                "X": 7.28,
                "Y": 8.0697
              }
            ]
          },
          {
            "CellId": "33,0",
            "CellName": "A34",
            "Value": "Balance at end of period",
            "Polygon": [
              {
                "X": 0.2449,
                "Y": 8.0614
              },
              {
                "X": 3.6048,
                "Y": 8.0614
              },
              {
                "X": 3.6048,
                "Y": 8.2273
              },
              {
                "X": 0.2366,
                "Y": 8.2273
              }
            ]
          },
          {
            "CellId": "33,1",
            "CellName": "B34",
            "Value": "$ (1,926)",
            "Polygon": [
              {
                "X": 3.6048,
                "Y": 8.0614
              },
              {
                "X": 5.2889,
                "Y": 8.0614
              },
              {
                "X": 5.2889,
                "Y": 8.2273
              },
              {
                "X": 3.6048,
                "Y": 8.2273
              }
            ]
          },
          {
            "CellId": "33,2",
            "CellName": "C34",
            "Value": "$ (912)",
            "Polygon": [
              {
                "X": 5.2889,
                "Y": 8.0614
              },
              {
                "X": 6.2845,
                "Y": 8.0614
              },
              {
                "X": 6.2845,
                "Y": 8.2273
              },
              {
                "X": 5.2889,
                "Y": 8.2273
              }
            ]
          },
          {
            "CellId": "33,3",
            "CellName": "D34",
            "Value": "$ —",
            "Polygon": [
              {
                "X": 6.2845,
                "Y": 8.0614
              },
              {
                "X": 7.28,
                "Y": 8.0697
              },
              {
                "X": 7.28,
                "Y": 8.2356
              },
              {
                "X": 6.2845,
                "Y": 8.2273
              }
            ]
          },
          {
            "CellId": "33,4",
            "CellName": "E34",
            "Value": "$\n(4)",
            "Polygon": [
              {
                "X": 7.28,
                "Y": 8.0697
              },
              {
                "X": 8.2506,
                "Y": 8.0697
              },
              {
                "X": 8.2506,
                "Y": 8.2356
              },
              {
                "X": 7.28,
                "Y": 8.2356
              }
            ]
          },
          {
            "CellId": "34,0",
            "CellName": "A35",
            "Value": "",
            "Polygon": [
              {
                "X": 0.2366,
                "Y": 8.2273
              },
              {
                "X": 3.6048,
                "Y": 8.2273
              },
              {
                "X": 3.6048,
                "Y": 8.3933
              },
              {
                "X": 0.2366,
                "Y": 8.3933
              }
            ]
          },
          {
            "CellId": "34,1",
            "CellName": "B35",
            "Value": "",
            "Polygon": [
              {
                "X": 3.6048,
                "Y": 8.2273
              },
              {
                "X": 5.2889,
                "Y": 8.2273
              },
              {
                "X": 5.2889,
                "Y": 8.3933
              },
              {
                "X": 3.6048,
                "Y": 8.3933
              }
            ]
          },
          {
            "CellId": "34,2",
            "CellName": "C35",
            "Value": "",
            "Polygon": [
              {
                "X": 5.2889,
                "Y": 8.2273
              },
              {
                "X": 6.2845,
                "Y": 8.2273
              },
              {
                "X": 6.2845,
                "Y": 8.3933
              },
              {
                "X": 5.2889,
                "Y": 8.3933
              }
            ]
          },
          {
            "CellId": "34,3",
            "CellName": "D35",
            "Value": "",
            "Polygon": [
              {
                "X": 6.2845,
                "Y": 8.2273
              },
              {
                "X": 7.28,
                "Y": 8.2356
              },
              {
                "X": 7.28,
                "Y": 8.4016
              },
              {
                "X": 6.2845,
                "Y": 8.3933
              }
            ]
          },
          {
            "CellId": "34,4",
            "CellName": "E35",
            "Value": "",
            "Polygon": [
              {
                "X": 7.28,
                "Y": 8.2356
              },
              {
                "X": 8.2506,
                "Y": 8.2356
              },
              {
                "X": 8.2506,
                "Y": 8.4016
              },
              {
                "X": 7.28,
                "Y": 8.4016
              }
            ]
          },
          {
            "CellId": "35,0",
            "CellName": "A36",
            "Value": "Total Accumulated other comprehensive loss",
            "Polygon": [
              {
                "X": 0.2366,
                "Y": 8.3933
              },
              {
                "X": 3.6048,
                "Y": 8.3933
              },
              {
                "X": 3.6048,
                "Y": 8.5592
              },
              {
                "X": 0.2366,
                "Y": 8.5426
              }
            ]
          },
          {
            "CellId": "35,1",
            "CellName": "B36",
            "Value": "$ (5,074)",
            "Polygon": [
              {
                "X": 3.6048,
                "Y": 8.3933
              },
              {
                "X": 5.2889,
                "Y": 8.3933
              },
              {
                "X": 5.2889,
                "Y": 8.5592
              },
              {
                "X": 3.6048,
                "Y": 8.5592
              }
            ]
          },
          {
            "CellId": "35,2",
            "CellName": "C36",
            "Value": "$ (5,878)",
            "Polygon": [
              {
                "X": 5.2889,
                "Y": 8.3933
              },
              {
                "X": 6.2845,
                "Y": 8.3933
              },
              {
                "X": 6.2845,
                "Y": 8.5675
              },
              {
                "X": 5.2889,
                "Y": 8.5592
              }
            ]
          },
          {
            "CellId": "35,3",
            "CellName": "D36",
            "Value": "$ (851)",
            "Polygon": [
              {
                "X": 6.2845,
                "Y": 8.3933
              },
              {
                "X": 7.28,
                "Y": 8.4016
              },
              {
                "X": 7.28,
                "Y": 8.5675
              },
              {
                "X": 6.2845,
                "Y": 8.5675
              }
            ]
          },
          {
            "CellId": "35,4",
            "CellName": "E36",
            "Value": "$ (974)",
            "Polygon": [
              {
                "X": 7.28,
                "Y": 8.4016
              },
              {
                "X": 8.2506,
                "Y": 8.4016
              },
              {
                "X": 8.2506,
                "Y": 8.5758
              },
              {
                "X": 7.28,
                "Y": 8.5675
              }
            ]
          }
        ]
      }
    }
  ]
}
'
);

-- =====================================================
-- 7. Insert Pages (with Name = PgNo-{pageNumber})
-- =====================================================
INSERT INTO Pages (PageKey, PdfPageNumber, Name)
SELECT DISTINCT
    page.key AS PageKey,
    CAST(page.key AS INTEGER) AS PdfPageNumber,
    'PgNo-' || page.key AS Name
FROM RawJson
JOIN json_each(RawJson.json, '$.Pages') AS p
JOIN json_each(p.value) AS page;

-- =====================================================
-- 8. Insert Cells (derive ExcelRow / ExcelColumn)
-- =====================================================
INSERT OR IGNORE INTO Cells (
    PageID,
    CellID,
    CellName,
    Value,
    ExcelRow,
    ExcelColumn
)
SELECT
    pages.PageID,
    json_extract(cell.value, '$.CellId'),
    json_extract(cell.value, '$.CellName'),
    json_extract(cell.value, '$.Value'),
    CAST(substr(json_extract(cell.value, '$.CellId'),
         1, instr(json_extract(cell.value, '$.CellId'), ',') - 1) AS INTEGER),
    CAST(substr(json_extract(cell.value, '$.CellId'),
         instr(json_extract(cell.value, '$.CellId'), ',') + 1) AS INTEGER)
FROM RawJson
JOIN json_each(RawJson.json, '$.Pages') AS p
JOIN json_each(p.value) AS page
JOIN Pages pages
     ON pages.PageKey = page.key
JOIN json_each(page.value, '$.Cells') AS cell;


-- =====================================================
-- 9. Insert PolygonData (ordered points)
-- =====================================================
INSERT INTO PolygonData (CellPK, PointOrder, X, Y)
SELECT
    c.CellPK,
    poly.key + 1,
    json_extract(poly.value, '$.X'),
    json_extract(poly.value, '$.Y')
FROM RawJson
JOIN json_each(RawJson.json, '$.Pages') AS p
JOIN json_each(p.value) AS page
JOIN Pages pages
     ON pages.PageKey = page.key
JOIN json_each(page.value, '$.Cells') AS cell
JOIN Cells c
     ON c.PageID = pages.PageID
    AND c.CellID = json_extract(cell.value, '$.CellId')
JOIN json_each(cell.value, '$.Polygon') AS poly;

COMMIT;
