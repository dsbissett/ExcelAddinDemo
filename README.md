# Excel Sqlite Datasource Demo
- [`sql.js`](https://sql.js.org/) is used to create a sqlite database
- The database is base64 encoded and stored in the Excel document's `customXml`
- The database is only saved on `CREATE`, `INSERT`, `Update`, and `DELETE`
- Read-only queries are read from memory

# File Storage Support
- PDF files are compressed using [Brotli WASM](https://github.com/httptoolkit/brotli-wasm)
  ⚠️ There are huge performance issues while files are being compressed
- PDF files are base64 encoded and persisted in the Excel document as a `customXml` part
- Unclear on the limits, but tested with a 8mb PDF and a total of 8mb of PDF data
- PDf files are displayed as thumbnails on PDF Viewer tab using [pdf.js](https://github.com/mozilla/pdf.js)

# Queries and results
## Querying seed data:
```sql
SELECT
    p.PageKey,
    p.PdfPageNumber,
    c.CellID,
    c.CellName,
    c.Value,
    json_group_array(
        json_object(
            'X', pd.X,
            'Y', pd.Y
        )
        ORDER BY pd.PointOrder
  ) AS Polygon
FROM Pages p
JOIN Cells c
    ON c.PageID = p.PageID
JOIN PolygonData pd
    ON pd.CellPK = c.CellPK
WHERE p.PageKey = '1.1'
GROUP BY
    c.CellPK
ORDER BY
    c.ExcelRow,
    c.ExcelColumn;
```
Results
| PageKey | PdfPageNumber | CellID | CellName | Value | Polygon |
|--------|---------------|--------|----------|-------|---------|
| 1.1 | 1 | 0,0 | A1 |  | [{"X":0.1922,"Y":4.8417},{"X":3.1042,"Y":4.8417},{"X":3.1042,"Y":5.3304},{"X":0.1922,"Y":5.3138}] |
| 1.1 | 1 | 0,1 | B1 | First quarter ended March 31, | [{"X":3.1042,"Y":4.8417},{"X":8.2499,"Y":4.85},{"X":8.2499,"Y":5.1647},{"X":3.1042,"Y":5.1647}] |
| 1.1 | 1 | 1,1 | B2 | 2022 | [{"X":3.1042,"Y":5.1647},{"X":6.4795,"Y":5.1647},{"X":6.4795,"Y":5.3304},{"X":3.1042,"Y":5.3304}] |
| 1.1 | 1 | 1,3 | D2 | 2021 | [{"X":6.4795,"Y":5.1647},{"X":8.2499,"Y":5.1647},{"X":8.2499,"Y":5.3304},{"X":6.4795,"Y":5.3304}] |
| 1.1 | 1 | 2,0 | A3 | Primary aluminum | [{"X":0.1922,"Y":5.3138},{"X":3.1042,"Y":5.3304},{"X":3.1042,"Y":5.4878},{"X":0.1922,"Y":5.4878}] |
| 1.1 | 1 | 2,1 | B3 | $ | [{"X":3.1042,"Y":5.3304},{"X":5.8177,"Y":5.3304},{"X":5.8177,"Y":5.4878},{"X":3.1042,"Y":5.4878}] |
| 1.1 | 1 | 2,2 | C3 | 2,447 | [{"X":5.8177,"Y":5.3304},{"X":6.4795,"Y":5.3304},{"X":6.4795,"Y":5.4878},{"X":5.8177,"Y":5.4878}] |
| 1.1 | 1 | 2,3 | D3 | $ 1,727 | [{"X":6.4795,"Y":5.3304},{"X":8.2499,"Y":5.3304},{"X":8.2499,"Y":5.4961},{"X":6.4795,"Y":5.4878}] |
| 1.1 | 1 | 3,0 | A4 | Alumina | [{"X":0.1922,"Y":5.4878},{"X":3.1042,"Y":5.4878},{"X":3.1042,"Y":5.6452},{"X":0.1922,"Y":5.6452}] |
| 1.1 | 1 | 3,1 | B4 |  | [{"X":3.1042,"Y":5.4878},{"X":5.8177,"Y":5.4878},{"X":5.8177,"Y":5.6452},{"X":3.1042,"Y":5.6452}] |
| 1.1 | 1 | 3,2 | C4 | 850 | [{"X":5.8177,"Y":5.4878},{"X":6.4795,"Y":5.4878},{"X":6.4795,"Y":5.6452},{"X":5.8177,"Y":5.6452}] |
| 1.1 | 1 | 3,3 | D4 | 760 | [{"X":6.4795,"Y":5.4878},{"X":8.2499,"Y":5.4961},{"X":8.2499,"Y":5.6452},{"X":6.4795,"Y":5.6452}] |
| 1.1 | 1 | 4,0 | A5 | Energy | [{"X":0.1922,"Y":5.6452},{"X":3.1042,"Y":5.6452},{"X":3.1042,"Y":5.8026},{"X":0.1922,"Y":5.8026}] |
| 1.1 | 1 | 4,1 | B5 |  | [{"X":3.1042,"Y":5.6452},{"X":5.8177,"Y":5.6452},{"X":5.8177,"Y":5.8026},{"X":3.1042,"Y":5.8026}] |
| 1.1 | 1 | 4,2 | C5 | 41 | [{"X":5.8177,"Y":5.6452},{"X":6.4795,"Y":5.6452},{"X":6.4795,"Y":5.8026},{"X":5.8177,"Y":5.8026}] |
| 1.1 | 1 | 4,3 | D5 | 39 | [{"X":6.4795,"Y":5.6452},{"X":8.2499,"Y":5.6452},{"X":8.2499,"Y":5.8109},{"X":6.4795,"Y":5.8026}] |
| 1.1 | 1 | 5,0 | A6 | Bauxite | [{"X":0.1922,"Y":5.8026},{"X":3.1042,"Y":5.8026},{"X":3.1042,"Y":5.96},{"X":0.1922,"Y":5.96}] |
| 1.1 | 1 | 5,1 | B6 |  | [{"X":3.1042,"Y":5.8026},{"X":5.8177,"Y":5.8026},{"X":5.8177,"Y":5.9683},{"X":3.1042,"Y":5.96}] |
| 1.1 | 1 | 5,2 | C6 | 28 | [{"X":5.8177,"Y":5.8026},{"X":6.4795,"Y":5.8026},{"X":6.4795,"Y":5.9683},{"X":5.8177,"Y":5.9683}] |
| 1.1 | 1 | 5,3 | D6 | 52 | [{"X":6.4795,"Y":5.8026},{"X":8.2499,"Y":5.8109},{"X":8.2499,"Y":5.9683},{"X":6.4795,"Y":5.9683}] |
| 1.1 | 1 | 6,0 | A7 | Flat-rolled aluminum(1) | [{"X":0.1922,"Y":5.96},{"X":3.1042,"Y":5.96},{"X":3.1042,"Y":6.1256},{"X":0.1922,"Y":6.1256}] |
| 1.1 | 1 | 6,1 | B7 |  | [{"X":3.1042,"Y":5.96},{"X":5.8177,"Y":5.9683},{"X":5.8177,"Y":6.1256},{"X":3.1042,"Y":6.1256}] |
| 1.1 | 1 | 6,2 | C7 | — | [{"X":5.8177,"Y":5.9683},{"X":6.4795,"Y":5.9683},{"X":6.4795,"Y":6.1256},{"X":5.8177,"Y":6.1256}] |
| 1.1 | 1 | 6,3 | D7 | 320 | [{"X":6.4795,"Y":5.9683},{"X":8.2499,"Y":5.9683},{"X":8.2499,"Y":6.1339},{"X":6.4795,"Y":6.1256}] |
| 1.1 | 1 | 7,0 | A8 | Other(2) | [{"X":0.1922,"Y":6.1256},{"X":3.1042,"Y":6.1256},{"X":3.1042,"Y":6.283},{"X":0.1922,"Y":6.283}] |
| 1.1 | 1 | 7,1 | B8 |  | [{"X":3.1042,"Y":6.1256},{"X":5.8177,"Y":6.1256},{"X":5.8259,"Y":6.2913},{"X":3.1042,"Y":6.283}] |
| 1.1 | 1 | 7,2 | C8 | (73) | [{"X":5.8177,"Y":6.1256},{"X":6.4795,"Y":6.1256},{"X":6.4795,"Y":6.2913},{"X":5.8259,"Y":6.2913}] |
| 1.1 | 1 | 7,3 | D8 | (28) | [{"X":6.4795,"Y":6.1256},{"X":8.2499,"Y":6.1339},{"X":8.2499,"Y":6.2913},{"X":6.4795,"Y":6.2913}] |
| 1.1 | 1 | 8,0 | A9 |  | [{"X":0.1922,"Y":6.283},{"X":3.1042,"Y":6.283},{"X":3.1042,"Y":6.4487},{"X":0.1922,"Y":6.4404}] |
| 1.1 | 1 | 8,1 | B9 | $ | [{"X":3.1042,"Y":6.283},{"X":5.8259,"Y":6.2913},{"X":5.8259,"Y":6.4487},{"X":3.1042,"Y":6.4487}] |
| 1.1 | 1 | 8,2 | C9 | 3,293 | [{"X":5.8259,"Y":6.2913},{"X":6.4795,"Y":6.2913},{"X":6.4795,"Y":6.4487},{"X":5.8259,"Y":6.4487}] |
| 1.1 | 1 | 8,3 | D9 | $ 2,870 | [{"X":6.4795,"Y":6.2913},{"X":8.2499,"Y":6.2913},{"X":8.2499,"Y":6.457},{"X":6.4795,"Y":6.4487}] |

## Compression performance (per file)
```sql
SELECT
  DataFileID,
  FileName,
  XmlPartName,
  RawFileSize,
  CompressedFileSize,
  (RawFileSize - CompressedFileSize) AS BytesSaved,
  ROUND(((RawFileSize - CompressedFileSize) * 100.0) / NULLIF(RawFileSize, 0), 2) AS PercentSaved,
  ROUND(RawFileSize / 1024.0, 2) AS RawKB,
  ROUND(CompressedFileSize / 1024.0, 2) AS CompressedKB,
  ROUND((RawFileSize - CompressedFileSize) / 1024.0, 2) AS SavedKB,
  ROUND(RawFileSize / 1048576.0, 2) AS RawMB,
  ROUND(CompressedFileSize / 1048576.0, 2) AS CompressedMB,
  ROUND((RawFileSize - CompressedFileSize) / 1048576.0, 2) AS SavedMB
FROM DataFiles
ORDER BY DataFileID;
```

## Results

| TotalRawBytes | TotalCompressedBytes | TotalBytesSaved | OverallPercentSaved | TotalRawKB | TotalCompressedKB | TotalSavedKB | TotalRawMB | TotalCompressedMB | TotalSavedMB |
|--------------:|---------------------:|----------------:|--------------------:|-----------:|------------------:|-------------:|-----------:|------------------:|-------------:|
| 11,189,448 | 8,545,281 | 2,644,167 | 23.63 | 10,927.20 | 8,345.00 | 2,582.19 | 10.67 | 8.15 | 2.52 |

## Compression Performance (totals)
```sql
SELECT
  SUM(RawFileSize) AS TotalRawBytes,
  SUM(CompressedFileSize) AS TotalCompressedBytes,
  SUM(RawFileSize - CompressedFileSize) AS TotalBytesSaved,
  ROUND((SUM(RawFileSize - CompressedFileSize) * 100.0) / NULLIF(SUM(RawFileSize), 0), 2) AS OverallPercentSaved,
  ROUND(SUM(RawFileSize) / 1024.0, 2) AS TotalRawKB,
  ROUND(SUM(CompressedFileSize) / 1024.0, 2) AS TotalCompressedKB,
  ROUND(SUM(RawFileSize - CompressedFileSize) / 1024.0, 2) AS TotalSavedKB,
  ROUND(SUM(RawFileSize) / 1048576.0, 2) AS TotalRawMB,
  ROUND(SUM(CompressedFileSize) / 1048576.0, 2) AS TotalCompressedMB,
  ROUND(SUM(RawFileSize - CompressedFileSize) / 1048576.0, 2) AS TotalSavedMB
FROM DataFiles;
```
## Results

| DataFileID | FileName | XmlPartName | RawFileSize | CompressedFileSize | BytesSaved | PercentSaved | RawKB | CompressedKB | SavedKB | RawMB | CompressedMB | SavedMB |
|-----------:|----------|-------------|------------:|-------------------:|-----------:|-------------:|------:|-------------:|--------:|------:|-------------:|--------:|
| 1 | FBLU01046_FMCK2729000_1Q25_FBA_Blue_InspectionCheck_Dig.pdf | dataFile-fblu01046-fmck2729000-1q25-fba-blue-inspectioncheck-dig-pdf-ogio3s | 185228 | 92954 | 92274 | 49.82 | 180.89 | 90.78 | 90.11 | 0.18 | 0.09 | 0.09 |
| 2 | FBLU01046_FMCK2728000_1Q25_FBA_Gold_InspectionCheck_Dig.pdf | dataFile-fblu01046-fmck2728000-1q25-fba-gold-inspectioncheck-dig-pdf-i9w5gd | 208007 | 103770 | 104237 | 50.11 | 203.13 | 101.34 | 101.79 | 0.20 | 0.10 | 0.10 |
| 3 | Autostacker-Brochure-2020.pdf | dataFile-autostacker-brochure-2020-pdf-34y3jk | 8643638 | 7120525 | 1523113 | 17.62 | 8441.05 | 6953.64 | 1487.42 | 8.24 | 6.79 | 1.45 |
| 4 | Information Fixodrop ES_INT_0.pdf | dataFile-information-fixodrop-es-int-0-pdf-e7pvds | 2152575 | 1228032 | 924543 | 42.95 | 2102.12 | 1199.25 | 902.87 | 2.05 | 1.17 | 0.88 |

## To run
- `npm install`
- `npm run start -- desktop --app excel`

## Notes
- Click `Seed database` button to seed a new database into the new Excel sheet
- Run queries by entering sqlite sql and hitting run

<img width="1275" height="764" alt="image" src="https://github.com/user-attachments/assets/75250988-d0a1-42c6-bd71-feebae5423b2" />
