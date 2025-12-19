# Excel Sqlite Datasource Demo
- [`sql.js`](https://sql.js.org/) is used to create a sqlite database
- The database is base64 encoded and stored in the Excel document's `customXml`
- The database is only saved on `CREATE`, `INSERT`, `Update`, and `DELETE`
- Read-only queries are read from memory

## To run
- `npm install`
- `npm run start -- desktop --app excel`

## Notes
- Click `Seed database` button to seed a new database into the new Excel sheet
- Run queries by entering sqlite sql and hitting run

<img width="1527" height="824" alt="demo" src="https://github.com/user-attachments/assets/7554ddff-efc8-430f-86f3-cfe7f1197702" />
