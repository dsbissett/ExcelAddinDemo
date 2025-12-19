/* global Office, Excel */

import { Component, NgZone } from "@angular/core";
import { NGXLogger } from "ngx-logger";
import { DataService } from "./services/data.service";
import { sql } from "@codemirror/lang-sql";
import seedScript from "../seed-script.sql";

@Component({
  selector: "app-root",
  templateUrl: "./app.component.html",
  styleUrls: ["./app.component.css"],
})
export class AppComponent {
  isReady = false;
  isExcelHost = false;
  isRunning = false;
  isCreatingDb = false;
  statusMessage = "";
  lastAddress = "";
  private readonly seedDisplayQuery = `SELECT
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
    c.ExcelColumn;`;
  private readonly seedPlaceholder = "Please seed the Sqlite database";
  sqlInput = this.seedDisplayQuery;
  sqlExtensions = [sql({ upperCaseKeywords: true })];
  queryResults: Array<{ columns: string[]; values: unknown[][] }> = [];
  hasDatabase = false;
  isSeeding = false;

  constructor(
    private zone: NgZone,
    private logger: NGXLogger,
    private dataService: DataService,
  ) {
    if (typeof Office !== "undefined" && Office.onReady) {
      Office.onReady((info) => {
        this.zone.run(() => {
          this.isExcelHost = info.host === Office.HostType.Excel;
          this.isReady = true;
          this.logger.debug("Office is ready");
          this.checkDatabasePresence();
        });
      });
    } else {
      this.zone.run(() => {
        this.isReady = true;
      });
    }
  }

  async run(): Promise<void> {
    if (!this.isExcelHost) {
      this.statusMessage = "Connect to Excel to run SQL.";
      return;
    }

    this.isRunning = true;
    this.statusMessage = "Running SQL...";

    try {
      const results = await this.dataService.execute(this.sqlInput);
      this.zone.run(() => {
        this.queryResults = results;
        this.statusMessage = results.length ? "Query completed." : "Query returned no results.";
        this.hasDatabase = true;
      });
    } catch (error) {
      console.error(error);
      this.zone.run(() => {
        this.statusMessage = `Error: ${error instanceof Error ? error.message : String(error)}`;
        this.queryResults = [];
      });
    } finally {
      this.zone.run(() => {
        this.isRunning = false;
      });
    }
  }

  async seedDatabase(): Promise<void> {
    if (!this.isExcelHost) {
      this.statusMessage = "Connect to Excel to seed the database.";
      return;
    }

    this.isSeeding = true;
    this.statusMessage = "Seeding database...";

    try {
      await this.dataService.seedDatabase(seedScript);
      this.zone.run(() => {
        this.hasDatabase = true;
        this.sqlInput = this.seedDisplayQuery;
        this.statusMessage = "Database seeded.";
      });
    } catch (error) {
      console.error(error);
      this.zone.run(() => {
        this.statusMessage = `Seed error: ${error instanceof Error ? error.message : String(error)}`;
      });
    } finally {
      this.zone.run(() => {
        this.isSeeding = false;
      });
    }
  }

  private async checkDatabasePresence(): Promise<void> {
    if (!this.isExcelHost) {
      this.hasDatabase = false;
      return;
    }

    try {
      const exists = await this.dataService.hasDatabase();
      this.zone.run(() => {
        this.hasDatabase = exists;
        if (!exists) {
          this.sqlInput = this.seedPlaceholder;
        } else if (!this.sqlInput.trim() || this.sqlInput === this.seedPlaceholder) {
          this.sqlInput = this.seedDisplayQuery;
        }
      });
    } catch (error) {
      this.logger.error("Failed to check database presence", error);
      this.zone.run(() => {
        this.hasDatabase = false;
      });
    }
  }

  async createDatabase(): Promise<void> {
    if (!this.isExcelHost) {
      this.statusMessage = "Connect to Excel to create the database.";
      return;
    }

    this.isCreatingDb = true;
    this.statusMessage = "Creating database...";

    try {
      await this.dataService.loadOrCreate();
      this.statusMessage = "Database is ready.";
      this.logger.info("Database created or loaded.");
    } catch (error) {
      this.logger.error("Failed to create database", error);
      this.statusMessage = `Database error: ${error instanceof Error ? error.message : String(error)}`;
    } finally {
      this.isCreatingDb = false;
    }
  }

}
