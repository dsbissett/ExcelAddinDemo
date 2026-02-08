/* global Office */

import { Component, NgZone, OnInit } from "@angular/core";
import { NGXLogger } from "ngx-logger";
import { DataService } from "../services/data.service";
import { sql, SQLite } from "@codemirror/lang-sql";

type QueryResult = { columns: string[]; values: unknown[][] };

@Component({
  selector: "app-admin-panel",
  templateUrl: "./admin-panel.html",
  styleUrls: ["./admin-panel.component.css"],
})
export class AdminPanelComponent implements OnInit {
  isReady = false;
  isExcelHost = false;
  isLoadingTables = false;
  isRunningQuery = false;
  hasDatabase = false;
  hasInitialized = false;
  tables: string[] = [];
  selectedTable: string | null = null;
  sqlInput = "";
  queryResults: QueryResult[] = [];
  statusMessage = "";
  tableCount = 0;
  recordCount = 0;
  readonly previewLimit = 100;
  sqlExtensions = [sql({ upperCaseKeywords: true, dialect: SQLite })];

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
          this.triggerInitialLoad();
        });
      });
    } else {
      this.isReady = true;
    }
  }

  ngOnInit(): void {
    this.triggerInitialLoad();
  }

  get queryPlaceholder(): string {
    return this.selectedTable
      ? `SELECT * FROM ${this.formatTableName(this.selectedTable)} LIMIT ${this.previewLimit};`
      : "SELECT * FROM table_name LIMIT 100;";
  }

  get emptyTablesMessage(): string {
    if (!this.isReady) {
      return "Waiting for Office to load.";
    }
    if (!this.isExcelHost) {
      return "Connect to Excel to view tables.";
    }
    if (!this.hasDatabase) {
      return "No database found. Seed the database first.";
    }
    return "No tables found.";
  }

  get statusText(): string {
    if (this.statusMessage) {
      return this.statusMessage;
    }
    if (!this.isReady) {
      return "Waiting for Office to load.";
    }
    if (!this.isExcelHost) {
      return "Connect to Excel to view data.";
    }
    if (!this.hasDatabase) {
      return "No database found.";
    }
    return "Connected to database.";
  }

  async refreshTables(): Promise<void> {
    if (!this.isReady) {
      this.statusMessage = "Waiting for Office to load.";
      return;
    }
    if (!this.isExcelHost) {
      this.statusMessage = "Connect to Excel to view tables.";
      return;
    }

    this.zone.run(() => {
      this.isLoadingTables = true;
      this.statusMessage = "Loading tables...";
    });

    try {
      const state = await this.dataService.getDatabaseState();
      this.zone.run(() => {
        this.hasDatabase = state.hasDatabase;
      });

      if (!state.hasDatabase) {
        this.zone.run(() => {
          this.tables = [];
          this.selectedTable = null;
          this.queryResults = [];
          this.tableCount = 0;
          this.recordCount = 0;
          this.statusMessage = "No database found. Seed the database to continue.";
        });
        return;
      }

      const tables = await this.dataService.listTables();
      const counts = await Promise.all(
        tables.map((table) => this.dataService.getTableRowCount(table)),
      );
      const totalRecords = counts.reduce((sum, count) => sum + count, 0);

      this.zone.run(() => {
        this.tables = tables;
        this.tableCount = tables.length;
        this.recordCount = totalRecords;
      });

      if (tables.length) {
        const nextSelection =
          this.selectedTable && tables.includes(this.selectedTable)
            ? this.selectedTable
            : tables[0];
        await this.selectTable(nextSelection);
      } else {
        this.zone.run(() => {
          this.selectedTable = null;
          this.sqlInput = "";
          this.queryResults = [];
          this.statusMessage = "No tables found.";
        });
      }
    } catch (error) {
      this.logger.error("refreshTables: failed", error);
      this.zone.run(() => {
        this.statusMessage = `Unable to load tables: ${
          error instanceof Error ? error.message : String(error)
        }`;
      });
    } finally {
      this.zone.run(() => {
        this.isLoadingTables = false;
      });
    }
  }

  async selectTable(tableName: string): Promise<void> {
    if (!tableName) {
      return;
    }

    this.zone.run(() => {
      this.selectedTable = tableName;
      this.sqlInput = `SELECT * FROM ${this.formatTableName(tableName)} LIMIT ${this.previewLimit};`;
      this.statusMessage = `Loading ${tableName}...`;
    });

    await this.loadTablePreview(tableName);
  }

  async runQuery(): Promise<void> {
    if (!this.isReady) {
      this.statusMessage = "Waiting for Office to load.";
      return;
    }
    if (!this.isExcelHost) {
      this.statusMessage = "Connect to Excel to run SQL.";
      return;
    }

    const sqlText = this.sqlInput.trim();
    if (!sqlText) {
      this.statusMessage = "Enter a SQL query to run.";
      return;
    }

    this.zone.run(() => {
      this.isRunningQuery = true;
      this.statusMessage = "Running SQL...";
    });

    try {
      const results = await this.dataService.execute(sqlText);
      this.zone.run(() => {
        this.queryResults = results;
        this.statusMessage = results.length ? "Query completed." : "Query returned no results.";
      });
    } catch (error) {
      this.logger.error("runQuery: failed", error);
      this.zone.run(() => {
        this.queryResults = [];
        this.statusMessage = `Query error: ${error instanceof Error ? error.message : String(error)}`;
      });
    } finally {
      this.zone.run(() => {
        this.isRunningQuery = false;
      });
    }
  }

  private async loadTablePreview(tableName: string): Promise<void> {
    this.zone.run(() => {
      this.isRunningQuery = true;
    });

    try {
      const preview = await this.dataService.previewTable(tableName, this.previewLimit);
      this.zone.run(() => {
        this.queryResults = preview.columns.length ? [preview] : [];
        this.statusMessage = `Loaded ${tableName}.`;
      });
    } catch (error) {
      this.logger.error("loadTablePreview: failed", error);
      this.zone.run(() => {
        this.queryResults = [];
        this.statusMessage = `Unable to load ${tableName}: ${
          error instanceof Error ? error.message : String(error)
        }`;
      });
    } finally {
      this.zone.run(() => {
        this.isRunningQuery = false;
      });
    }
  }

  private formatTableName(tableName: string): string {
    const escaped = tableName.replace(/"/g, "\"\"");
    return `"${escaped}"`;
  }

  private triggerInitialLoad(): void {
    if (this.hasInitialized || !this.isReady || !this.isExcelHost) {
      return;
    }
    this.hasInitialized = true;
    void this.refreshTables();
  }
}
