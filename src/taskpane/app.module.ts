import { NgModule } from "@angular/core";
import { BrowserModule } from "@angular/platform-browser";
import { FormsModule } from "@angular/forms";
import { LoggerModule, NgxLoggerLevel } from "ngx-logger";
import { CodeEditorModule } from "@acrodata/code-editor";

import { AppComponent } from "./app.component";

@NgModule({
  declarations: [AppComponent],
  imports: [
    BrowserModule,
    FormsModule,
    CodeEditorModule,
    LoggerModule.forRoot({
      level: NgxLoggerLevel.TRACE,
      serverLogLevel: NgxLoggerLevel.OFF,
      disableConsoleLogging: false,
    }),
  ],
  bootstrap: [AppComponent],
})
export class AppModule {}
