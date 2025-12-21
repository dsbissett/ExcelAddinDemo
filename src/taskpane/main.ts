/* global Office */

import { platformBrowserDynamic } from "@angular/platform-browser-dynamic";
import { AppModule } from "./app.module";
import "bootstrap/dist/css/bootstrap.min.css";
import "./taskpane.css";

const bootstrap = () => {
  platformBrowserDynamic()
    .bootstrapModule(AppModule)
    .catch((err) => console.error(err));
};

if (typeof Office !== "undefined" && Office.onReady) {
  Office.onReady().then(bootstrap);
} else {
  bootstrap();
}
