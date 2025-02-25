/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      /** Ajuster la mise en page de la feuille */
      let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

      // Set print area for selectedSheet to range "A:K"
      selectedSheet.pageLayout.setPrintArea("A:K");
      // Set ExcelScript.PageOrientation.landscape orientation for selectedSheet
      selectedSheet.pageLayout.orientation = Excel.PageOrientation.landscape;
      // Répéter seulement la rangée 5 sur toutes les pages
      selectedSheet.pageLayout.setPrintTitleRows("$5:$5");
      // Set Letter paperSize for selectedSheet
      selectedSheet.pageLayout.paperSize = Excel.PaperType["letter"];
      // Set FitAllColumnsOnOnePage scaling for selectedSheet
      selectedSheet.pageLayout.zoom = { horizontalFitToPages: 1, verticalFitToPages: 0, scale: null };

      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}
