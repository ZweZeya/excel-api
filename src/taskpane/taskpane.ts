/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// The initialize function must be run each time a new page is loaded
Office.onReady(() => {
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("run").onclick = run;
});

// export async function run() {
//   try {
//     await Excel.run(async (context) => {
//       /**
//        * Insert your Excel code here
//        */
//       const range = context.workbook.getSelectedRange();

//       // Read the range address
//       range.load("address");

//       // Update the fill color
//       range.format.fill.color = "yellow";

//       await context.sync();
//       console.log(`The range address was ${range.address}.`);
//     });
//   } catch (error) {
//     console.error(error);
//   }
// }

export async function run() {
  try {
    await Excel.run(async (ctx) => {
      // The top level object which contains related workbook objects such as worksheets, tables, and ranges
      const workbook = ctx.workbook;

      // Refreshes data connections associated to the workbook
      workbook.dataConnections.refreshAll();
    });
  } catch (err) {
    console.error(err);
  }
}
