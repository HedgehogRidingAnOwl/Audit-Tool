/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

import data from "./test.json";

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    var hr = document.createElement("div");
    hr.innerHTML = `
      <label for="target-condition">Target Condition:</label>
      <select id="target-condition">
        <option value="Pursuant to IDW PS 981 (4) and pursuant to AT 4.2 (2)  MaRisk Management has to define a risk strategy that is consistent with the Business Strategy and the risks resulting therefrom. ">Pursuant to IDW PS 981 (4) and pursuant to AT 4.2 (2)  MaRisk Management has to define a risk strategy that is consistent with the Business Strategy and the risks resulting therefrom. </option>
        <option value="Pursuant to AT 4.3 MaRisk a risk management system should be implemented consisting of policies and procedures, risk culture and risk monitoring functions depending on the riskiness of the business transacted. ">Pursuant to AT 4.3 MaRisk a risk management system should be implemented consisting of policies and procedures, risk culture and risk monitoring functions depending on the riskiness of the business transacted. </option>
        <option value="OPTION C">OPTION C</option>
        <option value="OPTION D">OPTION D</option>
      </select>`;
    document.body.insertBefore(hr, document.getElementById("actual"));

    //document.getElementById("input-form").style.display = "block";
    document.getElementById("accept").onclick = function() {
      createTable();
    };

  }
});

function createTable() {
  const targetCondition = document.getElementById("target-condition").value;
  const actualSituation = document.getElementById("actual-situation").value;
  const recommendation = document.getElementById("recommendation").value;
  const implementation = document.getElementById("implementation").value;
  const justification = document.getElementById("justification").value;

  const currentDate = new Date();
  const formattedDate = ("0" + currentDate.getDate()).slice(-2) + "/" + ("0" + (currentDate.getMonth() + 1)).slice(-2) + "/" + currentDate.getFullYear();

  Word.run(function(context) {
    const body = context.document.body;
    // Create a table with the values provided.
    const table = body.insertTable(4, 3, Word.InsertLocation.end, [
      ["Target Condition", targetCondition, ""],
      ["Actual Situation Prompt", actualSituation, ""],
      ["Recommendation", recommendation, ""],
      ["Implementation", implementation, formattedDate],
    ]);
    // Table formatting so only last row as 3 columns
    table.mergeCells(0, 1, 0, 2);
    table.mergeCells(1, 1, 1, 2);
    table.mergeCells(2, 1, 2, 2);
    // Date cell gets right justified formatting
    table.getCell(3, 2).horizontalAlignment = "Right";
    table.getCell(3, 2).verticalAlignment = "Bottom";

    // Add justification header and text
    body.insertParagraph("Justification:", Word.InsertLocation.end).font.bold = true;
    body.insertParagraph(justification, Word.InsertLocation.end).font.bold = false;

    //var mydata = JSON.parse(test);
    body.insertParagraph(data.dog, Word.InsertLocation.end);

    return context.sync();
  }).catch(function(error) {
    console.log(error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}