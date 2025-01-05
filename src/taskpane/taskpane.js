/* global Word console */

export async function insertText(text) {
  // Write text to the document.
  try {
    await Word.run(async (context) => {
      let body = context.document.body;
      body.insertParagraph(text, Word.InsertLocation.end);
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

export async function insertTable(targetCondition, actualSituation, recommendation, implementation, justification, image) {
  const currentDate = new Date();
  const formattedDate = ("0" + currentDate.getDate()).slice(-2) + "/" + ("0" + (currentDate.getMonth() + 1)).slice(-2) + "/" + currentDate.getFullYear();

  Word.run(async function(context) {
    //Alibi Paragraph for Table
    var selectionRange = context.document.getSelection();
    var paragraph = selectionRange.insertParagraph("", "Before");

    // Create a table with the values provided.
    const table = paragraph.insertTable(4, 3, "Before", [
      ["Target Condition", targetCondition, ""],
      ["Actual Situation Prompt", actualSituation, ""],
      ["Recommendation", recommendation, ""],
      ["Implementation", implementation, formattedDate],
    ]);
    // Table formatting so only last row as 3 columns
    table.mergeCells(0, 1, 0, 2);
    table.mergeCells(2, 1, 2, 2);
    // Date and traffic cell gets right justified formatting
    table.getCell(3, 2).horizontalAlignment = "Right";
    table.getCell(3, 2).verticalAlignment = "Bottom";
    table.getCell(1, 2).horizontalAlignment = "Right";
    table.getCell(1, 2).verticalAlignment = "Bottom";
    // Add traffic light
    if (image) {
      try {
        table.getCell(1, 2).body.insertInlinePictureFromBase64(image, "End");
        await context.sync();
      } catch (error) {
        console.log("Error inserting image: " + error.message);
      }
    }
    
    // Add justification
    paragraph.insertBreak("Line", "Before")
    paragraph.insertParagraph("Justification:", Word.InsertLocation.before).font.bold = true;

    return await context.sync();
  }).catch(function(error) {
    console.log(error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

export async function insertBullet(bullets) {
  Word.run(async function(context) {
    //Alibi Paragraph for Insertion
    var selectionRange = context.document.getSelection();
    let first = true;
    
    let list = undefined;
    // Insert bullet points
    bullets.forEach(bullet => {
      if(first) {
        const paragraph = selectionRange.insertParagraph(bullet.title, "Before");
        list = paragraph.startNewList();
        paragraph.listItem.level = 0;
        first = false;
      }
      else {
        let bulletParagraph = list.insertParagraph(bullet.title, "End");
        bulletParagraph.listItem.level = 0;
      }
      bullet.items.forEach(item => {
        let itemParagraph = list.insertParagraph(item, "End");
        itemParagraph.listItem.level = 1;
      });
    });

    return await context.sync();
  }).catch(function(error) {
    console.log(error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}
