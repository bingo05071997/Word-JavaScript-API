/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  return Word.run(async (context) => {
    /**
     * Insert your Word code here
     */

    // insert a paragraph at the end of the document.
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items");

    await context.sync();

    if (paragraphs.items.length === 0) {
      context.document.body.insertParagraph("No content found.", Word.InsertLocation.end);
      return;
    }

    // Get paragraph object
    const firstParagraph = paragraphs.items[0];
    // Load text content
    firstParagraph.load("text");
    await context.sync();

    const words = firstParagraph.text.split(/\s+/);

    if (words.length < 3) {
      context.document.body.insertParagraph("Not enough words for analysis.", Word.InsertLocation.end);
      return;
    }

    const firstWordRange = firstParagraph.search(words[0], { matchCase: false });
    const secondWordRange = firstParagraph.search(words[1], { matchCase: false });
    const thirdWordRange = firstParagraph.search(words[2], { matchCase: false });

    // Load formatting properties before accessing them
    firstWordRange.load("items/font/bold");
    secondWordRange.load("items/font/underline");
    thirdWordRange.load("items/font/size");

    await context.sync();

    // Check if the words were found before accessing properties
    const isFirstWordBold = firstWordRange.items.length > 0 ? firstWordRange.items[0].font.bold : "Not found";
    const isSecondWordUnderlined = secondWordRange.items.length > 0 ? secondWordRange.items[0].font.underline : "Not found";
    const thirdWordFontSize = thirdWordRange.items.length > 0 ? thirdWordRange.items[0].font.size : "Not found";


    // Insert the results in the Word document
    context.document.body.insertParagraph(
      `First word is bold: ${isFirstWordBold}\n` +
        `Second word is underlined: ${isSecondWordUnderlined}\n` +
        `Font size of third word: ${thirdWordFontSize}`,
      Word.InsertLocation.end
    );

    await context.sync();
  });
}
