/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  console.log('office ready',info.host )
  if (info.host === Office.HostType.Outlook) {
    console.log('prepare taskpane') 
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  /**
   * Insert your Outlook code here
   */

 insertImage();
}

function insertImage() {
  const imageDataUrl = 'https://www.thewowstyle.com/wp-content/uploads/2015/01/images-of-nature-4.jpg'

  // Create an HTML image element
  const imgElement = `<img src="${imageDataUrl}" alt="Inserted Image" />`;

  // Insert the image HTML at the cursor position
  Office.context.mailbox.item.body.setSelectedDataAsync(
      imgElement,
      { coercionType: Office.CoercionType.Html },
      function(asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log("Image inserted successfully.");
          } else {
              console.error("Failed to insert image: " + asyncResult.error.message);
          }
      }
  );
}