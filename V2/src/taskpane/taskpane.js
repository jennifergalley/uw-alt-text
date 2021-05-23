/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint || info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("action").onclick = action;
    document.getElementById("hide").onclick = hide;

    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, function(eventArgs) {
      eventArgs.document.getSelectedDataAsync(
        Office.CoercionType.Ooxml, // coercionType
        function(result) {
          console.log("HERE!");
          // TODO: do some decent XML parsing
          let m = result.value.match(/descr="[a-zA-z0-9\s]+"/);
          if (!m || m.length == 0) {
            console.log("no alt text");
          } else {
            console.log(m[0]);
            // TODO: use the decen XML parsing from above to put the 'right' alt text
            let newxml = result.value.replace('descr="dog"', 'descr="cat"');
            eventArgs.document.setSelectedDataAsync(newxml, { coercionType: Office.CoercionType.Ooxml }, function (asyncResult) {
              console.log("Done!");
              console.log(asyncResult);
              if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error(asyncResult.error.message);
              }
            });
          }
        }
      );
    });
  }
});

export async function run() {
  /**
   * Insert your PowerPoint code here
   */
  const options = { coercionType: Office.CoercionType.Text };

  await Office.context.document.setSelectedDataAsync(" ", options);
  await Office.context.document.setSelectedDataAsync("Hello World2!", options);
}

function hide2() {
  Office.addin.hide();
}

