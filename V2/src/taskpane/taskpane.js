/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

/* global document, Office */

const parser = new DOMParser();

/*
Gets docPr tag, see: https://c-rex.net/projects/samples/ooxml/e1/Part4/OOXML_P4_DOCX_docPr_topic_ID0ES32OB.html?hl=docpr
docPr contains the attribute "descr" which Word uses to store the alt text.
xmlStr Is the xlm string returned by the office api.
*/
function getDocPr(xmlStr) {
  let xml = parser.parseFromString(xmlStr, "application/xml");
          
  let tags = xml.getElementsByTagName("wp:docPr");
  if (tags.length > 1) {
    console.warn("Selected object with more than one cNvPr tag " + tags.length);
  }

  if (tags.length == 0) {
    console.log("Element may not have an alt text");
    return [null, xml];
  }
  
  return [tags[0], xml];
}

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint || info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("update-alt-text-button").onclick = updateAltText;
    // document.getElementById("dialog").onclick = openDialog;

    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, function(eventArgs) {
      eventArgs.document.getSelectedDataAsync(
        Office.CoercionType.Ooxml, // coercionType - only applies to Word
        function(result) {
          let altTextField = document.getElementById("curr-alt-text-input");
          altTextField.value = "";
          const [tag, xml] = getDocPr(result.value);

          // If it is an image selected
          if (tag != null) {
            $("#update-altext-container").show();
            let descr = tag.getAttribute("descr"); // the current alt text
            let name = tag.getAttribute("name");

            // If there is currently alt text
            if (descr) {
              altTextField.value = descr;
            } else {
              // If not, prompt the user to add it
              openDialog();
            }
          } else {
            $("#update-altext-container").hide();
          }
        }
      );
    });
  }
});

export async function updateAltText() {
  $("#spinner-container").show();
  $("#submit-container").hide();
  Office.context.document.getSelectedDataAsync(
    Office.CoercionType.Ooxml, // coercionType
    function(result) {
      const [tag, xml] = getDocPr(result.value);

      if (tag == null) {
        $("#spinner-container").hide();
        $("#submit-container").show();
        console.warn("This should not be null, perhaps the select element was not an image?")
        return;
      }

      let inputAltText = document.getElementById("curr-alt-text-input").value;
      tag.setAttribute("descr", inputAltText);
      let newxml = (new XMLSerializer()).serializeToString(xml);
      Office.context.document.setSelectedDataAsync(newxml, { coercionType: Office.CoercionType.Ooxml }, function (asyncResult) {
        console.log("Done!");
        console.log(asyncResult);
        $("#spinner-container").hide();
        $("#submit-container").show();
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error(asyncResult.error.message);
        }
      });
    }
  );
}