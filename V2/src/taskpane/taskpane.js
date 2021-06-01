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
let VISIBILITY_MODE = null; // Will update with listener

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

    Office.addin.onVisibilityModeChanged(function(message) {
      VISIBILITY_MODE = message.visibilityMode;
    });

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
            } 
            // Initially, VISIBILITY_MODE will be null, the first time it changes
            // it will be when the task pane is hidden, so is safe to assume it is open
            // when VISIBILITY_MODE == null the taskpane is open.
            else if (VISIBILITY_MODE != null && VISIBILITY_MODE != "Taskpane") {
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


// This function should get the paragraphs and do whatever we want with them:
// - compute similarity between current alt text and the whole doc (where to get this from???)
// - compute similarity between current alt text and each paragraph?
export async function updateParagraphsSimilarity() {
  return await Word.run(async (context) => {
    let paragraphs = context.document.body.paragraphs;
    paragraphs.load("text");
    await context.sync();

    let paragraphsText = [];
    paragraphs.items.forEach((item) => {
      let paragraph = item.text.trim();
      if (paragraph && paragraph.length > 0) {
          paragraphsText.push(paragraph);
      }
    });

    let body = {paragraphs: paragraphsText, alttext: document.getElementById("curr-alt-text-input").value};

    await context.sync(); // TODO: is this neccesary?
    return await fetch("http://localhost:5001/paragraph-similarity", {
          method: 'POST',
          headers: {'Content-Type': 'application/json', 'Access-Control-Allow-Origin':'*'},
          body: JSON.stringify(body)
      });
  });

}

export async function updateAltText() {
  $("#spinner-container").show();
  $("#submit-container").hide();
  
  try {
    updateParagraphsSimilarity().then(async function(response) {
      if (response.ok) {
        const data = await response.json(); // .sims is a sorted list of (similarity, paragraph)
        console.log(data);
        const [topScore, topParagraph] = data.sims[0]; // TODO: do something with the rest of paragraphs?
      
        if (topScore >= 0.9) {
          $("#top-paragraph-label").text(`The similarity (${topScore}) of this with the following paragraph is too high. Is this redundant?: ${topParagraph}`);
        } else if (topScore <= 0.1) {
          $("#top-paragraph-label").text(`The max similarity (${topScore}) seems too low, is the alt text relevant to the context?`);
        } else {
          $("#top-paragraph-label").text(`DEBUG (delete this later) (${topScore}) => ${topParagraph}`);
        }
      } else {
        $("#top-paragraph-label").hide();
      }      
    });
  } catch (e) {
    console.error(e);
    $("#top-paragraph-label").hide();
  }

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