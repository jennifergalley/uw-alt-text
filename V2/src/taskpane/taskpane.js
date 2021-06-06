/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";
import * as use from '@tensorflow-models/universal-sentence-encoder';


/* global document, Office */

const parser = new DOMParser();
let VISIBILITY_MODE = null; // Will update with listener
let universal_encoder = null;

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
    
    use.load().then(model => {  
      universal_encoder = model;
      
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
              
              // Analysis only happens when clicking on update alt text
              $("#top-paragraph-label").hide();

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
    });
  }
});

function dot(a, b){
  var hasOwnProperty = Object.prototype.hasOwnProperty;
  var sum = 0;
  for (var key in a) {
    if (hasOwnProperty.call(a, key) && hasOwnProperty.call(b, key)) {
      sum += a[key] * b[key]
    }
  }
  return sum
}

function similarity(a, b) {  
  var magnitudeA = Math.sqrt(dot(a, a));  
  var magnitudeB = Math.sqrt(dot(b, b));  
  if (magnitudeA && magnitudeB)  
    return dot(a, b) / (magnitudeA * magnitudeB);  
  else return false  
}

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

    return paragraphsText;
  });

}

export async function updateAltText() {
  $("#spinner-container").show();
  $("#submit-container").hide();
  $("#top-paragraph-label").hide();

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
      if (!inputAltText.endsWith(".")) {
        inputAltText += ".";
        document.getElementById("curr-alt-text-input").value = inputAltText;
      }
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

  try {
    let paragraphsText = await updateParagraphsSimilarity();
    const sentences = [document.getElementById("curr-alt-text-input").value].concat(paragraphsText);  
    universal_encoder.embed(sentences).then(async function(embeddings) {
      let embeds = embeddings.arraySync();
      let data = [];
      for (let i = 1; i < embeds.length; i++) {
        data.push([similarity(embeds[0], embeds[i]), paragraphsText[i - 1]]);
      }

      data.sort().reverse();
      
      console.log(data);
      const [topScore, topParagraph] = data[0]; // TODO: do something with the rest of paragraphs?
    
      if (topScore >= 0.9) {
        $("#top-paragraph-label").text(`The similarity of the alt text with the following paragraph is too high. Is the alt text redundant?: ${topParagraph}`);
        $("#top-paragraph-label").show();
      } else if (topScore <= 0.4) {
        $("#top-paragraph-label").text(`The similarity between the alt text and the text context seems low. Is the alt text relevant to the context?`);
        $("#top-paragraph-label").show();
      } else {
        //$("#top-paragraph-label").text(`DEBUG (delete this later) (${topScore}) => ${topParagraph}`);
        $("#top-paragraph-label").hide();
      }      
    });
  } catch (e) {
    console.error(e);
    $("#top-paragraph-label").hide();
  }
}