/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

var _count=0;

// /**
//  * Does a thing.
//  * @param event {Office.AddinCommands.Event}
//  */
function action(event) {
  // Your code goes here.
  _count++;
  // Office.addin.hide();

  // Office.addin.showAsTaskpane();
  document.getElementById("run").textContent="Go"+_count;

  // Be sure to indicate when the add-in command function is complete.
  // event.completed();
}

async function hide() {
  await Office.addin.hide();
}

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal();

// the add-in command functions need to be available in global scope
g.action = action;
g.hide = hide;