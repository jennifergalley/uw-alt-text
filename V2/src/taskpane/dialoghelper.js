/*Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
4  See LICENSE in the project root for license information */

var dialog;

function dialogCallback(asyncResult) {
    if (asyncResult.status == "failed") {
        // In addition to general system errors, there are 3 specific errors for 
        // displayDialogAsync that you can handle individually.
        switch (asyncResult.error.code) {
            case 12004:
                Office.context.document.setSelectedDataAsync("Domain is not trusted");
                break;
            case 12005:
                Office.context.document.setSelectedDataAsync("HTTPS is required");
                break;
            case 12007:
                Office.context.document.setSelectedDataAsync("A dialog is already opened.");
                break;
            default:
                Office.context.document.setSelectedDataAsync(asyncResult.error.message);
                break;
        }
    }
    else {
        dialog = asyncResult.value;

        /*Messages are sent by developers programatically from the dialog using office.context.ui.messageParent(...)*/
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, messageHandler);

        /*Events are sent by the platform in response to user actions or errors. For example, the dialog is closed via the 'x' button*/
        dialog.addEventHandler(Office.EventType.DialogEventReceived, eventHandler);
    }
}

// Handles the user's button click on the dialog box
function messageHandler(arg) {
    // Close the dialog box either way
    dialog.close();

    // If the user clicked OK, open the task pane (in case it isn't open)
    if (arg.message == "open-alt-text") {
        // Get verbose logs
        OfficeExtension.config.extendedErrorLogging = true;

        // This is throwing an error! WHY WON'T THIS WORK???!!!
        Office.addin.showAsTaskpane();
    }
}

function eventHandler(arg) {
    // In addition to general system errors, there are 2 specific errors 
    // and one event that you can handle individually.
    switch (arg.error) {
        case 12002:
            Office.context.document.setSelectedDataAsync("Cannot load URL, no such page or bad URL syntax.");
            break;
        case 12003:
            Office.context.document.setSelectedDataAsync("HTTPS is required.");
            break;
        case 12006:
            // The dialog was closed, typically because the user the pressed X button.
            Office.context.document.setSelectedDataAsync("Dialog closed by user");
            break;
        default:
            Office.context.document.setSelectedDataAsync("Undefined error in dialog window");
            break;
    }
}

function openDialog() {
    Office.context.ui.displayDialogAsync(
        "https://localhost:3000/src/taskpane/dialog.html",
        { height: 15, width: 15}, 
        dialogCallback);
}

function openDialogAsIframe() {
    //IMPORTANT: IFrame mode only works in Online (Web) clients. Desktop clients (Windows, IOS, Mac) always display as a pop-up inside of Office apps. 
    Office.context.ui.displayDialogAsync(
        "https://localhost:3000/src/taskpane/dialog.html",
        { height: 15, width: 15, displayInIframe: true }, 
        dialogCallback);
}