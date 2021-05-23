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