# uw-alt-text

## Resources used to generate the code
https://docs.microsoft.com/en-us/office/dev/add-ins/develop/configure-your-add-in-to-use-a-shared-runtime



## How to open the task pane:
Open it when they open the document - https://docs.microsoft.com/en-us/office/dev/add-ins/develop/automatically-open-a-task-pane-with-a-document

Open it programmtically, i.e. when they insert an image - https://docs.microsoft.com/en-us/office/dev/add-ins/develop/show-hide-add-in. In order to use this functionality, the app has to use the shared runtime, which is why we use the guide above to create this version of the code.


## TODO - Jenny - commands
1. figure out how to get a command working. the last step of https://docs.microsoft.com/en-us/office/dev/add-ins/develop/configure-your-add-in-to-use-a-shared-runtime is not working for me. can I run the action command another way? or just figure out how to start running taskpane.open or whatever?