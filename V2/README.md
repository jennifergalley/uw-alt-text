# uw-alt-text

## Resources used to generate the code
https://docs.microsoft.com/en-us/office/dev/add-ins/develop/configure-your-add-in-to-use-a-shared-runtime



## How to open the task pane:
Open it when they open the document - https://docs.microsoft.com/en-us/office/dev/add-ins/develop/automatically-open-a-task-pane-with-a-document

Open it programmtically, i.e. when they insert an image - https://docs.microsoft.com/en-us/office/dev/add-ins/develop/show-hide-add-in. In order to use this functionality, the app has to use the shared runtime, which is why we use the guide above to create this version of the code.

# Word vs Powerpoint
If you want to switch between word and powerpoint, go to package.json and edit `"app-to-debug"` to be either `"word"` or `"powerpoint"`.
