# Module 2: Excel API

In this module, you will add javascript code specific to Excel. This code will add a named selection range to a worksheet. That range is named with a binding ID, and just like in the Word code can then be the target of navigation to find it again.

## API's Used In This Module

- Excel.run 
- getSelectedRange
- names.range (worksheet names collection)

## Step 1: The Excel.run block

All Excel-specific API's must be enclosed in a callback function passed to Excel.run(). Variables must be used inside the callback function scope or there are limitations and special code that must be used to make them accessible. So let's stick with keeping everything in scope. 

1. Find the code in taskerWeb/home.js that has this comment:

```js
                // ====== START ======
                // Workshop module 2 code goes here:
```

Notice that we are now inside a switch statement that identifies the "Host" of this add-in as Excel!

2. Add the following Word.run block:

```js
                Excel.run(function (context) {
                });
```

Everything we do to the worksheet using Excel javascript API's will happen inside this block.

3. Now add the code to call getSelectedRange:

```js
                Excel.run(function (context) {

                    var selectedRange = context.workbook.getSelectedRange();

                    // ... more code ...

                }).catch(function (error) {
                    console.log("Error: " + error);
                    if (error instanceof OfficeExtension.Error) {
                        console.log("Debug info: " + JSON.stringify(error.debugInfo));
                        showNotification('addFromSelectionAsync', "Debug info: " + JSON.stringify(error.debugInfo));
                    }
                });
```
At this point, we have a Range object that is again a proxy object and needs to be synchronized with the Excel backend. We need the address property of the Range object.

4. Following the getSelectedRange call, add a call to load() so we can queue a backend request to get the address property:

```js
                Excel.run(function (context) {

                    var selectedRange = context.workbook.getSelectedRange();

                    selectedRange.load('address');

                    // ... more code ...

```

5. Before we can use the range object to add the name in uniqueBindId to the names collection, we need to synchronize the proxy object with Excel on the backend. After the selectedRange.load() call, replace the "// ... more code ..." with the following context.sync() call:

```js
                    return context.sync().then(function () {
                        console.log(selectedRange.address);
                        showNotification('addFromSelectionAsync', selectedRange.address);
                        context.workbook.names.add(uniqueBindId, selectedRange, title);
                    });
```
In this block, we output some debug info so you can see what is being added. Then we call the add() method on the workbook's names collection. We can use this collection later to look up this named range and navigate to it when the user selects the corresponding task. 

The final Excel block should look like this: 

```js
                Excel.run(function (context) {
                    var selectedRange = context.workbook.getSelectedRange();
                    selectedRange.load('address');
                    return context.sync().then(function () {
                        console.log(selectedRange.address);
                        showNotification('addFromSelectionAsync', selectedRange.address);
                        context.workbook.names.add(uniqueBindId, selectedRange, title);
                    });
                }).catch(function (error) {
                    console.log("Error: " + error);
                    if (error instanceof OfficeExtension.Error) {
                        console.log("Debug info: " + JSON.stringify(error.debugInfo));
                        showNotification('addFromSelectionAsync', "Debug info: " + JSON.stringify(error.debugInfo));
                    }
                });
```

We are now done adding the Excel-specific code to our add-in. 

7. When you test the add-in inside Excel, create a task and the selected range will be named. You will also notice that selecting the task in the task list view causes navigation to the content control. Cool!
