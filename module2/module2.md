# Module 2: Excel API

In this module, you will add javascript code specific to Excel. This code will load the worksheets collection of the workbook and add an onChanged event handler for each sheet. This will allow tasker to monitor changes to any cells and update tasks that reference those cells, notifying the user that the task may not be current.

## API's Used In This Module

- Excel.run 
- context.workbook.worksheets
- worksheet.onChanged

## Step 1: The Excel.run block

All Excel-specific API's must be enclosed in a callback function passed to Excel.run(). Variables must be used inside the callback function scope or there are limitations and special code that must be used to make them accessible. So let's stick with keeping everything in scope. 

1. Find the code in taskerWeb/home.js that has this comment:

```js
    // ====== START ======
    // Workshop module 2 code goes here:
```

and add the following: 

```js
        // Only move forward if we're in Excel
        if (Office.context.host == Office.HostType.Excel)
```

Notice that we are now inside a block that identifies the "Host" of this add-in as Excel!

2. Add the following Excel.run block:

```js
    Excel.run(function (context) {
    });
```

Everything we do to the worksheet using Excel javascript API's will happen inside this block.

3. Now add the code to get the worksheets collection of this workbook:

```js
    Excel.run(function (context) {

    var worksheets = context.workbook.worksheets;
    worksheets.load('items');

        // ... more code ...

}).catch(function (event) {
    console.log("Event register failed:" + event.message + ".");
});
```
At this point, we have a worksheets collection that is a proxy object and needs to be synchronized with the Excel backend. We need the items property of the worksheets collection so we've added the load() call to queue a batch request to fill it.


4. Before we can use the worksheets collection to register our event handlers, we need to synchronize the proxy object with Excel on the backend. Now, replace the "// ... more code ..." with the following context.sync() call:

```js
        return context.sync().then(function () {
            // ... register handler here ...
        });
```
5. Now replace the "// ... register handler here ..." comment with the following block: 

```js
        for (var i = 0; i < worksheets.items.length; i++) {
            console.log(worksheets.items[i].name);
            console.log(worksheets.items[i].index);
            worksheets.items[i].onChanged.add(handleSheetChange);
        }
        return context.sync()
            .then(function () {
                console.log("Event handler successfully registered for onChanged event for all worksheets.");
            });
```
In this block, we output some debug info so you can see the worksheet being worked on. Then we call the onChanged.add() method on each worksheet. The subsequent context.sync() call will make the event handler hookup effective to Excel's backend so it can notify us not only what happens here in this user session but also co-authoring remote sessions.

The final registerExcelEvents function should look like this: 

```js
    // ====== START ======
    // Workshop module 2 code goes here:

    // Only move forward if we're in Excel
    if (Office.context.host == Office.HostType.Excel)
        Excel.run(function (context) {
            var worksheets = context.workbook.worksheets;
            worksheets.load('items');
            return context.sync().then(function () {
                for (var i = 0; i < worksheets.items.length; i++) {
                    console.log(worksheets.items[i].name);
                    console.log(worksheets.items[i].index);
                    worksheets.items[i].onChanged.add(handleSheetChange);
                }
                return context.sync()
                    .then(function () {
                        console.log("Event handler successfully registered for onChanged event for all worksheets.");
                    });
            });
        }).catch(function (event) {
            console.log("Event register failed:" + event.message + ".");
        });

    // ===== END =====
```

We are now done adding the Excel-specific code to our add-in. 

6. Review the next function called handleSheetChange() to see what we do when the event fires. There are a lot of cool things that could be done here. In our sample implementation, we check to see that the range that was updated intersects with one of the ranges corresponding to a task. If so, then we mark that task as "dirty" and depending on the origin of the update ("Local" or this session, or "Remote" meaning another user session), we color the background yellow or red to alert the user that the task's relevant cell's have been modified.

7. When you test the add-in inside Excel, create a task and the selected range will be named. You will also notice that selecting the task in the task list view causes navigation to the content control. But now, modify a cell in the range you selected when creating the task. Notice it turns yellow in the list. If you have another user in your tenant, log in, open the workbook and modify that same range. Notice now the task back in your first session turns red. Cool!
