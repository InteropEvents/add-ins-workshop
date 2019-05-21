# Module 1: Word API

In this module, you will add JavaScript code specific to Word API for Office Add-ins. This code will add a content control to a selection in the Word edit session for a document. That content control is named with a binding ID and can then be the target of navigation to find it again.

## APIs Used In This Module

- [Word.run](https://docs.microsoft.com/en-us/office/dev/add-ins/reference/overview/word-add-ins-reference-overview#running-word-add-ins) 
- [addFromSelectionAsync](https://docs.microsoft.com/en-us/javascript/api/office/office.bindings?view=office-js#addfromselectionasync-bindingtype--callback-)
- [getSelection](https://docs.microsoft.com/en-us/javascript/api/word/word.document?view=office-js#getselection--)
- [range](https://docs.microsoft.com/en-us/javascript/api/word/word.range?view=office-js)
- [range.insertContentControl](https://docs.microsoft.com/en-us/javascript/api/word/word.range?view=office-js#insertcontentcontrol--)

## Step 1: The `Word.run` block

All Word-specific APIs must be enclosed in a callback function passed to `Word.run()`. Variables must be used inside the callback function scope or there are limitations and special code that must be used to make them accessible. So let's stick with keeping everything in scope. 

1. Find the code in taskerWeb/home.js that has this comment:

```js
                // ====== START ======
                // Workshop module 1 code goes here:
```

Notice that we are now inside a switch statement that identifies the "Host" of this add-in as Word!

2. **Add the following `Word.run` block:**

```js
                Word.run(function (context) {
                });
```

**Everything we do to the document using Word JavaScript APIs will happen inside this block.**

3. Now add the code to call `addFromSelectionAsync` in the `Word.run` block from the previous step so it looks like this:

```js
                Word.run(function (context) {

                    // Call addFromSelectionAsync to get the selection and add a binding in the document that we use to
                    // navigate to the selection again.
                    Office.context.document.bindings.addFromSelectionAsync(Office.BindingType.Text, { id: uniqueBindId }, function (asyncResult) {
                        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                            showNotification('addFromSelectionAsync', 'Action failed. Error: ' + asyncResult.error.message);
                        } else {
                            showNotification('addFromSelectionAsync', 'Added new binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
                        }
                    });

                // Word.run
                }).catch(function (error) {

                    var errormsg = 'Error: ' + JSON.stringify(error);
                    console.log(errormsg);
                    if (error instanceof OfficeExtension.Error) {
                        var errormsg = errormsg + ', Debug info: ' + JSON.stringify(error.debugInfo);
                        showNotification('addFromSelectionAsync', errormsg);
                        console.log(errormsg);
                    }

                });

```
At this point, we have made a binding, named with a previously computed value (`uniqueBindId`), which we can later use to navigate in the document. We also added a `catch` call on the result of `Word.run` so we know if something went wrong in any of the code in `Word.run`.

4. **Following the `addFromSelectionAsync` call in the previous step, add a call to `getSelection`** so we can create a content control around it.

```js
                    // Queue a command to get the current selection and then
                    // create a proxy range object with the results.
                    range = context.document.getSelection();
```

5. When an Office.js object is handed to us from an API call, this is usually a proxy object in which property values are not yet populated fully. We must call the object's `load()` method to queue a command to Office to load the property values. Let's **follow this call with a call to load the "style" property on the selection.** This will allow us to keep the same style when inserting the content control around the text.

```js
                    range.load("style");
```

6. Before we can use the `range` object to create a `ContentControl` and insert it into the document, we need to synchronize the proxy object with Word on the backend. **After the `range.load()` call, add the following `context.sync()` call:**

```js
                    // Synchronize the document state by executing the queued commands,
                    // and return a promise to indicate task completion.
                    return context.sync().then(function () {

                        // Queue a commmand to insert a content control around the selected text,
                        // and create a proxy content control object. We'll update the properties
                        // on the content control.
                        var myContentControl = range.insertContentControl();
                        myContentControl.tag = uniqueBindId;
                        myContentControl.title = title;
                        myContentControl.style = range.style;
                        myContentControl.cannotEdit = false;

                        //context.trackedObjects.remove(range);
                        return context.sync().then(function () {

                            //showNotification('addFromSelectionAsync', 'Wrapped a content control around the selected text.');
                            console.log('Wrapped a content control around the selected text.');

                        });
                    });
```

In this block, a lot is going on. We use the `range` object to `insertContentControl` and then set some properties on the content control. We tag the content control with our `uniqueBindId` so we can locate it and navigate to it. We give it a title, keep the style the same as the original `range` selection, and allow the user to edit the text within the content control. 

Notice that we have an embedded `context.sync()`. Because we made changes to a proxy object that came back from `insertContentControl`, we need to call `context.sync()` to synchronize with Word's backend. 

The `.then` function on the `context.sync()` allows us to run code when the asynchronous `context.sync` finishes. 

We are now done adding the Word-specific code to our add-in. 

7. When you test the add-in inside Word, create a task and you will now see a content control being added. You will also notice that selecting the task in the task list view causes navigation to the content control. Cool!
