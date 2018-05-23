# Module 1: Office.js for Word

In this module, you will add Office.js code specific to Word. This code will add a content control to a selection in the Word edit session for a document. That content control is named with a binding ID and can then be the target of navigation to find it again.

## API's Used In This Module

- Word.run 
- addFromSelectionAsync
- getSelection
- Range 
- range.insertContentControl

## Step 1: The Word.run block

All Word-specific API's must be enclosed in a callback function passed to Word.run(). Variables must be used inside the callback function scope or there are limitations and special code that must be used to make them accessible. So let's stick with keeping everything in scope. 

1. Find the code in taskerWeb/home.js that has this comment:

```js
                // ====== START ======
                // Workshop module 1 code goes here:
```

Notice that we are now inside a switch statement that identifies the "Host" of this add-in as Word!

2. Add the following Word.run block:

```js
                Word.run(function (context) {
                }
```

Everything we do to the document using Word javascript API's will happen inside this block.

3. Now add the code to call addFromSelectionAsync:

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

                }
```
At this point, we have made a binding, named with a value computed previously and with which we will later be able to navigate in the document.

4. Add a call to getSelection so we can create a content control around it.



