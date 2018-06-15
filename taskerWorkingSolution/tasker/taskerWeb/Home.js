// Per tenant variables to update when building against a new tenant.
// (1) AppId from Application Registration Portal
var azureAppId = "<<appid goes here>>";
// (2) Planner task's tenant-specific base URL. Get this from Planner with an open task.
var plannerTaskUrl = "https://tasks.office.com/jebosoft.onmicrosoft.com/en-US/Home/Task/";
// (3) Plan ID for the plan we created.
var planId = "<<plan id goes here>>";
// (4) Bucket ID for the tasks we create.
var bucketId = "<<bucket id goes here>>";

// Other globals
var gToken;
var authenticator;
var client;
var me;
var encodedDocUrl;
var docUrl;
var currentTasks;
var myPeople;

(function () {
    "use strict";

    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            if (OfficeHelpers.Authenticator.isAuthDialog()) return;

            // Initialize the FabricUI notification mechanism and hide it
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.init();
            //messageBanner.hideBanner();

            $('#button-text').text("Create task");
            $('#button-desc').text("Create a new task, marking the selected text with a content control.");

            $('#toggleAllDocs').click(loadContextData)

            $('#create-button').click(createTask);

            $('#signInBtn').click(SignInClick);

            $('#chkAutoOpenWithDoc').click(pinTaskerToDocument);
            $('#chkAutoOpenWithDoc').prop('checked', (Office.context.document.settings.get("Office.AutoShowTaskpaneWithDocument")));

            initializeComponents();

            document.getElementById("mytasks_list").addEventListener("click", function (e) {
                console.log(e.target.nodeName)
                if (e.target && (e.target.className.indexOf("ms-ListItem-secondaryText") == -1) && (e.target.className.indexOf("ms-Icon--More") == -1)) {
                    console.log(e.target.id);
                    goToTaskSelection(e.target.id.substr(2));
                }
            });
        });
    };

    function pinTaskerToDocument() {
        $('#chkAutoOpenWithDoc').prop('checked', (Office.context.document.settings.get("Office.AutoShowTaskpaneWithDocument")));
        Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", $('#chkAutoOpenWithDoc').is(':checked'));
        Office.context.document.settings.saveAsync();
    }

    function updateListView() {

        var alldocs = $('#toggleAllDocs')[0].checked;
        var numTasksCompleted = 0;
        var numTasksThisDocument = 0;
        var numTasksCompletedThisDocument = 0;

        var numTasks = 0;

        // clear out the list.
        $('#mytasks_list').empty();

        var liElements = currentTasks.value.map(function (task) {
            // TODO: chose this icon based on creator's relationship to me.
            var svgOrgOriginIcon = '"\.\\Images\\task-origin-icons\\org chart - team - square filled.svg"';
            var isUnread = '';
            var isUnreadStyle = '';
            var primaryTextStyle = 'color: gray';
            var secondaryTextStyle = '';
            var tertiaryTextStyle = '';
            var actionDiv = '';
            var title = task.title;
            var description = task.tasker_detail.description;
            var dueDate = new Date(task.dueDateTime);
            var liId = "otherdoc";

            if (task.planId != planId)
                return null;

            if (task.percentComplete == 0) {
                primaryTextStyle = '';
                isUnread = ' is-unread ';
                isUnreadStyle = 'color: #0078d7;font-weight:bold;'
            }

            numTasks++;

            if (task.percentComplete == 100)
                numTasksCompleted++;

            if (typeof task.tasker_detail.references[encodedDocUrl] == 'undefined') {
                var actionDivIcon = '';
                for (var key in task.tasker_detail.references) {
                    switch (task.tasker_detail.references[key].type) {
                        case Office.HostType.Word: {
                            actionDivIcon = '<img src = "\.\\Images\\task-origin-icons\\edit-docx.svg" />';
                        } break;
                        case Office.HostType.Excel: {
                            actionDivIcon = '<img src = "\.\\Images\\task-origin-icons\\edit-xlsx.svg" />';
                        } break;
                        case Office.HostType.PowerPoint: {
                            actionDivIcon = '<img src = "\.\\Images\\task-origin-icons\\edit-pptx.svg" />';
                        } break;
                    }
                }
                primaryTextStyle = 'color: lightgray';
                secondaryTextStyle = 'color: lightgray';
                tertiaryTextStyle = 'color: lightgray';
                actionDiv = '<div class="ms-ListItem-action"> <i class="ms-Icon">\
                    <a href="' + docUrl + '" target="_blank">' + actionDivIcon + '</a></i ></div >';
            }
            else {
                if (task.percentComplete == 100)
                    numTasksCompletedThisDocument++;
                numTasksThisDocument++;
                var selNotStart = (task.percentComplete == 0) ? "is-selected" : "";
                var selInProgress = (task.percentComplete == 50) ? "is-selected" : "";

                var liId = task.tasker_detail.references[encodedDocUrl].alias;
                actionDiv = '<div class="ms-ContextualMenuExample">\
                    <div class="ms-ListItem-action">\
                        <i class="ms-Icon ms-Icon--More"></i>\
                    </div>\
                    <ul id="cm' + liId + '" class="ms-ContextualMenu is-hidden">\
                        <li class="ms-ContextualMenu-item ms-ContextualMenu-item--header">STAGE</li>\
                        <li class="ms-ContextualMenu-item">\
                            <a class="ms-ContextualMenu-link tasker-notstartMenuItem ' + selNotStart + '" tabindex="1">Not Started</a>\
                        </li>\
                        <li class="ms-ContextualMenu-item">\
                            <a class="ms-ContextualMenu-link tasker-inprogressMenuItem ' + selInProgress + '" tabindex="1">In Progress</a>\
                        </li>\
                        <li class="ms-ContextualMenu-item">\
                            <a class="ms-ContextualMenu-link tasker-completedMenuItem " tabindex="1">Completed</a>\
                        </li>\
                        <li class="ms-ContextualMenu-item ms-ContextualMenu-item--divider"></li>\
                        <li class="ms-ContextualMenu-item ms-ContextualMenu-item--header">SEND</li>\
                        <li class="ms-ContextualMenu-item">\
                            <a class="ms-ContextualMenu-link" tabindex="1">Email</a>\
                        </li>\
                        <li class="ms-ContextualMenu-item">\
                            <a class="ms-ContextualMenu-link" tabindex="1">Teams Chat</a>\
                        </li>\
                    </ul>\
                </div>';
            }

            // BUG: Callout only sets up close button event for the first callout in the DOM. All others won't close on clicking the button.
            // This is documented: https://github.com/OfficeDev/office-ui-fabric-js/issues/333
            var calloutBlock = '<div class="ms-Callout ms-Callout--arrowLeft  ms-Callout--close is-hidden">\
                <div class="ms-Callout-main">\
                <button class="ms-Callout-close">\
                    <i class="ms-Icon ms-Icon--Clear"></i>\
                  </button>\
                  <div class="ms-Callout-header">\
                    <p class="ms-Callout-title">' + title + '</p>\
                  </div>\
                  <div class="ms-Callout-inner">\
                    <div class="ms-Callout-content">\
                      <p class="ms-Callout-subText">' + description + '</p>\
                    </div>\
                    <div class="ms-Callout-actions">\
                      <a class="ms-Link" title="Go to Planner" href="' + plannerTaskUrl + task.id + '" target="_blank">Go to Planner...</a>\
                    </div>\
                  </div>\
                </div>\
              </div>';

            return '<li id="li' + liId + '" class= "ms-ListItem ' + isUnread + ' ms-ListItem--image ms-CalloutTaskDescription" tabindex = "0" ><div  id="lm' + liId + '" class= "ms-ListItem-image" style = "background-color:white; width:20px;height:20px" >\
                <img id="oi' + liId + '" src = ' + svgOrgOriginIcon + ' style = "background-color:white; width:20px;height:20px" />\
                </div ><span  id="pr' + liId + '" class="ms-ListItem-primaryText" style="' + primaryTextStyle + isUnread + '">' + title + '</span>\
                <span  id="se' + liId + '" class= "ms-ListItem-secondaryText " style="' + secondaryTextStyle + '">' + description.substring(0, 30) + '... </span >\
                <span  id="te' + liId + '" class= "ms-ListItem-tertiaryText" style="' + tertiaryTextStyle + '"> Due: ' + dueDate.toDateString() + '</span >\
                <div  id="st' + liId + '" class="ms-ListItem-selectionTarget"></div>\
                <div class="ms-ListItem-actions">' + actionDiv + '</div >' + calloutBlock + '</li > ';
        });

        var index = 0;
        liElements.forEach(function (el) {
            if ((el != null) && (currentTasks.value[index].percentComplete != 100)) {
                if (typeof currentTasks.value[index].tasker_detail.references[encodedDocUrl] != 'undefined') {
                    $('#mytasks_list').append(el);
                }
            }
            index++;
        });

        if (alldocs) {

            index = 0;
            liElements.forEach(function (el) {
                if ((el != null) && (currentTasks.value[index].percentComplete != 100)) {
                    if (typeof currentTasks.value[index].tasker_detail.references[encodedDocUrl] == 'undefined') {
                        $('#mytasks_list').append(el);
                    }
                }
                index++;
            });

        }

        var ContextualMenuElements = document.querySelectorAll(".ms-ContextualMenuExample");
        for (var i = 0; i < ContextualMenuElements.length; i++) {
            var ButtonElement = ContextualMenuElements[i].querySelector(".ms-Icon--More");
            var ContextualMenuElement = ContextualMenuElements[i].querySelector(".ms-ContextualMenu");
            new fabric['ContextualMenu'](ContextualMenuElement, ButtonElement);
        }

        // pick out a certain list item 
        var inProgress = document.querySelectorAll(".tasker-inprogressMenuItem");
        inProgress.forEach(function (menuItem) {
            menuItem.addEventListener("click", function (e) {
                // e.target is our targetted element.
                console.log('<' + e.target.nodeName + '>');
                console.log("In Progress!");
                var task = currentTasks.value.find(function (t) {
                    return t.id == e.target.parentElement.parentElement.id.substr(8);
                });
                if (task.percentComplete != 50)
                    updateTask(task, '{ \"percentComplete\": 50 }');
            }, false);
        });
        // pick out a certain list item 
        var notStarted = document.querySelectorAll(".tasker-notstartMenuItem");
        notStarted.forEach(function (menuItem) {
            menuItem.addEventListener("click", function (e) {
                // e.target is our targetted element.
                console.log('<' + e.target.nodeName + '>');
                console.log("Not Started!");
                var task = currentTasks.value.find(function (t) {
                    return t.id == e.target.parentElement.parentElement.id.substr(8);
                });
                if (task.percentComplete != 0)
                    updateTask(task, "{ \"percentComplete\": 0 }");
            }, false);
        });
        // pick out a certain list item 
        var completed = document.querySelectorAll(".tasker-completedMenuItem");
        completed.forEach(function (menuItem) {
            menuItem.addEventListener("click", function (e) {
                // e.target is our targetted element.
                console.log('<' + e.target.nodeName + '>');
                console.log("Completed!");
                var task = currentTasks.value.find(function (t) {
                    return t.id == e.target.parentElement.parentElement.id.substr(8);
                });
                if (task.percentComplete != 100)
                    updateTask(task, '{ \"percentComplete\": 100 }');
            }, false);
        });

        var CalloutTaskDescriptions = document.querySelectorAll(".ms-CalloutTaskDescription");
        for (var i = 0; i < CalloutTaskDescriptions.length; i++) {
            var calloutTask = CalloutTaskDescriptions[i];
            var secondaryTextElement = calloutTask.querySelector(".ms-ListItem-secondaryText");
            var CalloutElement = calloutTask.querySelector(".ms-Callout");
            new fabric['Callout'](
                CalloutElement,
                secondaryTextElement,
                "right"
            );
        }
        var percentRatio = (alldocs ? (numTasksCompleted / numTasks) : (numTasksCompletedThisDocument / numTasksThisDocument));
        var percent = Number.parseFloat(percentRatio * 100).toFixed(1);
        if (alldocs) {
            $('#progressCompleteDescription').text(percent + '% tasks complete for all documents (' + numTasksCompleted + ' out of ' + numTasks + ' tasks).');
            $('#progressBarTooltip').text(percent + '% tasks complete for all documents (' + numTasksCompleted + ' out of ' + numTasks + ' tasks).');
        }
        //"ms-ProgressIndicator-progressBar"
        else {
            $('#progressCompleteDescription').text(percent + '% tasks complete for this document (' + numTasksCompletedThisDocument + ' out of ' + numTasksThisDocument + ' tasks).');
            $('#progressBarTooltip').text(percent + '% tasks complete for this document (' + numTasksCompletedThisDocument + ' out of ' + numTasksThisDocument + ' tasks).');
        }

        var ProgressIndicatorElements = document.querySelectorAll(".ms-ProgressIndicator");
        for (var i = 0; i < ProgressIndicatorElements.length; i++) {
            let ProgressIndicator = new fabric['ProgressIndicator'](ProgressIndicatorElements[i]);
            setTimeout(function () {
                //ProgressIndicator.setProgressPercent(alldocs ? (numTasksCompleted / numTasks) : (numTasksCompletedThisDocument / numTasksThisDocument));
                ProgressIndicator.setProgressPercent(percentRatio);
            }, 2000);
        }

        $('#mytasks_list').css({ "visibility": "hidden" });

        setTimeout(function () {
            $('#mytasks_list').css({ "visibility": "visible" });
        }, 100);
    }

    function updateTask(task, patchString) {
        $("#pivotContainer").find("*").prop('enabled', false);
        $('#spinnerAdding').css({ "display": "inline-block" });

        client
            .api('/planner/tasks/' + task.id)
            .header("If-Match", task["@odata.etag"])
            .header("Content-Type", "application/json")
            .patch(patchString).then(function (res) {
                // refresh the data
                loadContextData();

                $("#pivotContainer").find("*").prop('disabled', false);
                $('#spinnerAdding').css({ "display": "none" });

                showNotification('Success!', 'Update the new task');
            }).catch(function (err) {
                // catch any error that happened so far
                console.log("Error: " + err.message);
            });
    }

    function getPercentComplete() {
        if ($('#toggleAllDocs')[0].checked)
            return (4 / 10);
        else return (5 / 8);
    }

    function SignInClick() {

        $("#signInBtn").hide();
        $('#spinnerLoading').css({ "display": "inline-block" });
        authenticator = new OfficeHelpers.Authenticator();
        authenticator.endpoints.registerMicrosoftAuth(azureAppId, {
            redirectUrl: 'https://localhost:44382/Home.html',
            scope: 'files.readwrite.all User.Read.All Group.Read.All Group.ReadWrite.All People.Read Tasks.ReadWrite.Shared Tasks.ReadWrite Directory.Read.All Directory.ReadWrite.All Directory.AccessAsUser.All'
        });

        loadContextData();
    }

    function encodePlannerUrl(urlStr) {
        return encodeURIComponent(urlStr).replace(/[!'.()*]/g, function (c) {
            return '%' + c.charCodeAt(0).toString(16).toUpperCase();
        }).replace(/%2F/g, '/').replace(/%20/g, ' ');
    }

    function loadContextData() {
        authenticator
            .authenticate(OfficeHelpers.DefaultEndpoints.Microsoft, false)
            .then(function (token) {
                if (!token) {
                    console.log("ADAL error occurred: " + error);
                    return;
                }
                else {
                    //$("#signInBtn").text("SignOut");
                    gToken = token;
                    $('#spinnerLoading').css({ "display": "none" });
                    $('#pivotContainer').css({ "display": "inline-block" });

                    onAuthenticatedSdk(token);

                    // set the me global variable so I know who I am.
                    client
                        .api('/me')
                        .get((err, res) => {
                            console.log(res); // prints info about authenticated user
                            me = res;
                            $('#labelMe').text(me.mail);
                            $('#labelMe').css({ "display": "inline-block" });
                        });

                    // we'll need the document URL for creating or finding a unique task 
                    docUrl = Office.context.document.url;
                    encodedDocUrl = encodePlannerUrl(docUrl);

                    // use Graph's "social intelligence" to get the people I work with. 
                    client
                        .api('/me/people')
                        .get().then(function (res) {

                            console.log(res); // print out the people 
                            myPeople = res;

                            //updatePeoplePickerResultGroup();
                            var groupTitleDiv = '<div class="ms-PeoplePicker-resultGroupTitle"> Contacts </div >';
                            $('#myPeopleResultGroup').empty();
                            $('#myPeopleResultGroup').append(groupTitleDiv);

                            var divPeopleResults = myPeople.value.map(function (person) {
                                // TODO: chose this icon based on creator's relationship to me.
                                if ((person.personType.class == "Person") && (person.personType.subclass == "OrganizationUser")) {

                                    var email = person.scoredEmailAddresses[0].address;
                                    var fullName = person.displayName;
                                    var firstInitial = (person.givenName != null) ? person.givenName.substring(0, 1) : "?";
                                    var secondInitial = (person.surname != null) ? person.surname.substring(0, 1) : "?";
                                    var initials = firstInitial + secondInitial;
                                    var departmentOrJob = (person.department != null) ? person.department : ((person.jobTitle != null) ? person.jobTitle : "unknown");

                                    return '<div class="ms-PeoplePicker-result" tabindex="1"> \
                                        <div class="ms-Persona ms-Persona--xs"> \
                                            <div class="ms-Persona-imageArea"> \
                                                <div class="ms-Persona-initials ms-Persona-initials--blue">' + initials + '</div> \
                                            </div> \
                                            <div class="ms-Persona-presence"> \
                                            </div> \
                                            <div class="ms-Persona-details"> \
                                                <div class="ms-Persona-primaryText">' + fullName + '</div> \
                                                <div class="ms-Persona-secondaryText">' + person.id + '</div> \
                                            </div> \
                                        </div> \
                                        <button class="ms-PeoplePicker-resultAction"> \
                                            <i class="ms-Icon ms-Icon--Clear"></i> \
                                        </button> \
                                    </div>';
                                }
                                else
                                    return null;
                            });

                            divPeopleResults.forEach(function (per) {
                                if (per != null)
                                    $('#myPeopleResultGroup').append(per);
                            });

                            var myFirstInitial = (me.givenName != null) ? me.givenName.substring(0, 1) : "?";
                            var mySecondInitial = (me.surname != null) ? me.surname.substring(0, 1) : "?";
                            var meDiv = '<div class="ms-PeoplePicker-result" tabindex="1"> \
                                        <div class="ms-Persona ms-Persona--xs"> \
                                            <div class="ms-Persona-imageArea"> \
                                                <div class="ms-Persona-initials ms-Persona-initials--blue">' + myFirstInitial + mySecondInitial + '</div> \
                                            </div> \
                                            <div class="ms-Persona-presence"> \
                                            </div> \
                                            <div class="ms-Persona-details"> \
                                                <div class="ms-Persona-primaryText">' + me.displayName + '</div> \
                                                <div class="ms-Persona-secondaryText">' + me.id + '</div> \
                                            </div> \
                                        </div> \
                                        <button class="ms-PeoplePicker-resultAction"> \
                                            <i class="ms-Icon ms-Icon--Clear"></i> \
                                        </button> \
                                    </div>';

                            $('#myPeopleResultGroup').append(meDiv);

                            client
                                .api('/me/planner/tasks')
                                .get().then(function (res) {
                                    console.log(res); // print out the tasks collection
                                    currentTasks = res;

                                    // Take an array of promises and wait on them all
                                    return Promise.all(

                                        // Map array of tasks to an array of detail promises
                                        currentTasks.value.map(function (task) {
                                            // Make sure we have a details object.
                                            if (task.hasDescription)
                                                return client
                                                    .api('/planner/tasks/' + task.id + '/details')
                                                    .get();
                                            else
                                                // We should not get here.
                                                return Promise.resolve(null);
                                        }));

                                }).then(function (details) {

                                    var index = 0;
                                    details.forEach(function (detail) {
                                        // …and add to each of the corresponding tasks.
                                        currentTasks.value[index++].tasker_detail = detail;
                                    });

                                    // Update our pivot list with all the tasks.
                                    updateListView();

                                    //$('#divTaskListView').show(0, loadContextData);

                                }).catch(function (err) {
                                    // catch any error that happened so far
                                    console.log("Error: " + err.message);

                                }).then(function () {
                                })
                        });
                }
            })
            .catch(function (error) { /* handle error here */ });
    }

    // The polling function
    function poll(fn, timeout, interval) {
        var endTime = Number(new Date()) + (timeout || 2000);
        interval = interval || 100;

        var checkCondition = function (resolve, reject) {
            var ajax = fn();
            // dive into the ajax promise
            ajax.then(function (response) {
                // If the condition is met, we're done!
                if (response != null) {
                    resolve(response);
                }

                // If the condition isn't met but the timeout hasn't elapsed, go again
                else if (Number(new Date()) < endTime) {
                    setTimeout(checkCondition, interval, resolve, reject);
                }

                // Didn't match and too much time, reject!
                else {
                    reject(new Error('Timed out for ' + fn + ': ' + arguments));
                }

            }).catch(function () {
                setTimeout(checkCondition, interval, resolve, reject);
            });
        };

        return new Promise(checkCondition);
    }

    function createTask() {

        var assigneeMe = me.id;
        var newDueDate = getDateTimeOffset($('#newDueDate').val());

        // Create the new task to add.
        var newTask = {
            "bucketId": bucketId,
            "planId": planId,
            "title": $('#newTitle').val(),
            "dueDateTime": getDateTimeOffset($('#newDueDate').val()),
            "assignments": {
                //"4e98f8f1-bb03-4015-b8e0-19bb370949d8": {
                //    "@odata.type": "microsoft.graph.plannerAssignment",
                //    "orderHint": "String"
                //}
            }
        };

        var numAssignees = 0;
        // Add assignees from the People Picker search box.
        $('div.ms-PeoplePicker-searchBox').find('.ms-Persona-secondaryText').each(function () {
            numAssignees++;
            newTask.assignments[this.innerText] =
                {
                    "@odata.type": "#microsoft.graph.plannerAssignment",
                    "orderHint": " !"
                }
        });

        // if no assignees in the search box, just assign to me.
        if (numAssignees == 0) {
            newTask.assignments[assigneeMe] =
                {
                    "@odata.type": "#microsoft.graph.plannerAssignment",
                    "orderHint": " !"
                }
        }
        var resTask;

        $("#pivotContainer").find("*").prop('disabled', true);
        $('#spinnerAdding').css({ "display": "inline-block" });

        client
            .api('/planner/tasks')
            .post(newTask).then(function (res) {
                console.log(res)
                //sleep(2000);
                resTask = res;

                poll(function () {
                    //return axios.get('something.json');
                    return client
                        .api('/planner/tasks/' + res.id + '/details')
                        .get();
                }, 3000, 150).then(function (det) {
                    var newDescription = $('#newDescription').val();
                    // it just happens that the details object id is the task id.
                    var bindid = addLocation($('#newTitle').val(), det.id);

                    var reference = {
                        "previewType": "noPreview",
                        "description": newDescription,
                        "references": {
                        }
                    };
                    reference.references[encodedDocUrl] =
                        {
                            "@odata.type": "microsoft.graph.plannerExternalReference",
                            "alias": bindid,
                            "previewPriority": ' !',
                            "type": Office.context.host
                        };

                    return client
                        .api('/planner/tasks/' + resTask.id + '/details')
                        .header("If-Match", det["@odata.etag"])
                        .header("Content-Type", "application/json")
                        .patch(reference);

                }).then(function (res) {
                    // all done with patch Promise

                    //sleep(2000);
                    // refresh the data
                    loadContextData();

                    $("#pivotContainer").find("*").prop('disabled', false);
                    $('#spinnerAdding').css({ "display": "none" });

                    showNotification('Success!', 'Created a new task');

                    $('#newDueDate').val("");
                    $('#newTitle').val("");
                    $('#newDescription').text("");

                }).catch(function (err) {
                    // catch any error that happened so far
                    console.log("Error: " + err.message);

                });

            });
    }

    function goToTaskSelection(id) {
        if (id == "otherdoc")
            return;
        switch (Office.context.host) {
            case Office.HostType.Word: {
                Word.run(function (context) {

                    //Go to binding by id.
                    Office.context.document.goToByIdAsync(id, Office.GoToType.Binding, function (asyncResult) {
                        if (asyncResult.status == "failed") {
                            console.log("Action failed with error: " + asyncResult.error.message);
                        }
                        else {
                            console.log("Navigation successful");
                        }
                    });
                    return context.sync().then(function () {
                        console.log("done with goto");
                        // Create a proxy object for the content controls collection that contains a specific tag.
                        var contentControlsWithTag = context.document.contentControls.getByTag(id);

                        // Queue a command to load the text property for all of content controls with a specific tag.
                        context.load(contentControlsWithTag, 'text');

                        // Synchronize the document state by executing the queued commands,
                        // and return a promise to indicate task completion.
                        return context.sync().then(function () {
                            if (contentControlsWithTag.items.length === 0) {
                                console.log("There isn't a content control with a tag of Customer-Address in this document.");
                                showNotification('goToByIdAsync: Failure! Did you forget to implement addFromSelectionAsync?');
                                console.log('goToByIdAsync: Failure! Did you forget to implement addFromSelectionAsync?');
                            } else {
                                console.log('The first content control with the tag of Customer-Address has this text: ' + contentControlsWithTag.items[0].text);
                                contentControlsWithTag.items[0].select('Select');
                            }

                        });
                    })

                }).catch(function (error) {
                    var errormsg = 'Error: ' + JSON.stringify(error);
                    console.log(errormsg);
                    if (error instanceof OfficeExtension.Error) {
                        var errormsg = errormsg + ', Debug info: ' + JSON.stringify(error.debugInfo);
                        showNotification('goToByIdAsync: Failure! Did you forget to implement addFromSelectionAsync?');
                        console.log('goToByIdAsync: Failure! Did you forget to implement addFromSelectionAsync?');
                        console.log("Debug info: " + errormsg);
                    }
                });
            }
                break;
            case Office.HostType.Excel: {
                Excel.run(function (context) {
                    var names = context.workbook.names;
                    var range = names.getItem(id).getRange();
                    range.load('address');
                    return context.sync().then(function () {
                        console.log(range.address);
                        range.select();
                    });
                }).catch(function (error) {
                    console.log("Error: " + error);
                    if (error instanceof OfficeExtension.Error) {
                        showNotification('range.select: Failure! Did you forget to implement workbook.names.add?');
                        console.log('range.select: Failure! Did you forget to implement workbook.names.add?');
                        console.log("Debug info: " + JSON.stringify(error.debugInfo));
                    }
                });
            }
                break;
            case Office.HostType.PowerPoint: {
                //Go to slide by id.
                Office.context.document.goToByIdAsync(id, Office.GoToType.Slide, function (asyncResult) {
                    if (asyncResult.status == "failed") {
                        console.log("Action failed with error: " + asyncResult.error.message);
                    }
                    else {
                        console.log("Navigation successful");
                    }
                });
            }
                break;
            default:
                {

                }
        }
    }

    // Assuming we have a selection:
    // Choose for Word, Excel, PowerPoint 
    // Word: ContentControl (hidden) for the selected range.
    // Excel: named Range
    // PowerPoint: Slide id
    function addLocation(title, taskid) {
        var d = new Date();
        var n = d.getTime();
        var uniqueBindId = 'tasker' + taskid;
        var range = null;

        switch (Office.context.host) {
            case Office.HostType.Word: {
                // ====== START ======
                // Workshop module 1 code goes here:

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

                    // Queue a command to get the current selection and then
                    // create a proxy range object with the results.
                    range = context.document.getSelection();

                    range.load("style");

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

                // ===== END =====
            }
                break;
            case Office.HostType.Excel: {
                // ====== START ======
                // Workshop module 2 code goes here:

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

                // ===== END =====
            }
                break;
            case Office.HostType.PowerPoint: {
                Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange,
                    { valueFormat: "unformatted", filterType: "all" },
                    function (asyncResult) {
                        var error = asyncResult.error;
                        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                            showNotification('addFromSelectionAsyn', error.name + ": " + error.message);
                        }
                        else {
                            // Get selected data.
                            var dataValue = asyncResult.value;

                            // Example: Selected data is [{"id":256,"title":"Video/Flash Test","index":1}]
                            uniqueBindId += '###' + dataValue.slides[0].id;
                            showNotification('addFromSelectionAsyn', 'Selected data is ' + JSON.stringify(dataValue.slides));
                        }
                    });
            }
                break;
            default:
                {

                }
        }
        return uniqueBindId;
    }

    function displaySelectedText() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    showNotification('Error:', result.error.message);
                }
            });
    }

    //$$(Helper function for treating errors, $loc_script_taskpane_home_js_comment34$)$$
    function errorHandler(error) {
        // $$(Always be sure to catch any accumulated errors that bubble up from the Word.run execution., $loc_script_taskpane_home_js_comment35$)$$
        showNotification("Error:", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        //messageBanner.toggleExpansion();
    }

    function getDateTimeOffset(dateStr) {
        var dt = new Date(dateStr),
            current_date = dt.getDate(),
            current_month = dt.getMonth() + 1,
            current_year = dt.getFullYear(),
            current_hrs = dt.getHours(),
            current_mins = dt.getMinutes(),
            current_secs = dt.getSeconds(),
            current_datetime;

        // Add 0 before date, month, hrs, mins or secs if they are less than 0
        current_date = current_date < 10 ? '0' + current_date : current_date;
        current_month = current_month < 10 ? '0' + current_month : current_month;
        current_hrs = current_hrs < 10 ? '0' + current_hrs : current_hrs;
        current_mins = current_mins < 10 ? '0' + current_mins : current_mins;
        current_secs = current_secs < 10 ? '0' + current_secs : current_secs;

        // Current datetime
        // String such as 2016-07-16T19:20:30
        current_datetime = current_year + '-' + current_month + '-' + current_date + 'T' + current_hrs + ':' + current_mins + ':' + current_secs;

        var timezone_offset_min = new Date().getTimezoneOffset(),
            offset_hrs = parseInt(Math.abs(timezone_offset_min / 60)),
            offset_min = Math.abs(timezone_offset_min % 60),
            timezone_standard;

        if (offset_hrs < 10)
            offset_hrs = '0' + offset_hrs;

        if (offset_min < 10)
            offset_min = '0' + offset_min;

        // Add an opposite sign to the offset
        // If offset is 0, it means timezone is UTC
        if (timezone_offset_min < 0)
            timezone_standard = '+' + offset_hrs + ':' + offset_min;
        else if (timezone_offset_min > 0)
            timezone_standard = '-' + offset_hrs + ':' + offset_min;
        else if (timezone_offset_min == 0)
            timezone_standard = 'Z';

        // Timezone difference in hours and minutes
        // String such as +5:30 or -6:00 or Z
        console.log(timezone_standard);
        console.log(current_datetime + timezone_standard);
        return current_datetime + timezone_standard;
    }

    function sleep(ms) {
        var start = new Date().getTime();
        while (new Date().getTime() < start + ms);
    }

    function updatePeoplePickerResultGroup() {

        var groupTitleDiv = '<div class="ms-PeoplePicker-resultGroupTitle"> Contacts </div >';
        $('#myPeopleResultGroup').empty();
        $('#myPeopleResultGroup').append(groupTitleDiv);

        var liElements = myPeople.value.map(function (person) {
            // TODO: chose this icon based on creator's relationship to me.
            if (person.personType.class == "Person") {

                var email = person.scoredEmailAddresses[0].address;
                var fullName = person.displayName;
                var firstInitial = (person.givenName.substring(0, 1) != null) ? person.givenName.substring(0, 1) : " ";
                var secondInitial = (person.surname.substring(0, 1) != null) ? person.surname.substring(0, 1) : " ";
                var initials = firstInitial + secondInitial;
                var departmentOrJob = (person.department != null) ? person.department : ((person.jobTitle != null) ? person.jobTitle : " ");

                return '<div class="ms-PeoplePicker-result" tabindex="1"> \
                                        <div class="ms-Persona ms-Persona--xs"> \
                                            <div class="ms-Persona-imageArea"> \
                                                <div class="ms-Persona-initials ms-Persona-initials--blue">' + initials + '</div> \
                                            </div> \
                                            <div class="ms-Persona-presence"> \
                                            </div> \
                                            <div class="ms-Persona-details"> \
                                                <div class="ms-Persona-primaryText">' + fullName + '</div> \
                                                <div class="ms-Persona-secondaryText">' + departmentOrJob + '</div> \
                                            </div> \
                                        </div> \
                                        <button class="ms-PeoplePicker-resultAction"> \
                                            <i class="ms-Icon ms-Icon--Clear"></i> \
                                        </button> \
                                    </div>';
            }
            else
                return null;
        });

        var index = 0;
        liElements.forEach(function (el) {
            if (el != null)
                $('#myPeopleResultGroup').append(el);
            index++;
        });
    }
    function initializeComponents() {

        var PivotElements = document.querySelectorAll(".ms-Pivot");
        for (var i = 0; i < PivotElements.length; i++) {
            new fabric['Pivot'](PivotElements[i]);
        }

        var FacePileElements = document.querySelectorAll(".ms-FacePile");
        for (var i = 0; i < FacePileElements.length; i++) {
            new fabric['FacePile'](FacePileElements[i]);
        }

        var DatePickerElements = document.querySelectorAll(".ms-DatePicker");
        for (var i = 0; i < DatePickerElements.length; i++) {
            new fabric['DatePicker'](DatePickerElements[i]);
        }

        var TextFieldElements = document.querySelectorAll(".ms-TextField");
        for (var i = 0; i < TextFieldElements.length; i++) {
            new fabric['TextField'](TextFieldElements[i]);
        }

        var ToggleElements = document.querySelectorAll(".ms-Toggle");
        for (var i = 0; i < ToggleElements.length; i++) {
            new fabric['Toggle'](ToggleElements[i]);
        }

        var ProgressIndicatorElements = document.querySelectorAll(".ms-ProgressIndicator");
        for (var i = 0; i < ProgressIndicatorElements.length; i++) {
            let ProgressIndicator = new fabric['ProgressIndicator'](ProgressIndicatorElements[i]);
            setTimeout(function () {
                ProgressIndicator.setProgressPercent(getPercentComplete());
            }, 2000);
        }

        var ContextualMenuElements = document.querySelectorAll(".ms-ContextualMenuExample");
        for (var i = 0; i < ContextualMenuElements.length; i++) {
            var ButtonElement = ContextualMenuElements[i].querySelector(".ms-Icon--More");
            var ContextualMenuElement = ContextualMenuElements[i].querySelector(".ms-ContextualMenu");
            new fabric['ContextualMenu'](ContextualMenuElement, ButtonElement);
        }

        var CheckBoxElements = document.querySelectorAll(".ms-CheckBox");
        for (var i = 0; i < CheckBoxElements.length; i++) {
            new fabric['CheckBox'](CheckBoxElements[i]);
        }

        var SpinnerElements = document.querySelectorAll(".ms-Spinner");
        for (var i = 0; i < SpinnerElements.length; i++) {
            new fabric['Spinner'](SpinnerElements[i]);
        }

        var PeoplePickerElements = document.querySelectorAll(".ms-PeoplePicker");
        for (var i = 0; i < PeoplePickerElements.length; i++) {
            new fabric['PeoplePicker'](PeoplePickerElements[i]);
        }

    }
})();

function onAuthenticatedSdk(token, authWindow) {
    if (token) {
        if (authWindow) {
            // removeLoginButton();
            // authWindow.close();
        }

        gToken = token;

        if (client == null) {
            client = MicrosoftGraph.Client.init({
                authProvider: (done) => {
                    done(null, token.access_token); //first parameter takes an error if you can't get an access token
                }
            });
        }
    }
    else {
        console.log("Error signing in");
    }
}

