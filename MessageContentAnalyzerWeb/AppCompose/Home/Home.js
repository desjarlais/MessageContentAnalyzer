/// <reference path="../App.js" />
var item;
var table;

(function () {
    'use strict';

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            table = document.getElementById("debuglog2");
            item = Office.context.mailbox.item;
            
            $('#set-subject').click(setSubject);
            $('#get-subject').click(getSubject);

            $('#set-location').click(setLocation);
            $('#get-location').click(getLocation);

            $('#set-start-time').click(setStartTime);
            $('#get-start-time').click(getStartTime);

            $('#add-to-recipients').click(addToRecipients);
            $('#add-attachment').click(addAttachments);

            $('#add-bodytext').click(setItemBody);

            debugLog("init", "init completed");
        });
    };

    // debug output function
    function debugLog(debugType, debugValue) {
        var spanElement = "<span id=" + debugType + ">" + " " + debugValue + "</span>";

        var prevVal = document.getElementById('debuglog').innerHTML;
        var output = document.getElementById('debuglog');
        output.innerHTML = prevVal + spanElement;
    }

    // set the subject of the item
    function setSubject() {
        Office.cast.item.toItemCompose(item).subject.setAsync("Hello world!");
    }

    // get the subject of the item
    function getSubject() {
        Office.cast.item.toItemCompose(item).subject.getAsync(function (result) {
            app.showNotification('The current subject is', result.value)
        });
    }

    // Set the location of the item that the user is composing.
    function setLocation() {
        // check if item is appointment since you can't set location on message items
        if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
            Office.cast.item.toItemCompose(item).location.setAsync(
            'Conference room A',
            { asyncContext: { var1: 1, var2: 2 } },
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    app.showNotification(asyncResult.error.message);
                }
                else {
                    // Successfully set the location.
                    // Do whatever appropriate for your scenario
                    // using the arguments var1 and var2 as applicable.
                    app.showNotification("Location set for item.");
                }
            });
        }
        else {
            app.showNotification("Can't set location on message items.");
        }

    }

    // Get the location of the item that the user is composing.
    function getLocation() {
        if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
            Office.cast.item.toItemCompose(item).location.getAsync(
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    app.showNotification(asyncResult.error.message);
                }
                else {
                    // Successfully got the location, display it.
                    app.showNotification('The location is: ' + asyncResult.value);
                }
            });
        }
        else {
            app.showNotification("Message items don't have location.");
        }
    }    

    // Get the start time of the item that the user is composing.
    function getStartTime() {
        if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
            Office.cast.item.toItemCompose(item).start.getAsync(
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    app.showNotification(asyncResult.error.message);
                }
                else {
                    // Successfully got the start time, display it, first in UTC and 
                    // then convert the Date object to local time and display that.
                    app.showNotification('The start time in UTC is: ' + asyncResult.value.toString());
                    app.showNotification('The start time in local time is: ' + asyncResult.value.toLocaleString());
                }
            });
        }
        else {
            app.showNotification("Message items don't have start times.");
        }
    }

    // Set the start time of the item that the user is composing.
    function setStartTime() {
        var startDate = new Date("September 1, 2015 12:30:00");
        if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
            Office.cast.item.toItemCompose(item).start.setAsync(
            startDate,
            { asyncContext: { var1: 1, var2: 2 } },
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    app.showNotification(asyncResult.error.message);
                }
                else {
                    // Successfully set the start time.
                    // Do whatever appropriate for your scenario
                    // using the arguments var1 and var2 as applicable.
                    app.showNotification("Time set successfully.");
                }
            });
        }
        else {
            app.showNotification("Can't set start time on message items.");
        }
    }

    // add the currently logged in user to the recipient list
    function addToRecipients() {
        var addressToAdd = {
            displayName: Office.context.mailbox.userProfile.displayName,
            emailAddress: Office.context.mailbox.userProfile.emailAddress
        };

        if (item.itemType === Office.MailboxEnums.ItemType.Message) {
            Office.cast.item.toMessageCompose(item).to.addAsync([addressToAdd]);
        } else if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
            Office.cast.item.toAppointmentCompose(item).requiredAttendees.addAsync([addressToAdd]);
        }
        debugLog("addToRecipients", "recipients added");
    }

    // add generic image file as test attachment
    function addAttachments() {
        if (item.itemType === Office.MailboxEnums.ItemType.Message) {
            Office.cast.item.toMessageCompose(item).addFileAttachmentAsync("https://i.imgur.com/ucI9vyz.png", "image file", { asyncContext: null },
                function (asyncResult) {
                    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                        app.showNotification(asyncResult.error.message);
                    }
                    else {
                        app.showNotification('ID of added attachment: ' + asyncResult.value);
                    }
                });
        } else if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
            Office.cast.item.toAppointmentCompose(item).addFileAttachmentAsync("https://i.imgur.com/ucI9vyz.png", "image file");
        }
    }

    // Get the body type of the composed item, and set data in 
    // in the appropriate data type in the item body.
    function setItemBody() {
        item.body.getTypeAsync(
            function (result) {
                if (result.status == Office.AsyncResultStatus.Failed) {
                    app.showNotification(result.error.message);
                }
                else {
                    // Successfully got the type of item body.
                    // Set data of the appropriate type in body.
                    if (result.value == Office.MailboxEnums.BodyType.Html) {
                        // Body is of HTML type.
                        // Specify HTML in the coercionType parameter
                        // of setSelectedDataAsync.
                        item.body.setSelectedDataAsync(
                            '<b> Kindly note we now open 7 days a week.</b>',
                            {
                                coercionType: Office.CoercionType.Html,
                                asyncContext: { var3: 1, var4: 2 }
                            },
                            function (asyncResult) {
                                if (asyncResult.status ==
                                    Office.AsyncResultStatus.Failed) {
                                    app.showNotification(result.error.message);
                                }
                                else {
                                    // Successfully set data in item body.
                                    // Do whatever appropriate for your scenario,
                                    // using the arguments var3 and var4 as applicable.
                                    app.showNotification("HTML text added.");
                                }
                            });
                    }
                    else {
                        // Body is of text type. 
                        item.body.setSelectedDataAsync(
                            ' Kindly note we now open 7 days a week.',
                            {
                                coercionType: Office.CoercionType.Text,
                                asyncContext: { var3: 1, var4: 2 }
                            },
                            function (asyncResult) {
                                if (asyncResult.status ==
                                    Office.AsyncResultStatus.Failed) {
                                    app.showNotification(result.error.message);
                                }
                                else {
                                    // Successfully set data in item body.
                                    // Do whatever appropriate for your scenario,
                                    // using the arguments var3 and var4 as applicable.
                                    app.showNotification("Plain text added.");
                                }
                            });
                    }
                }
            });
    }
})();