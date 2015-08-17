/// <reference path="../App.js" />
var item;
var table;

(function () {
    'use strict';

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            table = document.getElementById("details");
            item = Office.context.mailbox.item;
            $('#set-subject').click(setSubject);
            $('#get-subject').click(getSubject);
            $('#add-to-recipients').click(addToRecipients);
            $('#add-attachment').click(addAttachments);
            $('#add-bodytext').click(setItemBody);
        });
    };

    function setSubject() {
        Office.cast.item.toItemCompose(item).subject.setAsync("Hello world!");
    }

    function getSubject() {
        Office.cast.item.toItemCompose(item).subject.getAsync(function (result) {
            app.showNotification('The current subject is', result.value)
        });
    }

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
    }

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